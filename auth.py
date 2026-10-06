"""Generate an O365-compatible MSAL cache through a public-client sign-in.

MSAL owns PKCE, nonce validation, token redemption, and the cache schema. This
module handles configuration, the manual browser round trip, and file storage.
Importing it never reads configuration, changes the environment, or uses the
network. Tokens and pasted redirect URLs must never appear in diagnostics.
"""

from __future__ import annotations

import argparse
import hmac
import json
import math
import os
import re
import sys
import tempfile
import webbrowser
from collections.abc import Sequence
from dataclasses import dataclass
from enum import IntEnum
from pathlib import Path
from typing import Any
from urllib.parse import parse_qsl, urlsplit
from uuid import UUID

import msal
from dotenv import dotenv_values
from msal.exceptions import MsalServiceError
from requests.exceptions import RequestException

SCOPES = ("Mail.ReadWrite", "Mail.Send", "User.Read")
REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"
TOKEN_CACHE_PATH = "o365_token.txt"
DEFAULT_ENV_FILE = Path(__file__).resolve().with_name(".env")
MAX_REDIRECT_LENGTH = 16_384
MAX_PASTE_ATTEMPTS = 3
RESERVED_SCOPES = {"openid", "profile", "offline_access"}


class ExitCode(IntEnum):
    """Stable process statuses for shell scripts and scheduled jobs."""

    SUCCESS = 0
    AUTHORIZATION = 1
    CONFIGURATION = 2
    NETWORK = 3
    STORAGE = 4
    CANCELLED = 130


class ConfigurationError(ValueError):
    """Invalid local settings; messages contain no credential values."""


class AuthorizationError(ValueError):
    """Invalid or unsuccessful authorization; safe to show to the operator."""


@dataclass(frozen=True)
class Configuration:
    """Validated settings resolved independently for each invocation."""

    client_id: str
    tenant_id: str
    scopes: tuple[str, ...]
    output: Path
    timeout: float

    @property
    def authority(self) -> str:
        """Use the public Microsoft cloud and a validated tenant path segment."""
        return f"https://login.microsoftonline.com/{self.tenant_id}"


def build_parser() -> argparse.ArgumentParser:
    """Keep the existing no-argument command while exposing common settings."""
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n", 1)[0])
    parser.add_argument(
        "--env-file",
        type=Path,
        default=None,
        help="Configuration file (default: .env beside auth.py).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(TOKEN_CACHE_PATH),
        help="Cache file (default: o365_token.txt in the current directory).",
    )
    parser.add_argument(
        "--scopes",
        nargs="+",
        default=SCOPES,
        metavar="SCOPE",
        help="Delegated Graph permissions (default: Mail.ReadWrite Mail.Send User.Read).",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=30.0,
        metavar="SECONDS",
        help="Network connect/read timeout (default: 30 seconds each).",
    )
    parser.add_argument(
        "--open-browser",
        action="store_true",
        help="Also try to open the sign-in URL in the default browser.",
    )
    return parser


def load_configuration(args: argparse.Namespace) -> Configuration:
    """Read one explicit file without mutating os.environ or searching parents.

    Process variables take precedence, even when explicitly empty. Interpolation
    is disabled: application identifiers should be literal values, and unrelated
    environment variables must not silently change them.
    """
    env_file = (args.env_file if args.env_file is not None else DEFAULT_ENV_FILE).expanduser()
    if args.env_file is not None and not env_file.is_file():
        raise ConfigurationError("The --env-file must point to an existing file.")
    if env_file.exists() and not env_file.is_file():
        raise ConfigurationError("The configuration path must be a regular file.")
    values = (
        dotenv_values(env_file, encoding="utf-8-sig", interpolate=False)
        if env_file.is_file()
        else {}
    )
    client_id = (os.environ.get("AZURE_CLIENT_ID", values.get("AZURE_CLIENT_ID")) or "").strip()
    tenant_id = (os.environ.get("AZURE_TENANT_ID", values.get("AZURE_TENANT_ID")) or "").strip()
    if not client_id or not tenant_id:
        raise ConfigurationError(
            "Set AZURE_CLIENT_ID and AZURE_TENANT_ID in the configuration file "
            "or process environment."
        )
    # UUID parsing rejects accidental names, URL fragments, and placeholder text.
    try:
        client_id = str(UUID(client_id))
    except ValueError:
        raise ConfigurationError("AZURE_CLIENT_ID must be an application UUID.") from None

    try:
        tenant_id = str(UUID(tenant_id))
    except ValueError:
        # Microsoft also supports tenant domains and these three account aliases.
        domain = re.fullmatch(
            r"(?=.{1,253}\Z)(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+"
            r"[A-Za-z](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?",
            tenant_id,
        )
        if tenant_id not in {"common", "organizations", "consumers"} and not domain:
            raise ConfigurationError(
                "AZURE_TENANT_ID must be a tenant UUID, domain, "
                "common, organizations, or consumers."
            ) from None

    scopes = tuple(dict.fromkeys(args.scopes))
    for scope in scopes:
        name = scope.removeprefix("https://graph.microsoft.com/")
        if name.lower() in RESERVED_SCOPES:
            raise ConfigurationError(
                "Do not request openid, profile, or offline_access explicitly; MSAL adds them."
            )
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9]*(?:\.[A-Za-z][A-Za-z0-9]*)+", name):
            raise ConfigurationError("Use explicit delegated Microsoft Graph permission names.")
    if not scopes:
        raise ConfigurationError("At least one delegated Microsoft Graph permission is required.")
    if not math.isfinite(args.timeout) or args.timeout <= 0:
        raise ConfigurationError("--timeout must be a finite positive number of seconds.")

    # Do not resolve the file itself: resolving a symlink would hide it from the
    # storage guard. Relative output paths intentionally retain the original CWD behavior.
    output = args.output.expanduser().absolute()
    validate_output_path(output)
    if output.resolve() == env_file.expanduser().resolve():
        raise ConfigurationError("The cache output must not replace the configuration file.")
    return Configuration(client_id, tenant_id, scopes, output, args.timeout)


def validate_output_path(path: Path) -> None:
    """Reject unusable destinations before sign-in and again before replacing."""
    if not path.parent.is_dir():
        raise ConfigurationError("The cache output directory must already exist.")
    if path.is_symlink() or (path.exists() and not path.is_file()):
        raise ConfigurationError(
            "The cache output must be a regular file, not a link or directory."
        )


def parse_authorization_response(final_url: str, expected_state: str) -> dict[str, str]:
    """Validate the native-client callback without echoing sensitive input.

    A substring such as 'code=' is insufficient: it can occur in a fragment or
    another value. Parse query pairs, reject ambiguous duplicates, and bind the
    response to this process's flow. MSAL will validate state and nonce again.
    """
    final_url = final_url.strip()
    if not final_url or len(final_url) > MAX_REDIRECT_LENGTH:
        raise AuthorizationError("Paste a complete redirect URL of at most 16,384 characters.")
    if any(
        character.isspace() or ord(character) < 32 or ord(character) == 127
        for character in final_url
    ):
        raise AuthorizationError("The redirect URL contains whitespace or control characters.")
    try:
        parsed = urlsplit(final_url)
        correct_endpoint = (
            parsed.scheme == "https"
            and parsed.hostname == "login.microsoftonline.com"
            and parsed.port in (None, 443)
            and parsed.username is None
            and parsed.password is None
            and parsed.path == "/common/oauth2/nativeclient"
            and not parsed.fragment
        )
        if not correct_endpoint:
            raise ValueError
        if re.search(r"%(?![0-9A-Fa-f]{2})", parsed.query):
            raise ValueError
        pairs = parse_qsl(
            parsed.query,
            keep_blank_values=True,
            strict_parsing=True,
            encoding="utf-8",
            errors="strict",
            max_num_fields=30,
        )
    except ValueError:
        raise AuthorizationError(
            "Paste the complete Microsoft native-client redirect URL, including its query."
        ) from None
    response: dict[str, str] = {}
    for key, value in pairs:
        if not key or key in response:
            raise AuthorizationError("The redirect query contains empty or duplicate parameters.")
        response[key] = value
    state = response.get("state", "")
    # Compare bytes so non-ASCII input also fails cleanly; compare_digest(str,
    # str) only accepts ASCII. The state must come from the current sign-in URL.
    if not state or not hmac.compare_digest(state.encode("utf-8"), expected_state.encode("utf-8")):
        raise AuthorizationError(
            "The redirect state does not match this sign-in. Use the current browser flow."
        )
    if ("code" in response) == ("error" in response):
        raise AuthorizationError("The redirect must contain exactly one of code or error.")
    if not response.get("code") and not response.get("error"):
        raise AuthorizationError("The authorization code or error value is empty.")
    return response


def describe_provider_error(result: dict[str, Any]) -> str:
    """Report useful identifiers without printing an untrusted error description.

    Provider messages and exception strings can contain personal data, codes,
    tokens, or terminal escape sequences. Only known OAuth errors, AADSTS codes,
    and a validated correlation UUID are included in the console output.
    """
    known_errors = {
        "access_denied",
        "invalid_request",
        "invalid_grant",
        "invalid_client",
        "invalid_scope",
        "unauthorized_client",
        "interaction_required",
        "consent_required",
        "login_required",
        "server_error",
        "temporarily_unavailable",
    }
    error = result.get("error")
    parts = [
        error if isinstance(error, str) and error in known_errors else "unspecified provider error"
    ]
    description = result.get("error_description")
    if isinstance(description, str):
        parts.extend(dict.fromkeys(re.findall(r"\bAADSTS[0-9]{4,10}\b", description)[:5]))
    correlation = result.get("correlation_id")
    if isinstance(correlation, str):
        try:
            parts.append(f"correlation ID: {UUID(correlation)}")
        except ValueError:
            pass
    return "; ".join(parts)


def save_cache(cache: msal.SerializableTokenCache, destination: Path) -> None:
    """Publish a complete cache using a same-directory atomic replacement.

    Validate the required sections before touching disk. A temporary file is
    written, flushed, and fsynced before os.replace publishes it. If any preceding
    operation fails, an existing cache remains intact. POSIX files receive mode
    0600; Windows access is governed by the destination directory's inherited ACL.
    """
    serialized = cache.serialize()
    payload = json.loads(serialized)
    if not isinstance(payload, dict) or any(
        not isinstance(payload.get(section), dict) or not payload[section]
        for section in ("AccessToken", "RefreshToken", "Account")
    ):
        raise AuthorizationError(
            "MSAL did not produce a reusable account and token cache. Sign in again."
        )
    validate_output_path(destination)
    temporary: Path | None = None
    try:
        # NamedTemporaryFile closes before os.replace, which is required on
        # Windows. The file lives beside the target to avoid cross-volume moves.
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=destination.parent,
            prefix=f".{destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as stream:
            temporary = Path(stream.name)
            if os.name != "nt":
                os.chmod(temporary, 0o600)
            stream.write(serialized)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, destination)
    finally:
        if temporary is not None:
            # Cleanup must not replace the primary error with a second error.
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                print(
                    "Warning: unable to remove a temporary cache file in the output directory.",
                    file=sys.stderr,
                )


def run_authorization(config: Configuration, *, open_browser: bool = False) -> None:
    """Execute one flow; never automatically retry a single-use code exchange."""
    cache = msal.SerializableTokenCache()
    app = msal.PublicClientApplication(
        config.client_id,
        authority=config.authority,
        token_cache=cache,
        timeout=config.timeout,
    )
    flow = app.initiate_auth_code_flow(
        scopes=list(config.scopes),
        redirect_uri=REDIRECT_URI,
        response_mode="query",
    )
    if not isinstance(flow, dict) or any(
        not isinstance(flow.get(field), str) or not flow[field] for field in ("auth_uri", "state")
    ):
        raise AuthorizationError(
            "MSAL could not create a sign-in flow. Check the application registration."
        )
    print("\nOpen this authorization URL in your browser:")
    print(flow["auth_uri"])
    if open_browser:
        try:
            opened = webbrowser.open(flow["auth_uri"])
        except (webbrowser.Error, OSError):
            opened = False
        if not opened:
            print("The browser could not be opened. Open the printed URL manually.")

    print("\nSign in, then copy the complete URL from the final browser address bar.")
    print("The native-client page may be blank. Keep this process running until you paste the URL.")
    for attempt in range(MAX_PASTE_ATTEMPTS):
        final_url = input("\nPaste the complete redirect URL: ")
        try:
            response = parse_authorization_response(final_url, flow["state"])
            break
        except AuthorizationError as error:
            if attempt == MAX_PASTE_ATTEMPTS - 1:
                raise
            print(f"Invalid redirect: {error} You can paste again.", file=sys.stderr)

    # Pass even an error redirect through MSAL's flow validation. The entire
    # original flow (including the PKCE verifier) stays in memory and is never saved.
    result = app.acquire_token_by_auth_code_flow(flow, auth_response=response)
    if not isinstance(result, dict) or not result.get("access_token"):
        details = (
            describe_provider_error(result) if isinstance(result, dict) else "invalid MSAL response"
        )
        raise AuthorizationError(
            f"Token acquisition failed: {details}. Check the app settings or sign in again."
        )
    save_cache(cache, config.output)
    print(f"\nSuccess: the MSAL token cache was saved to {config.output}.")
    print("Use this cache with the same client ID and tenant in your O365 application.")


def main(argv: Sequence[str] | None = None) -> int:
    """Map expected failures to safe stderr diagnostics and meaningful statuses."""
    args = build_parser().parse_args(argv)
    print("--- MSAL TOKEN CACHE GENERATOR ---")
    try:
        try:
            config = load_configuration(args)
        except ConfigurationError as error:
            print(f"Configuration error: {error}", file=sys.stderr)
            return ExitCode.CONFIGURATION
        except (OSError, UnicodeError):
            print(
                "Configuration error: unable to read the configuration file or inspect the output path.",
                file=sys.stderr,
            )
            return ExitCode.CONFIGURATION
        try:
            run_authorization(config, open_browser=args.open_browser)
        except AuthorizationError as error:
            print(f"Authorization failed: {error}", file=sys.stderr)
            return ExitCode.AUTHORIZATION
        except RequestException:
            print(
                "Network error: check connectivity, proxy settings, and TLS certificates, then restart for a fresh sign-in.",
                file=sys.stderr,
            )
            return ExitCode.NETWORK
        except MsalServiceError:
            print(
                "Microsoft identity service failed. Check the app registration and restart for a fresh sign-in.",
                file=sys.stderr,
            )
            return ExitCode.AUTHORIZATION
        except ConfigurationError as error:
            print(f"Cache output error: {error}", file=sys.stderr)
            return ExitCode.STORAGE
        except OSError:
            print(
                "Cache output error: unable to save the cache. Check directory permissions, free space, and file locks.",
                file=sys.stderr,
            )
            return ExitCode.STORAGE
        except (ValueError, KeyError, TypeError, RuntimeError):
            # MSAL's ValueError for mismatched state includes both state values.
            # Never print the raw exception or token response here.
            print(
                "Authorization failed: invalid application settings or an invalid MSAL response. Restart the sign-in.",
                file=sys.stderr,
            )
            return ExitCode.AUTHORIZATION
        return ExitCode.SUCCESS
    except EOFError:
        print(
            "Authorization cancelled: no redirect URL was provided on standard input.",
            file=sys.stderr,
        )
        return ExitCode.AUTHORIZATION
    except KeyboardInterrupt:
        print("\nAuthorization cancelled by the user.", file=sys.stderr)
        return ExitCode.CANCELLED


if __name__ == "__main__":
    raise SystemExit(main())
