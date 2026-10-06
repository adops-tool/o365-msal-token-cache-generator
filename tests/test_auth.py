"""Offline regression tests using temporary files and synthetic credentials.

MSAL transport is replaced before any application is constructed. Tests never
read the project's .env, touch a real cache, or send a request to Microsoft.
"""

from __future__ import annotations

import base64
import contextlib
import importlib
import io
import json
import os
import stat
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from urllib.parse import urlencode

import msal
from msal.exceptions import MsalServiceError
from requests.exceptions import ConnectionError, Timeout

import auth

CLIENT_ID = "11111111-1111-1111-1111-111111111111"
TENANT_ID = "22222222-2222-2222-2222-222222222222"
USER_ID = "33333333-3333-3333-3333-333333333333"
STATE = "synthetic-state"
CODE = "synthetic-authorization-code"
ACCESS_TOKEN = "synthetic-access-token"
REFRESH_TOKEN = "synthetic-refresh-token"
ENVIRONMENT = {"AZURE_CLIENT_ID": CLIENT_ID, "AZURE_TENANT_ID": TENANT_ID}


def redirect(**overrides: str) -> str:
    """Build only artificial responses; no real user URL is used in a test."""
    return auth.REDIRECT_URI + "?" + urlencode({"code": CODE, "state": STATE, **overrides})


def add_synthetic_tokens(cache: msal.SerializableTokenCache) -> None:
    """Let real MSAL create its own schema from a synthetic token response."""
    client_info = (
        base64.urlsafe_b64encode(json.dumps({"uid": USER_ID, "utid": TENANT_ID}).encode())
        .decode()
        .rstrip("=")
    )
    cache.add(
        {
            "client_id": CLIENT_ID,
            "scope": list(auth.SCOPES),
            "token_endpoint": f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
            "grant_type": "authorization_code",
            "response": {
                "access_token": ACCESS_TOKEN,
                "refresh_token": REFRESH_TOKEN,
                "client_info": client_info,
                "expires_in": 3600,
                "id_token_claims": {"oid": USER_ID, "preferred_username": "test@example.invalid"},
            },
        }
    )


class ConfigurationTests(unittest.TestCase):
    """Configuration validation happens before potentially expensive network work."""

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.env_file = self.root / ".env"
        self.env_file.write_text(
            f"AZURE_CLIENT_ID={CLIENT_ID}\nAZURE_TENANT_ID={TENANT_ID}\n", encoding="utf-8"
        )
        self.environment = patch.dict(os.environ, {}, clear=True)
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def settings(self, *extra: str) -> auth.Configuration:
        args = auth.build_parser().parse_args(
            [
                "--env-file",
                str(self.env_file),
                "--output",
                str(self.root / "cache.txt"),
                *extra,
            ]
        )
        return auth.load_configuration(args)

    def test_file_configuration_and_defaults(self) -> None:
        before = dict(os.environ)
        config = self.settings()
        self.assertEqual(config.client_id, CLIENT_ID)
        self.assertEqual(config.authority, f"https://login.microsoftonline.com/{TENANT_ID}")
        self.assertEqual(config.scopes, auth.SCOPES)
        self.assertEqual(config.timeout, 30)
        self.assertEqual(dict(os.environ), before)

    def test_process_environment_takes_precedence(self) -> None:
        with patch.dict(os.environ, {"AZURE_TENANT_ID": "organizations"}):
            self.assertEqual(self.settings().tenant_id, "organizations")

    def test_empty_environment_does_not_fall_back_to_file(self) -> None:
        with patch.dict(os.environ, {"AZURE_CLIENT_ID": " "}):
            with self.assertRaises(auth.ConfigurationError):
                self.settings()

    def test_configuration_is_reread_on_each_invocation(self) -> None:
        self.assertEqual(self.settings().tenant_id, TENANT_ID)
        self.env_file.write_text(f"AZURE_CLIENT_ID={CLIENT_ID}\nAZURE_TENANT_ID=common\n")
        self.assertEqual(self.settings().tenant_id, "common")

    def test_optional_default_env_file_can_be_missing(self) -> None:
        with patch.object(auth, "DEFAULT_ENV_FILE", self.root / "missing.env"):
            with patch.dict(os.environ, ENVIRONMENT):
                args = auth.build_parser().parse_args(["--output", str(self.root / "cache.txt")])
                self.assertEqual(auth.load_configuration(args).client_id, CLIENT_ID)

    def test_explicit_missing_file_is_rejected(self) -> None:
        self.env_file.unlink()
        with self.assertRaises(auth.ConfigurationError):
            self.settings()

    def test_utf8_bom_and_whitespace(self) -> None:
        self.env_file.write_text(
            f'AZURE_CLIENT_ID=" {CLIENT_ID} "\nAZURE_TENANT_ID={TENANT_ID}\n',
            encoding="utf-8-sig",
        )
        self.assertEqual(self.settings().client_id, CLIENT_ID)

    def test_identifiers_reject_placeholder_and_url_injection(self) -> None:
        for name in ENVIRONMENT:
            for value in (
                "your-identifier",
                "../../common",
                "common?x=1",
                "https://example.invalid",
            ):
                with self.subTest(name=name, value=value):
                    with patch.dict(os.environ, {name: value}):
                        with self.assertRaises(auth.ConfigurationError):
                            self.settings()

    def test_domains_and_account_aliases_remain_supported(self) -> None:
        for tenant in ("contoso.onmicrosoft.com", "common", "organizations", "consumers"):
            with self.subTest(tenant=tenant):
                with patch.dict(os.environ, {"AZURE_TENANT_ID": tenant}):
                    self.assertEqual(self.settings().tenant_id, tenant)

    def test_scope_deduplication_and_graph_resource_prefix(self) -> None:
        config = self.settings(
            "--scopes", "User.Read", "User.Read", "https://graph.microsoft.com/Mail.Read"
        )
        self.assertEqual(config.scopes, ("User.Read", "https://graph.microsoft.com/Mail.Read"))

    def test_reserved_and_invalid_scopes(self) -> None:
        for scope in (
            "openid",
            "profile",
            "offline_access",
            ".default",
            "Mail.Read Mail.Send",
            "https://example.invalid/User.Read",
        ):
            with self.subTest(scope=scope):
                with self.assertRaises(auth.ConfigurationError):
                    self.settings("--scopes", scope)

    def test_timeout_must_be_finite_and_positive(self) -> None:
        for timeout in ("0", "-1", "nan", "inf"):
            with self.subTest(timeout=timeout):
                with self.assertRaises(auth.ConfigurationError):
                    self.settings(f"--timeout={timeout}")

    def test_invalid_output_paths_and_config_overwrite(self) -> None:
        for output in (self.root, self.root / "missing" / "cache.txt", self.env_file):
            with self.subTest(output=output):
                with self.assertRaises(auth.ConfigurationError):
                    self.settings("--output", str(output))

    def test_normalized_path_cannot_replace_configuration(self) -> None:
        child = self.root / "child"
        child.mkdir()
        # A lexical comparison misses the '..' alias for the same file.
        with self.assertRaises(auth.ConfigurationError):
            self.settings("--output", str(child / ".." / ".env"))

    def test_import_does_not_load_configuration_or_construct_msal(self) -> None:
        with patch("dotenv.dotenv_values") as read_file:
            with patch("msal.PublicClientApplication") as construct:
                importlib.reload(auth)
                read_file.assert_not_called()
                construct.assert_not_called()
        # Restore imported function bindings after checking the mocked import.
        importlib.reload(auth)


class RedirectTests(unittest.TestCase):
    """Reject malformed callbacks before any one-time code can be redeemed."""

    def test_valid_response_preserves_extra_fields_and_encoded_code(self) -> None:
        response = auth.parse_authorization_response(
            "  "
            + redirect(code="code+value/=", session_state="session", client_info="info")
            + "  ",
            STATE,
        )
        self.assertEqual(response["code"], "code+value/=")
        self.assertEqual(response["session_state"], "session")
        self.assertEqual(response["client_info"], "info")

    def test_error_redirect_is_a_valid_response(self) -> None:
        url = auth.REDIRECT_URI + "?" + urlencode({"error": "access_denied", "state": STATE})
        self.assertEqual(auth.parse_authorization_response(url, STATE)["error"], "access_denied")

    def test_wrong_endpoint_fragment_and_userinfo(self) -> None:
        urls = [
            redirect().replace("https:", "http:"),
            redirect().replace("login.microsoftonline.com", "example.invalid"),
            redirect().replace(
                "login.microsoftonline.com", "login.microsoftonline.com.example.invalid"
            ),
            redirect().replace("nativeclient", "other"),
            redirect() + "#code=anything",
            redirect().replace("login.microsoftonline.com", "user@login.microsoftonline.com"),
            redirect().replace("login.microsoftonline.com", "login.microsoftonline.com:444"),
            redirect().replace("login.microsoftonline.com", "login.microsoftonline.com:invalid"),
        ]
        for url in urls:
            with self.subTest(url=url):
                with self.assertRaises(auth.AuthorizationError):
                    auth.parse_authorization_response(url, STATE)

    def test_duplicate_and_empty_query_keys(self) -> None:
        for suffix in ("&code=second", "&state=second", "&%63ode=second", "&=value", "&flag"):
            with self.subTest(suffix=suffix):
                with self.assertRaises(auth.AuthorizationError):
                    auth.parse_authorization_response(redirect() + suffix, STATE)

    def test_state_and_response_fields(self) -> None:
        queries = [
            {"code": CODE},
            {"code": CODE, "state": "wrong"},
            {"code": CODE, "state": "non-ascii-\u2603"},
            {"state": STATE},
            {"code": "", "state": STATE},
            {"error": "", "state": STATE},
            {"code": CODE, "error": "access_denied", "state": STATE},
            {"other": "code=anything", "state": STATE},
        ]
        for query in queries:
            with self.subTest(query=query):
                with self.assertRaises(auth.AuthorizationError):
                    auth.parse_authorization_response(
                        auth.REDIRECT_URI + "?" + urlencode(query), STATE
                    )

    def test_invalid_encoding_control_characters_and_limits(self) -> None:
        urls = [
            "",
            "x" * (auth.MAX_REDIRECT_LENGTH + 1),
            redirect() + "\nmore",
            redirect() + "\x1b",
            redirect() + "&extra=%GG",
            redirect() + "&extra=%FF",
            redirect() + "".join(f"&extra{i}=x" for i in range(31)),
        ]
        for url in urls:
            with self.subTest(length=len(url)):
                with self.assertRaises(auth.AuthorizationError):
                    auth.parse_authorization_response(url, STATE)

    def test_diagnostics_do_not_echo_pasted_url(self) -> None:
        with self.assertRaises(auth.AuthorizationError) as captured:
            auth.parse_authorization_response(redirect(state="wrong"), STATE)
        self.assertNotIn(CODE, str(captured.exception))


class StorageTests(unittest.TestCase):
    """Exercise the real filesystem, including failures before atomic publication."""

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.output = self.root / "o365_token.txt"
        self.cache = msal.SerializableTokenCache()
        add_synthetic_tokens(self.cache)

    def assert_no_temporary_files(self) -> None:
        self.assertEqual(list(self.root.glob(".*.tmp")), [])

    def test_cache_round_trip_and_replacement(self) -> None:
        self.output.write_text("previous cache")
        auth.save_cache(self.cache, self.output)
        restored = msal.SerializableTokenCache()
        restored.deserialize(self.output.read_text(encoding="utf-8"))
        tokens = list(restored.search(msal.TokenCache.CredentialType.REFRESH_TOKEN))
        self.assertEqual(tokens[0]["secret"], REFRESH_TOKEN)
        self.assert_no_temporary_files()

    def test_replace_failure_preserves_existing_cache_and_cleans_up(self) -> None:
        self.output.write_text("previous cache")
        with patch("auth.os.replace", side_effect=PermissionError("synthetic locked file")):
            with self.assertRaises(PermissionError):
                auth.save_cache(self.cache, self.output)
        self.assertEqual(self.output.read_text(), "previous cache")
        self.assert_no_temporary_files()

    def test_temporary_file_creation_failure_preserves_existing_cache(self) -> None:
        self.output.write_text("previous cache")
        with patch(
            "auth.tempfile.NamedTemporaryFile",
            side_effect=PermissionError("synthetic denied directory"),
        ):
            with self.assertRaises(PermissionError):
                auth.save_cache(self.cache, self.output)
        self.assertEqual(self.output.read_text(), "previous cache")
        self.assert_no_temporary_files()

    def test_fsync_failure_preserves_existing_cache_and_cleans_up(self) -> None:
        self.output.write_text("previous cache")
        with patch("auth.os.fsync", side_effect=OSError("synthetic full disk")):
            with self.assertRaises(OSError):
                auth.save_cache(self.cache, self.output)
        self.assertEqual(self.output.read_text(), "previous cache")
        self.assert_no_temporary_files()

    def test_incomplete_or_invalid_cache_never_replaces_existing_file(self) -> None:
        self.output.write_text("previous cache")
        for serialized in (
            "{}",
            "[]",
            "invalid json",
            '{"AccessToken": {"x": {}}, "RefreshToken": {}}',
        ):
            with self.subTest(serialized=serialized):
                cache = Mock(serialize=Mock(return_value=serialized))
                with self.assertRaises(ValueError):
                    auth.save_cache(cache, self.output)
                self.assertEqual(self.output.read_text(), "previous cache")
        self.assert_no_temporary_files()

    @unittest.skipIf(os.name == "nt", "POSIX file modes do not define Windows ACLs")
    def test_posix_cache_permissions(self) -> None:
        auth.save_cache(self.cache, self.output)
        self.assertEqual(stat.S_IMODE(self.output.stat().st_mode), 0o600)

    def test_output_symlink_is_rejected_without_touching_target(self) -> None:
        target = self.root / "target.txt"
        target.write_text("untouched")
        try:
            self.output.symlink_to(target)
        except OSError:
            self.skipTest("Creating symlinks requires privileges on this Windows machine")
        with self.assertRaises(auth.ConfigurationError):
            auth.save_cache(self.cache, self.output)
        self.assertEqual(target.read_text(), "untouched")

    def test_o365_can_load_the_real_msal_schema(self) -> None:
        try:
            from O365 import Account
            from O365.utils.token import FileSystemTokenBackend
        except ImportError:
            self.skipTest("Install requirements-dev.txt to test O365 interoperability")
        auth.save_cache(self.cache, self.output)
        backend = FileSystemTokenBackend(token_path=self.root, token_filename=self.output.name)
        account = Account(
            CLIENT_ID, auth_flow_type="public", tenant_id=TENANT_ID, token_backend=backend
        )
        # This loads the cache locally; it does not call Graph or refresh a token.
        self.assertTrue(account.is_authenticated)
        self.assertEqual(backend.get_refresh_token()["secret"], REFRESH_TOKEN)
        self.assertEqual(account.connection.username, "test@example.invalid")


class CommandTests(unittest.TestCase):
    """Check orchestration, statuses, cancellation, and credential-safe output."""

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.output = self.root / "o365_token.txt"
        self.env_file = self.root / "empty.env"
        self.env_file.write_text("")
        environment = patch.dict(os.environ, ENVIRONMENT, clear=True)
        environment.start()
        self.addCleanup(environment.stop)
        self.app = Mock()
        self.flow = {
            "auth_uri": "https://login.microsoftonline.com/synthetic-authorize",
            "state": STATE,
        }
        self.app.initiate_auth_code_flow.return_value = self.flow
        self.factory = patch(
            "auth.msal.PublicClientApplication", side_effect=self.create_app
        ).start()
        self.addCleanup(patch.stopall)
        self.cache = None
        self.app.acquire_token_by_auth_code_flow.side_effect = self.acquire

    def create_app(self, *args, **kwargs):
        self.cache = kwargs["token_cache"]
        return self.app

    def acquire(self, *args, **kwargs):
        add_synthetic_tokens(self.cache)
        return {"access_token": ACCESS_TOKEN}

    def invoke(self, inputs=None, *extra: str) -> int:
        self.stdout = io.StringIO()
        self.stderr = io.StringIO()
        with contextlib.redirect_stdout(self.stdout), contextlib.redirect_stderr(self.stderr):
            with patch("builtins.input", side_effect=[redirect()] if inputs is None else inputs):
                status = auth.main(
                    [
                        "--env-file",
                        str(self.env_file),
                        "--output",
                        str(self.output),
                        *extra,
                    ]
                )
        for credential in (CODE, ACCESS_TOKEN, REFRESH_TOKEN):
            self.assertNotIn(credential, self.stdout.getvalue() + self.stderr.getvalue())
        return status

    def test_success_uses_msal_flow_cache_and_timeout(self) -> None:
        self.assertEqual(self.invoke(None, "--timeout", "12"), auth.ExitCode.SUCCESS)
        self.assertTrue(self.output.is_file())
        self.assertEqual(self.factory.call_args.kwargs["timeout"], 12)
        self.app.initiate_auth_code_flow.assert_called_once_with(
            scopes=list(auth.SCOPES), redirect_uri=auth.REDIRECT_URI, response_mode="query"
        )
        self.app.acquire_token_by_auth_code_flow.assert_called_once_with(
            self.flow, auth_response={"code": CODE, "state": STATE}
        )

    def test_configuration_failure_never_constructs_app(self) -> None:
        with patch.dict(os.environ, {"AZURE_CLIENT_ID": ""}):
            self.assertEqual(self.invoke(), auth.ExitCode.CONFIGURATION)
        self.factory.assert_not_called()
        self.assertFalse(self.output.exists())

    def test_network_failures_at_every_stage(self) -> None:
        for stage in ("construction", "initiation", "exchange"):
            with self.subTest(stage=stage):
                self.factory.side_effect = self.create_app
                self.app.initiate_auth_code_flow.side_effect = None
                self.app.acquire_token_by_auth_code_flow.side_effect = self.acquire
                if stage == "construction":
                    self.factory.side_effect = ConnectionError(ACCESS_TOKEN)
                elif stage == "initiation":
                    self.app.initiate_auth_code_flow.side_effect = Timeout(CODE)
                else:
                    self.app.acquire_token_by_auth_code_flow.side_effect = Timeout(REFRESH_TOKEN)
                self.assertEqual(self.invoke(), auth.ExitCode.NETWORK)
                self.assertFalse(self.output.exists())

    def test_invalid_paste_can_be_corrected_without_restarting_flow(self) -> None:
        self.assertEqual(self.invoke(["not a URL", redirect(state="wrong"), redirect()]), 0)
        self.app.initiate_auth_code_flow.assert_called_once()
        self.app.acquire_token_by_auth_code_flow.assert_called_once()

    def test_repeated_invalid_pastes_do_not_exchange_or_change_cache(self) -> None:
        self.output.write_text("previous cache")
        self.assertEqual(self.invoke([redirect(state="wrong")] * 3), auth.ExitCode.AUTHORIZATION)
        self.app.acquire_token_by_auth_code_flow.assert_not_called()
        self.assertEqual(self.output.read_text(), "previous cache")

    def test_provider_denial_and_diagnostics(self) -> None:
        correlation = "44444444-4444-4444-4444-444444444444"
        result = {
            "error": "access_denied",
            "error_description": f"AADSTS65004 {CODE} {ACCESS_TOKEN} \x1b[31m user@example.invalid",
            "correlation_id": correlation,
        }
        self.app.acquire_token_by_auth_code_flow.side_effect = None
        self.app.acquire_token_by_auth_code_flow.return_value = result
        error_url = auth.REDIRECT_URI + "?" + urlencode({"error": "access_denied", "state": STATE})
        self.assertEqual(self.invoke([error_url]), auth.ExitCode.AUTHORIZATION)
        self.assertIn("AADSTS65004", self.stderr.getvalue())
        self.assertIn(correlation, self.stderr.getvalue())
        self.assertNotIn("user@example.invalid", self.stderr.getvalue())
        self.assertNotIn("\x1b", self.stderr.getvalue())
        self.app.acquire_token_by_auth_code_flow.assert_called_once()

    def test_malformed_or_empty_token_response(self) -> None:
        self.app.acquire_token_by_auth_code_flow.side_effect = None
        for result in (None, [], {}, {"access_token": ""}):
            with self.subTest(result=result):
                self.app.acquire_token_by_auth_code_flow.return_value = result
                self.assertEqual(self.invoke(), auth.ExitCode.AUTHORIZATION)
                self.assertFalse(self.output.exists())

    def test_missing_flow_fields_fail_before_input(self) -> None:
        for flow in (
            {},
            {"auth_uri": "url"},
            None,
            {"auth_uri": 1, "state": "s"},
            {"auth_uri": "url", "state": 1},
        ):
            with self.subTest(flow=flow):
                self.app.initiate_auth_code_flow.return_value = flow
                self.assertEqual(self.invoke([]), auth.ExitCode.AUTHORIZATION)
                self.app.acquire_token_by_auth_code_flow.assert_not_called()

    def test_sensitive_msal_exceptions_are_not_printed_or_retried(self) -> None:
        errors = [
            ValueError(CODE),
            RuntimeError(ACCESS_TOKEN),
            MsalServiceError(error="invalid_grant", error_description=REFRESH_TOKEN),
        ]
        for error in errors:
            with self.subTest(error=type(error).__name__):
                self.app.acquire_token_by_auth_code_flow.reset_mock()
                self.app.acquire_token_by_auth_code_flow.side_effect = error
                self.assertEqual(self.invoke(), auth.ExitCode.AUTHORIZATION)
                self.app.acquire_token_by_auth_code_flow.assert_called_once()

    def test_empty_msal_cache_is_not_published(self) -> None:
        self.app.acquire_token_by_auth_code_flow.side_effect = None
        self.app.acquire_token_by_auth_code_flow.return_value = {"access_token": ACCESS_TOKEN}
        self.assertEqual(self.invoke(), auth.ExitCode.AUTHORIZATION)
        self.assertFalse(self.output.exists())

    def test_storage_error_has_distinct_status_and_preserves_old_cache(self) -> None:
        self.output.write_text("previous cache")
        with patch("auth.os.replace", side_effect=PermissionError(ACCESS_TOKEN)):
            self.assertEqual(self.invoke(), auth.ExitCode.STORAGE)
        self.assertEqual(self.output.read_text(), "previous cache")

    def test_destination_changes_after_signin_are_reported_as_storage_failure(self) -> None:
        with patch(
            "auth.save_cache", side_effect=auth.ConfigurationError("Output directory disappeared.")
        ):
            self.assertEqual(self.invoke(), auth.ExitCode.STORAGE)

    def test_actual_process_returns_configuration_exit_status(self) -> None:
        # Use the real entry point, not a direct main() call. An explicitly empty
        # configuration stops before MSAL construction or any network operation.
        process = subprocess.run(
            [
                sys.executable,
                str(Path(auth.__file__).resolve()),
                "--env-file",
                str(self.env_file),
                "--output",
                str(self.output),
            ],
            env={**os.environ, "AZURE_CLIENT_ID": "", "AZURE_TENANT_ID": ""},
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
        self.assertEqual(process.returncode, auth.ExitCode.CONFIGURATION)
        self.assertIn("Configuration error", process.stderr)
        self.assertFalse(self.output.exists())

    def test_eof_and_keyboard_interrupt(self) -> None:
        self.assertEqual(self.invoke(EOFError()), auth.ExitCode.AUTHORIZATION)
        self.assertEqual(self.invoke(KeyboardInterrupt()), auth.ExitCode.CANCELLED)
        self.assertFalse(self.output.exists())

    def test_browser_failure_falls_back_to_printed_url(self) -> None:
        with patch("auth.webbrowser.open", side_effect=OSError("synthetic browser error")):
            self.assertEqual(self.invoke(None, "--open-browser"), 0)
        self.assertIn("Open the printed URL manually", self.stdout.getvalue())

    def test_browser_is_only_opened_on_request(self) -> None:
        with patch("auth.webbrowser.open") as browser:
            self.assertEqual(self.invoke(), 0)
            browser.assert_not_called()


if __name__ == "__main__":
    unittest.main()
