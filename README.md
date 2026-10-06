# O365 MSAL Token Cache Generator

A small Python command-line tool that signs a user in to Microsoft 365 and saves
an MSAL token cache for an O365 application. It uses a public client, requires no
client secret, and keeps the existing `python auth.py` workflow.

The cache contains credentials, including refresh tokens. It is plain JSON and
must be kept in a private directory. The tool creates the cache; your consuming
application is responsible for using it and saving subsequent token refreshes.

## Requirements

- Python 3.10 or newer.
- An internet connection to the Microsoft identity platform.
- A Microsoft Entra application registration configured for a public desktop client.
- A browser for the sign-in. The browser can run on another computer.

The direct dependencies are pinned in `requirements.txt`. Transitive dependencies
are still resolved by pip. The reviewed runtime dependency set can be installed
with `requirements-lock.txt` when a fully pinned version set is preferred.

## Quick start

Run these commands from the project directory. Using the virtual environment's
Python directly avoids activation and PowerShell execution-policy issues.

### Windows PowerShell

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-lock.txt

# Create the configuration only when it does not already exist.
if (-not (Test-Path -LiteralPath .env)) { Copy-Item .env.example .env }
# Edit .env and replace both example identifiers with your actual identifiers.
notepad .env

.venv\Scripts\python.exe auth.py
```

### Linux and macOS

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements-lock.txt

# Keep an existing local configuration.
if [ ! -e .env ]; then cp .env.example .env; fi
# Edit .env in your preferred text editor before running the tool.
.venv/bin/python auth.py
```

If `.env` is already configured, keep it. No migration is required for the two
existing settings:

```dotenv
AZURE_CLIENT_ID=your-application-client-id
AZURE_TENANT_ID=your-directory-tenant-id
```

1. Run the tool and open the printed Microsoft sign-in URL.
2. Sign in and approve the requested delegated permissions.
3. When the browser reaches the native-client redirect, copy its complete address.
   The page may be blank; the address bar contains the response.
4. Paste the URL into the waiting terminal and press Enter.
5. A successful run prints the absolute location of `o365_token.txt` and exits with status `0`.

Keep the same process running throughout sign-in: MSAL's state and PKCE verifier
live in that process. An old redirect cannot be used after restarting the tool.
Up to three paste attempts are allowed before any authorization code is exchanged.
An exchange failure requires a fresh run and sign-in.

## Microsoft Entra configuration

Use the [Microsoft public-client registration instructions](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-registration)
to register a desktop application in the intended tenant.

Configure this exact URI under the mobile/desktop application's redirect URIs:

```text
https://login.microsoftonline.com/common/oauth2/nativeclient
```

Ensure that the registration supports public-client authentication. Use the
application's **Application (client) ID** and **Directory (tenant) ID** in `.env`.
A client secret is unnecessary for this public-client flow.

The default delegated Microsoft Graph permissions are:

| Permission | Purpose |
| --- | --- |
| `Mail.ReadWrite` | Read and modify the signed-in user's mail. |
| `Mail.Send` | Send mail as the signed-in user. |
| `User.Read` | Read the signed-in user's profile. |

Choose fewer permissions if your consuming application needs less access. For
example, `--scopes User.Read Mail.Read` requests read-only mail access. Permissions
must match the registration and your tenant's consent policy. Admin approval may
be necessary. Do not supply `openid`, `profile`, or `offline_access`: MSAL adds its
reserved scopes, including the scope used to request refresh tokens.

## Command-line options

```powershell
.venv\Scripts\python.exe auth.py --help
.venv\Scripts\python.exe auth.py --open-browser
.venv\Scripts\python.exe auth.py --timeout 60
.venv\Scripts\python.exe auth.py --scopes User.Read Mail.Read
.venv\Scripts\python.exe auth.py --env-file .env.production --output service_token.txt
```

Use `.venv/bin/python` for the same commands on Linux/macOS.

| Option | Default | Behavior |
| --- | --- | --- |
| `--env-file PATH` | `.env` beside `auth.py` | Read this file; an explicitly selected file must exist. |
| `--output PATH` | `o365_token.txt` in the current directory | Save the entire cache at this location. The parent directory must exist. |
| `--scopes SCOPE ...` | `Mail.ReadWrite Mail.Send User.Read` | Request explicit delegated Graph permission names. Duplicate strings are removed. |
| `--timeout SECONDS` | `30` | Positive, finite timeout for network connection and socket reads, not the whole browser sign-in. |
| `--open-browser` | Disabled | Try to open the URL automatically; fall back to the printed URL on failure. |

Fully qualified scopes such as `https://graph.microsoft.com/Mail.Read` also work.
This tool targets the public Microsoft cloud and Microsoft Graph; other cloud
hosts, resources, redirect URIs, `.default`, and application-permission flows are
outside its scope.

Configuration is read when `main()` runs, rather than when the module is imported.
Process environment variables override file values. An explicitly empty process
variable is a configuration error and does not fall back to the file. The default
`.env` is optional when both identifiers are supplied through the environment.
Files use UTF-8, with an optional BOM. Values are literal: `${VARIABLE}` expansion
is deliberately disabled, and parent directories are not searched.

`AZURE_CLIENT_ID` must be a UUID. `AZURE_TENANT_ID` accepts a tenant UUID, a tenant
domain, or `common`, `organizations`, or `consumers`. The registration's supported
account types and tenant policies still determine which accounts can sign in.

Relative output paths remain relative to the current working directory, preserving
the original behavior. A successful run replaces the selected cache with the newly
signed-in account's cache; it does not merge accounts from a previous file.

## Using the cache with O365

The offline compatibility test covers O365 `2.1.10`. O365 is a developer-test
dependency, not a requirement for running the generator. Install the matching
version in the consuming application's environment if you use this example.

```python
"""Load an existing cache without starting another interactive sign-in."""

from pathlib import Path

from O365 import Account
from O365.utils.token import FileSystemTokenBackend

# Replace these identifiers with the same values used by the generator.
client_id = "your-application-client-id"
tenant_id = "your-directory-tenant-id"
cache_file = Path("o365_token.txt").resolve()

backend = FileSystemTokenBackend(
    token_path=cache_file.parent,
    token_filename=cache_file.name,
)
account = Account(
    client_id,
    auth_flow_type="public",
    tenant_id=tenant_id,
    token_backend=backend,
)

# This checks the local cache. It does not prove that Microsoft will accept a
# future refresh: revocation and tenant policy are evaluated by the service.
if not account.is_authenticated:
    raise RuntimeError("No reusable local cache was found. Run auth.py again.")

# Graph operations use this account, and O365 can renew tokens when required.
# The cache directory must remain writable so refreshed tokens can be saved.
```

See the [O365 token backend source](https://github.com/O365/python-o365/blob/master/O365/utils/token.py)
and [O365 account source](https://github.com/O365/python-o365/blob/master/O365/account.py)
for the underlying cache-loading behavior. Older O365 releases may use a different
cache schema or API; check the version used by your actual application.

## Error handling and cache storage

Normal instructions go to stdout. Errors go to stderr. The process status tells a
calling shell whether this run succeeded, even if an older cache still exists.

| Status | Meaning | Next step |
| --- | --- | --- |
| `0` | A complete reusable cache was saved. | Use the printed output path. |
| `1` | Authorization failed, the response/cache was invalid, or stdin ended. | Check registration/consent, then start a fresh sign-in. |
| `2` | Invalid arguments or local configuration. | Fix the arguments, file, identifiers, scopes, or destination. |
| `3` | Network or TLS failure. | Check connectivity, proxy configuration, and certificates; restart sign-in. |
| `4` | Cache storage failed after configuration validation. | Check permissions, disk space, directory changes, and file locks. |
| `130` | The user cancelled with Ctrl+C. | Start again when ready. |

Provider diagnostics include known OAuth error names, available `AADSTS` codes,
and a valid correlation ID. Raw exception messages, token responses, redirect
URLs, and arbitrary provider descriptions are not printed by this tool.

The redirect parser requires the correct HTTPS endpoint, a matching state, and
exactly one nonempty `code` or `error` field. It rejects duplicates, fragments,
invalid encoding, control characters, and excessively large queries. MSAL retains
control of PKCE and its own state and nonce checks. See the
[MSAL API documentation](https://msal-python.readthedocs.io/en/stable/).

The generator uses `response_mode=query` because the manual workflow needs the
response in the browser address bar. MSAL may display a warning recommending
`form_post`; that mode requires a callback receiver and cannot be pasted from the
address bar. Treat the redirect URL as sensitive, including browser history.

Before saving, the generator checks that the serialized cache contains nonempty
`AccessToken`, `RefreshToken`, and `Account` sections. It writes a temporary file
in the same directory, flushes it, calls `fsync`, closes it, and uses `os.replace`.
Failures before replacement preserve the existing cache. The temporary file is
removed when possible, and a cleanup failure is reported.

On POSIX systems the new file has mode `0600`. On Windows the file inherits the
output directory's ACL; `chmod` does not provide equivalent ACL protection. Use a
private directory. The file remains unencrypted for standard O365 compatibility.
Atomic replacement is not a multi-process lock or a guarantee against every
filesystem/power-loss scenario. Avoid concurrent generators and consumers writing
the same cache.

The `.gitignore` excludes `.env`, `.env.*` (except `.env.example`), the default
cache, `*_token.txt`, and atomic temporary files. If you choose a custom cache
filename outside those patterns, add it to your own ignore rules. Ignoring a file
does not make it safe to share or remove it from any existing Git history.

## Troubleshooting

| Symptom | Check |
| --- | --- |
| Configuration error | Replace example values and check environment-variable overrides. Use `--env-file` if needed. |
| State mismatch | Copy the full redirect from the current flow; keep the original terminal process running. |
| `AADSTS50011` | Confirm the exact native-client URI in the app registration. |
| `AADSTS700016` | Confirm the application exists in the selected tenant and both identifiers are correct. |
| `AADSTS65001` or consent denied | Check requested delegated scopes and the tenant's consent policy. |
| `invalid_grant` | The code may be expired or consumed. Restart the generator and sign in again. |
| Network error | Check internet access, proxy settings, TLS certificates, and `--timeout`. |
| Cache output error | Check the parent directory, free space, permissions, and Windows file locks. |
| Consumer later needs sign-in | Refresh tokens can expire or be revoked; policies can require interaction. Generate a fresh cache. |

A refresh token is not a promise of permanent unattended access. Microsoft can
require sign-in again. No live authorization or token refresh is performed by the
automated tests.

## Development and verification

```powershell
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pip check
.venv\Scripts\python.exe -m ruff check .
.venv\Scripts\python.exe -m ruff format --check .
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Use `.venv/bin/python` on Linux/macOS. Tests use temporary directories and
synthetic credentials. They exercise configuration, callback parsing, real MSAL
PKCE/state/nonce behavior through an offline HTTP transport, cancellation, network
errors, atomic-save failures, and real O365 cache loading. Platform-specific tests
are skipped when their required filesystem capability is unavailable.

GitHub Actions is configured for Windows and Linux on Python 3.10, 3.12, and 3.14.
The workflow runs when the project is placed in a GitHub repository and pushed;
adding the file locally does not mean those hosted checks have already run.

| File | Responsibility |
| --- | --- |
| `auth.py` | Configuration, CLI, callback validation, MSAL orchestration, atomic persistence. |
| `.env.example` | English configuration template with placeholder identifiers. |
| `requirements.txt` | Pinned direct runtime dependencies. |
| `requirements-lock.txt` | Reviewed direct and transitive runtime version set. |
| `requirements-dev.txt` | Runtime dependencies, Ruff, and O365 interoperability dependency. |
| `pyproject.toml` | Ruff linting and formatting configuration. |
| `tests/test_auth.py` | Configuration, validation, CLI, filesystem, and O365 regression tests. |
| `tests/test_msal_integration.py` | Real MSAL protocol behavior with synthetic HTTP responses. |
| `.github/workflows/checks.yml` | Hosted test matrix. |
| `cmd_commands.txt` | Windows setup and verification commands. |
| `CODE_REVIEW.md` | Findings, changes, compatibility notes, and verification limits. |
| `LICENSE` | Original Apache License 2.0 text. |

## License and project links

Copyright 2025 adops-tool. This project uses the [Apache License 2.0](LICENSE).
The original license text and project attribution are retained.

[Original repository](https://github.com/adops-tool/o365-msal-token-cache-generator) ·
[Original gist](https://gist.github.com/OstinUA/e1c6ab1453ff8db09c973b273e75647a)

Support and community links from the original project:
[Patreon](https://www.patreon.com/OstinFCT), [Ko-fi](https://ko-fi.com/fctostin),
[Boosty](https://boosty.to/ostinfct), [YouTube](https://www.youtube.com/@FCT-Ostin),
[Telegram](https://t.me/FCTostin).
