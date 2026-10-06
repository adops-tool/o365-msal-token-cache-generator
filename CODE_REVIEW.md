# Code review and reliability changes

Reviewed on 2026-10-06. This directory contains a Python authentication utility,
not a browser extension. The review covers the application, dependency files,
setup commands, ignore rules, and documentation. The existing `.env` was checked
for configuration validity without printing its identifiers or changing the file.
The Apache license text is retained unchanged.

## Findings and resolutions

Severity describes the original impact: **High** can lose credential data or hide
a failed operation; **Medium** causes avoidable failures, unsafe diagnostics, or
misleading behavior; **Low** primarily affects setup or maintainability.

| Severity | Original issue | Resolution |
| --- | --- | --- |
| High | `open(..., "w")` truncated an existing cache before the new write completed. Disk errors or interruption could destroy the usable cache. | Validate the new cache, write to a same-directory temporary file, flush and fsync it, close it, then atomically replace the target. Failure tests verify preservation and cleanup. |
| High | Handled failures returned `None`; the executable therefore exited with status `0`. An older cache could make a failed run appear successful. | Return explicit statuses and use `SystemExit(main())`. A subprocess test verifies the actual process status. |
| Medium | MSAL construction and flow initiation occurred outside the exception handler. Requests transport failures and filesystem errors were not handled. | Protect the entire run, classify configuration/authentication/network/storage failures, and report them to stderr. |
| Medium | Network requests had no application-specified timeout. | Pass a validated, finite timeout to MSAL, with a default of 30 seconds and a `--timeout` override. This bounds connect/read waits, not total sign-in duration. |
| Medium | URL validation only searched for the substring `code=`. It did not verify the redirect endpoint or distinguish a provider denial from malformed input. | Parse and validate the exact HTTPS native-client endpoint, expected state, and code/error alternatives before passing the response to MSAL. |
| Medium | Duplicate parameters silently kept the first value, hiding ambiguous callbacks. | Reject duplicate keys, including percent-encoded aliases, and preserve unambiguous optional fields such as `session_state` and `client_info`. |
| Medium | Printed exceptions and provider descriptions could contain state, personal data, credential fragments, or terminal escapes. Real MSAL nonce errors can also include sensitive values. | Avoid raw exception/response output. Display only known OAuth errors, bounded AADSTS codes, and validated correlation UUIDs. Handle MSAL RuntimeError safely. |
| Medium | No guarantee that a successful response contained a reusable persisted cache. | Require nonempty access-token, refresh-token, and account sections before touching the output file. |
| Medium | Configuration loaded at import time and settings became stale for later calls. The comment claiming import safety was inaccurate. | Read configuration per invocation without mutating process variables; imports make no configuration or network calls. |
| Medium | Nonempty but invalid identifiers could reach authority discovery. `.env` discovery could depend on parent directories. | Validate the client UUID, supported tenant forms, scopes, and output path. Read a predictable `.env` beside the script or an explicit `--env-file`. |
| Medium | Dependencies were completely unpinned. Installing the same source later could select materially different releases. | Pin direct dependencies and add a complete reviewed runtime version set in `requirements-lock.txt`. Version pins do not provide wheel hashes. |
| Low | Setup created `venv` but activated `.venv`, and installation occurred before activation. | Use `.venv` consistently and invoke its Python directly for both pip and the generator. |
| Low | The README described tests and CI templates that did not exist as executable project files, alongside speculative deployment examples. | Replace it with documentation for implemented behavior; add actual offline tests, Ruff configuration, and a GitHub Actions workflow. |
| Low | README examples included invalid test exception construction and a device-flow call without flow initiation. Long examples made executable behavior difficult to find. | Remove unsupported recipes, verify current MSAL/O365 APIs, and provide focused setup, usage, troubleshooting, and cache-loading examples. |
| Low | Custom output/scopes required source edits. An invalid paste forced the user to start again. | Add CLI options and up to three local paste-validation attempts. Exchange the one-time code only once at the application layer. |

The default delegated permissions were retained for compatibility. They include
mail modification and sending; applications that only read mail should explicitly
request a smaller scope set. No artificial claim of faster Microsoft sign-in or
permanent refresh-token validity is made.

## Architecture after the review

The project remains a single application module with three direct runtime
dependencies. The standard library supplies argument parsing, immutable
configuration, URL handling, and atomic filesystem operations. No framework,
database, daemon, local callback server, or client secret was added.

`auth.py` separates six responsibilities:

1. `build_parser()` exposes supported CLI settings while preserving no-argument use.
2. `load_configuration()` resolves and validates settings for this invocation.
3. `parse_authorization_response()` checks the browser callback before redemption.
4. `run_authorization()` delegates the protocol and cache schema to MSAL.
5. `save_cache()` validates and publishes the serialized cache.
6. `main()` converts expected failures and cancellation into process statuses.

Each function has an English docstring. Comments explain protocol boundaries,
state lifetime, configuration precedence, credential-safe diagnostics, Windows
file replacement, and the limits of file permissions. README text, configuration
examples, setup instructions, tests, and CI comments are in English.

## Compatibility and migration

- Existing `python auth.py`, `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, native-client
  redirect, default scopes, and `o365_token.txt` remain supported.
- The default output remains in the current working directory. The default `.env`
  is beside `auth.py`; parent-directory discovery is no longer used.
- The existing local configuration passed validation. Its contents were not
  changed, and no real cache was generated during the review.
- The new tested dependency set requires Python 3.10 or newer. Python 3.9 users
  must upgrade before installing it.
- Configuration values are literal; dotenv interpolation is disabled. Explicitly
  empty process settings are errors, even when the file contains values.
- `main(argv=None)` now returns an integer status. Programmatic callers should
  supply a list such as `auth.main([])` to avoid inheriting unrelated command-line
  arguments and should check its return value.
- Dynamic module constants `CLIENT_ID`, `TENANT_ID`, and `AUTHORITY` are replaced
  by invocation-scoped configuration. External code that patched those globals
  should set environment values or use `Configuration`/`run_authorization()`.
- A successful run replaces the selected account cache; it does not merge old
  accounts. A failed run does not deliberately delete the previous cache.
- Runtime dependencies are pinned separately from O365, which is installed only
  for interoperability tests. O365 2.1.10 successfully loads a cache populated by
  real MSAL with synthetic token data.

## Verification

Checks performed locally on Windows with Python 3.14:

- Offline regression suite: 49 tests discovered, 47 passed, and two skipped for
  Windows filesystem limitations. This includes actual MSAL protocol methods and actual
  O365 cache loading. HTTP is replaced by synthetic transport responses; real
  Microsoft services and accounts are never used by these tests.
- Ruff linting and formatting checks over project files, with virtual environments
  and original-file backups explicitly excluded.
- `pip check` for installed dependency consistency.
- A fresh-install dry run for the fully pinned runtime dependency set.
- A Windows Python 3.10 wheel-resolution dry run for the developer dependency set.
  This is an installation compatibility check, not execution under Python 3.10.
- Existing configuration validation without printing local identifiers.
- `pip-audit` 2.10.1 with `--disable-pip --no-deps -r requirements-lock.txt`: no known
  vulnerabilities were reported for the eleven pinned runtime packages at review
  time. All runtime dependencies are listed explicitly in that audited file.

See the test files for individual scenarios. Platform-specific POSIX permissions
and actual symlink creation are checked only when the platform permits them. The
GitHub Actions matrix covers Windows/Linux and Python 3.10/3.12/3.14, but hosted
checks have not run from this local directory.

During local tooling setup, an initial Ruff exclusion replaced its built-in
exclusions and allowed formatting of the newly created environment's dependencies.
The configuration now extends exclusions correctly. All affected dependencies
and pip were reinstalled from their wheels before verification continued; the
user's global Python installation was not modified. The environment then passed
the real MSAL/O365 tests and `pip check`.

## Practical limits

- A real end-to-end browser sign-in, MFA/consent, Graph access, and refresh-token
  renewal require the user's Microsoft account and tenant. They were not performed.
  Offline tests cannot prove that a particular tenant policy will allow access.
- Cache JSON is unencrypted. POSIX files use `0600`; Windows inherits the parent
  directory ACL. Atomic replacement may replace a previous file's individual ACL
  with permissions inherited by the new file, so the directory must be private.
- Atomic replacement protects against partial application writes. It is not an
  interprocess lock or a universal durability guarantee for network filesystems
  or sudden power loss. Do not write one cache from multiple processes concurrently.
- A hard process kill can leave an unpublished temporary credential file. Normal
  exceptions and Ctrl+C attempt cleanup; cleanup failures produce a warning.
- This manual redirect flow uses query parameters. MSAL recommends `form_post`
  for server-received callbacks, but that response mode cannot be copied from an
  address bar. Redirect URLs and browser history should be treated as sensitive.
- Tokens can be revoked or expire. The utility cannot bypass tenant policies or
  guarantee indefinite unattended operation.
- Pins cover package versions, not distribution hashes. Vulnerability databases
  change; rerun the audit when updating dependencies.

## Original files and attribution

This folder was not a Git checkout when inspected. Before changing existing files,
the original `auth.py`, `README.md`, `requirements.txt`, `cmd_commands.txt`, and
`.gitignore` were copied into `.review-backup/`. That local folder is ignored and
contains no copy of `.env` or a token cache. It provides original-file comparison
and recovery material, not version-control history.

The original Apache license, copyright attribution from the README, repository
and gist links, and author support/community links were retained.

## References checked

- [MSAL Python API](https://msal-python.readthedocs.io/en/stable/)
- [Microsoft token acquisition guide](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens)
- [O365 token backend source](https://github.com/O365/python-o365/blob/master/O365/utils/token.py)
- [O365 account source](https://github.com/O365/python-o365/blob/master/O365/account.py)
- Installed MSAL 1.39.0 and O365 2.1.10 implementation, exercised by the tests.
- Published package metadata on PyPI for the reviewed dependency versions.
