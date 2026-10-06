"""Exercise real MSAL PKCE and nonce behavior with an entirely offline transport."""

from __future__ import annotations

import base64
import hashlib
import io
import json
import tempfile
import time
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch
from urllib.parse import parse_qs, urlencode

import msal
from test_auth import ACCESS_TOKEN, CLIENT_ID, CODE, REFRESH_TOKEN, TENANT_ID, USER_ID

import auth


def base64url(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


class OfflineResponse:
    """Minimal response contract required by the real MSAL HTTP client wrapper."""

    status_code = 200
    headers = {}

    def __init__(self, data: dict) -> None:
        self.text = json.dumps(data)

    def json(self) -> dict:
        return json.loads(self.text)


class OfflineTransport:
    """Allow only discovery and token exchange; unexpected endpoints fail the test."""

    def __init__(self) -> None:
        self.posts = []
        self.token_response = {}

    def get(self, url: str, **kwargs) -> OfflineResponse:
        if "/.well-known/openid-configuration" in url:
            authority = f"https://login.microsoftonline.com/{TENANT_ID}"
            return OfflineResponse(
                {
                    "authorization_endpoint": authority + "/oauth2/v2.0/authorize",
                    "token_endpoint": authority + "/oauth2/v2.0/token",
                    "issuer": authority + "/v2.0",
                }
            )
        if "/discovery/instance" in url:
            return OfflineResponse(
                {
                    "metadata": [
                        {
                            "preferred_network": "login.microsoftonline.com",
                            "preferred_cache": "login.microsoftonline.com",
                            "aliases": ["login.microsoftonline.com"],
                        }
                    ]
                }
            )
        raise AssertionError(f"Unexpected offline GET endpoint: {url}")

    def post(self, url: str, data=None, **kwargs) -> OfflineResponse:
        if not url.endswith("/oauth2/v2.0/token"):
            raise AssertionError(f"Unexpected offline POST endpoint: {url}")
        self.posts.append(data.copy())
        return OfflineResponse(self.token_response)


class RealMsalTests(unittest.TestCase):
    """A synthetic provider response travels through the SDK's actual flow code."""

    def setUp(self) -> None:
        self.transport = OfflineTransport()
        self.cache = msal.SerializableTokenCache()
        self.app = msal.PublicClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=self.cache,
            http_client=self.transport,
            instance_discovery=False,
        )
        self.flow = self.app.initiate_auth_code_flow(
            scopes=list(auth.SCOPES), redirect_uri=auth.REDIRECT_URI, response_mode="query"
        )
        self.query = parse_qs(self.flow["auth_uri"].split("?", 1)[1])
        # The unsigned token exists only inside this artificial transport. This
        # checks MSAL's cache/nonce machinery, not provider signature verification.
        claims = {
            "aud": CLIENT_ID,
            "iss": f"https://login.microsoftonline.com/{TENANT_ID}/v2.0",
            "iat": int(time.time()),
            "exp": int(time.time()) + 3600,
            "oid": USER_ID,
            "sub": USER_ID,
            "tid": TENANT_ID,
            "preferred_username": "test@example.invalid",
            "nonce": self.query["nonce"][0],
        }
        self.transport.token_response = {
            "access_token": ACCESS_TOKEN,
            "refresh_token": REFRESH_TOKEN,
            "token_type": "Bearer",
            "expires_in": 3600,
            "scope": " ".join(auth.SCOPES),
            "client_info": base64url(json.dumps({"uid": USER_ID, "utid": TENANT_ID}).encode()),
            "id_token": self.id_token(claims),
        }
        self.claims = claims

    @staticmethod
    def id_token(claims: dict) -> str:
        return base64url(b'{"alg":"none"}') + "." + base64url(json.dumps(claims).encode()) + "."

    def response(self, state=None) -> dict[str, str]:
        url = (
            auth.REDIRECT_URI
            + "?"
            + urlencode(
                {
                    "code": CODE,
                    "state": self.flow["state"] if state is None else state,
                }
            )
        )
        return auth.parse_authorization_response(url, self.flow["state"])

    def test_real_exchange_uses_pkce_and_produces_reusable_cache(self) -> None:
        self.assertEqual(self.query["response_mode"], ["query"])
        self.assertEqual(self.query["code_challenge_method"], ["S256"])
        self.assertIn("offline_access", self.query["scope"][0].split())
        result = self.app.acquire_token_by_auth_code_flow(self.flow, self.response())
        self.assertEqual(result["access_token"], ACCESS_TOKEN)
        self.assertEqual(len(self.transport.posts), 1)
        request = self.transport.posts[0]
        self.assertEqual(request["code"], CODE)
        self.assertEqual(request["redirect_uri"], auth.REDIRECT_URI)
        self.assertEqual(
            base64url(hashlib.sha256(request["code_verifier"].encode()).digest()),
            self.query["code_challenge"][0],
        )
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "o365_token.txt"
            auth.save_cache(self.cache, output)
            restored = msal.SerializableTokenCache()
            restored.deserialize(output.read_text(encoding="utf-8"))
            self.assertTrue(list(restored.search(msal.TokenCache.CredentialType.ACCOUNT)))
            self.assertTrue(list(restored.search(msal.TokenCache.CredentialType.REFRESH_TOKEN)))

    def test_real_msal_rejects_state_mismatch_before_token_post(self) -> None:
        with self.assertRaises(ValueError):
            self.app.acquire_token_by_auth_code_flow(self.flow, {"code": CODE, "state": "wrong"})
        self.assertEqual(self.transport.posts, [])

    def test_real_msal_nonce_failure_is_reported_without_saving_or_leaking(self) -> None:
        claims = {**self.claims, "nonce": CODE}
        self.transport.token_response["id_token"] = self.id_token(claims)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "o365_token.txt"
            config = auth.Configuration(CLIENT_ID, TENANT_ID, auth.SCOPES, output, 30)
            stdout, stderr = io.StringIO(), io.StringIO()
            # Reuse the actual MSAL instance and actual initiate/exchange methods.
            # Its cache remains real and all HTTP requests use OfflineTransport.
            with patch("auth.load_configuration", return_value=config):
                with patch("auth.msal.PublicClientApplication", return_value=self.app):
                    with patch("auth.msal.SerializableTokenCache", return_value=self.cache):
                        with patch(
                            "builtins.input",
                            side_effect=lambda _: (
                                auth.REDIRECT_URI
                                + "?"
                                + urlencode(
                                    {
                                        "code": CODE,
                                        # run_authorization starts a new flow. Capture its
                                        # state from the real initiate method below.
                                        "state": initiated["state"],
                                    }
                                )
                            ),
                        ):
                            initiated = {}
                            original = self.app.initiate_auth_code_flow

                            def initiate(*args, **kwargs):
                                flow = original(*args, **kwargs)
                                initiated.update(flow)
                                return flow

                            with patch.object(
                                self.app, "initiate_auth_code_flow", side_effect=initiate
                            ):
                                with redirect_stdout(stdout), redirect_stderr(stderr):
                                    status = auth.main([])
            self.assertEqual(status, auth.ExitCode.AUTHORIZATION)
            self.assertFalse(output.exists())
            self.assertEqual(len(self.transport.posts), 1)
            for sensitive in (CODE, ACCESS_TOKEN, REFRESH_TOKEN):
                self.assertNotIn(sensitive, stdout.getvalue() + stderr.getvalue())


if __name__ == "__main__":
    unittest.main()
