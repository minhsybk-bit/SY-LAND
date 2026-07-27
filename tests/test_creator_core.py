from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import parse_qs, urlparse

from creator_api.media_validation import valid_video_signature
from creator_api.errors import public_error
from creator_api.preflight import parse_env_file, validate_environment
from creator_api.signing import signed_download_url, valid_download_signature
from creator_api.source_policy import SourceUrlError, parse_source_url, platform_for_host
from creator_api.translation_validation import validated_translations


class SourcePolicyTests(unittest.TestCase):
    def test_supported_hosts_and_subdomains(self) -> None:
        self.assertEqual(platform_for_host("youtube.com"), "YouTube")
        self.assertEqual(platform_for_host("www.tiktok.com"), "TikTok")
        self.assertEqual(platform_for_host("m.bilibili.com"), "Bilibili")
        self.assertIsNone(platform_for_host("youtube.com.attacker.example"))

    def test_only_https_without_credentials(self) -> None:
        self.assertEqual(parse_source_url("https://youtu.be/abc"), ("youtu.be", "YouTube"))
        for value in (
            "http://youtube.com/watch?v=abc",
            "https://user:pass@youtube.com/watch?v=abc",
            "https://youtube.com:8443/watch?v=abc",
            "https://youtube.com:not-a-port/watch?v=abc",
            "https://example.com/video",
            "file:///etc/passwd",
        ):
            with self.subTest(value=value), self.assertRaises(SourceUrlError):
                parse_source_url(value)


class MediaValidationTests(unittest.TestCase):
    def test_mp4_and_ebml_signatures(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            mp4 = root / "video.mp4"
            mp4.write_bytes(b"\x00\x00\x00\x18ftypisom")
            webm = root / "video.webm"
            webm.write_bytes(b"\x1aE\xdf\xa3\x9fB\x86")
            fake = root / "fake.mp4"
            fake.write_bytes(b"<script>alert(1)</script>")
            self.assertTrue(valid_video_signature(mp4, ".mp4"))
            self.assertTrue(valid_video_signature(webm, ".webm"))
            self.assertFalse(valid_video_signature(fake, ".mp4"))
            self.assertFalse(valid_video_signature(mp4, ".avi"))


class DownloadSigningTests(unittest.TestCase):
    def setUp(self) -> None:
        self.settings = SimpleNamespace(
            public_api_url="https://creator-api.example.com",
            signed_url_ttl_seconds=900,
            download_signing_secret="a" * 64,
        )

    def test_signed_url_and_tamper_protection(self) -> None:
        url = signed_download_url(self.settings, "job-1", "user-1")
        parsed = urlparse(url)
        query = parse_qs(parsed.query)
        expires = int(query["expires"][0])
        signature = query["signature"][0]
        self.assertGreater(expires, int(time.time()))
        self.assertTrue(
            valid_download_signature(self.settings, "job-1", "user-1", expires, signature)
        )
        self.assertFalse(
            valid_download_signature(self.settings, "job-2", "user-1", expires, signature)
        )
        self.assertFalse(
            valid_download_signature(self.settings, "job-1", "other-user", expires, signature)
        )

    def test_expired_signature_is_rejected(self) -> None:
        self.assertFalse(
            valid_download_signature(
                self.settings,
                "job-1",
                "user-1",
                int(time.time()) - 1,
                "invalid",
            )
        )


class CreatorPreflightTests(unittest.TestCase):
    def valid_values(self) -> dict[str, str]:
        return {
            "ENVIRONMENT": "staging",
            "PUBLIC_API_URL": "https://creator-staging.syland.vn",
            "ALLOWED_ORIGINS": "https://minhsybk-bit.github.io",
            "SUPABASE_URL": "https://abc123.supabase.co",
            "SUPABASE_ANON_KEY": "anon-value",
            "SUPABASE_SERVICE_ROLE_KEY": "service-role-value",
            "OPENAI_API_KEY": "openai-value",
            "GEMINI_API_KEY": "",
            "REDIS_URL": "redis://redis:6379/0",
            "CREATOR_STORAGE_DIR": "/data/creator",
            "DOWNLOAD_SIGNING_SECRET": "a" * 64,
            "SIGNED_URL_TTL_SECONDS": "900",
            "MAX_UPLOAD_BYTES": "524288000",
            "MAX_VIDEO_SECONDS": "1800",
            "OUTPUT_RETENTION_HOURS": "24",
        }

    def test_valid_staging_environment(self) -> None:
        self.assertEqual(validate_environment(self.valid_values()), [])

    def test_placeholders_wildcard_and_missing_ai_key_are_rejected(self) -> None:
        values = self.valid_values()
        values.update(
            {
                "PUBLIC_API_URL": "https://creator-api.ten-mien-cua-ban.vn",
                "ALLOWED_ORIGINS": "*",
                "OPENAI_API_KEY": "",
                "DOWNLOAD_SIGNING_SECRET": "REPLACE_WITH_SECRET",
            }
        )
        errors = validate_environment(values)
        self.assertTrue(any("PUBLIC_API_URL" in error for error in errors))
        self.assertTrue(any("ALLOWED_ORIGINS" in error for error in errors))
        self.assertTrue(any("OPENAI_API_KEY" in error for error in errors))
        self.assertTrue(any("DOWNLOAD_SIGNING_SECRET" in error for error in errors))

    def test_parse_env_file_does_not_expand_secret_values(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ".env.creator"
            path.write_text("ENVIRONMENT=staging\nSECRET='literal-$VALUE'\n", encoding="utf-8")
            self.assertEqual(parse_env_file(path)["SECRET"], "literal-$VALUE")


class TranslationValidationTests(unittest.TestCase):
    def test_translation_rows_must_match_expected_ids(self) -> None:
        self.assertEqual(
            validated_translations(
                [{"id": 1, "vi": "Xin chào"}, {"id": 2, "vi": "Cảm ơn"}],
                {1, 2},
            ),
            {1: "Xin chào", 2: "Cảm ơn"},
        )
        invalid_rows = (
            [{"id": 1, "vi": ""}],
            [{"id": 3, "vi": "Ngoài phạm vi"}],
            [{"id": 1, "vi": "Một"}, {"id": 1, "vi": "Hai"}],
            [{"id": "không-phải-số", "vi": "Lỗi"}],
            [{"id": 1}],
        )
        for rows in invalid_rows:
            with self.subTest(rows=rows), self.assertRaises(RuntimeError):
                validated_translations(rows, {1, 2})


class PublicErrorTests(unittest.TestCase):
    def test_secrets_and_internal_paths_are_not_returned(self) -> None:
        sensitive = RuntimeError(
            "Provider failed with sk-secret-value while reading /data/creator/private-file.mp4"
        )
        message = public_error(sensitive)
        self.assertNotIn("sk-secret-value", message)
        self.assertNotIn("/data/creator", message)
        self.assertEqual(
            public_error(RuntimeError("CREATOR_QUOTA_EXCEEDED: details")),
            "Bạn đã dùng hết số phút Creator của tháng này.",
        )


class CreatorMigrationSecurityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.sql = Path("SUPABASE_CREATOR.sql").read_text(encoding="utf-8").lower()

    def test_browser_roles_are_read_only(self) -> None:
        self.assertNotIn('create policy "creator_project_owner_insert"', self.sql)
        self.assertNotIn('create policy "creator_project_owner_update"', self.sql)
        for table in (
            "creator_projects",
            "creator_jobs",
            "creator_usage_ledger",
            "creator_plan_limits",
        ):
            self.assertIn(f"revoke all on public.{table} from anon, authenticated;", self.sql)
            self.assertIn(f"grant select on public.{table} to authenticated;", self.sql)

    def test_monthly_usage_counts_unrefunded_reservations(self) -> None:
        self.assertIn("reserve_entry.event_type = 'reserve'", self.sql)
        self.assertIn("refund_entry.event_type = 'refund'", self.sql)
        self.assertIn("and not exists", self.sql)

    def test_usage_summary_is_server_only_and_never_unlimited(self) -> None:
        self.assertIn("creator_usage_summary", self.sql)
        self.assertIn("'remainingminutes'", self.sql)
        self.assertIn(
            "revoke all on function public.creator_usage_summary(uuid) from public, anon, authenticated;",
            self.sql,
        )
        self.assertIn(
            "grant execute on function public.creator_usage_summary(uuid) to service_role;",
            self.sql,
        )
        summary = self.sql[self.sql.index("create or replace function public.creator_usage_summary") :]
        summary = summary[: summary.index("create or replace function public.creator_reserve_minutes")]
        self.assertNotIn("unlimited", summary)
        self.assertNotIn("không giới hạn", summary)


if __name__ == "__main__":
    unittest.main()
