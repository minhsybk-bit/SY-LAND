from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import parse_qs, urlparse

from creator_api.media_validation import valid_video_signature
from creator_api.signing import signed_download_url, valid_download_signature
from creator_api.source_policy import SourceUrlError, parse_source_url, platform_for_host


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


if __name__ == "__main__":
    unittest.main()
