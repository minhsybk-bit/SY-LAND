"use client";

import { useEffect, useMemo, useState } from "react";

type ReleaseAsset = { name: string; size: number; browser_download_url: string; digest?: string | null };
type GithubRelease = { tag_name: string; name: string; html_url: string; published_at: string; assets: ReleaseAsset[] };

const FALLBACK: GithubRelease = {
  tag_name: "v11.2.0",
  name: "SỸ LAND 11.2.0",
  html_url: "https://github.com/minhsybk-bit/SY-LAND/releases/tag/v11.2.0",
  published_at: "2026-07-18T00:00:00Z",
  assets: [{ name: "SYLAND_Setup_11.2.0.zip", size: 136288437, browser_download_url: "https://github.com/minhsybk-bit/SY-LAND/releases/download/v11.2.0/SYLAND_Setup_11.2.0.zip", digest: "sha256:0661eba7ac01eaa1e893b007280a2846722e44ce771639f9cd60e0f806a07890" }],
};

export default function SoftwareRelease() {
  const [release, setRelease] = useState<GithubRelease>(FALLBACK);
  const [live, setLive] = useState(false);

  useEffect(() => {
    const controller = new AbortController();
    fetch("https://api.github.com/repos/minhsybk-bit/SY-LAND/releases/latest", { headers: { Accept: "application/vnd.github+json" }, signal: controller.signal })
      .then((response) => response.ok ? response.json() : Promise.reject(new Error("Không đọc được bản phát hành")))
      .then((data: GithubRelease) => { if (data?.tag_name && data.assets?.length) { setRelease(data); setLive(true); } })
      .catch(() => { /* Dùng thông tin dự phòng để nút tải luôn hoạt động. */ });
    return () => controller.abort();
  }, []);

  const installer = useMemo(() => release.assets.find((asset) => asset.name.toLowerCase().endsWith(".exe")) || release.assets[0] || FALLBACK.assets[0], [release]);
  const version = release.tag_name.replace(/^v/i, "");
  const size = `${(installer.size / 1024 / 1024).toLocaleString("vi-VN", { maximumFractionDigits: 1 })} MB`;
  const checksum = installer.digest?.replace(/^sha256:/i, "") || (version === "11.2.0" ? FALLBACK.assets[0].digest!.replace("sha256:", "") : "Xem tại trang phát hành GitHub");

  return (
    <section className="software-release" id="tai-phan-mem" aria-labelledby="release-title">
      <div className="release-copy">
        <p className="section-kicker">Tiện ích hỗ trợ làm sạch CSDL đất đai SỸ LAND</p>
        <h2 id="release-title">Xử lý hồ sơ ngay trên máy tính.<br />Chủ động và riêng tư hơn.</h2>
        <p>Bản cài đặt dành cho người dùng cần xử lý số lượng hồ sơ lớn, làm việc ngoại tuyến và sử dụng các công cụ chuyên sâu trên Windows.</p>
        <div className="release-actions">
          <a className="button button-primary" href={installer.browser_download_url}>Tải SỸ LAND {version} <span aria-hidden="true">↓</span></a>
          <a className="text-link" href={release.html_url} target="_blank" rel="noreferrer">Xem ghi chú phát hành <span aria-hidden="true">›</span></a>
        </div>
        <p className="release-caution"><span aria-hidden="true">i</span> Chỉ tải từ website chính thức này. Windows có thể yêu cầu xác nhận khi cài đặt phần mềm mới.</p>
      </div>
      <aside className="release-card" aria-label="Thông tin phiên bản phần mềm">
        <div className="release-card-top"><span className="windows-mark" aria-hidden="true">⊞</span><div><small>PHIÊN BẢN ỔN ĐỊNH {live ? "· ĐÃ ĐỒNG BỘ" : ""}</small><strong>SỸ LAND {version}</strong></div><span className="stable-badge">Mới nhất</span></div>
        <dl>
          <div><dt>Hệ điều hành</dt><dd>Windows 10/11</dd></div>
          <div><dt>Tệp cài đặt</dt><dd>{installer.name}</dd></div>
          <div><dt>Dung lượng</dt><dd>{size}</dd></div>
          <div><dt>Ngày phát hành</dt><dd>{new Date(release.published_at).toLocaleDateString("vi-VN")}</dd></div>
        </dl>
        <div className="checksum"><span>SHA-256</span><code>{checksum}</code></div>
        <ul className="release-trust">
          <li><span aria-hidden="true">✓</span> Tự đồng bộ bản phát hành mới nhất</li>
          <li><span aria-hidden="true">✓</span> Có thể kiểm tra tính toàn vẹn bộ cài</li>
          <li><span aria-hidden="true">✓</span> Không tự động gửi hồ sơ lên máy chủ</li>
        </ul>
      </aside>
    </section>
  );
}
