"use client";

import { useEffect, useMemo, useState } from "react";

type ReleaseAsset = { name: string; size: number; browser_download_url: string; digest?: string | null };
type GithubRelease = { tag_name: string; name: string; html_url: string; published_at: string; assets: ReleaseAsset[] };

const FALLBACK: GithubRelease = {
  tag_name: "v11.6.3",
  name: "SỸ LAND 11.5.0",
  html_url: "https://github.com/minhsybk-bit/SY-LAND/releases/tag/v11.6.3",
  published_at: "2026-07-19T00:00:00Z",
  assets: [{ name: "SYLAND_Setup_11.6.3.zip", size: 138156330, browser_download_url: "https://github.com/minhsybk-bit/SY-LAND/releases/download/v11.6.3/SYLAND_Setup_11.6.3.zip", digest: "sha256:ecfc1b0ce1023beb29438a39c34d91b3f74571105c17a46f3b7d5ebcb14849c3" }],
};

export default function SoftwareRelease() {
  const [release, setRelease] = useState<GithubRelease>(FALLBACK);
  const [live, setLive] = useState(false);
  const [verifyStatus, setVerifyStatus] = useState<"idle" | "checking" | "match" | "mismatch" | "unavailable">("idle");
  const [verifiedFile, setVerifiedFile] = useState("");
  const [calculatedHash, setCalculatedHash] = useState("");

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
  const checksum = installer.digest?.replace(/^sha256:/i, "") || (version === FALLBACK.tag_name.replace(/^v/i, "") ? FALLBACK.assets[0].digest!.replace("sha256:", "") : "Xem tại trang phát hành GitHub");

  async function verifyFile(file?: File) {
    if (!file) return;
    setVerifiedFile(file.name); setCalculatedHash(""); setVerifyStatus("checking");
    try {
      const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
      const hash = Array.from(new Uint8Array(digest)).map((value) => value.toString(16).padStart(2, "0")).join("");
      setCalculatedHash(hash);
      setVerifyStatus(/^[a-f0-9]{64}$/i.test(checksum) ? (hash.toLowerCase() === checksum.toLowerCase() ? "match" : "mismatch") : "unavailable");
    } catch (reason) { console.error(reason); setVerifyStatus("unavailable"); }
  }

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
        <div className="release-verifier">
          <div><strong>Kiểm tra tệp đã tải</strong><small>SHA-256 được tính cục bộ, tệp không rời khỏi thiết bị.</small></div>
          <label className={verifyStatus === "checking" ? "disabled" : ""}><input type="file" accept=".zip,.exe,application/zip,application/x-msdownload" disabled={verifyStatus === "checking"} onChange={(event) => { void verifyFile(event.target.files?.[0]); event.target.value = ""; }} />{verifyStatus === "checking" ? "Đang kiểm tra…" : "Chọn ZIP hoặc EXE"}</label>
          {verifyStatus !== "idle" && verifyStatus !== "checking" && <div className={`verify-result ${verifyStatus}`}><span aria-hidden="true">{verifyStatus === "match" ? "✓" : verifyStatus === "mismatch" ? "!" : "i"}</span><p><b>{verifyStatus === "match" ? "Tệp chính xác, mã SHA-256 khớp." : verifyStatus === "mismatch" ? "Không khớp — không nên cài đặt tệp này." : "Đã tính mã nhưng chưa có mã chính thức để đối chiếu."}</b><small>{verifiedFile}</small>{calculatedHash && <code>{calculatedHash}</code>}</p></div>}
        </div>
        <ul className="release-trust">
          <li><span aria-hidden="true">✓</span> Tự đồng bộ bản phát hành mới nhất</li>
          <li><span aria-hidden="true">✓</span> Có thể kiểm tra tính toàn vẹn bộ cài</li>
          <li><span aria-hidden="true">✓</span> Không tự động gửi hồ sơ lên máy chủ</li>
        </ul>
      </aside>
    </section>
  );
}
