"use client";

import { useState } from "react";

type Check = { label: string; ok: boolean; note: string };

export default function SystemCheck() {
  const [checks, setChecks] = useState<Check[]>([]);
  const [testedAt, setTestedAt] = useState("");

  async function runChecks() {
    const results: Check[] = [];
    results.push({ label: "Đọc và tải tệp", ok: typeof FileReader !== "undefined" && typeof Blob !== "undefined", note: "Cần để xử lý Word, PDF, Excel ngay trên thiết bị." });
    results.push({ label: "Lưu cấu hình cục bộ", ok: (() => { try { localStorage.setItem("sy-land-self-test", "1"); localStorage.removeItem("sy-land-self-test"); return true; } catch { return false; } })(), note: "Dùng cho cấu hình địa bàn, tài khoản thử nghiệm và lịch sử." });
    results.push({ label: "Kiểm tra SHA-256", ok: Boolean(globalThis.crypto?.subtle), note: "Dùng để xác minh bộ cài tải về." });
    results.push({ label: "Xử lý ảnh và OCR", ok: (() => { try { return Boolean(document.createElement("canvas").getContext("2d")); } catch { return false; } })(), note: "Cần cho PDF scan, ảnh thu nhỏ và chuyển đổi ảnh." });
    results.push({ label: "WebAssembly", ok: typeof WebAssembly !== "undefined", note: "Giúp các thư viện xử lý tài liệu hoạt động ổn định." });
    results.push({ label: "Bộ nhớ trình duyệt", ok: typeof navigator !== "undefined" && "storage" in navigator, note: "Tài liệu lớn cần đủ bộ nhớ trống và RAM." });
    setChecks(results);
    setTestedAt(new Date().toLocaleString("vi-VN"));
  }

  const ready = checks.length > 0 && checks.every((check) => check.ok);
  return <section className="system-check" aria-labelledby="system-check-title">
    <div className="system-check-head"><div><p className="section-kicker">Kiểm tra trước khi sử dụng</p><h3 id="system-check-title">Thiết bị của bạn có sẵn sàng?</h3><p>Không đọc nội dung hồ sơ và không gửi thông tin thiết bị lên máy chủ.</p></div><button type="button" className="button button-primary" onClick={() => void runChecks()}>{checks.length ? "Kiểm tra lại" : "Kiểm tra thiết bị"}<span aria-hidden="true">→</span></button></div>
    {checks.length > 0 && <><div className={`system-check-summary ${ready ? "ready" : "warning"}`}><strong>{ready ? "Thiết bị sẵn sàng" : "Cần kiểm tra trình duyệt"}</strong><span>{checks.filter((check) => check.ok).length}/{checks.length} tiêu chí đạt · {testedAt}</span></div><div className="system-check-grid">{checks.map((check) => <article key={check.label} className={check.ok ? "ok" : "failed"}><span aria-hidden="true">{check.ok ? "✓" : "!"}</span><div><strong>{check.label}</strong><small>{check.note}</small></div></article>)}</div>{!ready && <p className="system-check-help">Hãy cập nhật Chrome, Edge hoặc Cốc Cốc lên phiên bản mới, cho phép lưu dữ liệu trang và đóng bớt ứng dụng trước khi xử lý PDF lớn.</p>}</>}
  </section>;
}
