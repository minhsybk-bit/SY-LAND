"use client";

import { useEffect, useMemo, useState } from "react";

const NEEDS = ["OCR PDF scan", "Đối chiếu Excel", "Chuẩn hóa tên tệp", "Kho hồ sơ dùng chung", "Phân quyền người dùng", "Triển khai máy chủ riêng"];

type Draft = { unit: string; contact: string; volume: number; people: number; needs: string[]; note: string };
const EMPTY: Draft = { unit: "", contact: "", volume: 300, people: 1, needs: ["OCR PDF scan", "Đối chiếu Excel"], note: "" };

export default function TrialRequest() {
  const [draft, setDraft] = useState<Draft>(EMPTY);
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      try {
        const stored = localStorage.getItem("sy-land-trial-draft");
        if (stored) setDraft({ ...EMPTY, ...JSON.parse(stored) });
      } catch { /* Bỏ qua dữ liệu cũ không hợp lệ. */ }
    }, 0);
    return () => window.clearTimeout(timer);
  }, []);

  const plan = useMemo(() => draft.people > 5 || draft.volume > 2500 ? "Đơn vị" : draft.people > 1 || draft.volume > 300 ? "Văn phòng" : "Cá nhân", [draft.people, draft.volume]);

  function update(patch: Partial<Draft>) {
    setDraft((current) => ({ ...current, ...patch }));
    setSaved(false);
  }

  function saveDraft() {
    localStorage.setItem("sy-land-trial-draft", JSON.stringify(draft));
    setSaved(true);
  }

  function downloadRequest() {
    const content = [
      "PHIẾU NHU CẦU DÙNG THỬ SỸ LAND",
      `Ngày tạo: ${new Date().toLocaleDateString("vi-VN")}`,
      `Đơn vị/cá nhân: ${draft.unit || "Chưa nhập"}`,
      `Thông tin liên hệ: ${draft.contact || "Chưa nhập"}`,
      `Số hồ sơ dự kiến/tháng: ${draft.volume.toLocaleString("vi-VN")}`,
      `Số người sử dụng: ${draft.people}`,
      `Nhu cầu: ${draft.needs.join(", ") || "Chưa chọn"}`,
      `Gói đề xuất: ${plan}`,
      `Ghi chú: ${draft.note || "Không có"}`,
      "",
      "Lưu ý: Phiếu này được tạo cục bộ trên thiết bị và chưa được gửi tự động.",
    ].join("\r\n");
    const url = URL.createObjectURL(new Blob(["\uFEFF", content], { type: "text/plain;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `PHIEU_NHU_CAU_SY_LAND_${new Date().toISOString().slice(0, 10)}.txt`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  return (
    <section className="trial-request" aria-labelledby="trial-title">
      <div className="trial-heading"><div><p className="section-kicker">Đăng ký dùng thử</p><h3 id="trial-title">Tạo phiếu nhu cầu triển khai</h3></div><p>Điền thông tin để xác định phạm vi phù hợp. Phiếu chỉ lưu trên thiết bị này và có thể tải xuống để gửi sau.</p></div>
      <div className="trial-grid">
        <div className="trial-form">
          <label>Đơn vị hoặc cá nhân<input value={draft.unit} onChange={(event) => update({ unit: event.target.value })} placeholder="Ví dụ: Văn phòng dịch vụ đất đai" /></label>
          <label>Thông tin liên hệ<input value={draft.contact} onChange={(event) => update({ contact: event.target.value })} placeholder="Số điện thoại hoặc email" /></label>
          <label>Hồ sơ mỗi tháng<input type="number" min="1" value={draft.volume} onChange={(event) => update({ volume: Math.max(1, Number(event.target.value) || 1) })} /></label>
          <label>Số người sử dụng<input type="number" min="1" value={draft.people} onChange={(event) => update({ people: Math.max(1, Number(event.target.value) || 1) })} /></label>
          <fieldset><legend>Nhu cầu chính</legend>{NEEDS.map((need) => <label className="need-option" key={need}><input type="checkbox" checked={draft.needs.includes(need)} onChange={() => update({ needs: draft.needs.includes(need) ? draft.needs.filter((item) => item !== need) : [...draft.needs, need] })} /><span>{need}</span></label>)}</fieldset>
          <label className="trial-note">Ghi chú<textarea value={draft.note} onChange={(event) => update({ note: event.target.value })} placeholder="Mô tả loại hồ sơ, khó khăn hoặc yêu cầu riêng…" /></label>
        </div>
        <aside className="trial-summary"><span>GÓI PHÙ HỢP DỰ KIẾN</span><strong>{plan}</strong><p>{draft.volume.toLocaleString("vi-VN")} hồ sơ/tháng · {draft.people} người dùng</p><ul>{draft.needs.slice(0, 4).map((need) => <li key={need}>✓ {need}</li>)}</ul><button type="button" className="button button-primary" onClick={saveDraft}>{saved ? "Đã lưu trên thiết bị ✓" : "Lưu bản nháp"}</button><button type="button" className="button trial-download" onClick={downloadRequest}>Tải phiếu nhu cầu (.txt) ↓</button><small>Website chưa gửi thông tin lên máy chủ hoặc tự động liên hệ.</small></aside>
      </div>
    </section>
  );
}
