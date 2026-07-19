"use client";

import { FormEvent, useState } from "react";

const SUPABASE_URL = String(import.meta.env.VITE_SUPABASE_URL || "").trim().replace(/\/$/, "");
const SUPABASE_ANON_KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "");

export default function SupportCenter() {
  const [form, setForm] = useState({ title: "", category: "Xử lý PDF", priority: "Bình thường", details: "" });
  const [sending, setSending] = useState(false);
  const [message, setMessage] = useState("");

  async function submit(event: FormEvent) {
    event.preventDefault(); setMessage("");
    if (form.title.trim().length < 5 || form.details.trim().length < 10) { setMessage("Hãy nhập tiêu đề và mô tả vấn đề đủ chi tiết."); return; }
    let session: any = null;
    try { session = JSON.parse(localStorage.getItem("sy-land-auth-session") || "null"); } catch { /* Bỏ qua. */ }
    if (!session?.accessToken) { setMessage("Hãy đăng nhập tài khoản SỸ LAND trước khi gửi hỗ trợ."); return; }
    setSending(true);
    try {
      const response = await fetch(`${SUPABASE_URL}/rest/v1/support_tickets`, {
        method: "POST",
        headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${session.accessToken}`, "Content-Type": "application/json", Prefer: "return=representation" },
        body: JSON.stringify({ title: form.title.trim(), category: form.category, priority: form.priority, details: form.details.trim(), app_version: "Web" }),
      });
      const data = await response.json().catch(() => null);
      if (!response.ok) throw new Error(data?.message || "Không gửi được yêu cầu hỗ trợ.");
      const ticket = Array.isArray(data) ? data[0] : null;
      setForm({ title: "", category: "Xử lý PDF", priority: "Bình thường", details: "" });
      setMessage(`Đã gửi yêu cầu${ticket?.ticket_code ? ` ${ticket.ticket_code}` : ""}. SỸ LAND sẽ theo dõi và phản hồi theo thông tin tài khoản.`);
    } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không gửi được yêu cầu hỗ trợ."); }
    finally { setSending(false); }
  }

  return (
    <section className="support-center" id="ho-tro" aria-labelledby="support-title">
      <div className="support-copy"><p className="section-kicker">Hỗ trợ sau bán hàng</p><h2 id="support-title">Gửi yêu cầu kỹ thuật<br />có mã theo dõi.</h2><p>Mô tả lỗi hoặc nhu cầu hỗ trợ. Yêu cầu được gắn với tài khoản, lưu trạng thái xử lý và chuyển thẳng tới khu vực quản trị SỸ LAND.</p><ul><li>Phân loại theo Word, PDF, Excel, tài khoản hoặc bản quyền</li><li>Mức ưu tiên rõ ràng</li><li>Không đính kèm hồ sơ chứa dữ liệu cá nhân tại biểu mẫu này</li></ul></div>
      <form className="support-form" onSubmit={submit}><label>Tiêu đề<input value={form.title} onChange={(event) => setForm((value) => ({ ...value, title: event.target.value }))} placeholder="Ví dụ: Không đọc được số tờ trong PDF scan" /></label><div><label>Nhóm hỗ trợ<select value={form.category} onChange={(event) => setForm((value) => ({ ...value, category: event.target.value }))}><option>Xử lý PDF</option><option>Word và văn bản</option><option>Excel và đối chiếu</option><option>Tài khoản</option><option>Bản quyền</option><option>Cài đặt Windows</option><option>Khác</option></select></label><label>Mức ưu tiên<select value={form.priority} onChange={(event) => setForm((value) => ({ ...value, priority: event.target.value }))}><option>Thấp</option><option>Bình thường</option><option>Cao</option><option>Khẩn cấp</option></select></label></div><label>Mô tả chi tiết<textarea value={form.details} onChange={(event) => setForm((value) => ({ ...value, details: event.target.value }))} placeholder="Các bước thực hiện, kết quả mong muốn và thông báo lỗi…" /></label><button type="submit" disabled={sending}>{sending ? "Đang gửi…" : "Gửi yêu cầu hỗ trợ"}</button>{message && <p>{message}</p>}</form>
    </section>
  );
}
