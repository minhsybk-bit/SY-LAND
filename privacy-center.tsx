"use client";

import { FormEvent, useEffect, useState } from "react";

type PrivacyRequest = {
  id: string;
  requestCode: string;
  requestType: "Xuất dữ liệu" | "Xóa tài khoản";
  status: "Mới" | "Đang xác minh" | "Đang xử lý" | "Hoàn tất" | "Từ chối";
  note: string;
  adminNote: string;
  createdAt: string;
  updatedAt: string;
};

const SUPABASE_URL = String(import.meta.env.VITE_SUPABASE_URL || "").trim().replace(/\/$/, "");
const SUPABASE_ANON_KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "");
const SESSION_KEY = "sy-land-auth-session";

function getSession() {
  try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); }
  catch { return null; }
}

async function dataRequest(path: string, token: string, options: { method?: "GET" | "POST" | "PATCH"; payload?: unknown } = {}) {
  const response = await fetch(`${SUPABASE_URL}/rest/v1${path}`, {
    method: options.method || "GET",
    headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: "return=representation" },
    ...(options.payload === undefined ? {} : { body: JSON.stringify(options.payload) }),
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) throw new Error(data?.message || data?.hint || "Không kết nối được máy chủ quyền dữ liệu.");
  return data;
}

function saveJson(data: unknown, filename: string) {
  const url = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" }));
  const link = document.createElement("a"); link.href = url; link.download = filename; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function mapRequest(item: any): PrivacyRequest {
  return { id: item.id, requestCode: item.request_code, requestType: item.request_type, status: item.status, note: item.note || "", adminNote: item.admin_note || "", createdAt: item.created_at, updatedAt: item.updated_at };
}

export default function PrivacyCenter() {
  const [requests, setRequests] = useState<PrivacyRequest[]>([]);
  const [note, setNote] = useState("");
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const session = typeof window === "undefined" ? null : getSession();
  const isAdmin = session?.account?.role === "admin";

  async function loadRequests() {
    const session = getSession();
    if (!session?.accessToken) return;
    try {
      const rows = await dataRequest("/privacy_requests?select=id,request_code,request_type,status,note,admin_note,created_at,updated_at&order=created_at.desc", session.accessToken);
      setRequests(Array.isArray(rows) ? rows.map(mapRequest) : []);
    } catch { /* Bảng có thể chưa được quản trị viên kích hoạt. */ }
  }

  useEffect(() => { loadRequests(); }, []);

  async function exportMyData() {
    const session = getSession();
    if (!session?.accessToken || !session?.account) { setMessage("Hãy đăng nhập tài khoản SỸ LAND trước khi tải dữ liệu."); return; }
    setBusy(true); setMessage("Đang tổng hợp dữ liệu tài khoản…");
    try {
      const token = session.accessToken;
      const endpoints = [
        ["profile", `/profiles?id=eq.${encodeURIComponent(session.account.id)}&select=id,full_name,email,role,terms_version,terms_accepted_at,created_at`],
        ["licenses", "/licenses?select=id,code,customer,email,plan,expires_at,status,max_devices,created_at"],
        ["devices", "/device_activations?select=id,license_id,device_name,app_version,status,first_seen_at,last_seen_at"],
        ["consultations", "/leads?select=id,unit,contact,monthly_volume,people,needs,proposed_plan,note,status,created_at"],
        ["supportTickets", "/support_tickets?select=id,ticket_code,title,category,priority,details,status,app_version,created_at,updated_at"],
        ["privacyRequests", "/privacy_requests?select=id,request_code,request_type,status,note,admin_note,created_at,updated_at"],
      ] as const;
      const entries = await Promise.all(endpoints.map(async ([key, path]) => {
        try { return [key, await dataRequest(path, token)] as const; }
        catch (reason) { return [key, { unavailable: reason instanceof Error ? reason.message : "Không truy xuất được" }] as const; }
      }));
      saveJson({ product: "SỸ LAND", exportedAt: new Date().toISOString(), account: session.account, ...Object.fromEntries(entries), notice: "Tệp hồ sơ Word, PDF, Excel không nằm trong bản xuất vì được xử lý cục bộ trên thiết bị." }, `SYLAND_DU_LIEU_CA_NHAN_${new Date().toISOString().slice(0, 10)}.json`);
      setMessage("Đã tạo tệp dữ liệu cá nhân. Hãy bảo quản tệp tại nơi an toàn.");
    } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không xuất được dữ liệu."); }
    finally { setBusy(false); }
  }

  async function submitDeletion(event: FormEvent) {
    event.preventDefault();
    const session = getSession();
    if (!session?.accessToken) { setMessage("Hãy đăng nhập trước khi gửi yêu cầu xóa tài khoản."); return; }
    if (!window.confirm("Gửi yêu cầu xóa tài khoản? Tài khoản chưa bị xóa ngay; SỸ LAND sẽ xác minh trước khi xử lý.")) return;
    setBusy(true); setMessage("");
    try {
      const rows = await dataRequest("/privacy_requests", session.accessToken, { method: "POST", payload: { request_type: "Xóa tài khoản", note: note.trim() } });
      const created = Array.isArray(rows) ? rows[0] : null;
      setNote(""); await loadRequests();
      setMessage(`Đã tiếp nhận yêu cầu${created?.request_code ? ` ${created.request_code}` : ""}. Bạn vẫn có thể đăng nhập trong thời gian xác minh.`);
    } catch (reason) { setMessage(reason instanceof Error ? `${reason.message} Quản trị viên cần chạy SUPABASE_PRIVACY_REQUESTS.sql một lần.` : "Không gửi được yêu cầu."); }
    finally { setBusy(false); }
  }

  async function updateRequest(item: PrivacyRequest, status: PrivacyRequest["status"]) {
    const currentSession = getSession();
    if (!isAdmin || !currentSession?.accessToken) return;
    const adminNote = window.prompt("Phản hồi hiển thị cho người dùng (có thể để trống):", item.adminNote) ?? item.adminNote;
    setBusy(true); setMessage("");
    try {
      await dataRequest(`/privacy_requests?id=eq.${encodeURIComponent(item.id)}`, currentSession.accessToken, { method: "PATCH", payload: { status, admin_note: adminNote.trim() } });
      await loadRequests(); setMessage(`Đã cập nhật ${item.requestCode} sang trạng thái “${status}”.`);
    } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không cập nhật được yêu cầu."); }
    finally { setBusy(false); }
  }

  return (
    <section className="privacy-center" aria-labelledby="privacy-center-title">
      <div className="privacy-intro"><p className="section-kicker">Quyền dữ liệu cá nhân</p><h2 id="privacy-center-title">Tự xem, tải và yêu cầu xóa dữ liệu.</h2><p>Tài khoản đã đăng nhập có thể tải bản sao dữ liệu do máy chủ SỸ LAND lưu giữ. Yêu cầu xóa luôn qua bước xác minh để tránh mất tài khoản ngoài ý muốn.</p><ul><li>Xuất JSON có thời điểm và nguồn rõ ràng</li><li>Không đưa hồ sơ xử lý cục bộ vào bản xuất</li><li>Có mã và trạng thái theo dõi yêu cầu xóa</li></ul></div>
      <div className="privacy-actions"><article><span>01</span><h3>Tải dữ liệu của tôi</h3><p>Bao gồm hồ sơ tài khoản, bản quyền, thiết bị, yêu cầu tư vấn, phiếu hỗ trợ và lịch sử yêu cầu quyền dữ liệu.</p><button type="button" onClick={exportMyData} disabled={busy}>Tải tệp JSON</button></article><form onSubmit={submitDeletion}><span>02</span><h3>Yêu cầu xóa tài khoản</h3><p>Gửi yêu cầu để quản trị viên xác minh bản quyền, nghĩa vụ hỗ trợ và danh tính trước khi xóa.</p><label>Lý do hoặc lưu ý (không bắt buộc)<textarea value={note} onChange={(event) => setNote(event.target.value)} maxLength={1000} placeholder="Thông tin giúp SỸ LAND xác minh yêu cầu…" /></label><button type="submit" disabled={busy}>Gửi yêu cầu xóa</button></form></div>
      {message && <p className="privacy-message" role="status">{message}</p>}
      {requests.length > 0 && <div className="privacy-history"><h3>{isAdmin ? "Quản trị yêu cầu quyền dữ liệu" : "Yêu cầu gần đây"}</h3>{requests.map((item) => <article key={item.id}><div><strong>{item.requestCode}</strong><span>{item.requestType}</span><small>{new Date(item.createdAt).toLocaleString("vi-VN")}</small></div>{isAdmin ? <select aria-label={`Trạng thái ${item.requestCode}`} value={item.status} disabled={busy} onChange={(event) => updateRequest(item, event.target.value as PrivacyRequest["status"])}><option>Mới</option><option>Đang xác minh</option><option>Đang xử lý</option><option>Hoàn tất</option><option>Từ chối</option></select> : <b data-status={item.status}>{item.status}</b>}{item.adminNote && <p>Phản hồi: {item.adminNote}</p>}</article>)}</div>}
    </section>
  );
}
