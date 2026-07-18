"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";

type Account = { id: string; name: string; email: string; passwordHash: string; salt: string; role: "admin" | "user"; createdAt: string };
type BugReport = { id: string; title: string; area: string; severity: string; steps: string; expected: string; actual: string; createdAt: string; status: "Mới" | "Đang kiểm tra" | "Đã xử lý" };

const ACCOUNT_KEY = "sy-land-test-accounts";
const SESSION_KEY = "sy-land-test-session";
const BUG_KEY = "sy-land-bug-reports";

function bytesToHex(bytes: Uint8Array) { return Array.from(bytes).map((value) => value.toString(16).padStart(2, "0")).join(""); }
function randomSalt() { const bytes = crypto.getRandomValues(new Uint8Array(16)); return bytesToHex(bytes); }
async function hashPassword(password: string, salt: string) { return bytesToHex(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(`${salt}:${password}`)))); }
function downloadJson(data: unknown, name: string) { const url = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: "application/json" })); const anchor = document.createElement("a"); anchor.href = url; anchor.download = name; anchor.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); }

export default function AccountPortal() {
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [sessionId, setSessionId] = useState("");
  const [mode, setMode] = useState<"login" | "register" | "setup">("login");
  const [name, setName] = useState("Nguyễn Minh Sỹ");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [message, setMessage] = useState("");
  const [reports, setReports] = useState<BugReport[]>([]);
  const [report, setReport] = useState({ title: "", area: "Xử lý PDF", severity: "Trung bình", steps: "", expected: "", actual: "" });
  const current = useMemo(() => accounts.find((account) => account.id === sessionId) || null, [accounts, sessionId]);

  useEffect(() => {
    try {
      const storedAccounts = JSON.parse(localStorage.getItem(ACCOUNT_KEY) || "[]") as Account[];
      setAccounts(storedAccounts);
      setReports(JSON.parse(localStorage.getItem(BUG_KEY) || "[]"));
      setSessionId(localStorage.getItem(SESSION_KEY) || "");
      setMode(storedAccounts.some((account) => account.role === "admin") ? "login" : "setup");
    } catch { setMode("setup"); }
  }, []);

  function saveAccounts(next: Account[]) { setAccounts(next); localStorage.setItem(ACCOUNT_KEY, JSON.stringify(next)); }

  async function submitAccount(event: FormEvent) {
    event.preventDefault(); setMessage("");
    const normalizedEmail = email.trim().toLocaleLowerCase();
    if (!normalizedEmail || !password) { setMessage("Hãy nhập email và mật khẩu."); return; }
    if (mode === "login") {
      const account = accounts.find((item) => item.email === normalizedEmail);
      if (!account || await hashPassword(password, account.salt) !== account.passwordHash) { setMessage("Email hoặc mật khẩu chưa đúng."); return; }
      localStorage.setItem(SESSION_KEY, account.id); setSessionId(account.id); setPassword(""); setMessage("Đăng nhập thành công."); return;
    }
    if (password.length < 8) { setMessage("Mật khẩu cần ít nhất 8 ký tự."); return; }
    if (password !== confirm) { setMessage("Mật khẩu xác nhận chưa khớp."); return; }
    if (accounts.some((item) => item.email === normalizedEmail)) { setMessage("Email này đã được đăng ký trên thiết bị."); return; }
    const salt = randomSalt();
    const role: Account["role"] = mode === "setup" && !accounts.some((item) => item.role === "admin") ? "admin" : "user";
    const account: Account = { id: crypto.randomUUID(), name: name.trim() || (role === "admin" ? "Nguyễn Minh Sỹ" : "Người dùng SỸ LAND"), email: normalizedEmail, passwordHash: await hashPassword(password, salt), salt, role, createdAt: new Date().toISOString() };
    const next = [...accounts, account]; saveAccounts(next); localStorage.setItem(SESSION_KEY, account.id); setSessionId(account.id); setPassword(""); setConfirm(""); setMessage(role === "admin" ? "Đã tạo tài khoản quản trị cục bộ." : "Đăng ký thành công.");
  }

  function logout() { localStorage.removeItem(SESSION_KEY); setSessionId(""); setMode("login"); setMessage(""); }

  function saveReport(event: FormEvent) {
    event.preventDefault();
    if (!report.title.trim() || !report.actual.trim()) { setMessage("Hãy nhập tên lỗi và kết quả thực tế."); return; }
    const item: BugReport = { id: crypto.randomUUID(), ...report, createdAt: new Date().toISOString(), status: "Mới" };
    const next = [item, ...reports]; setReports(next); localStorage.setItem(BUG_KEY, JSON.stringify(next)); setReport({ title: "", area: "Xử lý PDF", severity: "Trung bình", steps: "", expected: "", actual: "" }); setMessage("Đã lưu báo cáo lỗi trên thiết bị.");
  }

  function updateReport(id: string, status: BugReport["status"]) { const next = reports.map((item) => item.id === id ? { ...item, status } : item); setReports(next); localStorage.setItem(BUG_KEY, JSON.stringify(next)); }

  if (current) return (
    <section className="account-portal" id="tai-khoan" aria-labelledby="account-title">
      <div className="account-header"><div><p className="section-kicker">Khu vực kiểm thử</p><h2 id="account-title">Xin chào, {current.name}</h2><p>{current.role === "admin" ? "Tài khoản quản trị cục bộ · Toàn quyền kiểm thử trên thiết bị này" : "Tài khoản người dùng thử nghiệm"}</p></div><button type="button" onClick={logout}>Đăng xuất</button></div>
      {current.role === "admin" ? <div className="admin-grid">
        <aside className="admin-menu"><span>QUẢN TRỊ KIỂM THỬ</span><a href="#minh-hoa">Xử lý Word/PDF/Excel</a><a href="#cong-cu-pdf">Bộ công cụ PDF</a><a href="#chuc-nang-phan-mem">Kiểm tra chức năng</a><a href="#tai-phan-mem">Kiểm tra bản phát hành</a><button type="button" onClick={() => downloadJson(reports, `SYLAND_BAO_CAO_LOI_${new Date().toISOString().slice(0,10)}.json`)} disabled={!reports.length}>Xuất báo cáo lỗi JSON</button><small>Dữ liệu tài khoản và lỗi chỉ tồn tại trên trình duyệt này.</small></aside>
        <div className="bug-workspace">
          <div className="admin-stats"><div><strong>{accounts.length}</strong><span>Tài khoản cục bộ</span></div><div><strong>{reports.length}</strong><span>Lỗi đã ghi nhận</span></div><div><strong>{reports.filter((item) => item.status === "Mới").length}</strong><span>Lỗi mới</span></div></div>
          <form className="bug-form" onSubmit={saveReport}><h3>Ghi nhận lỗi kiểm thử</h3><label>Tên lỗi<input value={report.title} onChange={(event) => setReport((value) => ({ ...value, title: event.target.value }))} placeholder="Mô tả ngắn gọn" /></label><div><label>Khu vực<select value={report.area} onChange={(event) => setReport((value) => ({ ...value, area: event.target.value }))}><option>Xử lý PDF</option><option>Word/PDF/Excel</option><option>Địa bàn và mã xã</option><option>Đổi tên hàng loạt</option><option>Đối chiếu Excel</option><option>Đăng nhập</option><option>Giao diện</option><option>Khác</option></select></label><label>Mức độ<select value={report.severity} onChange={(event) => setReport((value) => ({ ...value, severity: event.target.value }))}><option>Thấp</option><option>Trung bình</option><option>Cao</option><option>Nghiêm trọng</option></select></label></div><label>Các bước tái hiện<textarea value={report.steps} onChange={(event) => setReport((value) => ({ ...value, steps: event.target.value }))} placeholder="1. Mở… 2. Chọn… 3. Nhấn…" /></label><label>Kết quả mong muốn<textarea value={report.expected} onChange={(event) => setReport((value) => ({ ...value, expected: event.target.value }))} /></label><label>Kết quả thực tế<textarea value={report.actual} onChange={(event) => setReport((value) => ({ ...value, actual: event.target.value }))} /></label><button type="submit">Lưu báo cáo lỗi</button></form>
          <div className="bug-list"><h3>Danh sách lỗi gần đây</h3>{!reports.length ? <p>Chưa có lỗi được ghi nhận.</p> : reports.slice(0, 20).map((item) => <article key={item.id}><div><span className={`severity ${item.severity.toLocaleLowerCase("vi-VN")}`}>{item.severity}</span><strong>{item.title}</strong><small>{item.area} · {new Date(item.createdAt).toLocaleString("vi-VN")}</small></div><select value={item.status} onChange={(event) => updateReport(item.id, event.target.value as BugReport["status"])}><option>Mới</option><option>Đang kiểm tra</option><option>Đã xử lý</option></select></article>)}</div>
        </div>
      </div> : <div className="user-test-panel"><h3>Tài khoản dùng thử đã sẵn sàng</h3><p>Bạn có thể sử dụng các công cụ công khai trên website. Chức năng đồng bộ tài khoản nhiều thiết bị sẽ có ở bản thương mại.</p><a className="button button-primary" href="#minh-hoa">Mở công cụ</a></div>}
      {message && <p className="account-message">{message}</p>}
    </section>
  );

  return (
    <section className="account-portal account-auth" id="tai-khoan" aria-labelledby="account-title">
      <div className="auth-copy"><p className="section-kicker">Tài khoản kiểm thử</p><h2 id="account-title">{mode === "setup" ? "Thiết lập quản trị viên SỸ LAND" : mode === "register" ? "Đăng ký tài khoản dùng thử" : "Đăng nhập SỸ LAND"}</h2><p>{mode === "setup" ? "Thiết lập mật khẩu quản trị lần đầu cho anh Nguyễn Minh Sỹ trên thiết bị này." : "Đây là chế độ kiểm thử cục bộ. Tài khoản chưa đồng bộ sang máy khác."}</p><ul><li>✓ Mật khẩu không được ghi trong mã website</li><li>✓ Chỉ lưu mã băm trên trình duyệt</li><li>✓ Admin có khu vực ghi nhận và xuất báo cáo lỗi</li></ul></div>
      <form className="auth-form" onSubmit={submitAccount}>{mode !== "login" && <label>Họ và tên<input value={name} onChange={(event) => setName(event.target.value)} autoComplete="name" /></label>}<label>Email<input type="email" value={email} onChange={(event) => setEmail(event.target.value)} autoComplete="email" placeholder="Nhập email quản trị" /></label><label>Mật khẩu<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete={mode === "login" ? "current-password" : "new-password"} placeholder="Tối thiểu 8 ký tự" /></label>{mode !== "login" && <label>Xác nhận mật khẩu<input type="password" value={confirm} onChange={(event) => setConfirm(event.target.value)} autoComplete="new-password" /></label>}<button type="submit">{mode === "setup" ? "Tạo tài khoản admin" : mode === "register" ? "Đăng ký" : "Đăng nhập"}</button>{message && <p className="account-message">{message}</p>}{mode !== "setup" && <p className="auth-switch">{mode === "login" ? <>Chưa có tài khoản? <button type="button" onClick={() => { setMode("register"); setMessage(""); }}>Đăng ký dùng thử</button></> : <>Đã có tài khoản? <button type="button" onClick={() => { setMode("login"); setMessage(""); }}>Đăng nhập</button></>}</p>}</form>
    </section>
  );
}
