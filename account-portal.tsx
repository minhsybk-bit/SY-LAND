"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";

type Account = { id: string; name: string; email: string; passwordHash: string; salt: string; role: "admin" | "user"; createdAt: string };
type BugReport = { id: string; title: string; area: string; severity: string; steps: string; expected: string; actual: string; createdAt: string; status: "Mới" | "Đang kiểm tra" | "Đã xử lý" };
type License = { id: string; code: string; customer: string; email: string; plan: "Cá nhân" | "Văn phòng" | "Đơn vị"; expiresAt: string; status: "Hoạt động" | "Đã khóa"; createdAt: string };

const ACCOUNT_KEY = "sy-land-test-accounts";
const SESSION_KEY = "sy-land-test-session";
const BUG_KEY = "sy-land-bug-reports";
const LICENSE_KEY = "sy-land-test-licenses";
const ACTIVATION_KEY = "sy-land-test-activations";
const REMOTE_SESSION_KEY = "sy-land-auth-session";
const SUPABASE_URL = String(import.meta.env.VITE_SUPABASE_URL || "").trim().replace(/\/$/, "");
const SUPABASE_ANON_KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "");
const AUTH_INPUT_PRESENT = Boolean(SUPABASE_URL || SUPABASE_ANON_KEY);
const SUPABASE_URL_VALID = /^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(SUPABASE_URL);
const REMOTE_AUTH = Boolean(SUPABASE_URL && SUPABASE_ANON_KEY && SUPABASE_URL_VALID);
const AUTH_CONFIG_ERROR = (AUTH_INPUT_PRESENT && !REMOTE_AUTH) || (import.meta.env.PROD && !REMOTE_AUTH);

function bytesToHex(bytes: Uint8Array) { return Array.from(bytes).map((value) => value.toString(16).padStart(2, "0")).join(""); }
function randomSalt() { const bytes = crypto.getRandomValues(new Uint8Array(16)); return bytesToHex(bytes); }
async function hashPassword(password: string, salt: string) { return bytesToHex(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(`${salt}:${password}`)))); }
function downloadJson(data: unknown, name: string) { const url = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: "application/json" })); const anchor = document.createElement("a"); anchor.href = url; anchor.download = name; anchor.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); }
function generateLicenseCode() { const bytes = crypto.getRandomValues(new Uint8Array(10)); const value = Array.from(bytes).map((item) => (item % 36).toString(36).toUpperCase()).join(""); return `SYL-${value.slice(0, 5)}-${value.slice(5)}`; }
async function remoteAuth(path: string, payload: Record<string, unknown>, options: { method?: "POST" | "PUT"; accessToken?: string } = {}) {
  if (!REMOTE_AUTH) throw new Error("Máy chủ tài khoản chưa được cấu hình đúng.");
  const response = await fetch(`${SUPABASE_URL}/auth/v1${path}`, {
    method: options.method || "POST",
    headers: { apikey: SUPABASE_ANON_KEY, "Content-Type": "application/json", ...(options.accessToken ? { Authorization: `Bearer ${options.accessToken}` } : {}) },
    body: JSON.stringify(payload),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.msg || data.message || data.error_description || "Yêu cầu xác thực thất bại.");
  return data;
}

async function remoteData(path: string, accessToken: string, options: { method?: "GET" | "POST" | "PATCH"; payload?: unknown } = {}) {
  const response = await fetch(`${SUPABASE_URL}/rest/v1${path}`, {
    method: options.method || "GET",
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    ...(options.payload === undefined ? {} : { body: JSON.stringify(options.payload) }),
  });
  const data = await response.json().catch(() => null);
  if (!response.ok) throw new Error(data?.message || data?.hint || "Không đồng bộ được dữ liệu tài khoản.");
  return data;
}

function accountFromRemoteUser(user: any): Account {
  return { id: user.id, name: user.user_metadata?.full_name || user.email?.split("@")[0] || "Người dùng SỸ LAND", email: user.email || "", passwordHash: "", salt: "", role: user.app_metadata?.role === "admin" ? "admin" : "user", createdAt: user.created_at || new Date().toISOString() };
}

export default function AccountPortal() {
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [sessionId, setSessionId] = useState("");
  const [remoteToken, setRemoteToken] = useState("");
  const [mode, setMode] = useState<"login" | "register" | "setup" | "forgot" | "reset">("login");
  const [name, setName] = useState("Nguyễn Minh Sỹ");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [message, setMessage] = useState("");
  const [reports, setReports] = useState<BugReport[]>([]);
  const [licenses, setLicenses] = useState<License[]>([]);
  const [licenseDraft, setLicenseDraft] = useState({ customer: "", email: "", plan: "Cá nhân" as License["plan"], months: 12 });
  const [activationCode, setActivationCode] = useState("");
  const [activations, setActivations] = useState<Record<string, string>>({});
  const [report, setReport] = useState({ title: "", area: "Xử lý PDF", severity: "Trung bình", steps: "", expected: "", actual: "" });
  const current = useMemo(() => accounts.find((account) => account.id === sessionId) || null, [accounts, sessionId]);
  const activeLicense = useMemo(() => {
    if (!current) return null;
    if (REMOTE_AUTH) return licenses.find((item) => item.email === current.email && item.status === "Hoạt động" && new Date(item.expiresAt) >= new Date()) || null;
    const code = activations[current.id];
    const license = licenses.find((item) => item.code === code);
    return license && license.status === "Hoạt động" && new Date(license.expiresAt) >= new Date() ? license : null;
  }, [current, activations, licenses]);

  async function hydrateRemoteAccount(user: any, accessToken: string) {
    let account = accountFromRemoteUser(user);
    try {
      const profiles = await remoteData(`/profiles?id=eq.${encodeURIComponent(user.id)}&select=id,full_name,email,role,created_at`, accessToken);
      const profile = Array.isArray(profiles) ? profiles[0] : null;
      if (profile) account = { ...account, name: profile.full_name || account.name, email: profile.email || account.email, role: profile.role === "admin" ? "admin" : "user", createdAt: profile.created_at || account.createdAt };
      const rows = await remoteData(`/licenses?select=id,code,customer,email,plan,expires_at,status,created_at&order=created_at.desc`, accessToken);
      if (Array.isArray(rows)) setLicenses(rows.map((item: any) => ({ id: item.id, code: item.code, customer: item.customer, email: item.email, plan: item.plan, expiresAt: item.expires_at, status: item.status, createdAt: item.created_at })));
    } catch (reason) {
      setMessage(reason instanceof Error ? `${reason.message} Hãy chạy tệp SUPABASE_SCHEMA.sql.` : "Chưa đồng bộ được hồ sơ máy chủ.");
    }
    setRemoteToken(accessToken);
    return account;
  }

  useEffect(() => {
    async function restoreSession() {
      try {
      const storedAccounts = JSON.parse(localStorage.getItem(ACCOUNT_KEY) || "[]") as Account[];
      const remoteSession = JSON.parse(localStorage.getItem(REMOTE_SESSION_KEY) || "null");
      const remoteAccount = remoteSession?.account as Account | undefined;
      setAccounts(remoteAccount ? [...storedAccounts.filter((item) => item.id !== remoteAccount.id), remoteAccount] : storedAccounts);
      setReports(JSON.parse(localStorage.getItem(BUG_KEY) || "[]"));
      setLicenses(JSON.parse(localStorage.getItem(LICENSE_KEY) || "[]"));
      setActivations(JSON.parse(localStorage.getItem(ACTIVATION_KEY) || "{}"));
      setSessionId(REMOTE_AUTH ? "" : remoteAccount?.id || localStorage.getItem(SESSION_KEY) || "");
      setMode(REMOTE_AUTH ? "login" : storedAccounts.some((account) => account.role === "admin") ? "login" : "setup");
      if (REMOTE_AUTH) {
        const hash = new URLSearchParams(location.hash.replace(/^#/, ""));
        const recoveryToken = hash.get("access_token");
        if (hash.get("type") === "recovery" && recoveryToken) {
          localStorage.setItem(REMOTE_SESSION_KEY, JSON.stringify({ accessToken: recoveryToken, refreshToken: hash.get("refresh_token") || "", account: remoteAccount || null }));
          history.replaceState(null, "", location.pathname + location.search);
          setMode("reset");
          setMessage("Hãy đặt mật khẩu mới cho tài khoản.");
          return;
        }
        if (remoteSession?.refreshToken) {
          try {
            const data = await remoteAuth("/token?grant_type=refresh_token", { refresh_token: remoteSession.refreshToken });
            const account = await hydrateRemoteAccount(data.user, data.access_token);
            localStorage.setItem(REMOTE_SESSION_KEY, JSON.stringify({ accessToken: data.access_token, refreshToken: data.refresh_token, account }));
            setAccounts((items) => [...items.filter((item) => item.id !== account.id), account]);
            setSessionId(account.id);
          } catch {
            localStorage.removeItem(REMOTE_SESSION_KEY);
            setMessage("Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.");
          }
        }
      }
      } catch { setMode(REMOTE_AUTH ? "login" : "setup"); }
    }
    void restoreSession();
  }, []);

  function saveAccounts(next: Account[]) { setAccounts(next); localStorage.setItem(ACCOUNT_KEY, JSON.stringify(next)); }

  async function submitAccount(event: FormEvent) {
    event.preventDefault(); setMessage("");
    const normalizedEmail = email.trim().toLocaleLowerCase();
    if (AUTH_CONFIG_ERROR) { setMessage("Cấu hình tài khoản chưa đúng. Project URL phải có dạng https://xxxxx.supabase.co và cần đủ anon key."); return; }
    if (mode === "forgot" && REMOTE_AUTH) {
      if (!normalizedEmail) { setMessage("Hãy nhập email đã đăng ký."); return; }
      try {
        await remoteAuth("/recover", { email: normalizedEmail, redirect_to: `${location.origin}${location.pathname}#tai-khoan` });
        setMessage("Nếu email tồn tại, hướng dẫn đặt lại mật khẩu đã được gửi. Hãy kiểm tra cả thư rác.");
      } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không gửi được email khôi phục."); }
      return;
    }
    if (mode === "reset" && REMOTE_AUTH) {
      if (password.length < 8) { setMessage("Mật khẩu cần ít nhất 8 ký tự."); return; }
      if (password !== confirm) { setMessage("Mật khẩu xác nhận chưa khớp."); return; }
      try {
        const stored = JSON.parse(localStorage.getItem(REMOTE_SESSION_KEY) || "null");
        if (!stored?.accessToken) throw new Error("Liên kết khôi phục đã hết hạn. Hãy yêu cầu lại email đặt mật khẩu.");
        await remoteAuth("/user", { password }, { method: "PUT", accessToken: stored.accessToken });
        localStorage.removeItem(REMOTE_SESSION_KEY); setPassword(""); setConfirm(""); setMode("login"); setMessage("Đã đổi mật khẩu. Bạn có thể đăng nhập lại.");
      } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không đổi được mật khẩu."); }
      return;
    }
    if (!normalizedEmail || !password) { setMessage("Hãy nhập email và mật khẩu."); return; }
    if (REMOTE_AUTH) {
      try {
        if (mode === "login") {
          const data = await remoteAuth("/token?grant_type=password", { email: normalizedEmail, password });
          const account = await hydrateRemoteAccount(data.user, data.access_token);
          localStorage.setItem(REMOTE_SESSION_KEY, JSON.stringify({ accessToken: data.access_token, refreshToken: data.refresh_token, account }));
          setAccounts((current) => [...current.filter((item) => item.id !== account.id), account]); setSessionId(account.id); setPassword(""); setMessage("Đăng nhập tài khoản dùng chung thành công."); return;
        }
        if (password.length < 8) { setMessage("Mật khẩu cần ít nhất 8 ký tự."); return; }
        if (password !== confirm) { setMessage("Mật khẩu xác nhận chưa khớp."); return; }
        const data = await remoteAuth("/signup", { email: normalizedEmail, password, data: { full_name: name.trim() } });
        setPassword(""); setConfirm(""); setMode("login");
        setMessage(data.access_token ? "Đăng ký thành công. Hãy đăng nhập." : "Đã đăng ký. Hãy kiểm tra email xác nhận rồi đăng nhập."); return;
      } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không kết nối được máy chủ tài khoản."); return; }
    }
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

  function logout() { localStorage.removeItem(SESSION_KEY); localStorage.removeItem(REMOTE_SESSION_KEY); setRemoteToken(""); setSessionId(""); setMode("login"); setMessage(""); }

  function saveReport(event: FormEvent) {
    event.preventDefault();
    if (!report.title.trim() || !report.actual.trim()) { setMessage("Hãy nhập tên lỗi và kết quả thực tế."); return; }
    const item: BugReport = { id: crypto.randomUUID(), ...report, createdAt: new Date().toISOString(), status: "Mới" };
    const next = [item, ...reports]; setReports(next); localStorage.setItem(BUG_KEY, JSON.stringify(next)); setReport({ title: "", area: "Xử lý PDF", severity: "Trung bình", steps: "", expected: "", actual: "" }); setMessage("Đã lưu báo cáo lỗi trên thiết bị.");
  }

  function updateReport(id: string, status: BugReport["status"]) { const next = reports.map((item) => item.id === id ? { ...item, status } : item); setReports(next); localStorage.setItem(BUG_KEY, JSON.stringify(next)); }

  async function createLicense(event: FormEvent) {
    event.preventDefault();
    if (!licenseDraft.customer.trim() || !/^\S+@\S+\.\S+$/.test(licenseDraft.email.trim())) { setMessage("Hãy nhập tên khách hàng và email hợp lệ."); return; }
    const expires = new Date();
    expires.setMonth(expires.getMonth() + Math.max(1, Math.min(36, licenseDraft.months)));
    const item: License = { id: crypto.randomUUID(), code: generateLicenseCode(), customer: licenseDraft.customer.trim(), email: licenseDraft.email.trim().toLocaleLowerCase(), plan: licenseDraft.plan, expiresAt: expires.toISOString(), status: "Hoạt động", createdAt: new Date().toISOString() };
    if (REMOTE_AUTH) {
      if (!remoteToken || current?.role !== "admin") { setMessage("Chỉ quản trị viên máy chủ được cấp bản quyền."); return; }
      try {
        const rows = await remoteData("/licenses", remoteToken, { method: "POST", payload: { code: item.code, customer: item.customer, email: item.email, plan: item.plan, expires_at: item.expiresAt, status: item.status } });
        const saved = Array.isArray(rows) ? rows[0] : null;
        const remoteItem = saved ? { ...item, id: saved.id, createdAt: saved.created_at || item.createdAt } : item;
        setLicenses((values) => [remoteItem, ...values]); setLicenseDraft({ customer: "", email: "", plan: "Cá nhân", months: 12 }); setMessage("Đã cấp bản quyền trên máy chủ."); return;
      } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không tạo được bản quyền."); return; }
    }
    const next = [item, ...licenses];
    setLicenses(next); localStorage.setItem(LICENSE_KEY, JSON.stringify(next));
    setLicenseDraft({ customer: "", email: "", plan: "Cá nhân", months: 12 }); setMessage("Đã tạo mã bản quyền kiểm thử.");
  }

  async function toggleLicense(id: string) {
    if (REMOTE_AUTH) {
      const selected = licenses.find((item) => item.id === id);
      if (!selected || !remoteToken || current?.role !== "admin") return;
      const status: License["status"] = selected.status === "Hoạt động" ? "Đã khóa" : "Hoạt động";
      try { await remoteData(`/licenses?id=eq.${encodeURIComponent(id)}`, remoteToken, { method: "PATCH", payload: { status } }); setLicenses((items) => items.map((item) => item.id === id ? { ...item, status } : item)); setMessage("Đã cập nhật trạng thái bản quyền."); } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không cập nhật được bản quyền."); }
      return;
    }
    const next = licenses.map((item) => item.id === id ? { ...item, status: item.status === "Hoạt động" ? "Đã khóa" as const : "Hoạt động" as const } : item);
    setLicenses(next); localStorage.setItem(LICENSE_KEY, JSON.stringify(next));
  }

  function activateLicense(event: FormEvent) {
    event.preventDefault();
    if (!current) return;
    const normalizedCode = activationCode.trim().toUpperCase();
    const license = licenses.find((item) => item.code === normalizedCode);
    if (!license) { setMessage("Không tìm thấy mã bản quyền trên thiết bị này."); return; }
    if (license.email !== current.email) { setMessage("Mã bản quyền được cấp cho email khác."); return; }
    if (license.status !== "Hoạt động") { setMessage("Mã bản quyền đã bị khóa."); return; }
    if (new Date(license.expiresAt) < new Date()) { setMessage("Mã bản quyền đã hết hạn."); return; }
    const next = { ...activations, [current.id]: license.code };
    setActivations(next); localStorage.setItem(ACTIVATION_KEY, JSON.stringify(next)); setActivationCode(""); setMessage("Kích hoạt bản quyền kiểm thử thành công.");
  }

  if (current) return (
    <section className="account-portal" id="tai-khoan" aria-labelledby="account-title">
      <div className="account-header"><div><p className="section-kicker">Trung tâm tài khoản</p><h2 id="account-title">Xin chào, {current.name}</h2><p>{current.role === "admin" ? `Quản trị viên SỸ LAND · ${REMOTE_AUTH ? "Tài khoản máy chủ" : "Chế độ cục bộ"}` : `Tài khoản SỸ LAND · ${REMOTE_AUTH ? "Đã đồng bộ" : "Chế độ cục bộ"}`}</p></div><button type="button" onClick={logout}>Đăng xuất</button></div>
      {current.role === "admin" ? <div className="admin-grid">
        <aside className="admin-menu"><span>QUẢN TRỊ KIỂM THỬ</span><a href="#minh-hoa">Xử lý Word/PDF/Excel</a><a href="#cong-cu-pdf">Bộ công cụ PDF</a><a href="#chuc-nang-phan-mem">Kiểm tra chức năng</a><a href="#tai-phan-mem">Kiểm tra bản phát hành</a><button type="button" onClick={() => downloadJson(reports, `SYLAND_BAO_CAO_LOI_${new Date().toISOString().slice(0,10)}.json`)} disabled={!reports.length}>Xuất báo cáo lỗi JSON</button><button type="button" onClick={() => downloadJson(licenses, `SYLAND_BAN_QUYEN_${new Date().toISOString().slice(0,10)}.json`)} disabled={!licenses.length}>Xuất bản quyền JSON</button><small>Dữ liệu tài khoản, bản quyền và lỗi chỉ tồn tại trên trình duyệt này.</small></aside>
        <div className="bug-workspace">
          <div className="admin-stats"><div><strong>{accounts.length}</strong><span>Tài khoản cục bộ</span></div><div><strong>{licenses.filter((item) => item.status === "Hoạt động" && new Date(item.expiresAt) >= new Date()).length}</strong><span>Bản quyền hoạt động</span></div><div><strong>{reports.length}</strong><span>Lỗi đã ghi nhận</span></div><div><strong>{reports.filter((item) => item.status === "Mới").length}</strong><span>Lỗi mới</span></div></div>
          <section className="license-workspace"><div className="license-heading"><div><h3>Quản lý bản quyền kiểm thử</h3><p>Tạo và thử quy trình cấp mã trước khi kết nối máy chủ thương mại.</p></div><span>CỤC BỘ</span></div><form className="license-form" onSubmit={createLicense}><label>Khách hàng<input value={licenseDraft.customer} onChange={(event) => setLicenseDraft((value) => ({ ...value, customer: event.target.value }))} placeholder="Họ tên hoặc đơn vị" /></label><label>Email<input type="email" value={licenseDraft.email} onChange={(event) => setLicenseDraft((value) => ({ ...value, email: event.target.value }))} placeholder="khachhang@example.com" /></label><label>Gói<select value={licenseDraft.plan} onChange={(event) => setLicenseDraft((value) => ({ ...value, plan: event.target.value as License["plan"] }))}><option>Cá nhân</option><option>Văn phòng</option><option>Đơn vị</option></select></label><label>Số tháng<input type="number" min={1} max={36} value={licenseDraft.months} onChange={(event) => setLicenseDraft((value) => ({ ...value, months: Number(event.target.value) || 1 }))} /></label><button type="submit">Tạo mã bản quyền</button></form><div className="license-list">{!licenses.length ? <p>Chưa có mã bản quyền kiểm thử.</p> : licenses.slice(0, 20).map((item) => <article key={item.id}><div><strong>{item.customer}</strong><code>{item.code}</code><small>{item.email} · Gói {item.plan} · Hết hạn {new Date(item.expiresAt).toLocaleDateString("vi-VN")}</small></div><button type="button" className={item.status === "Hoạt động" ? "active" : "locked"} onClick={() => toggleLicense(item.id)}>{item.status}</button></article>)}</div><p className="license-warning">Mã tạo ở đây chỉ dùng kiểm thử giao diện, chưa đủ an toàn để bán. Bản thương mại phải xác minh mã trên máy chủ.</p></section>
          <form className="bug-form" onSubmit={saveReport}><h3>Ghi nhận lỗi kiểm thử</h3><label>Tên lỗi<input value={report.title} onChange={(event) => setReport((value) => ({ ...value, title: event.target.value }))} placeholder="Mô tả ngắn gọn" /></label><div><label>Khu vực<select value={report.area} onChange={(event) => setReport((value) => ({ ...value, area: event.target.value }))}><option>Xử lý PDF</option><option>Word/PDF/Excel</option><option>Địa bàn và mã xã</option><option>Đổi tên hàng loạt</option><option>Đối chiếu Excel</option><option>Đăng nhập</option><option>Giao diện</option><option>Khác</option></select></label><label>Mức độ<select value={report.severity} onChange={(event) => setReport((value) => ({ ...value, severity: event.target.value }))}><option>Thấp</option><option>Trung bình</option><option>Cao</option><option>Nghiêm trọng</option></select></label></div><label>Các bước tái hiện<textarea value={report.steps} onChange={(event) => setReport((value) => ({ ...value, steps: event.target.value }))} placeholder="1. Mở… 2. Chọn… 3. Nhấn…" /></label><label>Kết quả mong muốn<textarea value={report.expected} onChange={(event) => setReport((value) => ({ ...value, expected: event.target.value }))} /></label><label>Kết quả thực tế<textarea value={report.actual} onChange={(event) => setReport((value) => ({ ...value, actual: event.target.value }))} /></label><button type="submit">Lưu báo cáo lỗi</button></form>
          <div className="bug-list"><h3>Danh sách lỗi gần đây</h3>{!reports.length ? <p>Chưa có lỗi được ghi nhận.</p> : reports.slice(0, 20).map((item) => <article key={item.id}><div><span className={`severity ${item.severity.toLocaleLowerCase("vi-VN")}`}>{item.severity}</span><strong>{item.title}</strong><small>{item.area} · {new Date(item.createdAt).toLocaleString("vi-VN")}</small></div><select value={item.status} onChange={(event) => updateReport(item.id, event.target.value as BugReport["status"])}><option>Mới</option><option>Đang kiểm tra</option><option>Đã xử lý</option></select></article>)}</div>
        </div>
      </div> : <div className="user-test-panel"><h3>{activeLicense ? `Bản quyền ${activeLicense.plan} đang hoạt động` : "Tài khoản SỸ LAND đã sẵn sàng"}</h3>{activeLicense ? <div className="activated-license"><span>ĐÃ KÍCH HOẠT</span><strong>{activeLicense.code}</strong><p>Cấp cho {activeLicense.email} · Hết hạn {new Date(activeLicense.expiresAt).toLocaleDateString("vi-VN")}</p></div> : <><p>Bạn có thể sử dụng các công cụ công khai hoặc nhập mã bản quyền do quản trị viên cấp.</p><form className="activation-form" onSubmit={activateLicense}><label>Mã bản quyền<input value={activationCode} onChange={(event) => setActivationCode(event.target.value.toUpperCase())} placeholder="SYL-XXXXX-XXXXX" /></label><button type="submit">Kích hoạt mã</button></form></>}<a className="button button-primary" href="#minh-hoa">Mở công cụ</a><small>{REMOTE_AUTH ? "Tài khoản được đồng bộ giữa website và phần mềm SỸ LAND." : "Chế độ cục bộ chưa đồng bộ dữ liệu giữa các thiết bị."}</small></div>}
      {message && <p className="account-message">{message}</p>}
    </section>
  );

  return (
    <section className="account-portal account-auth" id="tai-khoan" aria-labelledby="account-title">
      <div className="auth-copy"><p className="section-kicker">{REMOTE_AUTH ? "Tài khoản SỸ LAND" : AUTH_CONFIG_ERROR ? "Cần hoàn tất cấu hình" : "Chế độ tài khoản cục bộ"}</p><h2 id="account-title">{mode === "setup" ? "Thiết lập quản trị viên SỸ LAND" : mode === "register" ? "Đăng ký tài khoản SỸ LAND" : mode === "forgot" ? "Khôi phục mật khẩu" : mode === "reset" ? "Đặt mật khẩu mới" : "Đăng nhập SỸ LAND"}</h2><p>{AUTH_CONFIG_ERROR ? "Website chưa nhận được cấu hình máy chủ tài khoản. Đăng ký và đăng nhập tạm khóa để tránh tạo tài khoản không đồng bộ." : REMOTE_AUTH ? "Một tài khoản đăng nhập được trên cả website và phần mềm SỸ LAND." : mode === "setup" ? "Thiết lập mật khẩu quản trị lần đầu cho anh Nguyễn Minh Sỹ trên thiết bị này." : "Chế độ phát triển cục bộ đang hoạt động."}</p><ul><li>✓ Mật khẩu không được ghi trong mã website</li><li>✓ {REMOTE_AUTH ? "Tự khôi phục phiên đăng nhập qua HTTPS" : AUTH_CONFIG_ERROR ? "Cần thêm hai GitHub Actions secrets" : "Chỉ lưu mã băm trên trình duyệt"}</li><li>✓ Có quy trình quên và đặt lại mật khẩu</li></ul></div>
      <form className="auth-form" onSubmit={submitAccount}>{(mode === "register" || mode === "setup") && <label>Họ và tên<input value={name} onChange={(event) => setName(event.target.value)} autoComplete="name" /></label>}{mode !== "reset" && <label>Email<input type="email" value={email} onChange={(event) => setEmail(event.target.value)} autoComplete="email" placeholder="Nhập email đã đăng ký" /></label>}{mode !== "forgot" && <label>Mật khẩu<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete={mode === "login" ? "current-password" : "new-password"} placeholder="Tối thiểu 8 ký tự" /></label>}{mode !== "login" && mode !== "forgot" && <label>Xác nhận mật khẩu<input type="password" value={confirm} onChange={(event) => setConfirm(event.target.value)} autoComplete="new-password" /></label>}<button type="submit" disabled={AUTH_CONFIG_ERROR}>{mode === "setup" ? "Tạo tài khoản admin" : mode === "register" ? "Tạo tài khoản" : mode === "forgot" ? "Gửi email khôi phục" : mode === "reset" ? "Lưu mật khẩu mới" : "Đăng nhập"}</button>{message && <p className="account-message">{message}</p>}{mode === "login" && REMOTE_AUTH && <button className="auth-secondary" type="button" onClick={() => { setMode("forgot"); setMessage(""); }}>Quên mật khẩu?</button>}{mode !== "setup" && mode !== "reset" && <p className="auth-switch">{mode === "login" ? <>Chưa có tài khoản? <button type="button" onClick={() => { setMode("register"); setMessage(""); }}>Đăng ký tài khoản</button></> : <>Đã có tài khoản? <button type="button" onClick={() => { setMode("login"); setMessage(""); }}>Đăng nhập</button></>}</p>}</form>
    </section>
  );
}
