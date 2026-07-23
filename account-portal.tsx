"use client";

import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import { TERMS_VERSION } from "./legal-center";

type Account = { id: string; name: string; email: string; passwordHash: string; salt: string; role: "admin" | "user"; createdAt: string };
type BugReport = { id: string; title: string; area: string; severity: string; steps: string; expected: string; actual: string; createdAt: string; status: "Mới" | "Đang kiểm tra" | "Đã xử lý" };
type License = { id: string; code: string; customer: string; email: string; plan: "Cá nhân" | "Văn phòng" | "Đơn vị"; expiresAt: string; status: "Hoạt động" | "Đã khóa"; createdAt: string; maxDevices?: number };
type DeviceActivation = { id: string; licenseId: string; deviceHash: string; deviceName: string; appVersion: string; status: "Hoạt động" | "Đã hủy"; firstSeenAt: string; lastSeenAt: string };
type AuditEvent = { id: string; action: string; entityType: string; entityId: string; details: Record<string, unknown>; createdAt: string };
type Lead = { id: string; unit: string; contact: string; monthlyVolume: number; people: number; needs: string[]; proposedPlan: string; note: string; status: "Mới" | "Đang liên hệ" | "Đã chuyển đổi" | "Đã đóng"; createdAt: string };
type SupportTicket = { id: string; ticketCode: string; title: string; category: string; priority: string; details: string; status: "Mới" | "Đang xử lý" | "Chờ người dùng" | "Đã giải quyết" | "Đã đóng"; appVersion: string; createdAt: string };
type PaymentNotice = { id: string; orderCode: string; plan: string; status: string; licenseCode: string; createdAt: string; confirmedAt: string };

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

function authRedirectUrl() {
  const basePath = String(import.meta.env.BASE_URL || "/SY-LAND/");
  return `${new URL(basePath, location.origin).toString()}#tai-khoan`;
}

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
  const [acceptedTerms, setAcceptedTerms] = useState(false);
  const [message, setMessage] = useState("");
  const [reports, setReports] = useState<BugReport[]>([]);
  const [licenses, setLicenses] = useState<License[]>([]);
  const [devices, setDevices] = useState<DeviceActivation[]>([]);
  const [auditEvents, setAuditEvents] = useState<AuditEvent[]>([]);
  const [leads, setLeads] = useState<Lead[]>([]);
  const [tickets, setTickets] = useState<SupportTicket[]>([]);
  const [paymentNotices, setPaymentNotices] = useState<PaymentNotice[]>([]);
  const [licenseDraft, setLicenseDraft] = useState({ customer: "", email: "", plan: "Cá nhân" as License["plan"], months: 12, maxDevices: 1 });
  const [activationCode, setActivationCode] = useState("");
  const configInputRef = useRef<HTMLInputElement>(null);
  const [configMessage, setConfigMessage] = useState("");
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
      const rows = await remoteData(`/licenses?select=id,code,customer,email,plan,expires_at,status,created_at,max_devices&order=created_at.desc`, accessToken);
      if (Array.isArray(rows)) setLicenses(rows.map((item: any) => ({ id: item.id, code: item.code, customer: item.customer, email: item.email, plan: item.plan, expiresAt: item.expires_at, status: item.status, createdAt: item.created_at, maxDevices: item.max_devices })));
      const deviceRows = await remoteData(`/device_activations?select=id,license_id,device_hash,device_name,app_version,status,first_seen_at,last_seen_at&order=last_seen_at.desc`, accessToken);
      if (Array.isArray(deviceRows)) setDevices(deviceRows.map((item: any) => ({ id: item.id, licenseId: item.license_id, deviceHash: item.device_hash, deviceName: item.device_name, appVersion: item.app_version, status: item.status, firstSeenAt: item.first_seen_at, lastSeenAt: item.last_seen_at })));
      const paymentRows = await remoteData(`/payment_orders?select=id,order_code,plan,status,license_code,created_at,confirmed_at&order=created_at.desc&limit=500`, accessToken);
      if (Array.isArray(paymentRows)) setPaymentNotices(paymentRows.map((item: any) => ({ id: item.id, orderCode: item.order_code, plan: item.plan, status: item.status, licenseCode: item.license_code || "", createdAt: item.created_at, confirmedAt: item.confirmed_at || "" })));
      if (profile?.role === "admin") {
        const eventRows = await remoteData(`/audit_events?select=id,action,entity_type,entity_id,details,created_at&order=created_at.desc&limit=500`, accessToken);
        if (Array.isArray(eventRows)) setAuditEvents(eventRows.map((item: any) => ({ id: item.id, action: item.action, entityType: item.entity_type, entityId: item.entity_id, details: item.details || {}, createdAt: item.created_at })));
        const leadRows = await remoteData(`/leads?select=id,unit,contact,monthly_volume,people,needs,proposed_plan,note,status,created_at&order=created_at.desc&limit=500`, accessToken);
        if (Array.isArray(leadRows)) setLeads(leadRows.map((item: any) => ({ id: item.id, unit: item.unit, contact: item.contact, monthlyVolume: item.monthly_volume, people: item.people, needs: item.needs || [], proposedPlan: item.proposed_plan, note: item.note || "", status: item.status, createdAt: item.created_at })));
        const ticketRows = await remoteData(`/support_tickets?select=id,ticket_code,title,category,priority,details,status,app_version,created_at&order=created_at.desc&limit=500`, accessToken);
        if (Array.isArray(ticketRows)) setTickets(ticketRows.map((item: any) => ({ id: item.id, ticketCode: item.ticket_code, title: item.title, category: item.category, priority: item.priority, details: item.details, status: item.status, appVersion: item.app_version, createdAt: item.created_at })));
      }
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

  useEffect(() => {
    if (!REMOTE_AUTH || !remoteToken || !current) return;
    async function refreshEntitlements() {
      try {
        const [licenseRows, paymentRows] = await Promise.all([
          remoteData(`/licenses?select=id,code,customer,email,plan,expires_at,status,created_at,max_devices&order=created_at.desc`, remoteToken),
          remoteData(`/payment_orders?select=id,order_code,plan,status,license_code,created_at,confirmed_at&order=created_at.desc&limit=500`, remoteToken),
        ]);
        if (Array.isArray(licenseRows)) setLicenses(licenseRows.map((item: any) => ({ id: item.id, code: item.code, customer: item.customer, email: item.email, plan: item.plan, expiresAt: item.expires_at, status: item.status, createdAt: item.created_at, maxDevices: item.max_devices })));
        if (Array.isArray(paymentRows)) setPaymentNotices(paymentRows.map((item: any) => ({ id: item.id, orderCode: item.order_code, plan: item.plan, status: item.status, licenseCode: item.license_code || "", createdAt: item.created_at, confirmedAt: item.confirmed_at || "" })));
      } catch { /* Giữ dữ liệu hiện tại nếu mạng tạm gián đoạn. */ }
    }
    const onPaymentUpdated = () => void refreshEntitlements();
    window.addEventListener("syland-payment-updated", onPaymentUpdated);
    const timer = window.setInterval(refreshEntitlements, 30000);
    void refreshEntitlements();
    return () => { window.removeEventListener("syland-payment-updated", onPaymentUpdated); window.clearInterval(timer); };
  }, [current, remoteToken]);

  function saveAccounts(next: Account[]) { setAccounts(next); localStorage.setItem(ACCOUNT_KEY, JSON.stringify(next)); }

  async function submitAccount(event: FormEvent) {
    event.preventDefault(); setMessage("");
    const normalizedEmail = email.trim().toLocaleLowerCase();
    if (AUTH_CONFIG_ERROR) { setMessage("Cấu hình tài khoản chưa đúng. Project URL phải có dạng https://xxxxx.supabase.co và cần đủ anon key."); return; }
    if (mode === "forgot" && REMOTE_AUTH) {
      if (!normalizedEmail) { setMessage("Hãy nhập email đã đăng ký."); return; }
      try {
        await remoteAuth("/recover", { email: normalizedEmail, redirect_to: authRedirectUrl() });
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
        if (!acceptedTerms) { setMessage("Hãy đồng ý Chính sách bảo mật và Điều khoản sử dụng."); return; }
        const redirectTo = encodeURIComponent(authRedirectUrl());
        const data = await remoteAuth(`/signup?redirect_to=${redirectTo}`, { email: normalizedEmail, password, data: { full_name: name.trim(), terms_version: TERMS_VERSION, terms_accepted_at: new Date().toISOString() } });
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
    if (!acceptedTerms) { setMessage("Hãy đồng ý Chính sách bảo mật và Điều khoản sử dụng."); return; }
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
    const item: License = { id: crypto.randomUUID(), code: generateLicenseCode(), customer: licenseDraft.customer.trim(), email: licenseDraft.email.trim().toLocaleLowerCase(), plan: licenseDraft.plan, expiresAt: expires.toISOString(), status: "Hoạt động", createdAt: new Date().toISOString(), maxDevices: Math.max(1, Math.min(100, licenseDraft.maxDevices)) };
    if (REMOTE_AUTH) {
      if (!remoteToken || current?.role !== "admin") { setMessage("Chỉ quản trị viên máy chủ được cấp bản quyền."); return; }
      try {
        const rows = await remoteData("/licenses", remoteToken, { method: "POST", payload: { code: item.code, customer: item.customer, email: item.email, plan: item.plan, expires_at: item.expiresAt, status: item.status, max_devices: item.maxDevices } });
        const saved = Array.isArray(rows) ? rows[0] : null;
        const remoteItem = saved ? { ...item, id: saved.id, createdAt: saved.created_at || item.createdAt } : item;
        setLicenses((values) => [remoteItem, ...values]); setLicenseDraft({ customer: "", email: "", plan: "Cá nhân", months: 12, maxDevices: 1 }); setMessage("Đã cấp bản quyền trên máy chủ."); return;
      } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không tạo được bản quyền."); return; }
    }
    const next = [item, ...licenses];
    setLicenses(next); localStorage.setItem(LICENSE_KEY, JSON.stringify(next));
    setLicenseDraft({ customer: "", email: "", plan: "Cá nhân", months: 12, maxDevices: 1 }); setMessage("Đã tạo mã bản quyền kiểm thử.");
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

  async function toggleDevice(id: string) {
    const selected = devices.find((item) => item.id === id);
    if (!selected || !remoteToken || current?.role !== "admin") return;
    const status: DeviceActivation["status"] = selected.status === "Hoạt động" ? "Đã hủy" : "Hoạt động";
    try {
      await remoteData(`/device_activations?id=eq.${encodeURIComponent(id)}`, remoteToken, { method: "PATCH", payload: { status } });
      setDevices((items) => items.map((item) => item.id === id ? { ...item, status } : item));
      setMessage(status === "Đã hủy" ? "Đã hủy quyền sử dụng của thiết bị." : "Đã mở lại thiết bị.");
    } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không cập nhật được thiết bị."); }
  }

  async function updateLead(id: string, status: Lead["status"]) {
    if (!remoteToken || current?.role !== "admin") return;
    try {
      await remoteData(`/leads?id=eq.${encodeURIComponent(id)}`, remoteToken, { method: "PATCH", payload: { status } });
      setLeads((items) => items.map((item) => item.id === id ? { ...item, status } : item));
      setMessage("Đã cập nhật trạng thái khách hàng.");
    } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không cập nhật được yêu cầu."); }
  }

  async function updateTicket(id: string, status: SupportTicket["status"]) {
    if (!remoteToken || current?.role !== "admin") return;
    try {
      await remoteData(`/support_tickets?id=eq.${encodeURIComponent(id)}`, remoteToken, { method: "PATCH", payload: { status } });
      setTickets((items) => items.map((item) => item.id === id ? { ...item, status } : item));
      setMessage("Đã cập nhật trạng thái hỗ trợ.");
    } catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không cập nhật được yêu cầu hỗ trợ."); }
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

  async function copyLicense(code: string) {
    try {
      if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(code);
      else {
        const area = document.createElement("textarea"); area.value = code; area.style.position = "fixed"; area.style.opacity = "0";
        document.body.appendChild(area); area.select();
        if (!document.execCommand("copy")) throw new Error("copy blocked");
        area.remove();
      }
      setMessage("Đã sao chép mã bản quyền vào bộ nhớ tạm.");
    } catch { setMessage("Trình duyệt không cho phép sao chép tự động. Hãy chọn mã và sao chép thủ công."); }
  }

  async function activateFromNotice(code: string) {
    if (!current) return;
    const license = licenses.find((item) => item.code === code && item.email === current.email);
    if (!license || license.status !== "Hoạt động" || new Date(license.expiresAt) < new Date()) { setMessage("Không thể kích hoạt: mã không hợp lệ, đã khóa hoặc hết hạn."); return; }
    const next = { ...activations, [current.id]: code };
    setActivations(next); localStorage.setItem(ACTIVATION_KEY, JSON.stringify(next)); setActivationCode("");
    setMessage("Đã kích hoạt bản quyền thành công trên tài khoản SỸ LAND.");
  }

  function exportConfiguration() {
    let locationProfile: unknown = null;
    let desktopSettings: unknown = null;
    try { locationProfile = JSON.parse(localStorage.getItem("sy-land-location-profile") || "null"); } catch { /* Bỏ qua cấu hình cũ bị lỗi. */ }
    try { desktopSettings = JSON.parse(localStorage.getItem("sy-land-desktop-settings") || "null"); } catch { /* Bỏ qua cấu hình cũ bị lỗi. */ }
    const configuration = {
      format: "syland-config",
      schema_version: 1,
      product: "Tiện ích hỗ trợ làm sạch CSDL đất đai SỸ LAND",
      source: "website",
      exported_at: new Date().toISOString(),
      user: current ? { name: current.name, email: current.email } : null,
      settings: { location_profile: locationProfile, app_settings: desktopSettings },
      security: { contains_password: false, contains_session_token: false, contains_license_code: false },
    };
    downloadJson(configuration, `SYLAND_CAU_HINH_${new Date().toISOString().slice(0, 10)}.json`);
    setConfigMessage("Đã xuất cấu hình tương thích. Tệp không chứa mật khẩu, phiên đăng nhập hoặc mã bản quyền.");
  }

  async function importConfiguration(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]; event.target.value = "";
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) { setConfigMessage("Tệp cấu hình vượt quá 2 MB."); return; }
    try {
      const data = JSON.parse(await file.text());
      if (!data || typeof data !== "object") throw new Error("Tệp không phải cấu hình hợp lệ.");
      const format = String(data.format || data.product || "").toLocaleLowerCase("vi-VN");
      if (!format.includes("syland") && !format.includes("sỹ land")) throw new Error("Tệp không thuộc định dạng cấu hình SỸ LAND.");
      const locationProfile = data.settings?.location_profile ?? data.settings?.locationProfile ?? data.location_profile ?? data.locationProfile;
      if (locationProfile && typeof locationProfile === "object") localStorage.setItem("sy-land-location-profile", JSON.stringify(locationProfile));
      const desktopSettings = data.settings?.app_settings ?? data.settings?.desktop_settings;
      if (desktopSettings && typeof desktopSettings === "object") localStorage.setItem("sy-land-desktop-settings", JSON.stringify(desktopSettings));
      window.dispatchEvent(new CustomEvent("syland-config-imported", { detail: { source: data.source || "software" } }));
      setConfigMessage(locationProfile || desktopSettings ? "Đã nhập cấu hình dùng chung. Hãy mở lại công cụ để áp dụng." : "Tệp hợp lệ nhưng chưa có cấu hình để nhập.");
    } catch (reason) { setConfigMessage(reason instanceof Error ? reason.message : "Không đọc được tệp cấu hình."); }
  }

  if (current) return (
    <section className="account-portal" id="tai-khoan" aria-labelledby="account-title">
      <div className="account-header"><div><p className="section-kicker">Trung tâm tài khoản</p><h2 id="account-title">Xin chào, {current.name}</h2><p>{current.role === "admin" ? `Quản trị viên SỸ LAND · ${REMOTE_AUTH ? "Tài khoản máy chủ" : "Chế độ cục bộ"}` : `Tài khoản SỸ LAND · ${REMOTE_AUTH ? "Đã đồng bộ" : "Chế độ cục bộ"}`}</p>{current.role === "admin" && <div className="admin-full-access"><b>FULL ACCESS</b><span>Không giới hạn gói dịch vụ · toàn quyền kiểm thử và quản trị</span></div>}</div><button type="button" onClick={logout}>Đăng xuất</button></div>
      {current.role === "admin" ? <div className="admin-grid">
        <aside className="admin-menu"><span>QUẢN TRỊ · FULL ACCESS</span><a href="#minh-hoa">Xử lý Word/PDF/Excel</a><a href="#cong-cu-pdf">Bộ công cụ PDF</a><a href="#chuc-nang-phan-mem">Kiểm tra chức năng</a><a href="#tai-phan-mem">Kiểm tra bản phát hành</a><button type="button" onClick={() => downloadJson(reports, `SYLAND_BAO_CAO_LOI_${new Date().toISOString().slice(0,10)}.json`)} disabled={!reports.length}>Xuất báo cáo lỗi JSON</button><button type="button" onClick={() => downloadJson(licenses, `SYLAND_BAN_QUYEN_${new Date().toISOString().slice(0,10)}.json`)} disabled={!licenses.length}>Xuất bản quyền JSON</button><small>Mở toàn bộ công cụ, đối soát, bản quyền, thiết bị, hỗ trợ và nhật ký. Giới hạn kỹ thuật của trình duyệt vẫn được giữ để tránh treo máy.</small></aside>
        <div className="bug-workspace">
          {REMOTE_AUTH && <section className="license-workspace"><div className="license-heading"><div><h3>Thông báo thanh toán</h3><p>Đơn người dùng đã báo chuyển khoản và đang chờ quản trị viên đối soát.</p></div><span>{paymentNotices.filter((item) => item.status === "Chờ xác nhận").length} CHỜ XÁC NHẬN</span></div><div className="license-list">{!paymentNotices.length ? <p>Chưa có đơn thanh toán.</p> : paymentNotices.map((item) => <article key={item.id}><div><strong>{item.orderCode} · Gói {item.plan}</strong>{item.licenseCode && <code>{item.licenseCode}</code>}<small>{item.status === "Đã thanh toán" ? "Đã xác nhận" : item.status} · {new Date(item.createdAt).toLocaleString("vi-VN")}</small></div><span className={item.status === "Đã thanh toán" ? "active" : "locked"}>{item.status === "Đã thanh toán" ? "Đã xác nhận" : item.status}</span></article>)}</div><a className="button button-primary" href="#thanh-toan">Mở đối soát và xác nhận tiền vào</a></section>}
          {REMOTE_AUTH && <section className="license-workspace"><div className="license-heading"><div><h3>Yêu cầu hỗ trợ kỹ thuật</h3><p>Tiếp nhận và theo dõi vấn đề của người dùng đã đăng nhập.</p></div><span>{tickets.filter((item) => item.status === "Mới").length} MỚI</span></div><div className="license-list">{!tickets.length ? <p>Chưa có yêu cầu hỗ trợ.</p> : tickets.map((item) => <article key={item.id}><div><strong>{item.ticketCode} · {item.title}</strong><code>{item.priority}</code><small>{item.category} · {item.appVersion || "—"} · {new Date(item.createdAt).toLocaleString("vi-VN")}</small></div><select value={item.status} onChange={(event) => void updateTicket(item.id, event.target.value as SupportTicket["status"])}><option>Mới</option><option>Đang xử lý</option><option>Chờ người dùng</option><option>Đã giải quyết</option><option>Đã đóng</option></select></article>)}</div></section>}
          {REMOTE_AUTH && <section className="license-workspace"><div className="license-heading"><div><h3>Yêu cầu tư vấn và mua bản quyền</h3><p>Theo dõi khách hàng từ lúc gửi nhu cầu đến khi hoàn tất.</p></div><span>{leads.filter((item) => item.status === "Mới").length} MỚI</span></div><div className="license-list">{!leads.length ? <p>Chưa có yêu cầu tư vấn.</p> : leads.map((item) => <article key={item.id}><div><strong>{item.unit}</strong><code>{item.proposedPlan}</code><small>{item.contact} · {item.monthlyVolume.toLocaleString("vi-VN")} hồ sơ/tháng · {item.people} người · {new Date(item.createdAt).toLocaleString("vi-VN")}</small></div><select value={item.status} onChange={(event) => void updateLead(item.id, event.target.value as Lead["status"])}><option>Mới</option><option>Đang liên hệ</option><option>Đã chuyển đổi</option><option>Đã đóng</option></select></article>)}</div></section>}
          {REMOTE_AUTH && <section className="license-workspace"><div className="license-heading"><div><h3>Thiết bị đã kích hoạt</h3><p>Hủy thiết bị mất quyền hoặc mở lại thiết bị hợp lệ.</p></div><span>{devices.filter((item) => item.status === "Hoạt động").length} MÁY</span></div><div className="license-list">{!devices.length ? <p>Chưa có thiết bị đăng nhập phần mềm.</p> : devices.map((item) => { const license = licenses.find((value) => value.id === item.licenseId); return <article key={item.id}><div><strong>{item.deviceName || "Windows PC"}</strong><code>{item.deviceHash.slice(0, 8)}••••</code><small>{license?.email || license?.code || "Chưa xác định bản quyền"} · Bản {item.appVersion || "—"} · Gần nhất {new Date(item.lastSeenAt).toLocaleString("vi-VN")}</small></div><button type="button" className={item.status === "Hoạt động" ? "active" : "locked"} onClick={() => toggleDevice(item.id)}>{item.status}</button></article>; })}</div></section>}
          {REMOTE_AUTH && <section className="license-workspace"><div className="license-heading"><div><h3>Nhật ký quản trị</h3><p>50 sự kiện mới nhất liên quan tới bản quyền và thiết bị.</p></div><span>KIỂM SOÁT</span></div><div className="license-list">{!auditEvents.length ? <p>Chưa có sự kiện quản trị.</p> : auditEvents.map((item) => <article key={item.id}><div><strong>{item.action}</strong><small>{item.entityType} · {new Date(item.createdAt).toLocaleString("vi-VN")}</small></div></article>)}</div></section>}
          <div className="admin-stats"><div><strong>{accounts.length}</strong><span>Tài khoản đã tải</span></div><div><strong>{licenses.filter((item) => item.status === "Hoạt động" && new Date(item.expiresAt) >= new Date()).length}</strong><span>Bản quyền hoạt động</span></div><div><strong>{devices.filter((item) => item.status === "Hoạt động").length}</strong><span>Thiết bị hoạt động</span></div><div><strong>{REMOTE_AUTH ? leads.filter((item) => item.status === "Mới").length : reports.filter((item) => item.status === "Mới").length}</strong><span>{REMOTE_AUTH ? "Yêu cầu mới" : "Lỗi mới"}</span></div></div>
          <section className="license-workspace"><div className="license-heading"><div><h3>Quản lý bản quyền {REMOTE_AUTH ? "máy chủ" : "kiểm thử"}</h3><p>{REMOTE_AUTH ? "Cấp bản quyền theo email và giới hạn thiết bị sử dụng." : "Tạo và thử quy trình cấp mã trên thiết bị."}</p></div><span>{REMOTE_AUTH ? "ĐỒNG BỘ" : "CỤC BỘ"}</span></div><form className="license-form" onSubmit={createLicense}><label>Khách hàng<input value={licenseDraft.customer} onChange={(event) => setLicenseDraft((value) => ({ ...value, customer: event.target.value }))} placeholder="Họ tên hoặc đơn vị" /></label><label>Email<input type="email" value={licenseDraft.email} onChange={(event) => setLicenseDraft((value) => ({ ...value, email: event.target.value }))} placeholder="khachhang@example.com" /></label><label>Gói<select value={licenseDraft.plan} onChange={(event) => setLicenseDraft((value) => ({ ...value, plan: event.target.value as License["plan"] }))}><option>Cá nhân</option><option>Văn phòng</option><option>Đơn vị</option></select></label><label>Số tháng<input type="number" min={1} max={36} value={licenseDraft.months} onChange={(event) => setLicenseDraft((value) => ({ ...value, months: Number(event.target.value) || 1 }))} /></label><label>Số thiết bị<input type="number" min={1} max={100} value={licenseDraft.maxDevices} onChange={(event) => setLicenseDraft((value) => ({ ...value, maxDevices: Number(event.target.value) || 1 }))} /></label><button type="submit">Tạo mã bản quyền</button></form><div className="license-list">{!licenses.length ? <p>Chưa có mã bản quyền.</p> : licenses.map((item) => <article key={item.id}><div><strong>{item.customer}</strong><code>{item.code}</code><small>{item.email} · Gói {item.plan} · {item.maxDevices || 1} thiết bị · Hết hạn {new Date(item.expiresAt).toLocaleDateString("vi-VN")}</small></div><button type="button" className={item.status === "Hoạt động" ? "active" : "locked"} onClick={() => toggleLicense(item.id)}>{item.status}</button></article>)}</div>{!REMOTE_AUTH && <p className="license-warning">Mã cục bộ chỉ dùng kiểm tra giao diện; bản thương mại phải xác minh trên máy chủ.</p>}</section>
          <form className="bug-form" onSubmit={saveReport}><h3>Ghi nhận lỗi kiểm thử</h3><label>Tên lỗi<input value={report.title} onChange={(event) => setReport((value) => ({ ...value, title: event.target.value }))} placeholder="Mô tả ngắn gọn" /></label><div><label>Khu vực<select value={report.area} onChange={(event) => setReport((value) => ({ ...value, area: event.target.value }))}><option>Xử lý PDF</option><option>Word/PDF/Excel</option><option>Địa bàn và mã xã</option><option>Đổi tên hàng loạt</option><option>Đối chiếu Excel</option><option>Đăng nhập</option><option>Giao diện</option><option>Khác</option></select></label><label>Mức độ<select value={report.severity} onChange={(event) => setReport((value) => ({ ...value, severity: event.target.value }))}><option>Thấp</option><option>Trung bình</option><option>Cao</option><option>Nghiêm trọng</option></select></label></div><label>Các bước tái hiện<textarea value={report.steps} onChange={(event) => setReport((value) => ({ ...value, steps: event.target.value }))} placeholder="1. Mở… 2. Chọn… 3. Nhấn…" /></label><label>Kết quả mong muốn<textarea value={report.expected} onChange={(event) => setReport((value) => ({ ...value, expected: event.target.value }))} /></label><label>Kết quả thực tế<textarea value={report.actual} onChange={(event) => setReport((value) => ({ ...value, actual: event.target.value }))} /></label><button type="submit">Lưu báo cáo lỗi</button></form>
          <div className="bug-list"><h3>Danh sách lỗi gần đây</h3>{!reports.length ? <p>Chưa có lỗi được ghi nhận.</p> : reports.slice(0, 20).map((item) => <article key={item.id}><div><span className={`severity ${item.severity.toLocaleLowerCase("vi-VN")}`}>{item.severity}</span><strong>{item.title}</strong><small>{item.area} · {new Date(item.createdAt).toLocaleString("vi-VN")}</small></div><select value={item.status} onChange={(event) => updateReport(item.id, event.target.value as BugReport["status"])}><option>Mới</option><option>Đang kiểm tra</option><option>Đã xử lý</option></select></article>)}</div>
        </div>
      </div> : <div className="user-test-panel"><div className="user-profile-summary"><div className="user-avatar" aria-hidden="true">{current.name.trim().charAt(0).toUpperCase()}</div><div><span>HỒ SƠ NGƯỜI DÙNG</span><strong>{current.name}</strong><small>{current.email} · Thành viên từ {new Date(current.createdAt).toLocaleDateString("vi-VN")}</small></div><a href="#thanh-toan">Quản lý gói</a></div><div className="account-notification-head"><div><span>TRUNG TÂM THÔNG BÁO</span><h3>{activeLicense ? `Bản quyền ${activeLicense.plan} đang hoạt động` : "Tài khoản SỸ LAND đã sẵn sàng"}</h3></div><b>{paymentNotices.filter((item) => item.licenseCode).length}</b></div>{paymentNotices.length > 0 && <div className="account-notifications">{paymentNotices.slice(0, 8).map((item) => <article className={item.licenseCode ? "license-ready" : ""} key={item.id}><span aria-hidden="true">{item.licenseCode ? "✓" : "…"}</span><div><strong>{item.licenseCode ? "Mã bản quyền đã được cấp" : `Đơn ${item.status.toLocaleLowerCase("vi-VN")}`}</strong><small>{item.orderCode} · Gói {item.plan} · {new Date(item.createdAt).toLocaleString("vi-VN")}</small>{item.licenseCode && <code>{item.licenseCode}</code>}</div>{item.licenseCode && <div className="notification-actions"><button type="button" onClick={() => void copyLicense(item.licenseCode)}>Sao chép</button><button type="button" onClick={() => void activateFromNotice(item.licenseCode)}>Kích hoạt ngay</button></div>}</article>)}</div>}{activeLicense ? <div className="activated-license"><span>ĐÃ KÍCH HOẠT TRÊN TÀI KHOẢN</span><strong>{activeLicense.code}</strong><p>Cấp cho {activeLicense.email} · {activeLicense.maxDevices || 1} thiết bị · Hết hạn {new Date(activeLicense.expiresAt).toLocaleDateString("vi-VN")}</p><div className="activated-actions"><button type="button" onClick={() => void copyLicense(activeLicense.code)}>Sao chép mã</button><button type="button" onClick={() => void activateFromNotice(activeLicense.code)}>Mở phần mềm và kích hoạt</button></div></div> : <><p>Bạn có thể sử dụng các công cụ công khai hoặc nhập mã bản quyền do quản trị viên cấp.</p><form className="activation-form" onSubmit={activateLicense}><label>Mã bản quyền<input value={activationCode} onChange={(event) => setActivationCode(event.target.value.toUpperCase())} placeholder="SYL-XXXXX-XXXXX" /></label><button type="submit">Kích hoạt mã</button></form></>}<section className="configuration-transfer"><div><span>ĐỒNG BỘ CẤU HÌNH</span><h3>Chuyển cấu hình giữa website và phần mềm</h3><p>Xuất địa bàn, đơn vị hành chính và tùy chọn nghiệp vụ ra JSON; sau đó nhập vào SỸ LAND trên thiết bị khác. Không xuất mật khẩu, phiên đăng nhập hoặc mã bản quyền.</p></div><input ref={configInputRef} type="file" accept="application/json,.json" onChange={(event) => void importConfiguration(event)} /><div className="configuration-actions"><button type="button" onClick={exportConfiguration}>Xuất cấu hình</button><button type="button" onClick={() => configInputRef.current?.click()}>Nhập từ phần mềm</button></div>{configMessage && <p role="status">{configMessage}</p>}<small>Định dạng dùng chung: syland-config · schema_version 1 · UTF-8 JSON</small></section><div className="account-quick-actions"><a href="#minh-hoa">Mở công cụ</a><a href="#cong-cu-pdf">Công cụ PDF</a><a href="#tai-phan-mem">Tải phần mềm</a><a href="#thanh-toan">Thanh toán</a></div><small>{REMOTE_AUTH ? "Bản quyền được gắn theo email và tự đồng bộ giữa website với phần mềm SỸ LAND khi đăng nhập." : "Chế độ cục bộ chưa đồng bộ dữ liệu giữa các thiết bị."}</small></div>}
      {message && <p className="account-message">{message}</p>}
    </section>
  );

  return (
    <section className="account-portal account-auth" id="tai-khoan" aria-labelledby="account-title">
      <div className="auth-copy"><p className="section-kicker">{REMOTE_AUTH ? "Tài khoản SỸ LAND" : AUTH_CONFIG_ERROR ? "Cần hoàn tất cấu hình" : "Chế độ tài khoản cục bộ"}</p><h2 id="account-title">{mode === "setup" ? "Thiết lập quản trị viên SỸ LAND" : mode === "register" ? "Đăng ký tài khoản SỸ LAND" : mode === "forgot" ? "Khôi phục mật khẩu" : mode === "reset" ? "Đặt mật khẩu mới" : "Đăng nhập SỸ LAND"}</h2><p>{AUTH_CONFIG_ERROR ? "Website chưa nhận được cấu hình máy chủ tài khoản. Đăng ký và đăng nhập tạm khóa để tránh tạo tài khoản không đồng bộ." : REMOTE_AUTH ? "Một tài khoản đăng nhập được trên cả website và phần mềm SỸ LAND." : mode === "setup" ? "Thiết lập mật khẩu quản trị lần đầu cho anh Nguyễn Minh Sỹ trên thiết bị này." : "Chế độ phát triển cục bộ đang hoạt động."}</p><ul><li>✓ Mật khẩu không được ghi trong mã website</li><li>✓ {REMOTE_AUTH ? "Tự khôi phục phiên đăng nhập qua HTTPS" : AUTH_CONFIG_ERROR ? "Cần thêm hai GitHub Actions secrets" : "Chỉ lưu mã băm trên trình duyệt"}</li><li>✓ Có quy trình quên và đặt lại mật khẩu</li></ul></div>
      <form className="auth-form" onSubmit={submitAccount}>{(mode === "register" || mode === "setup") && <label>Họ và tên<input value={name} onChange={(event) => setName(event.target.value)} autoComplete="name" /></label>}{mode !== "reset" && <label>Email<input type="email" value={email} onChange={(event) => setEmail(event.target.value)} autoComplete="email" placeholder="Nhập email đã đăng ký" /></label>}{mode !== "forgot" && <label>Mật khẩu<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete={mode === "login" ? "current-password" : "new-password"} placeholder="Tối thiểu 8 ký tự" /></label>}{mode !== "login" && mode !== "forgot" && <label>Xác nhận mật khẩu<input type="password" value={confirm} onChange={(event) => setConfirm(event.target.value)} autoComplete="new-password" /></label>}{(mode === "register" || mode === "setup") && <label className="terms-consent"><input type="checkbox" checked={acceptedTerms} onChange={(event) => setAcceptedTerms(event.target.checked)} /><span>Đồng ý <a href="#phap-ly">Chính sách bảo mật và Điều khoản sử dụng</a> phiên bản {TERMS_VERSION}.</span></label>}<button type="submit" disabled={AUTH_CONFIG_ERROR}>{mode === "setup" ? "Tạo tài khoản admin" : mode === "register" ? "Tạo tài khoản" : mode === "forgot" ? "Gửi email khôi phục" : mode === "reset" ? "Lưu mật khẩu mới" : "Đăng nhập"}</button>{message && <p className="account-message">{message}</p>}{mode === "login" && REMOTE_AUTH && <button className="auth-secondary" type="button" onClick={() => { setMode("forgot"); setMessage(""); }}>Quên mật khẩu?</button>}{mode !== "setup" && mode !== "reset" && <p className="auth-switch">{mode === "login" ? <>Chưa có tài khoản? <button type="button" onClick={() => { setMode("register"); setMessage(""); }}>Đăng ký tài khoản</button></> : <>Đã có tài khoản? <button type="button" onClick={() => { setMode("login"); setMessage(""); }}>Đăng nhập</button></>}</p>}</form>
    </section>
  );
}
