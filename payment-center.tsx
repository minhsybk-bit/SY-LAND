"use client";

import { useEffect, useMemo, useState } from "react";
import { PAYMENT_CONFIG, PAYMENT_PLANS } from "./payment-config";

type Payment = { id: string; orderCode: string; plan: string; amount: number; transferContent: string; status: "Chờ thanh toán" | "Chờ xác nhận" | "Đã thanh toán" | "Từ chối" | "Đã hủy"; licenseCode: string; createdAt: string };
const URL = String(import.meta.env.VITE_SUPABASE_URL || "").trim().replace(/\/$/, "");
const KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "");

function session() { try { return JSON.parse(localStorage.getItem("sy-land-auth-session") || "null"); } catch { return null; } }
async function rest(path: string, token: string, options: { method?: "GET" | "POST" | "PATCH"; payload?: unknown } = {}) {
  const response = await fetch(`${URL}/rest/v1${path}`, { method: options.method || "GET", headers: { apikey: KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: "return=representation" }, ...(options.payload === undefined ? {} : { body: JSON.stringify(options.payload) }) });
  const data = await response.json().catch(() => null); if (!response.ok) throw new Error(data?.message || data?.hint || "Không kết nối được máy chủ thanh toán."); return data;
}
function mapPayment(item: any): Payment { return { id: item.id, orderCode: item.order_code, plan: item.plan, amount: item.amount, transferContent: item.transfer_content, status: item.status, licenseCode: item.license_code || "", createdAt: item.created_at }; }

export default function PaymentCenter() {
  const [planId, setPlanId] = useState<(typeof PAYMENT_PLANS)[number]["id"]>("personal");
  const [payments, setPayments] = useState<Payment[]>([]);
  const [active, setActive] = useState<Payment | null>(null);
  const [message, setMessage] = useState("");
  const [busy, setBusy] = useState(false);
  const currentSession = typeof window === "undefined" ? null : session();
  const isAdmin = currentSession?.account?.role === "admin";
  const configured = Boolean(PAYMENT_CONFIG.bankBin && PAYMENT_CONFIG.accountNumber && PAYMENT_CONFIG.bankName);
  const plan = PAYMENT_PLANS.find((item) => item.id === planId) || PAYMENT_PLANS[0];
  const qrUrl = useMemo(() => active && configured ? `https://img.vietqr.io/image/${encodeURIComponent(PAYMENT_CONFIG.bankBin)}-${encodeURIComponent(PAYMENT_CONFIG.accountNumber)}-compact2.png?amount=${active.amount}&addInfo=${encodeURIComponent(active.transferContent)}&accountName=${encodeURIComponent(PAYMENT_CONFIG.accountName)}` : "", [active, configured]);

  async function load() {
    const auth = session(); if (!auth?.accessToken) return;
    try { const rows = await rest("/payment_orders?select=id,order_code,plan,amount,transfer_content,status,license_code,created_at&order=created_at.desc&limit=100", auth.accessToken); setPayments(Array.isArray(rows) ? rows.map(mapPayment) : []); }
    catch { /* Schema có thể chưa được cài. */ }
  }
  useEffect(() => { load(); }, []);

  async function createOrder() {
    const auth = session(); setMessage("");
    if (!auth?.accessToken) { setMessage("Hãy đăng nhập tài khoản SỸ LAND trước khi đăng ký gói."); return; }
    if (!configured) { setMessage("Thanh toán chưa mở vì quản trị viên chưa cấu hình tài khoản ngân hàng."); return; }
    setBusy(true);
    try {
      const rows = await rest("/payment_orders", auth.accessToken, { method: "POST", payload: { plan: plan.name, amount: plan.amount, duration_months: plan.months, max_devices: plan.maxDevices } });
      const order = mapPayment(Array.isArray(rows) ? rows[0] : rows); setActive(order); await load(); setMessage("Đã tạo đơn. Chuyển đúng số tiền và nội dung hiển thị bên dưới.");
    } catch (reason) { setMessage(reason instanceof Error ? `${reason.message} Hãy chạy SUPABASE_PAYMENTS.sql một lần.` : "Không tạo được đơn."); }
    finally { setBusy(false); }
  }

  async function markTransferred() {
    const auth = session(); if (!auth?.accessToken || !active) return; setBusy(true);
    try { await rest(`/payment_orders?id=eq.${encodeURIComponent(active.id)}`, auth.accessToken, { method: "PATCH", payload: { status: "Chờ xác nhận" } }); await load(); setActive((value) => value ? { ...value, status: "Chờ xác nhận" } : value); setMessage("Đã ghi nhận thông báo chuyển khoản. Mã bản quyền sẽ tự tạo ngay khi admin xác nhận tiền vào tài khoản."); }
    catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không cập nhật được đơn."); } finally { setBusy(false); }
  }

  async function adminConfirm(item: Payment, status: "Đã thanh toán" | "Từ chối") {
    const auth = session(); if (!isAdmin || !auth?.accessToken) return;
    if (status === "Đã thanh toán" && !window.confirm(`Xác nhận đã nhận ${item.amount.toLocaleString("vi-VN")}đ cho ${item.orderCode}? Mã bản quyền sẽ được cấp tự động.`)) return;
    setBusy(true);
    try { await rest(`/payment_orders?id=eq.${encodeURIComponent(item.id)}`, auth.accessToken, { method: "PATCH", payload: { status } }); await load(); setMessage(status === "Đã thanh toán" ? "Đã xác nhận. Hệ thống đã tự tạo mã bản quyền." : "Đã từ chối đơn thanh toán."); }
    catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không xác nhận được đơn."); } finally { setBusy(false); }
  }

  return <section className="payment-center" id="thanh-toan" aria-labelledby="payment-title"><div className="payment-heading"><div><p className="section-kicker">Thanh toán bản quyền</p><h2 id="payment-title">Chuyển khoản rõ nội dung.<br />Cấp mã sau đối soát.</h2></div><p>Hệ thống tạo nội dung chuyển khoản riêng cho từng đơn. Admin chỉ xác nhận sau khi tiền thực tế vào tài khoản; mã bản quyền được cơ sở dữ liệu tự cấp, không thể tạo giả từ trình duyệt.</p></div>
    {!configured && <p className="payment-warning">Chưa mở thanh toán: cần điền ngân hàng, mã BIN và số tài khoản trong <code>app/payment-config.ts</code>.</p>}
    <div className="payment-grid"><div className="payment-plans">{PAYMENT_PLANS.map((item) => <button type="button" className={item.id === planId ? "selected" : ""} key={item.id} onClick={() => setPlanId(item.id)}><span>{item.name}</span><strong>{item.amount.toLocaleString("vi-VN")}đ<small>/ tháng</small></strong><p>{item.description}</p></button>)}<button type="button" className="create-order" disabled={busy || !configured} onClick={createOrder}>Tạo đơn và mã QR</button></div>
      <div className="payment-qr">{active ? <><span>ĐƠN {active.orderCode}</span>{qrUrl && <img src={qrUrl} alt={`Mã QR chuyển khoản đơn ${active.orderCode}`} />}<dl><div><dt>Ngân hàng</dt><dd>{PAYMENT_CONFIG.bankName}</dd></div><div><dt>Số tài khoản</dt><dd>{PAYMENT_CONFIG.accountNumber}</dd></div><div><dt>Chủ tài khoản</dt><dd>{PAYMENT_CONFIG.accountName}</dd></div><div><dt>Số tiền</dt><dd>{active.amount.toLocaleString("vi-VN")}đ</dd></div><div><dt>Nội dung</dt><dd><code>{active.transferContent}</code></dd></div></dl><button type="button" disabled={busy || active.status !== "Chờ thanh toán"} onClick={markTransferred}>{active.status === "Chờ thanh toán" ? "Tôi đã chuyển khoản" : active.status}</button>{active.licenseCode && <div className="payment-license"><small>MÃ BẢN QUYỀN ĐÃ CẤP</small><strong>{active.licenseCode}</strong></div>}</> : <div className="payment-empty"><b>QR</b><p>Chọn gói và tạo đơn để hiển thị mã chuyển khoản chính xác.</p></div>}</div></div>
    {message && <p className="payment-message" role="status">{message}</p>}
    {payments.length > 0 && <div className="payment-history"><h3>{isAdmin ? "Đối soát thanh toán" : "Đơn của tôi"}</h3>{payments.map((item) => <article key={item.id}><div><strong>{item.orderCode}</strong><span>{item.plan} · {item.amount.toLocaleString("vi-VN")}đ · {new Date(item.createdAt).toLocaleString("vi-VN")}</span>{item.licenseCode && <code>{item.licenseCode}</code>}</div><b data-payment-status={item.status}>{item.status}</b>{isAdmin && item.status === "Chờ xác nhận" && <div className="payment-admin-actions"><button type="button" disabled={busy} onClick={() => adminConfirm(item, "Đã thanh toán")}>Xác nhận tiền vào</button><button type="button" disabled={busy} onClick={() => adminConfirm(item, "Từ chối")}>Từ chối</button></div>}</article>)}</div>}
  </section>;
}
