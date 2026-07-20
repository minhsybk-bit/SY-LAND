"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { BILLING_CYCLES, PAYMENT_CONFIG, PAYMENT_PLANS, planTotal } from "./payment-config";

type Payment = { id: string; orderCode: string; plan: string; amount: number; transferContent: string; status: "Chờ thanh toán" | "Chờ xác nhận" | "Đã thanh toán" | "Từ chối" | "Đã hủy"; licenseCode: string; createdAt: string; confirmedAt: string };
const URL = String(import.meta.env.VITE_SUPABASE_URL || "").trim().replace(/\/$/, "");
const KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "");

function session() { try { return JSON.parse(localStorage.getItem("sy-land-auth-session") || "null"); } catch { return null; } }
async function rest(path: string, token: string, options: { method?: "GET" | "POST" | "PATCH"; payload?: unknown } = {}) {
  const response = await fetch(`${URL}/rest/v1${path}`, { method: options.method || "GET", headers: { apikey: KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: "return=representation" }, ...(options.payload === undefined ? {} : { body: JSON.stringify(options.payload) }) });
  const data = await response.json().catch(() => null); if (!response.ok) throw new Error(data?.message || data?.hint || "Không kết nối được máy chủ thanh toán."); return data;
}
function mapPayment(item: any): Payment { return { id: item.id, orderCode: item.order_code, plan: item.plan, amount: item.amount, transferContent: item.transfer_content, status: item.status, licenseCode: item.license_code || "", createdAt: item.created_at, confirmedAt: item.confirmed_at || "" }; }
function paymentStatusLabel(status: Payment["status"]) { return status === "Đã thanh toán" ? "Đã xác nhận" : status; }

export default function PaymentCenter() {
  const [planId, setPlanId] = useState<(typeof PAYMENT_PLANS)[number]["id"]>("personal");
  const [billingMonths, setBillingMonths] = useState(1);
  const [people, setPeople] = useState(1);
  const [volume, setVolume] = useState(200);
  const [budget, setBudget] = useState(300000);
  const [payments, setPayments] = useState<Payment[]>([]);
  const [active, setActive] = useState<Payment | null>(null);
  const [message, setMessage] = useState("");
  const [busy, setBusy] = useState(false);
  const knownPayments = useRef<Record<string, string>>({});
  const currentSession = typeof window === "undefined" ? null : session();
  const isAdmin = currentSession?.account?.role === "admin";
  const configured = Boolean(PAYMENT_CONFIG.bankBin && PAYMENT_CONFIG.accountNumber && PAYMENT_CONFIG.bankName);
  const plan = PAYMENT_PLANS.find((item) => item.id === planId) || PAYMENT_PLANS[0];
  const total = planTotal(plan.amount, billingMonths);
  const fullPrice = plan.amount * billingMonths;
  const saving = fullPrice - total;
  const recommendation = people <= 1 && volume <= 300 ? "personal" : people <= 5 && volume <= 2500 ? "office" : "custom";
  const qrUrl = useMemo(() => active && configured ? `https://img.vietqr.io/image/${encodeURIComponent(PAYMENT_CONFIG.bankBin)}-${encodeURIComponent(PAYMENT_CONFIG.accountNumber)}-compact2.png?amount=${active.amount}&addInfo=${encodeURIComponent(active.transferContent)}&accountName=${encodeURIComponent(PAYMENT_CONFIG.accountName)}` : "", [active, configured]);

  async function load() {
    const auth = session(); if (!auth?.accessToken) return;
    try {
      const rows = await rest("/payment_orders?select=id,order_code,plan,amount,transfer_content,status,license_code,created_at,confirmed_at&order=created_at.desc&limit=100", auth.accessToken);
      const next = Array.isArray(rows) ? rows.map(mapPayment) : [];
      const paid = next.find((item) => item.status === "Đã thanh toán" && item.licenseCode && knownPayments.current[item.id] && knownPayments.current[item.id] !== "Đã thanh toán");
      knownPayments.current = Object.fromEntries(next.map((item) => [item.id, item.status]));
      setPayments(next);
      setActive((value) => paid || (value ? next.find((item) => item.id === value.id) || value : value));
      if (paid) setMessage(`Thanh toán ${paid.orderCode} đã được xác nhận thành công. Mã bản quyền đã gửi vào Trung tâm tài khoản.`);
      window.dispatchEvent(new Event("syland-payment-updated"));
    }
    catch { /* Schema có thể chưa được cài. */ }
  }
  useEffect(() => { void load(); const timer = window.setInterval(() => void load(), 30000); return () => window.clearInterval(timer); }, []);

  async function createOrder() {
    const auth = session(); setMessage("");
    if (!auth?.accessToken) { setMessage("Hãy đăng nhập tài khoản SỸ LAND trước khi đăng ký gói."); return; }
    if (!configured) { setMessage("Thanh toán chưa mở vì quản trị viên chưa cấu hình tài khoản ngân hàng."); return; }
    setBusy(true);
    try {
      const rows = await rest("/payment_orders", auth.accessToken, { method: "POST", payload: { plan: plan.name, amount: total, duration_months: billingMonths, max_devices: plan.maxDevices } });
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
    const question = status === "Đã thanh toán"
      ? `Xác nhận đã nhận ${item.amount.toLocaleString("vi-VN")}đ cho ${item.orderCode}? Mã bản quyền sẽ được cấp tự động.`
      : `Từ chối đơn ${item.orderCode}? Người dùng sẽ thấy trạng thái đơn bị từ chối.`;
    if (!window.confirm(question)) { setMessage("Đã hủy thao tác, trạng thái đơn không thay đổi."); return; }
    setBusy(true);
    try {
      const rows = await rest("/rpc/admin_confirm_payment", auth.accessToken, { method: "POST", payload: { p_order_id: item.id, p_status: status } });
      const row = Array.isArray(rows) ? rows[0] : rows;
      if (!row) throw new Error("Máy chủ không trả về đơn đã cập nhật. Hãy kiểm tra quyền quản trị và chính sách RLS.");
      const updated = mapPayment(row);
      if (status === "Đã thanh toán" && (!updated?.licenseCode || updated.status !== "Đã thanh toán")) throw new Error("Máy chủ chưa tạo được mã bản quyền. Hãy chạy lại SUPABASE_PAYMENTS.sql rồi xác nhận lại.");
      setPayments((items) => items.map((value) => value.id === updated.id ? updated : value));
      setActive((value) => value?.id === updated.id ? updated : value);
      window.dispatchEvent(new Event("syland-payment-updated"));
      await load();
      setMessage(status === "Đã thanh toán" ? `Đã xác nhận thanh toán và cấp mã ${updated.licenseCode}. Người dùng sẽ nhận trong Trung tâm tài khoản.` : "Đã từ chối đơn thanh toán.");
    }
    catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không xác nhận được đơn."); } finally { setBusy(false); }
  }

  function applyRecommendation() {
    if (recommendation === "custom") { document.getElementById("tu-van")?.scrollIntoView({ behavior: "smooth" }); setMessage("Quy mô vượt gói Văn phòng. Hãy gửi nhu cầu để nhận báo giá theo mức sử dụng thực tế."); return; }
    setPlanId(recommendation); setMessage(`Đã chọn gói ${recommendation === "personal" ? "Cá nhân" : "Văn phòng"} theo nhu cầu đã nhập.`);
  }

  function openAccountCenter() {
    window.dispatchEvent(new Event("syland-payment-updated"));
    document.getElementById("tai-khoan")?.scrollIntoView({ behavior: "smooth" });
    setMessage("Đã chuyển đến Trung tâm tài khoản. Mở thông báo mới để nhận và kích hoạt mã bản quyền.");
  }

  return <section className="payment-center" id="thanh-toan" aria-labelledby="payment-title"><div className="payment-heading"><div><p className="section-kicker">Chọn gói và thanh toán</p><h2 id="payment-title">Đúng nhu cầu.<br />Không mua thừa.</h2></div><p>Nhập quy mô sử dụng để SỸ LAND đề xuất gói kinh tế phù hợp. Người dùng vẫn được xem đầy đủ giá, thời hạn, mức tiết kiệm và chủ động đổi lựa chọn trước khi tạo đơn.</p></div>
    {!configured && <p className="payment-warning">Chưa mở thanh toán: cần điền ngân hàng, mã BIN và số tài khoản trong <code>app/payment-config.ts</code>.</p>}
    <div className="plan-advisor"><div><label>Số người sử dụng<input type="number" min="1" max="500" value={people} onChange={(event) => setPeople(Math.max(1, Number(event.target.value) || 1))} /></label><label>Hồ sơ dự kiến/tháng<input type="number" min="1" value={volume} onChange={(event) => setVolume(Math.max(1, Number(event.target.value) || 1))} /></label><label>Ngân sách/tháng<input type="number" min="0" step="50000" value={budget} onChange={(event) => setBudget(Math.max(0, Number(event.target.value) || 0))} /></label></div><aside><span>GỢI Ý PHÙ HỢP</span><strong>{recommendation === "personal" ? "Gói Cá nhân" : recommendation === "office" ? "Gói Văn phòng" : "Gói Đơn vị"}</strong><p>{recommendation === "personal" ? "Đáp ứng tối đa 300 hồ sơ/tháng cho 01 người dùng." : recommendation === "office" ? "Phù hợp nhóm tối đa 05 người và 2.500 hồ sơ/tháng." : "Cần báo giá riêng để không phải trả cho hạn mức không sử dụng."}</p>{recommendation !== "custom" && budget < (recommendation === "personal" ? 199000 : 1490000) && <small>Ngân sách hiện thấp hơn giá tháng. Có thể tiếp tục dùng thử hoặc liên hệ để cân đối phạm vi.</small>}<button type="button" onClick={applyRecommendation}>{recommendation === "custom" ? "Nhận báo giá riêng" : "Chọn gói đề xuất"}</button></aside></div>
    <div className="billing-cycles" aria-label="Chu kỳ thanh toán">{BILLING_CYCLES.map((cycle) => <button type="button" className={billingMonths === cycle.months ? "selected" : ""} key={cycle.months} onClick={() => setBillingMonths(cycle.months)}><strong>{cycle.label}</strong><span>{cycle.discount ? `Tiết kiệm ${cycle.discount}%` : "Linh hoạt"}</span></button>)}</div>
    <div className="payment-grid"><div className="payment-plans">{PAYMENT_PLANS.map((item) => { const itemTotal = planTotal(item.amount, billingMonths); return <button type="button" className={item.id === planId ? "selected" : ""} key={item.id} onClick={() => setPlanId(item.id)}><span>{item.name}</span><strong>{itemTotal.toLocaleString("vi-VN")}đ<small>/ {billingMonths === 1 ? "tháng" : `${billingMonths} tháng`}</small></strong>{billingMonths > 1 && <em>{item.amount.toLocaleString("vi-VN")}đ/tháng · tiết kiệm {(item.amount * billingMonths - itemTotal).toLocaleString("vi-VN")}đ</em>}<p>{item.description}</p></button>})}<div className="payment-total"><span>Tổng thanh toán</span><strong>{total.toLocaleString("vi-VN")}đ</strong>{saving > 0 && <small>Đã giảm {saving.toLocaleString("vi-VN")}đ</small>}</div><button type="button" className="create-order" disabled={busy || !configured} onClick={createOrder}>Tạo đơn và mã QR</button></div>
      <div className="payment-qr">{active ? <><span>ĐƠN {active.orderCode}</span>{qrUrl && !active.licenseCode && <img src={qrUrl} alt={`Mã QR chuyển khoản đơn ${active.orderCode}`} />}<dl><div><dt>Ngân hàng</dt><dd>{PAYMENT_CONFIG.bankName}</dd></div><div><dt>Số tài khoản</dt><dd>{PAYMENT_CONFIG.accountNumber}</dd></div><div><dt>Chủ tài khoản</dt><dd>{PAYMENT_CONFIG.accountName}</dd></div><div><dt>Số tiền</dt><dd>{active.amount.toLocaleString("vi-VN")}đ</dd></div><div><dt>Nội dung</dt><dd><code>{active.transferContent}</code></dd></div></dl><button type="button" disabled={busy || active.status !== "Chờ thanh toán"} onClick={markTransferred}>{active.status === "Chờ thanh toán" ? "Tôi đã chuyển khoản" : active.status}</button>{active.licenseCode && <div className="payment-license"><small>THANH TOÁN THÀNH CÔNG</small><p>Mã bản quyền đã được gửi an toàn đến Trung tâm tài khoản.</p><button type="button" onClick={openAccountCenter}>Mở Trung tâm tài khoản</button></div>}</> : <div className="payment-empty"><b>QR</b><p>Chọn gói và tạo đơn để hiển thị mã chuyển khoản chính xác.</p></div>}</div></div>
    {message && <p className="payment-message" role="status">{message}</p>}
    {payments.length > 0 && <div className="payment-history"><h3>{isAdmin ? "Đối soát thanh toán" : "Đơn của tôi"}</h3>{payments.map((item) => <article key={item.id}><div><strong>{item.orderCode}</strong><span>{item.plan} · {item.amount.toLocaleString("vi-VN")}đ · {new Date(item.createdAt).toLocaleString("vi-VN")}</span></div><b data-payment-status={item.status}>{paymentStatusLabel(item.status)}</b>{isAdmin && item.status === "Chờ xác nhận" && <div className="payment-admin-actions"><button type="button" disabled={busy} onClick={() => adminConfirm(item, "Đã thanh toán")}>Xác nhận đã nhận tiền</button><button type="button" disabled={busy} onClick={() => adminConfirm(item, "Từ chối")}>Từ chối đơn</button></div>}</article>)}</div>}
  </section>;
}
