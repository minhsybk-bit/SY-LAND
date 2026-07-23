"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { BILLING_CYCLES, EMPTY_PAYMENT_CONFIG, PAYMENT_PLANS, planTotal } from "./payment-config";
import type { PaymentConfig } from "./payment-config";

type Payment = { id: string; orderCode: string; plan: string; amount: number; seatCount: number; transferContent: string; status: "Chờ thanh toán" | "Chờ xác nhận" | "Đã thanh toán" | "Từ chối" | "Đã hủy"; licenseCode: string; createdAt: string; confirmedAt: string };
type ActivePlan = { plan: string; expiresAt: string; seatCount: number; maxParcelsPerRun: number | null };
const URL = String(import.meta.env.VITE_SUPABASE_URL || "").trim().replace(/\/$/, "");
const KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "");

function session() { try { return JSON.parse(localStorage.getItem("sy-land-auth-session") || "null"); } catch { return null; } }
async function rest(path: string, token: string, options: { method?: "GET" | "POST" | "PATCH"; payload?: unknown; prefer?: string } = {}) {
  const response = await fetch(`${URL}/rest/v1${path}`, { method: options.method || "GET", headers: { apikey: KEY, Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: options.prefer || "return=representation" }, ...(options.payload === undefined ? {} : { body: JSON.stringify(options.payload) }) });
  const data = await response.json().catch(() => null);
  if (!response.ok) {
    const detail = [data?.message, data?.details, data?.hint].filter(Boolean).join(" · ");
    throw new Error(detail || `Máy chủ thanh toán trả về lỗi ${response.status}.`);
  }
  return data;
}
function mapPayment(item: any): Payment { return { id: item.id, orderCode: item.order_code, plan: item.plan, amount: item.amount, seatCount: Number(item.seat_count || item.max_devices || 1), transferContent: item.transfer_content, status: item.status, licenseCode: item.license_code || "", createdAt: item.created_at, confirmedAt: item.confirmed_at || "" }; }
function paymentStatusLabel(status: Payment["status"]) { return status === "Đã thanh toán" ? "Đã xác nhận" : status; }

export default function PaymentCenter() {
  const [planId, setPlanId] = useState<(typeof PAYMENT_PLANS)[number]["id"]>("plus");
  const [billingMonths, setBillingMonths] = useState(1);
  const [people, setPeople] = useState(1);
  const [volume, setVolume] = useState(200);
  const [budget, setBudget] = useState(300000);
  const [payments, setPayments] = useState<Payment[]>([]);
  const [active, setActive] = useState<Payment | null>(null);
  const [message, setMessage] = useState("");
  const [busy, setBusy] = useState(false);
  const [paymentConfig, setPaymentConfig] = useState<PaymentConfig>(EMPTY_PAYMENT_CONFIG);
  const [configDraft, setConfigDraft] = useState<PaymentConfig>(EMPTY_PAYMENT_CONFIG);
  const [activePlan, setActivePlan] = useState<ActivePlan | null>(null);
  const knownPayments = useRef<Record<string, string>>({});
  const currentSession = typeof window === "undefined" ? null : session();
  const isAdmin = currentSession?.account?.role === "admin";
  const configured = Boolean(paymentConfig.bankBin && paymentConfig.accountNumber && paymentConfig.bankName && paymentConfig.accountName);
  const plan = PAYMENT_PLANS.find((item) => item.id === planId) || PAYMENT_PLANS[0];
  const officeSeats = Math.max(2, Math.min(500, people));
  const monthlyPrice = plan.id === "office" ? plan.amount * officeSeats : plan.amount;
  const total = planTotal(monthlyPrice, billingMonths);
  const fullPrice = monthlyPrice * billingMonths;
  const saving = fullPrice - total;
  const recommendation = people >= 2 ? "office" : volume <= 80 ? "go" : volume <= 140 ? "plus" : "pro";
  const qrUrl = useMemo(() => active && configured ? `https://img.vietqr.io/image/${encodeURIComponent(paymentConfig.bankBin)}-${encodeURIComponent(paymentConfig.accountNumber)}-compact2.png?amount=${active.amount}&addInfo=${encodeURIComponent(active.transferContent)}&accountName=${encodeURIComponent(paymentConfig.accountName)}` : "", [active, configured, paymentConfig]);

  async function load() {
    const auth = session(); if (!auth?.accessToken) return;
    try {
      const settings = await rest("/payment_settings?select=bank_bin,bank_name,account_number,account_name,support_phone&id=eq.primary&limit=1", auth.accessToken);
      if (Array.isArray(settings) && settings[0]) {
        const nextConfig = { bankBin: settings[0].bank_bin || "", bankName: settings[0].bank_name || "", accountNumber: settings[0].account_number || "", accountName: settings[0].account_name || "", supportPhone: settings[0].support_phone || "" };
        setPaymentConfig(nextConfig); setConfigDraft(nextConfig);
      }
      const rows = await rest("/payment_orders?select=id,order_code,plan,amount,seat_count,max_devices,transfer_content,status,license_code,created_at,confirmed_at&order=created_at.desc&limit=100", auth.accessToken);
      const next = Array.isArray(rows) ? rows.map(mapPayment) : [];
      const previous = knownPayments.current;
      const paid = next.find((item) => item.status === "Đã thanh toán" && item.licenseCode && knownPayments.current[item.id] && !knownPayments.current[item.id].startsWith("Đã thanh toán:"));
      const changed = Object.keys(previous).length > 0 && (
        next.length !== Object.keys(previous).length ||
        next.some((item) => previous[item.id] !== `${item.status}:${item.licenseCode}`)
      );
      knownPayments.current = Object.fromEntries(next.map((item) => [item.id, `${item.status}:${item.licenseCode}`]));
      setPayments(next);
      setActive((value) => paid || (value ? next.find((item) => item.id === value.id) || value : value));
      if (auth.account?.role === "admin") {
        setActivePlan({ plan: "Quản trị viên · Full Access", expiresAt: "", seatCount: 1, maxParcelsPerRun: null });
      } else {
        const licenseRows = await rest(`/licenses?select=plan,expires_at,seat_count,max_parcels_per_run&status=eq.${encodeURIComponent("Hoạt động")}&expires_at=gt.${encodeURIComponent(new Date().toISOString())}&order=expires_at.desc&limit=1`, auth.accessToken);
        const license = Array.isArray(licenseRows) ? licenseRows[0] : null;
        setActivePlan(license ? { plan: license.plan, expiresAt: license.expires_at || "", seatCount: Number(license.seat_count || 1), maxParcelsPerRun: license.max_parcels_per_run == null ? null : Number(license.max_parcels_per_run) } : null);
      }
      if (paid) setMessage(`Thanh toán ${paid.orderCode} đã được xác nhận thành công. Mã bản quyền đã gửi vào Trung tâm tài khoản.`);
      if (changed) window.dispatchEvent(new Event("syland-payment-updated"));
    }
    catch { /* Schema có thể chưa được cài. */ }
  }
  useEffect(() => {
    const refresh = () => { if (document.visibilityState === "visible") void load(); };
    refresh();
    const timer = window.setInterval(refresh, 60000);
    document.addEventListener("visibilitychange", refresh);
    return () => { window.clearInterval(timer); document.removeEventListener("visibilitychange", refresh); };
  }, []);
  useEffect(() => {
    const selectPlan = (event: Event) => {
      const next = (event as CustomEvent<{ plan?: string }>).detail?.plan;
      if (next === "plus" || next === "pro") setPlanId(next);
    };
    window.addEventListener("syland-select-plan", selectPlan);
    return () => window.removeEventListener("syland-select-plan", selectPlan);
  }, []);

  async function savePaymentConfig() {
    const auth = session(); if (!isAdmin || !auth?.accessToken) return;
    if (!configDraft.bankBin.trim() || !configDraft.bankName.trim() || !configDraft.accountNumber.trim() || !configDraft.accountName.trim()) {
      setMessage("Hãy nhập đủ mã BIN, ngân hàng, số tài khoản và tên chủ tài khoản."); return;
    }
    if (!/^\d{6}$/.test(configDraft.bankBin.trim())) { setMessage("Mã BIN VietQR phải gồm đúng 6 chữ số."); return; }
    if (!/^\d{6,30}$/.test(configDraft.accountNumber.trim())) { setMessage("Số tài khoản chỉ gồm chữ số, dài từ 6 đến 30 ký tự."); return; }
    setBusy(true);
    try {
      const payload = { p_bank_bin: configDraft.bankBin.trim(), p_bank_name: configDraft.bankName.trim(), p_account_number: configDraft.accountNumber.trim(), p_account_name: configDraft.accountName.trim().toUpperCase(), p_support_phone: configDraft.supportPhone.trim() };
      const rows = await rest("/rpc/admin_save_payment_settings", auth.accessToken, { method: "POST", payload });
      const saved = Array.isArray(rows) ? rows[0] : rows;
      if (!saved?.bank_bin || !saved?.account_number) throw new Error("Máy chủ chưa trả về cấu hình vừa lưu.");
      const nextConfig = { bankBin: saved.bank_bin, bankName: saved.bank_name, accountNumber: saved.account_number, accountName: saved.account_name, supportPhone: saved.support_phone || "" };
      setPaymentConfig(nextConfig);
      setConfigDraft(nextConfig);
      setMessage("Đã lưu cấu hình nhận tiền an toàn trong Supabase. Thông tin không nằm trong mã nguồn GitHub.");
    } catch (reason) { setMessage(reason instanceof Error ? `Không lưu được: ${reason.message} Hãy chạy toàn bộ tệp SUPABASE_REPAIR_AUTH_PAYMENTS.sql trong một New query.` : "Không lưu được cấu hình nhận tiền."); }
    finally { setBusy(false); }
  }

  async function createOrder() {
    const auth = session(); setMessage("");
    if (!auth?.accessToken) { setMessage("Hãy đăng nhập tài khoản SỸ LAND trước khi đăng ký gói."); return; }
    if (!configured) { setMessage("Thanh toán chưa mở vì quản trị viên chưa cấu hình tài khoản ngân hàng."); return; }
    setBusy(true);
    try {
      const seats = plan.id === "office" ? officeSeats : 1;
      const rows = await rest("/payment_orders", auth.accessToken, { method: "POST", payload: { plan: plan.name, amount: total, duration_months: billingMonths, seat_count: seats, max_devices: seats } });
      const order = mapPayment(Array.isArray(rows) ? rows[0] : rows); setActive(order); await load(); setMessage("Đã tạo đơn. Chuyển đúng số tiền và nội dung hiển thị bên dưới.");
    } catch (reason) {
      const detail = reason instanceof Error ? reason.message : "";
      const outdatedPlans = /Gói thanh toán không hợp lệ|payment_orders_(plan|amount|duration_months|max_devices|seat_count)_check/i.test(detail);
      setMessage(outdatedPlans
        ? "Máy chủ thanh toán đang dùng cấu trúc cũ. Quản trị viên cần chạy toàn bộ tệp SUPABASE_REPAIR_AUTH_PAYMENTS.sql trong một New query của Supabase SQL Editor."
        : detail || "Không tạo được đơn thanh toán. Vui lòng thử lại.");
    }
    finally { setBusy(false); }
  }

  async function markTransferred() {
    const auth = session(); if (!auth?.accessToken || !active) return; setBusy(true);
    try { await rest(`/payment_orders?id=eq.${encodeURIComponent(active.id)}`, auth.accessToken, { method: "PATCH", payload: { status: "Chờ xác nhận" } }); await load(); setActive((value) => value ? { ...value, status: "Chờ xác nhận" } : value); setMessage("Đã ghi nhận thông báo chuyển khoản. Mã bản quyền sẽ tự tạo ngay khi admin xác nhận tiền vào tài khoản."); }
    catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không cập nhật được đơn."); } finally { setBusy(false); }
  }

  async function cancelPayment(item: Payment) {
    const auth = session();
    if (!auth?.accessToken || !["Chờ thanh toán", "Chờ xác nhận"].includes(item.status)) return;
    if (!window.confirm(`Hủy giao dịch ${item.orderCode}?\n\nĐơn đã hủy sẽ không được cấp mã bản quyền. Thao tác này không ảnh hưởng các gói đã kích hoạt.`)) {
      setMessage("Đã giữ nguyên giao dịch."); return;
    }
    setBusy(true);
    try {
      const rows = await rest("/rpc/cancel_my_payment", auth.accessToken, { method: "POST", payload: { p_order_id: item.id } });
      const row = Array.isArray(rows) ? rows[0] : rows;
      if (!row) throw new Error("Không tìm thấy giao dịch có thể hủy.");
      const updated = mapPayment(row);
      setPayments((items) => items.map((value) => value.id === updated.id ? updated : value));
      setActive((value) => value?.id === updated.id ? updated : value);
      setMessage(`Đã hủy giao dịch ${updated.orderCode}.`);
      window.dispatchEvent(new Event("syland-payment-updated"));
    } catch (reason) {
      setMessage(reason instanceof Error ? `${reason.message} Hãy chạy tệp SUPABASE_ADD_CANCEL_PAYMENT.sql một lần.` : "Không hủy được giao dịch.");
    } finally { setBusy(false); }
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
      if (status === "Đã thanh toán" && (!updated?.licenseCode || updated.status !== "Đã thanh toán")) throw new Error("Máy chủ chưa tạo được mã bản quyền. Hãy chạy toàn bộ SUPABASE_REPAIR_AUTH_PAYMENTS.sql rồi xác nhận lại.");
      setPayments((items) => items.map((value) => value.id === updated.id ? updated : value));
      setActive((value) => value?.id === updated.id ? updated : value);
      window.dispatchEvent(new Event("syland-payment-updated"));
      await load();
      setMessage(status === "Đã thanh toán" ? `Đã xác nhận thanh toán và cấp mã ${updated.licenseCode}. Người dùng sẽ nhận trong Trung tâm tài khoản.` : "Đã từ chối đơn thanh toán.");
    }
    catch (reason) { setMessage(reason instanceof Error ? reason.message : "Không xác nhận được đơn."); } finally { setBusy(false); }
  }

  function applyRecommendation() {
    setPlanId(recommendation); setMessage(`Đã chọn gói ${recommendation === "office" ? "Văn phòng" : recommendation[0].toUpperCase() + recommendation.slice(1)} theo nhu cầu đã nhập.`);
  }

  function openAccountCenter() {
    window.dispatchEvent(new Event("syland-payment-updated"));
    document.getElementById("tai-khoan")?.scrollIntoView({ behavior: "smooth" });
    setMessage("Đã chuyển đến Trung tâm tài khoản. Mở thông báo mới để nhận và kích hoạt mã bản quyền.");
  }

  return <section className="payment-center" aria-labelledby="payment-title"><div className="payment-heading"><div><p className="section-kicker">Chọn gói và thanh toán</p><h2 id="payment-title">Đúng nhu cầu.<br />Không mua thừa.</h2></div><p>Nhập quy mô sử dụng để SỸ LAND đề xuất gói kinh tế phù hợp. Người dùng vẫn được xem đầy đủ giá, thời hạn, mức tiết kiệm và chủ động đổi lựa chọn trước khi tạo đơn.</p></div>
    {currentSession?.accessToken && <div className="current-plan-banner"><div><span>GÓI ĐANG SỬ DỤNG</span><strong>{activePlan ? activePlan.plan : "Chưa có gói trả phí"}</strong></div><p>{activePlan ? `${activePlan.seatCount} tài khoản · ${activePlan.maxParcelsPerRun == null ? "Không giới hạn thửa/lượt" : `${activePlan.maxParcelsPerRun} thửa/lượt`}${activePlan.expiresAt ? ` · Hết hạn ${new Date(activePlan.expiresAt).toLocaleDateString("vi-VN")}` : ""}` : "Tài khoản đang ở chế độ trải nghiệm. Chọn gói bên dưới để mở rộng hạn mức."}</p></div>}
    {!configured && <p className="payment-warning">Thanh toán đang tạm khóa vì chưa có cấu hình nhận tiền trong Supabase. Quản trị viên đăng nhập và nhập thông tin bên dưới một lần.</p>}
    {isAdmin && <div className="payment-secure-config"><header><div><b>Cấu hình nhận tiền bảo mật</b><small>Lưu trong Supabase; không ghi vào GitHub. Người chưa đăng nhập không đọc được.</small></div><span>{configured ? "ĐÃ CẤU HÌNH" : "CHƯA CẤU HÌNH"}</span></header><div><label>Mã BIN VietQR<input value={configDraft.bankBin} inputMode="numeric" maxLength={6} onChange={(event) => setConfigDraft((value) => ({ ...value, bankBin: event.target.value.replace(/\D/g, "") }))} placeholder="6 chữ số" /></label><label>Tên ngân hàng<input value={configDraft.bankName} onChange={(event) => setConfigDraft((value) => ({ ...value, bankName: event.target.value }))} placeholder="Tên ngân hàng" /></label><label>Số tài khoản<input value={configDraft.accountNumber} inputMode="numeric" onChange={(event) => setConfigDraft((value) => ({ ...value, accountNumber: event.target.value.replace(/\D/g, "") }))} placeholder="Số tài khoản nhận tiền" /></label><label>Chủ tài khoản<input value={configDraft.accountName} onChange={(event) => setConfigDraft((value) => ({ ...value, accountName: event.target.value }))} placeholder="VIẾT HOA KHÔNG DẤU" /></label><label>Số hỗ trợ (tùy chọn)<input value={configDraft.supportPhone} onChange={(event) => setConfigDraft((value) => ({ ...value, supportPhone: event.target.value }))} placeholder="Chỉ dùng số hỗ trợ kinh doanh" /></label><button type="button" disabled={busy} onClick={() => void savePaymentConfig()}>{busy ? "Đang lưu…" : "Lưu an toàn"}</button></div>{message && <p className="payment-config-message" role="status">{message}</p>}</div>}
    <div className="plan-advisor"><div><label>Số người sử dụng<input type="number" min="1" max="500" value={people} onChange={(event) => setPeople(Math.max(1, Math.min(500, Number(event.target.value) || 1)))} /></label><label>Số thửa cần xử lý/lượt<input type="number" min="1" value={volume} onChange={(event) => setVolume(Math.max(1, Number(event.target.value) || 1))} /></label><label>Ngân sách/tháng<input type="number" min="0" step="50000" value={budget} onChange={(event) => setBudget(Math.max(0, Number(event.target.value) || 0))} /></label></div><aside><span>GỢI Ý PHÙ HỢP</span><strong>Gói {recommendation === "office" ? "Văn phòng" : recommendation[0].toUpperCase() + recommendation.slice(1)}</strong><p>{recommendation === "go" ? "Tối đa 80 thửa/lượt và 40% công cụ." : recommendation === "plus" ? "Tối đa 140 thửa/lượt và 70% công cụ." : recommendation === "pro" ? "Tối đa 200 thửa/lượt, đầy đủ công cụ." : people >= 5 ? `${people} tài khoản, đầy đủ công cụ và không giới hạn thửa/lượt.` : `${people} tài khoản, quyền tương đương Plus.`}</p>{budget < (recommendation === "go" ? 99000 : recommendation === "plus" ? 199000 : recommendation === "pro" ? 399000 : 298000 * officeSeats) && <small>Ngân sách hiện thấp hơn giá tháng. Có thể chọn gói thấp hơn hoặc liên hệ để được tư vấn.</small>}<button type="button" onClick={applyRecommendation}>Chọn gói đề xuất</button></aside></div>
    <div className="billing-cycles" aria-label="Chu kỳ thanh toán">{BILLING_CYCLES.map((cycle) => <button type="button" className={billingMonths === cycle.months ? "selected" : ""} key={cycle.months} onClick={() => setBillingMonths(cycle.months)}><strong>{cycle.label}</strong><span>{cycle.discount ? `Tiết kiệm ${cycle.discount}%` : "Linh hoạt"}</span></button>)}</div>
    <div className="payment-grid"><div className="payment-plans">{PAYMENT_PLANS.map((item) => { const itemMonthly = item.id === "office" ? item.amount * officeSeats : item.amount; const itemTotal = planTotal(itemMonthly, billingMonths); return <button type="button" className={item.id === planId ? "selected" : ""} key={item.id} onClick={() => setPlanId(item.id)}><span>{item.name}</span><strong>{itemTotal.toLocaleString("vi-VN")}đ<small>/ {billingMonths === 1 ? "tháng" : `${billingMonths} tháng`}</small></strong>{item.id === "office" && <em>{officeSeats} tài khoản × {item.amount.toLocaleString("vi-VN")}đ/tháng</em>}{billingMonths > 1 && <em>Tiết kiệm {(itemMonthly * billingMonths - itemTotal).toLocaleString("vi-VN")}đ</em>}<p>{item.description}</p></button>})}{plan.id === "office" && <label className="office-seat-picker">Số tài khoản Văn phòng<input type="number" min="2" max="500" value={officeSeats} onChange={(event) => setPeople(Math.max(2, Math.min(500, Number(event.target.value) || 2)))} /><small>2–4 tài khoản: quyền Plus · từ 5 tài khoản: đầy đủ, không giới hạn/lượt.</small></label>}<div className="payment-total"><span>Tổng thanh toán</span><strong>{total.toLocaleString("vi-VN")}đ</strong>{saving > 0 && <small>Đã giảm {saving.toLocaleString("vi-VN")}đ</small>}</div><button type="button" className="create-order" disabled={busy || !configured} onClick={createOrder}>Tạo đơn và mã QR</button></div>
      <div className="payment-qr">{active ? <><span>ĐƠN {active.orderCode}</span>{qrUrl && !active.licenseCode && active.status !== "Đã hủy" && <img src={qrUrl} alt={`Mã QR chuyển khoản đơn ${active.orderCode}`} />}<dl><div><dt>Ngân hàng</dt><dd>{paymentConfig.bankName}</dd></div><div><dt>Số tài khoản</dt><dd>{paymentConfig.accountNumber}</dd></div><div><dt>Chủ tài khoản</dt><dd>{paymentConfig.accountName}</dd></div><div><dt>Số tiền</dt><dd>{active.amount.toLocaleString("vi-VN")}đ</dd></div><div><dt>Nội dung</dt><dd><code>{active.transferContent}</code></dd></div></dl><div className="payment-order-actions"><button type="button" disabled={busy || active.status !== "Chờ thanh toán"} onClick={markTransferred}>{active.status === "Chờ thanh toán" ? "Tôi đã chuyển khoản" : active.status}</button>{!isAdmin && ["Chờ thanh toán", "Chờ xác nhận"].includes(active.status) && <button type="button" className="cancel-payment" disabled={busy} onClick={() => void cancelPayment(active)}>Hủy giao dịch</button>}</div>{active.licenseCode && <div className="payment-license"><small>THANH TOÁN THÀNH CÔNG</small><p>Mã bản quyền đã được gửi an toàn đến Trung tâm tài khoản.</p><button type="button" onClick={openAccountCenter}>Mở Trung tâm tài khoản</button></div>}</> : <div className="payment-empty"><b>QR</b><p>Chọn gói và tạo đơn để hiển thị mã chuyển khoản chính xác.</p></div>}</div></div>
    {message && <p className="payment-message" role="status">{message}</p>}
    {payments.length > 0 && <div className="payment-history"><h3>{isAdmin ? "Đối soát thanh toán" : "Đơn của tôi"}</h3>{payments.map((item) => <article key={item.id}><div><strong>{item.orderCode}</strong><span>{item.plan}{item.seatCount > 1 ? ` · ${item.seatCount} tài khoản` : ""} · {item.amount.toLocaleString("vi-VN")}đ · {new Date(item.createdAt).toLocaleString("vi-VN")}</span></div><b data-payment-status={item.status}>{paymentStatusLabel(item.status)}</b>{isAdmin && item.status === "Chờ xác nhận" && <div className="payment-admin-actions"><button type="button" disabled={busy} onClick={() => adminConfirm(item, "Đã thanh toán")}>Xác nhận đã nhận tiền</button><button type="button" disabled={busy} onClick={() => adminConfirm(item, "Từ chối")}>Từ chối đơn</button></div>}{!isAdmin && ["Chờ thanh toán", "Chờ xác nhận"].includes(item.status) && <div className="payment-user-actions"><button type="button" disabled={busy} onClick={() => void cancelPayment(item)}>Hủy giao dịch</button></div>}</article>)}</div>}
  </section>;
}
