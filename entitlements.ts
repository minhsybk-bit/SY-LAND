import { useEffect, useState } from "react";

export type SylandPlan = "Dùng thử" | "Go" | "Plus" | "Pro" | "Văn phòng" | "Đơn vị";
export type EntitlementSnapshot = {
  role: "admin" | "user" | "guest";
  plan: SylandPlan;
  seats: number;
  featurePercent: number;
  maxParcelsPerRun: number | null;
  maxFileSizeMB: number;
  maxTotalUploadMB: number;
  maxUsesPerDay: number | null;
  fullTools: boolean;
  reason: string;
};

const KEY = "sy-land-entitlements";
const EVENT = "sy-land-entitlements-updated";

export function resolveEntitlements(input: { role?: string; plan?: string; seats?: number; licenseMaxParcels?: number | null }): EntitlementSnapshot {
  const role = input.role === "admin" ? "admin" : input.role === "user" ? "user" : "guest";
  const seats = Math.max(1, Math.floor(Number(input.seats) || 1));
  const raw = String(input.plan || "Dùng thử").toLocaleLowerCase("vi-VN");
  const plan: SylandPlan = raw.includes("văn phòng") ? "Văn phòng" : raw.includes("đơn vị") ? "Đơn vị" : raw.includes("pro") ? "Pro" : raw.includes("plus") || raw === "cá nhân" ? "Plus" : raw.includes("go") ? "Go" : "Dùng thử";
  if (role === "admin") return { role, plan: "Đơn vị", seats, featurePercent: 100, maxParcelsPerRun: null, maxFileSizeMB: 100, maxTotalUploadMB: 2000, maxUsesPerDay: null, fullTools: true, reason: "Quản trị viên SỸ LAND · không giới hạn" };

  let featurePercent = 25, accountLimit: number | null = role === "guest" ? 5 : 50, maxFileSizeMB = role === "guest" ? 10 : 20;
  let maxTotalUploadMB = role === "guest" ? 25 : 150, maxUsesPerDay: number | null = role === "guest" ? 3 : 10;
  let fullTools = false, reason = role === "guest" ? "Khách chưa đăng nhập · 3 lượt/ngày · 5 tệp/lượt · 10 MB/tệp" : "Tài khoản dùng thử · 10 lượt/ngày";
  if (plan === "Go") { featurePercent = 40; accountLimit = 80; maxTotalUploadMB = 500; maxUsesPerDay = null; reason = "Cá nhân Go · tối đa 80 thửa/lần"; }
  if (plan === "Plus") { featurePercent = 70; accountLimit = 140; maxFileSizeMB = 30; maxTotalUploadMB = 750; maxUsesPerDay = null; reason = "Cá nhân Plus · tối đa 140 thửa/lần"; }
  if (plan === "Pro") { featurePercent = 100; accountLimit = 200; maxFileSizeMB = 50; maxTotalUploadMB = 1000; maxUsesPerDay = null; fullTools = true; reason = "Cá nhân Pro · đầy đủ công cụ, tối đa 200 thửa/lần"; }
  if (plan === "Văn phòng" && seats < 5) { featurePercent = 70; accountLimit = 140; maxFileSizeMB = 30; maxTotalUploadMB = 750; maxUsesPerDay = null; reason = `Văn phòng ${seats} tài khoản · quyền tương đương Plus`; }
  if (plan === "Văn phòng" && seats >= 5) { featurePercent = 100; accountLimit = null; maxFileSizeMB = 100; maxTotalUploadMB = 2000; maxUsesPerDay = null; fullTools = true; reason = `Văn phòng ${seats} tài khoản · không giới hạn thửa/lần`; }
  if (plan === "Đơn vị") { featurePercent = 100; accountLimit = null; maxFileSizeMB = 100; maxTotalUploadMB = 2000; maxUsesPerDay = null; fullTools = true; reason = "Gói Đơn vị · không giới hạn thửa/lần"; }
  const licenseLimit = Number(input.licenseMaxParcels);
  const hasLicenseLimit = Number.isFinite(licenseLimit) && licenseLimit > 0;
  const maxParcelsPerRun = hasLicenseLimit ? (accountLimit == null ? Math.floor(licenseLimit) : Math.min(accountLimit, Math.floor(licenseLimit))) : accountLimit;
  if (hasLicenseLimit) reason += ` · mã bản quyền giới hạn ${Math.floor(licenseLimit)} thửa/lần`;
  return { role, plan, seats, featurePercent, maxParcelsPerRun, maxFileSizeMB, maxTotalUploadMB, maxUsesPerDay, fullTools, reason };
}

const USAGE_KEY = "sy-land-daily-usage";
function todayKey() { return new Date().toLocaleDateString("en-CA"); }
export function readDailyUsage() {
  try {
    const value = JSON.parse(localStorage.getItem(USAGE_KEY) || "null");
    return value?.date === todayKey() ? Math.max(0, Number(value.count) || 0) : 0;
  } catch { return 0; }
}
export function consumeDailyUsage(entitlements: EntitlementSnapshot) {
  if (entitlements.maxUsesPerDay == null) return { allowed: true, used: 0, remaining: null as number | null };
  const used = readDailyUsage();
  if (used >= entitlements.maxUsesPerDay) return { allowed: false, used, remaining: 0 };
  const next = used + 1;
  localStorage.setItem(USAGE_KEY, JSON.stringify({ date: todayKey(), count: next }));
  window.dispatchEvent(new Event("sy-land-usage-updated"));
  return { allowed: true, used: next, remaining: Math.max(0, entitlements.maxUsesPerDay - next) };
}

export function publishEntitlements(value: EntitlementSnapshot) {
  localStorage.setItem(KEY, JSON.stringify(value));
  window.dispatchEvent(new CustomEvent(EVENT, { detail: value }));
}
function readEntitlements() {
  try { const value = JSON.parse(localStorage.getItem(KEY) || "null"); if (value?.role && "maxParcelsPerRun" in value && "maxFileSizeMB" in value) return value as EntitlementSnapshot; } catch {}
  return resolveEntitlements({});
}
export function useEntitlements() {
  const [value, setValue] = useState<EntitlementSnapshot>(() => readEntitlements());
  useEffect(() => {
    const refresh = () => setValue(readEntitlements());
    const receive = (event: Event) => setValue((event as CustomEvent<EntitlementSnapshot>).detail || readEntitlements());
    window.addEventListener("storage", refresh); window.addEventListener(EVENT, receive);
    return () => { window.removeEventListener("storage", refresh); window.removeEventListener(EVENT, receive); };
  }, []);
  return value;
}
