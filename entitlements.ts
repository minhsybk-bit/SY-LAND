import { useEffect, useState } from "react";

export type SylandPlan = "Dùng thử" | "Go" | "Plus" | "Pro" | "Văn phòng" | "Đơn vị";
export type EntitlementSnapshot = {
  role: "admin" | "user" | "guest";
  plan: SylandPlan;
  seats: number;
  featurePercent: number;
  maxParcelsPerRun: number | null;
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
  if (role === "admin") return { role, plan: "Đơn vị", seats, featurePercent: 100, maxParcelsPerRun: null, fullTools: true, reason: "Quản trị viên SỸ LAND · không giới hạn" };

  let featurePercent = 25, accountLimit: number | null = 50, fullTools = false, reason = "Tài khoản dùng thử";
  if (plan === "Go") { featurePercent = 40; accountLimit = 80; reason = "Cá nhân Go · tối đa 80 thửa/lần"; }
  if (plan === "Plus") { featurePercent = 70; accountLimit = 140; reason = "Cá nhân Plus · tối đa 140 thửa/lần"; }
  if (plan === "Pro") { featurePercent = 100; accountLimit = 200; fullTools = true; reason = "Cá nhân Pro · đầy đủ công cụ, tối đa 200 thửa/lần"; }
  if (plan === "Văn phòng" && seats < 5) { featurePercent = 70; accountLimit = 140; reason = `Văn phòng ${seats} tài khoản · quyền tương đương Plus`; }
  if (plan === "Văn phòng" && seats >= 5) { featurePercent = 100; accountLimit = null; fullTools = true; reason = `Văn phòng ${seats} tài khoản · không giới hạn thửa/lần`; }
  if (plan === "Đơn vị") { featurePercent = 100; accountLimit = null; fullTools = true; reason = "Gói Đơn vị · không giới hạn thửa/lần"; }
  const licenseLimit = Number(input.licenseMaxParcels);
  const hasLicenseLimit = Number.isFinite(licenseLimit) && licenseLimit > 0;
  const maxParcelsPerRun = hasLicenseLimit ? (accountLimit == null ? Math.floor(licenseLimit) : Math.min(accountLimit, Math.floor(licenseLimit))) : accountLimit;
  if (hasLicenseLimit) reason += ` · mã bản quyền giới hạn ${Math.floor(licenseLimit)} thửa/lần`;
  return { role, plan, seats, featurePercent, maxParcelsPerRun, fullTools, reason };
}

export function publishEntitlements(value: EntitlementSnapshot) {
  localStorage.setItem(KEY, JSON.stringify(value));
  window.dispatchEvent(new CustomEvent(EVENT, { detail: value }));
}
function readEntitlements() {
  try { const value = JSON.parse(localStorage.getItem(KEY) || "null"); if (value?.role && "maxParcelsPerRun" in value) return value as EntitlementSnapshot; } catch {}
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
