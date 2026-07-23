// Thông tin nhận tiền không đặt trong mã nguồn/GitHub. Quản trị viên nhập
// một lần trên website và dữ liệu được lưu trong bảng payment_settings.
export type PaymentConfig = {
  bankBin: string;
  bankName: string;
  accountNumber: string;
  accountName: string;
  supportPhone: string;
};
export const EMPTY_PAYMENT_CONFIG: PaymentConfig = {
  bankBin: "", bankName: "", accountNumber: "", accountName: "", supportPhone: "",
};

export const PAYMENT_PLANS = [
  { id: "go", name: "Go", amount: 99000, months: 1, maxDevices: 1, maxParcels: 80, featurePercent: 40, description: "01 tài khoản · 80 thửa/lượt · 40% bộ công cụ" },
  { id: "plus", name: "Plus", amount: 199000, months: 1, maxDevices: 1, maxParcels: 140, featurePercent: 70, description: "01 tài khoản · 140 thửa/lượt · 70% bộ công cụ" },
  { id: "pro", name: "Pro", amount: 399000, months: 1, maxDevices: 1, maxParcels: 200, featurePercent: 100, description: "01 tài khoản · 200 thửa/lượt · đầy đủ công cụ" },
  { id: "office", name: "Văn phòng", amount: 298000, months: 1, maxDevices: 5, maxParcels: null, featurePercent: 100, perSeat: true, description: "298.000đ/tài khoản · từ 5 tài khoản: đầy đủ, không giới hạn/lượt" },
] as const;

export const BILLING_CYCLES = [
  { months: 1, label: "Hàng tháng", discount: 0 },
  { months: 6, label: "06 tháng", discount: 10 },
  { months: 12, label: "12 tháng", discount: 20 },
] as const;

export function planTotal(baseAmount: number, months: number) {
  const discount = months === 12 ? 20 : months === 6 ? 10 : 0;
  return Math.round(baseAmount * months * (100 - discount) / 100);
}
