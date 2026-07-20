// Điền đúng thông tin nhận tiền trước khi mở thanh toán công khai.
// bankBin: mã BIN ngân hàng theo VietQR (ví dụ MB = 970422, Vietcombank = 970436).
export const PAYMENT_CONFIG = {
  bankBin: "970418",
  bankName: "BIDV - Ngân hàng TMCP Đầu tư và Phát triển Việt Nam",
  accountNumber: "3950549732",
  accountName: "NGUYEN MINH SY",
  supportPhone: "0972560335",
};

export const PAYMENT_PLANS = [
  { id: "personal", name: "Cá nhân", amount: 199000, months: 1, maxDevices: 1, description: "01 người dùng · 01 thiết bị · 300 hồ sơ/tháng" },
  { id: "office", name: "Văn phòng", amount: 1490000, months: 1, maxDevices: 5, description: "Tối đa 05 người dùng · 05 thiết bị · 2.500 hồ sơ/tháng" },
] as const;
