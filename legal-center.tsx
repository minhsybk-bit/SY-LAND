"use client";

import { useState } from "react";

export const TERMS_VERSION = "2026-07-19";

export default function LegalCenter() {
  const [open, setOpen] = useState<"privacy" | "terms" | "data" | null>(null);
  return (
    <section className="legal-center" aria-labelledby="legal-title">
      <div className="legal-heading"><div><p className="section-kicker">Minh bạch và an toàn</p><h2 id="legal-title">Dữ liệu thuộc quyền kiểm soát của người dùng.</h2></div><p>SỸ LAND công bố rõ dữ liệu nào được xử lý trên máy, dữ liệu nào được gửi lên máy chủ và trách nhiệm kiểm tra kết quả nghiệp vụ.</p></div>
      <div className="legal-grid">
        <article><span>01</span><h3>Chính sách bảo mật</h3><p>Tài khoản, yêu cầu tư vấn và trạng thái bản quyền được lưu trên máy chủ. Hồ sơ Word, PDF và Excel mặc định xử lý cục bộ trong trình duyệt hoặc phần mềm.</p><button type="button" onClick={() => setOpen(open === "privacy" ? null : "privacy")}>Xem chính sách</button></article>
        <article><span>02</span><h3>Điều khoản sử dụng</h3><p>SỸ LAND là công cụ hỗ trợ. Người dùng và chuyên viên chịu trách nhiệm kiểm tra nguồn, cảnh báo và kết quả trước khi sử dụng trong hồ sơ chính thức.</p><button type="button" onClick={() => setOpen(open === "terms" ? null : "terms")}>Xem điều khoản</button></article>
        <article><span>03</span><h3>Xử lý và lưu giữ dữ liệu</h3><p>Không tự động tải GCN, CCCD hoặc hồ sơ địa chính lên máy chủ. Phiếu hỗ trợ không được dùng để gửi tài liệu chứa dữ liệu cá nhân.</p><button type="button" onClick={() => setOpen(open === "data" ? null : "data")}>Xem nguyên tắc</button></article>
      </div>
      {open && <div className="legal-detail" role="region" aria-live="polite"><button type="button" onClick={() => setOpen(null)} aria-label="Đóng">×</button>{open === "privacy" ? <><h3>Chính sách bảo mật SỸ LAND</h3><ul><li>Thu thập tối thiểu thông tin tài khoản, email, tên hiển thị, bản quyền, thiết bị đã băm và yêu cầu hỗ trợ.</li><li>Không bán dữ liệu cá nhân hoặc dùng hồ sơ nghiệp vụ để quảng cáo.</li><li>Mã thiết bị là giá trị băm; không công khai serial phần cứng thô.</li><li>Người dùng có thể yêu cầu chỉnh sửa hoặc xóa thông tin tài khoản theo quy trình hỗ trợ.</li><li>Không đưa service role key vào website hoặc phần mềm người dùng.</li></ul></> : open === "terms" ? <><h3>Điều khoản sử dụng</h3><ul><li>Không sử dụng công cụ để sửa đổi, giả mạo hoặc che giấu thông tin hồ sơ.</li><li>Kết quả OCR và đối chiếu có thể cần sửa; chuyên viên phải xác nhận trước khi xuất.</li><li>Bản quyền được cấp theo email, thời hạn, gói và số thiết bị.</li><li>Không chia sẻ tài khoản hoặc vượt giới hạn thiết bị đã mua.</li><li>SỸ LAND có thể khóa bản quyền vi phạm nhưng không tự xóa hồ sơ cục bộ của người dùng.</li></ul></> : <><h3>Nguyên tắc xử lý dữ liệu hồ sơ</h3><ul><li>Word, PDF, Excel và OCR trên web mặc định thực hiện ngay trên thiết bị.</li><li>Chỉ thông tin tài khoản, bản quyền, yêu cầu tư vấn và hỗ trợ được gửi lên Supabase.</li><li>Không tải hồ sơ có GCN, CCCD hoặc thông tin nhạy cảm vào biểu mẫu hỗ trợ.</li><li>Sao lưu tệp gốc trước khi đổi tên, tách, gộp hoặc làm sạch dữ liệu hàng loạt.</li><li>Nhật ký quản trị chỉ dành cho admin và được bảo vệ bằng Row Level Security.</li></ul></>}<small>Phiên bản điều khoản: {TERMS_VERSION}</small></div>}
    </section>
  );
}
