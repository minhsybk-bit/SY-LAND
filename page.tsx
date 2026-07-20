import FileProcessor from "./file-processor";
import ValueCalculator from "./value-calculator";
import TrialRequest from "./trial-request";
import ChatAssistant from "./chat-assistant";
import SoftwareRelease from "./software-release";
import PdfToolkit from "./pdf-toolkit";
import AccountPortal from "./account-portal";
import SystemCheck from "./system-check";
import SupportCenter from "./support-center";
import LegalCenter from "./legal-center";
import PrivacyCenter from "./privacy-center";
import PaymentCenter from "./payment-center";

const features = [
  {
    code: "01",
    title: "Soạn thảo văn bản hành chính",
    text: "Tạo dự thảo công văn, báo cáo, biên bản và tờ trình theo nội dung nghiệp vụ; hỗ trợ định dạng theo Nghị định 30.",
    tag: "Word · PDF",
  },
  {
    code: "02",
    title: "Đối chiếu dữ liệu địa chính",
    text: "So khớp mã xã, số tờ, số thửa, diện tích và tên chủ giữa nhiều bảng; đánh dấu rõ trường hợp thiếu hoặc sai khác.",
    tag: "Excel · Kiểm tra trùng",
  },
  {
    code: "03",
    title: "Chuẩn hóa hồ sơ hàng loạt",
    text: "Đổi tên PDF, Word theo cấu trúc thống nhất; phân loại hồ sơ có giấy, chưa có giấy và các nhóm tài liệu liên quan.",
    tag: "PDF · Tên tệp",
  },
  {
    code: "04",
    title: "Đọc và trích xuất thông tin",
    text: "Nhận diện dữ liệu trong văn bản scan, trích lục và biểu mẫu để giảm thao tác nhập lại thủ công.",
    tag: "OCR · Trích xuất",
  },
  {
    code: "05",
    title: "Tổng hợp báo cáo chuyên ngành",
    text: "Tạo bảng tổng hợp, tỷ lệ hoàn thành và danh sách cần xử lý tiếp từ dữ liệu đã được kiểm tra.",
    tag: "Báo cáo · Biểu đồ",
  },
  {
    code: "06",
    title: "Kiểm soát trước khi xuất",
    text: "Giữ chuyên viên ở vị trí quyết định: xem nguồn, rà soát cảnh báo và xác nhận kết quả trước khi sử dụng.",
    tag: "Có thể kiểm tra",
  },
];

const workflow = [
  {
    number: "01",
    title: "Tiếp nhận hồ sơ",
    text: "Chọn Word, PDF, Excel hoặc thư mục hồ sơ cần xử lý.",
    note: "Đầu vào linh hoạt",
  },
  {
    number: "02",
    title: "AI đọc và đối chiếu",
    text: "Trích xuất trường dữ liệu, so khớp quy tắc và chỉ ra điểm cần chú ý.",
    note: "Hiển thị nguồn kiểm tra",
  },
  {
    number: "03",
    title: "Chuyên viên rà soát",
    text: "Kiểm tra nội dung, sửa thông tin và xác nhận phương án xử lý.",
    note: "Con người quyết định",
  },
  {
    number: "04",
    title: "Xuất kết quả",
    text: "Tải văn bản, bảng tổng hợp hoặc bộ hồ sơ đã chuẩn hóa để tiếp tục nghiệp vụ.",
    note: "Kết quả dùng được",
  },
];

const plans = [
  {
    name: "Cá nhân",
    price: "199.000đ",
    cadence: "/ tháng",
    description: "Cho chuyên viên xử lý hồ sơ độc lập.",
    features: ["300 hồ sơ mỗi tháng", "Word, PDF, Excel và OCR", "Đối chiếu file tổng", "Xuất ZIP, CSV và Excel"],
    action: "Dùng thử công cụ",
    featured: false,
  },
  {
    name: "Văn phòng",
    price: "1.490.000đ",
    cadence: "/ tháng",
    description: "Cho nhóm nhỏ cần dùng chung quy trình.",
    features: ["Tối đa 5 người dùng", "2.500 hồ sơ mỗi tháng", "Kho hồ sơ và lịch sử xử lý", "Hỗ trợ thiết lập quy tắc riêng"],
    action: "Đăng ký quan tâm",
    featured: true,
  },
  {
    name: "Đơn vị",
    price: "Liên hệ",
    cadence: "theo nhu cầu",
    description: "Triển khai riêng, phân quyền và tích hợp dữ liệu.",
    features: ["Người dùng và hạn mức tùy chỉnh", "Máy chủ và vùng dữ liệu riêng", "Nhật ký thao tác, sao lưu", "Đào tạo và hỗ trợ vận hành"],
    action: "Xem lộ trình",
    featured: false,
  },
];

const softwareModules = [
  { code: "01", title: "Chuẩn hóa và đổi tên hồ sơ", status: "available", label: "Đang sử dụng", items: ["Đổi tên PDF, Word theo Mã xã · Số tờ · Số thửa", "Đổi hậu tố GT, TBXN, DDK hàng loạt", "Khai báo tỉnh, xã, thôn/xóm cũ và mới trên toàn quốc", "Chuẩn hóa tờ đất rừng 110000, 210000, 310000 cho RSX/RPH/RDD"] },
  { code: "02", title: "OCR và trích xuất dữ liệu", status: "available", label: "Đang sử dụng", items: ["OCR tiếng Việt cho PDF scan trên thiết bị", "Đọc mã xã, số tờ, số thửa và diện tích", "OCR tùy chọn tối đa 5 PDF trong một lượt hàng loạt", "Cho chuyên viên sửa và xác nhận trước khi xuất"] },
  { code: "03", title: "PDF và tài liệu văn phòng", status: "available", label: "Đang sử dụng", items: ["Tách, ghép, xoay, sắp xếp và xóa trang PDF", "Chuyển khổ A3 hoặc khổ khác sang A4", "PDF sang JPG/PNG và JPG/PNG/WebP sang PDF", "Nén ảnh, OCR, so sánh và làm sạch metadata"] },
  { code: "04", title: "Tổng hợp và kiểm tra Excel", status: "available", label: "Đang sử dụng", items: ["Làm sạch khoảng trắng, dòng trống và dòng trùng", "Tự nhận diện cột mã xã, số tờ và số thửa", "Phát hiện khóa địa chính thiếu, sai hoặc trùng", "Xuất dữ liệu sạch kèm sheet báo cáo kiểm tra"] },
  { code: "05", title: "Đối chiếu CSDL đất đai", status: "available", label: "Đang sử dụng", items: ["So khớp mã xã, số tờ, số thửa với Excel tổng", "Phát hiện trùng, thiếu và không khớp", "Lọc hoặc di chuyển hồ sơ không trùng", "Giữ nguyên dữ liệu gốc và xuất báo cáo mới"] },
  { code: "06", title: "Phân loại và vận hành", status: "available", label: "Đang sử dụng", items: ["Tự tạo cấu trúc thư mục Có GCN/Chưa có giấy trong ZIP", "Phân nhóm GT, TBXN, DDK và kèm nhật ký CSV", "Tải bộ cài Windows từ GitHub Release", "Kiểm tra SHA-256 và tự hiển thị phiên bản mới"] },
];

export default function Home() {
  return (
    <main>
      <section className="hero" aria-labelledby="hero-title">
        <div className="parcel-lines" aria-hidden="true" />

        <header className="site-header page-shell">
          <a className="brand" href="#top" aria-label="SỸ LAND - Trang chủ">
            <span className="brand-mark" aria-hidden="true"><i /><i /><i /><i /></span>
            <span>SỸ LAND</span>
          </a>

          <nav className="main-nav" aria-label="Điều hướng chính">
            <a className="nav-direct" href="#top">Tổng quan</a>
            <details className="nav-group">
              <summary>Công cụ <span aria-hidden="true">⌄</span></summary>
              <div className="nav-dropdown">
                <a href="#minh-hoa"><b>Xử lý tệp</b><small>Word, PDF, Excel và dữ liệu địa chính</small></a>
                <a href="#cong-cu-pdf"><b>Bộ công cụ PDF</b><small>Tách, ghép, kiểm tra, OCR và tối ưu</small></a>
                <a href="#chuc-nang-phan-mem"><b>Phần mềm SỸ LAND</b><small>Các chức năng dành cho Windows</small></a>
              </div>
            </details>
            <details className="nav-group">
              <summary>Sản phẩm <span aria-hidden="true">⌄</span></summary>
              <div className="nav-dropdown">
                <a href="#tai-phan-mem"><b>Tải phần mềm</b><small>Phiên bản mới và hướng dẫn cập nhật</small></a>
                <a href="#tai-khoan"><b>Tài khoản</b><small>Đăng ký, đăng nhập và quản lý bản quyền</small></a>
              </div>
            </details>
            <details className="nav-group">
              <summary>Dịch vụ <span aria-hidden="true">⌄</span></summary>
              <div className="nav-dropdown nav-dropdown-right">
                <a href="#goi-dich-vu"><b>Gói sử dụng</b><small>So sánh quyền lợi theo nhu cầu</small></a>
                <a href="#thanh-toan"><b>Đăng ký và thanh toán</b><small>Chọn gói, quét QR và nhận mã</small></a>
              </div>
            </details>
            <a className="nav-direct" href="#gioi-thieu">Giới thiệu</a>
          </nav>

          <div className="header-tools">
            <a className="header-action" href="#minh-hoa">Mở công cụ</a>
            <details className="mobile-menu">
              <summary aria-label="Mở menu"><span /><span /><span /></summary>
              <div>
                <a href="#top"><b>Tổng quan</b><small>Trang chủ SỸ LAND</small></a>
                <p>Công cụ</p>
                <a href="#minh-hoa">Xử lý Word, PDF, Excel</a>
                <a href="#cong-cu-pdf">Bộ công cụ PDF</a>
                <a href="#chuc-nang-phan-mem">Phần mềm SỸ LAND</a>
                <p>Tài khoản và dịch vụ</p>
                <a href="#tai-phan-mem">Tải phần mềm</a>
                <a href="#tai-khoan">Tài khoản</a>
                <a href="#goi-dich-vu">Gói sử dụng</a>
                <a href="#thanh-toan">Đăng ký và thanh toán</a>
                <a href="#gioi-thieu">Giới thiệu</a>
              </div>
            </details>
          </div>
        </header>

        <div className="hero-grid page-shell" id="top">
          <div className="hero-copy">
            <p className="eyebrow"><span className="ai-glyph" aria-hidden="true">AI</span>Tiện ích hỗ trợ làm sạch CSDL đất đai SỸ LAND</p>
            <h1 id="hero-title">Quản lý đất đai<br />rõ ràng hơn.<br /><span>Nhanh hơn.</span></h1>
            <p className="hero-description">
              Hỗ trợ soạn thảo văn bản hành chính, xử lý dữ liệu địa chính,
              chuẩn hóa tên tệp và tổng hợp báo cáo — trong một quy trình rõ
              ràng, có thể kiểm tra.
            </p>
            <div className="hero-actions">
              <a className="button button-primary" href="#quy-trinh">Khám phá quy trình <span aria-hidden="true">→</span></a>
              <a className="text-link" href="#tinh-nang">Xem tính năng <span aria-hidden="true">›</span></a>
            </div>
          </div>

          <div className="workflow-preview" aria-label="Minh họa quy trình xử lý hồ sơ">
            <div className="flow-step"><span className="step-number">1</span><strong>Hồ sơ đầu vào</strong></div>
            <div className="flow-step"><span className="step-number">2</span><strong>AI đối chiếu</strong></div>
            <div className="flow-step"><span className="step-number">3</span><strong>Báo cáo hoàn thiện</strong></div>

            <article className="preview-card map-card">
              <div className="map-north">N<br /><span>▲</span></div>
              <div className="parcel-map" aria-hidden="true">
                <span className="lot lot-a">233</span><span className="lot lot-b">234</span><span className="lot lot-c">237</span>
                <span className="parcel-shape"><b>235</b><small>1.253,6 m²</small></span>
              </div>
              <div className="file-strip"><span>▱</span> Trích lục thửa đất.pdf</div>
            </article>

            <span className="flow-arrow arrow-one" aria-hidden="true">→</span>

            <article className="preview-card document-card">
              <div className="document-type">W</div>
              <h2>DỰ THẢO VĂN BẢN</h2>
              <div className="document-lines" aria-hidden="true"><i /><i /><i className="short" /><i /><i className="checked" /><i /><i className="short" /></div>
              <div className="file-strip"><span>▤</span> Dự thảo quyết định.docx</div>
            </article>

            <span className="flow-arrow arrow-two" aria-hidden="true">→</span>

            <article className="preview-card report-card">
              <div className="checkmark" aria-hidden="true">✓</div>
              <h2>ĐỐI CHIẾU HỢP LỆ</h2>
              <div className="report-lines" aria-hidden="true"><i /><i /><i /></div>
              <div className="file-strip"><span>▧</span> Báo cáo tổng hợp.pdf</div>
            </article>
          </div>
        </div>
      </section>

      <section className="features section-pad" id="tinh-nang" aria-labelledby="features-title">
        <div className="page-shell">
          <div className="section-heading split-heading">
            <div><p className="section-kicker">Năng lực cốt lõi</p><h2 id="features-title">Một trợ lý cho những việc<br />thường ngày nhưng tốn thời gian.</h2></div>
            <p>Thiết kế xoay quanh quy trình thực tế của chuyên viên: dữ liệu đầu vào đa dạng, cần kiểm tra kỹ và kết quả phải tiếp tục chỉnh sửa được.</p>
          </div>
          <div className="feature-grid">
            {features.map((feature) => (
              <article className="feature-card" key={feature.code}>
                <div className="feature-top"><span>{feature.code}</span><span className="feature-arrow" aria-hidden="true">↗</span></div>
                <h3>{feature.title}</h3>
                <p>{feature.text}</p>
                <small>{feature.tag}</small>
              </article>
            ))}
          </div>
        </div>
      </section>

      <section className="process section-pad" id="quy-trinh" aria-labelledby="process-title">
        <div className="page-shell process-shell">
          <div className="process-intro">
            <p className="section-kicker">Quy trình làm việc</p>
            <h2 id="process-title">AI hỗ trợ từng bước.<br />Chuyên viên giữ quyền quyết định.</h2>
            <p>Hệ thống không biến hồ sơ thành “hộp đen”. Mỗi kết quả đều đi cùng trạng thái, điểm cần kiểm tra và bước xác nhận trước khi xuất.</p>
            <a className="text-link" href="#minh-hoa">Mở công cụ xử lý tệp <span aria-hidden="true">›</span></a>
          </div>
          <div className="process-list">
            {workflow.map((step) => (
              <article className="process-item" key={step.number}>
                <span className="process-number">{step.number}</span>
                <div><h3>{step.title}</h3><p>{step.text}</p></div>
                <small>{step.note}</small>
              </article>
            ))}
          </div>
        </div>
      </section>

      <section className="demo section-pad" id="minh-hoa" aria-labelledby="demo-title">
        <div className="page-shell">
          <div className="section-heading demo-heading">
            <div><p className="section-kicker">Công cụ xử lý tệp</p><h2 id="demo-title">Đọc Word, PDF, Excel thật.<br />Kiểm tra trước khi sử dụng.</h2></div>
            <p className="prototype-note live-note"><span aria-hidden="true">●</span> Công cụ đang hoạt động · Xử lý cục bộ</p>
          </div>
          <FileProcessor />
          <PdfToolkit />

          <section className="software-capabilities" id="chuc-nang-phan-mem" aria-labelledby="software-capabilities-title">
            <div className="section-heading split-heading">
              <div><p className="section-kicker">Phương án xây dựng phần mềm</p><h2 id="software-capabilities-title">Một bộ công cụ cho toàn bộ<br />quy trình làm sạch CSDL đất đai.</h2></div>
              <p>Các chức năng được công bố theo đúng trạng thái phát triển. Chức năng “đang hoàn thiện” và “theo lộ trình” sẽ chỉ phát hành sau khi kiểm thử đủ độ ổn định và chính xác.</p>
            </div>
            <div className="software-module-grid">
              {softwareModules.map((module) => <article className="software-module" key={module.code}>
                <div className="module-top"><span>{module.code}</span><span className={`module-status ${module.status}`}>{module.label}</span></div>
                <h3>{module.title}</h3>
                <ul>{module.items.map((item) => <li key={item}><span aria-hidden="true">✓</span>{item}</li>)}</ul>
              </article>)}
            </div>
            <div className="software-principles"><strong>Nguyên tắc phát triển</strong><span>Không tự bịa dữ liệu</span><span>Không ghi đè tệp gốc</span><span>Có tiến độ và nhật ký</span><span>Chuyên viên xác nhận kết quả</span></div>
          </section>

          <SoftwareRelease />
          <SystemCheck />
          <AccountPortal />
          <SupportCenter />
          <LegalCenter />
          <PrivacyCenter />

          <section className="pricing" id="goi-dich-vu" aria-labelledby="pricing-title">
            <div className="section-heading split-heading pricing-heading">
              <div><p className="section-kicker">MVP thương mại</p><h2 id="pricing-title">Bắt đầu nhỏ.<br />Mở rộng khi cần.</h2></div>
              <p>Chọn theo số người và số hồ sơ thực tế. Có thể thanh toán từng tháng hoặc tiết kiệm hơn khi chọn chu kỳ 6–12 tháng; quy mô lớn được báo giá riêng.</p>
            </div>
            <div className="pricing-grid">
              {plans.map((plan) => <article className={`pricing-card ${plan.featured ? "featured" : ""}`} key={plan.name}>
                {plan.featured && <span className="plan-badge">Phù hợp để bắt đầu</span>}
                <p className="plan-name">{plan.name}</p>
                <div className="plan-price"><strong>{plan.price}</strong><span>{plan.cadence}</span></div>
                <p className="plan-description">{plan.description}</p>
                <ul>{plan.features.map((feature) => <li key={feature}><span aria-hidden="true">✓</span>{feature}</li>)}</ul>
                <a className={`button ${plan.featured ? "button-primary" : "plan-button"}`} href={plan.name === "Đơn vị" ? "#tu-van" : "#thanh-toan"}>{plan.name === "Đơn vị" ? "Nhận báo giá riêng" : "Chọn gói này"}<span aria-hidden="true">→</span></a>
              </article>)}
            </div>
            <PaymentCenter />
            <ValueCalculator />
            <TrialRequest />
            <div className="commercial-roadmap" id="lo-trinh">
              <div><span>01</span><strong>Công cụ đang hoạt động</strong><small>Xử lý hồ sơ cục bộ và giữ tệp gốc</small></div>
              <div><span>02</span><strong>Tài khoản và bản quyền chung</strong><small>Đăng nhập trên website và phần mềm cài đặt</small></div>
              <div><span>03</span><strong>Thanh toán có đối soát</strong><small>QR theo đơn; tự cấp mã khi admin xác nhận</small></div>
            </div>
          </section>

          <div className="vision" id="gioi-thieu">
            <div><p className="section-kicker">Định hướng dự án</p><h2>Giảm thao tác lặp lại.<br />Tăng thời gian cho chuyên môn.</h2></div>
            <div className="vision-copy"><p>Tiện ích hỗ trợ làm sạch CSDL đất đai SỸ LAND được định hướng như một lớp hỗ trợ an toàn cho nghiệp vụ hành chính: làm nhanh phần tổng hợp, giữ rõ nguồn dữ liệu và luôn dành bước rà soát cuối cho chuyên viên.</p><a className="button button-light" href="#top">Trở lại đầu trang <span aria-hidden="true">↑</span></a></div>
          </div>
        </div>
      </section>

      <footer className="site-footer">
        <div className="page-shell"><a className="brand footer-brand" href="#top"><span className="brand-mark" aria-hidden="true"><i /><i /><i /><i /></span><span>SỸ LAND</span></a><p>Tiện ích hỗ trợ làm sạch CSDL đất đai SỸ LAND</p><span>2026 · Phiên bản chính thức</span></div>
      </footer>
      <ChatAssistant />
    </main>
  );
}
