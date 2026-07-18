"use client";

import { useMemo, useState } from "react";

export default function ValueCalculator() {
  const [files, setFiles] = useState(300);
  const [minutes, setMinutes] = useState(8);
  const [people, setPeople] = useState(1);

  const result = useMemo(() => {
    const savedMinutes = files * Math.max(0, minutes - 2);
    const hours = Math.round(savedMinutes / 60);
    const workDays = Math.round((hours / 8) * 10) / 10;
    const value = hours * 60000;
    const plan = people > 5 || files > 2500 ? "Đơn vị" : people > 1 || files > 300 ? "Văn phòng" : "Cá nhân";
    return { hours, workDays, value, plan };
  }, [files, minutes, people]);

  return (
    <section className="value-calculator" aria-labelledby="calculator-title">
      <div className="calculator-copy">
        <p className="section-kicker">Ước tính hiệu quả</p>
        <h3 id="calculator-title">Một tháng có thể tiết kiệm bao nhiêu?</h3>
        <p>Kết quả chỉ là ước tính tham khảo, dựa trên khoảng 2 phút xử lý mỗi hồ sơ bằng công cụ và giá trị thời gian 60.000đ/giờ.</p>
        <div className="calculator-fields">
          <label>Số hồ sơ mỗi tháng<input type="number" min="1" max="100000" value={files} onChange={(event) => setFiles(Math.max(1, Number(event.target.value) || 1))} /></label>
          <label>Phút xử lý thủ công/hồ sơ<input type="number" min="2" max="180" value={minutes} onChange={(event) => setMinutes(Math.max(2, Number(event.target.value) || 2))} /></label>
          <label>Số người sử dụng<input type="number" min="1" max="1000" value={people} onChange={(event) => setPeople(Math.max(1, Number(event.target.value) || 1))} /></label>
        </div>
      </div>
      <div className="calculator-result" aria-live="polite">
        <span>HIỆU QUẢ DỰ KIẾN / THÁNG</span>
        <div><strong>{result.hours.toLocaleString("vi-VN")}</strong><small>giờ tiết kiệm</small></div>
        <div><strong>{result.workDays.toLocaleString("vi-VN")}</strong><small>ngày công</small></div>
        <div><strong>{result.value.toLocaleString("vi-VN")}đ</strong><small>giá trị thời gian</small></div>
        <p>Gói đề xuất <b>{result.plan}</b></p>
        <a className="button button-primary" href="#goi-dich-vu">So sánh các gói <span aria-hidden="true">↑</span></a>
      </div>
    </section>
  );
}
