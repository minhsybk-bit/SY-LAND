"use client";

import { FormEvent, useMemo, useRef, useState } from "react";

type Message = { id: number; role: "bot" | "user"; text: string };

const QUICK_QUESTIONS = ["Tải phần mềm", "Hướng dẫn cài đặt", "Xử lý PDF, Word, Excel", "Bảng giá"];

function answerQuestion(question: string) {
  const q = question.toLocaleLowerCase("vi-VN").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d");
  if (/tai|download|phan mem|11\.1/.test(q)) return "Bạn có thể tải SỸ LAND 11.1.0 tại mục ‘Tải phần mềm’. Bộ cài dành cho Windows 10/11, dung lượng khoảng 130,6 MB.";
  if (/cai dat|setup|windows|smart.?screen/.test(q)) return "Tải tệp SYLAND_Setup_11.1.0.exe, kiểm tra đúng tên và mã SHA-256 trên website, sau đó mở bộ cài. Windows có thể yêu cầu bạn xác nhận trước khi tiếp tục.";
  if (/pdf|word|excel|ocr|tep|file/.test(q)) return "SỸ LAND hỗ trợ DOCX, PDF, XLSX, XLS và CSV; có OCR tiếng Việt cho PDF scan, trích số tờ, số thửa, diện tích và đối chiếu Excel tổng.";
  if (/gia|phi|goi|ban quyen|mua/.test(q)) return "Website đang công bố mức giá dự kiến cho gói Cá nhân, Văn phòng và Đơn vị. Để nhận tư vấn bản quyền phù hợp, hãy liên hệ anh Nguyễn Minh Sỹ qua Zalo 0972560335.";
  if (/bao mat|rieng tu|du lieu|may chu/.test(q)) return "Công cụ web hiện xử lý tệp cục bộ trên thiết bị và không tự động tải nội dung hồ sơ lên máy chủ. Người dùng vẫn cần rà soát kết quả trước khi sử dụng chính thức.";
  if (/ma xa|02140|02143|02146|thon/.test(q)) return "Quy tắc đang hỗ trợ mã xã cũ 02140 (Văn Lang), 02143 (Lương Thượng) và 02146 (Kim Hỷ), ưu tiên nhận diện theo tên thôn và địa chỉ thửa đất.";
  if (/110000|210000|rsx|rph|rdd|lam nghiep/.test(q)) return "Các loại đất RSX, RPH, RDD có thể dùng số tờ lâm nghiệp 110000 hoặc 210000. SỸ LAND giữ nguyên các số tờ 6 chữ số này và không coi là lỗi.";
  if (/lien he|zalo|facebook|tu van|sy/.test(q)) return "Bạn có thể liên hệ anh Nguyễn Minh Sỹ qua Zalo 0972560335 hoặc Facebook facebook.com/nguyensybk để được hỗ trợ trực tiếp.";
  if (/xin chao|chao|hello|hi\b/.test(q)) return "Xin chào! Tôi là trợ lý tự động của SỸ LAND. Bạn có thể hỏi về tải phần mềm, cài đặt, xử lý hồ sơ, bảng giá hoặc quyền riêng tư.";
  return "Tôi chưa có câu trả lời chắc chắn cho nội dung này. Bạn vui lòng liên hệ anh Nguyễn Minh Sỹ qua Zalo 0972560335 hoặc Facebook để được tư vấn chính xác.";
}

export default function ChatAssistant() {
  const [open, setOpen] = useState(false);
  const [input, setInput] = useState("");
  const [messages, setMessages] = useState<Message[]>([{ id: 1, role: "bot", text: "Xin chào! Tôi là trợ lý SỸ LAND. Tôi có thể giúp gì cho bạn?" }]);
  const nextId = useRef(2);
  const latest = useMemo(() => messages.slice(-8), [messages]);

  function ask(text: string) {
    const value = text.trim();
    if (!value) return;
    const userId = nextId.current++;
    const botId = nextId.current++;
    setMessages((current) => [...current, { id: userId, role: "user", text: value }, { id: botId, role: "bot", text: answerQuestion(value) }]);
    setInput("");
  }

  function submit(event: FormEvent) { event.preventDefault(); ask(input); }

  return (
    <div className="support-hub">
      <div className="social-links" aria-label="Liên hệ SỸ LAND">
        <a className="social-button zalo-button" href="https://zalo.me/0972560335" target="_blank" rel="noreferrer" aria-label="Liên hệ Zalo Nguyễn Minh Sỹ"><span>Z</span><b>Zalo</b></a>
        <a className="social-button facebook-button" href="https://www.facebook.com/nguyensybk/" target="_blank" rel="noreferrer" aria-label="Liên hệ Facebook Nguyễn Minh Sỹ"><span>f</span><b>Facebook</b></a>
      </div>

      {open && <section className="chat-panel" aria-label="Trợ lý tự động SỸ LAND">
        <header><div><span className="chat-avatar">AI</span><div><strong>Trợ lý SỸ LAND</strong><small><i /> Trả lời tự động</small></div></div><button type="button" onClick={() => setOpen(false)} aria-label="Đóng chatbot">×</button></header>
        <div className="chat-privacy">Không gửi tệp hoặc thông tin cá nhân nhạy cảm vào khung chat.</div>
        <div className="chat-messages" aria-live="polite">{latest.map((message) => <p className={message.role} key={message.id}>{message.text}</p>)}</div>
        <div className="quick-questions">{QUICK_QUESTIONS.map((question) => <button type="button" key={question} onClick={() => ask(question)}>{question}</button>)}</div>
        <form onSubmit={submit}><input value={input} onChange={(event) => setInput(event.target.value)} maxLength={300} placeholder="Nhập câu hỏi…" aria-label="Nhập câu hỏi cho trợ lý" /><button type="submit" aria-label="Gửi câu hỏi">➜</button></form>
        <footer><span>Cần hỗ trợ thêm?</span><a href="https://zalo.me/0972560335" target="_blank" rel="noreferrer">Chat Zalo với Nguyễn Minh Sỹ</a></footer>
      </section>}

      <button className={`chat-launcher ${open ? "active" : ""}`} type="button" onClick={() => setOpen((value) => !value)} aria-expanded={open} aria-label={open ? "Đóng trợ lý SỸ LAND" : "Mở trợ lý SỸ LAND"}><span>{open ? "×" : "AI"}</span><b>{open ? "Đóng" : "Hỏi SỸ LAND"}</b></button>
    </div>
  );
}
