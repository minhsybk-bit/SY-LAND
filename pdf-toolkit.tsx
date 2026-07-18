"use client";

import { ChangeEvent, useMemo, useRef, useState } from "react";

type Mode = "split" | "merge" | "rotate" | "organize" | "resize" | "annotate" | "optimize" | "compare" | "sanitize" | "ocr" | "compress";
type PagePreview = { page: number; image: string };
type CompareResult = { pagesA: number; pagesB: number; changedPages: number[]; samePages: number; textCoverage: number };

function downloadBlob(bytes: Uint8Array, name: string, type = "application/pdf") {
  const blob = new Blob([bytes as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadText(content: string, name: string, type = "text/plain;charset=utf-8") {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const anchor = document.createElement("a"); anchor.href = url; anchor.download = name; anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function baseName(name: string) { return name.replace(/\.pdf$/i, "").replace(/[^a-zA-Z0-9À-ỹ_-]+/g, "_"); }

function parseRanges(input: string, total: number) {
  const groups: number[][] = [];
  input.split(",").map((part) => part.trim()).filter(Boolean).forEach((part) => {
    const match = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!match) throw new Error(`Khoảng trang “${part}” không hợp lệ.`);
    const start = Number(match[1]);
    const end = Number(match[2] || match[1]);
    if (start < 1 || end < start || end > total) throw new Error(`Khoảng ${part} nằm ngoài 1–${total}.`);
    groups.push(Array.from({ length: end - start + 1 }, (_, index) => start - 1 + index));
  });
  if (!groups.length) throw new Error("Hãy nhập trang cần tách, ví dụ 1-3, 5, 8-10.");
  return groups;
}

export default function PdfToolkit() {
  const [mode, setMode] = useState<Mode>("split");
  const [files, setFiles] = useState<File[]>([]);
  const [ranges, setRanges] = useState("1");
  const [splitKind, setSplitKind] = useState<"ranges" | "each" | "selected">("ranges");
  const [angle, setAngle] = useState(90);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [previews, setPreviews] = useState<PagePreview[]>([]);
  const [removedPages, setRemovedPages] = useState<Set<number>>(new Set());
  const [watermark, setWatermark] = useState("SY LAND");
  const [addWatermark, setAddWatermark] = useState(true);
  const [addNumbers, setAddNumbers] = useState(true);
  const [numberStart, setNumberStart] = useState(1);
  const [numberPosition, setNumberPosition] = useState<"bottom" | "top">("bottom");
  const [compareResult, setCompareResult] = useState<CompareResult | null>(null);
  const [compressQuality, setCompressQuality] = useState<"small" | "balanced" | "clear">("balanced");
  const inputRef = useRef<HTMLInputElement>(null);
  const totalSize = useMemo(() => files.reduce((sum, file) => sum + file.size, 0), [files]);

  async function renderPreviews(file: File) {
    setNotice("Đang tạo ảnh xem trước…");
    try {
      const pdfjs = await import("pdfjs-dist");
      const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
      pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
      const document = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
      const next: PagePreview[] = [];
      const limit = Math.min(document.numPages, 80);
      for (let pageNumber = 1; pageNumber <= limit; pageNumber++) {
        const page = await document.getPage(pageNumber);
        const base = page.getViewport({ scale: 1 });
        const viewport = page.getViewport({ scale: Math.min(.35, 190 / base.width) });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
        await page.render({ canvas, canvasContext: canvas.getContext("2d")!, viewport }).promise;
        next.push({ page: pageNumber, image: canvas.toDataURL("image/jpeg", .72) });
      }
      setPreviews(next); setRemovedPages(new Set());
      setNotice(document.numPages > 80 ? `Hiển thị 80/${document.numPages} trang đầu. Tệp quá dài để sắp xếp trực quan toàn bộ.` : `Đã tải ${document.numPages} trang. Chọn trang cần xóa hoặc thay đổi thứ tự.`);
    } catch (reason) { console.error(reason); setNotice("Không tạo được ảnh xem trước của PDF này."); }
  }

  async function chooseFiles(event: ChangeEvent<HTMLInputElement>) {
    const selected = Array.from(event.target.files || []).filter((file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"));
    setFiles(mode === "merge" ? selected.slice(0, 30) : mode === "compare" ? selected.slice(0, 2) : selected.slice(0, 1));
    setNotice(selected.length ? "" : "Hãy chọn tệp PDF hợp lệ.");
    event.target.value = "";
    if (mode === "organize" && selected[0]) await renderPreviews(selected[0]);
  }

  function move(index: number, direction: -1 | 1) {
    setFiles((current) => {
      const target = index + direction;
      if (target < 0 || target >= current.length) return current;
      const next = [...current];
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  }

  function changeMode(next: Mode) { setMode(next); setFiles([]); setPreviews([]); setRemovedPages(new Set()); setCompareResult(null); setNotice(""); }

  function movePage(index: number, direction: -1 | 1) { setPreviews((current) => { const target = index + direction; if (target < 0 || target >= current.length) return current; const next = [...current]; [next[index], next[target]] = [next[target], next[index]]; return next; }); }
  function toggleRemoved(page: number) { setRemovedPages((current) => { const next = new Set(current); if (next.has(page)) next.delete(page); else next.add(page); return next; }); }

  async function run() {
    if (!files.length) { setNotice("Hãy chọn tệp PDF trước khi xử lý."); return; }
    setBusy(true); setNotice("Đang xử lý trên thiết bị…");
    try {
      if (mode === "compare") {
        if (files.length !== 2) throw new Error("Hãy chọn đúng hai tệp PDF để so sánh.");
        const pdfjs = await import("pdfjs-dist");
        const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
        pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
        const documents = await Promise.all(files.map(async (file) => pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise));
        const texts: string[][] = [];
        for (const document of documents) {
          const pages: string[] = [];
          for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber++) {
            const content = await (await document.getPage(pageNumber)).getTextContent();
            pages.push(content.items.map((item) => "str" in item ? item.str : "").join(" ").replace(/\s+/g, " ").trim().toLocaleLowerCase("vi"));
          }
          texts.push(pages);
        }
        const maximum = Math.max(texts[0].length, texts[1].length);
        const changedPages: number[] = [];
        let samePages = 0; let pagesWithText = 0;
        for (let index = 0; index < maximum; index++) {
          const left = texts[0][index] || ""; const right = texts[1][index] || "";
          if (left || right) pagesWithText++;
          if (left === right && index < texts[0].length && index < texts[1].length) samePages++; else changedPages.push(index + 1);
        }
        setCompareResult({ pagesA: texts[0].length, pagesB: texts[1].length, changedPages, samePages, textCoverage: maximum ? Math.round(pagesWithText / maximum * 100) : 0 });
        setNotice(changedPages.length ? `Phát hiện ${changedPages.length} trang khác nhau hoặc bị thêm/bớt.` : "Hai tệp có nội dung văn bản trùng khớp theo từng trang.");
        return;
      }
      if (mode === "ocr") {
        const pdfjs = await import("pdfjs-dist");
        const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
        const { createWorker } = await import("tesseract.js");
        pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
        const source = await pdfjs.getDocument({ data: new Uint8Array(await files[0].arrayBuffer()) }).promise;
        if (source.numPages > 40) throw new Error("OCR trực tiếp hỗ trợ tối đa 40 trang mỗi lượt.");
        const worker = await createWorker("vie");
        const output: string[] = [`OCR SỸ LAND`, `Tệp nguồn: ${files[0].name}`, `Số trang: ${source.numPages}`, ""];
        try {
          for (let pageNumber = 1; pageNumber <= source.numPages; pageNumber++) {
            setNotice(`Đang OCR trang ${pageNumber}/${source.numPages}…`);
            const page = await source.getPage(pageNumber);
            const base = page.getViewport({ scale: 1 });
            const viewport = page.getViewport({ scale: Math.min(2.2, 2200 / Math.max(base.width, base.height)) });
            const canvas = document.createElement("canvas"); canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
            await page.render({ canvas, canvasContext: canvas.getContext("2d", { willReadFrequently: true })!, viewport }).promise;
            const result = await worker.recognize(canvas);
            output.push(`===== TRANG ${pageNumber} =====`, result.data.text.trim(), "");
          }
        } finally { await worker.terminate(); }
        downloadText(output.join("\n"), `${baseName(files[0].name)}_OCR.txt`);
        setNotice(`Đã OCR ${source.numPages} trang và xuất tệp văn bản. Cần kiểm tra lại tên riêng, số thửa và diện tích.`);
        return;
      }
      const { PDFDocument, StandardFonts, degrees, rgb } = await import("pdf-lib");
      if (mode === "merge") {
        const output = await PDFDocument.create();
        for (const file of files) {
          const source = await PDFDocument.load(await file.arrayBuffer(), { ignoreEncryption: false });
          const pages = await output.copyPages(source, source.getPageIndices());
          pages.forEach((page) => output.addPage(page));
        }
        downloadBlob(await output.save(), `SYLAND_NOI_${files.length}_TEP.pdf`);
        setNotice(`Đã nối ${files.length} tệp thành công.`);
      } else if (mode === "rotate") {
        const document = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false });
        document.getPages().forEach((page) => page.setRotation(degrees((page.getRotation().angle + angle) % 360)));
        downloadBlob(await document.save(), `${baseName(files[0].name)}_XOAY_${angle}.pdf`);
        setNotice(`Đã xoay ${document.getPageCount()} trang ${angle}°.`);
      } else if (mode === "organize") {
        const source = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false });
        const order = previews.filter((item) => !removedPages.has(item.page)).map((item) => item.page - 1);
        if (!order.length) throw new Error("Không còn trang để xuất.");
        const output = await PDFDocument.create();
        (await output.copyPages(source, order)).forEach((page) => output.addPage(page));
        downloadBlob(await output.save(), `${baseName(files[0].name)}_SAP_XEP.pdf`);
        setNotice(`Đã xuất ${order.length} trang; loại bỏ ${removedPages.size} trang.`);
      } else if (mode === "resize") {
        const source = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false });
        const output = await PDFDocument.create();
        for (const sourcePage of source.getPages()) {
          const embedded = await output.embedPage(sourcePage);
          const landscape = sourcePage.getWidth() > sourcePage.getHeight();
          const targetWidth = landscape ? 841.89 : 595.28;
          const targetHeight = landscape ? 595.28 : 841.89;
          const margin = 18;
          const scale = Math.min((targetWidth - margin * 2) / embedded.width, (targetHeight - margin * 2) / embedded.height);
          const width = embedded.width * scale; const height = embedded.height * scale;
          const page = output.addPage([targetWidth, targetHeight]);
          page.drawPage(embedded, { x: (targetWidth - width) / 2, y: (targetHeight - height) / 2, width, height });
        }
        downloadBlob(await output.save(), `${baseName(files[0].name)}_A4.pdf`);
        setNotice(`Đã chuyển ${source.getPageCount()} trang về khổ A4 và giữ đúng tỷ lệ nội dung.`);
      } else if (mode === "annotate") {
        if (!addWatermark && !addNumbers) throw new Error("Hãy chọn ít nhất một nội dung cần thêm.");
        const document = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false });
        const regular = await document.embedFont(StandardFonts.Helvetica);
        const bold = await document.embedFont(StandardFonts.HelveticaBold);
        const safeWatermark = watermark.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/Đ/g, "D").replace(/đ/g, "d").replace(/[^\x20-\x7E]/g, "").trim() || "SY LAND";
        document.getPages().forEach((page, index) => {
          const width = page.getWidth(); const height = page.getHeight();
          if (addWatermark) {
            const fontSize = Math.max(28, Math.min(62, width / 8));
            const textWidth = bold.widthOfTextAtSize(safeWatermark, fontSize);
            page.drawText(safeWatermark, { x: (width - textWidth * .82) / 2, y: height / 2, size: fontSize, font: bold, color: rgb(.08, .32, .23), opacity: .16, rotate: degrees(-32) });
          }
          if (addNumbers) {
            const label = `${numberStart + index}`; const fontSize = 10; const labelWidth = regular.widthOfTextAtSize(label, fontSize);
            page.drawText(label, { x: (width - labelWidth) / 2, y: numberPosition === "bottom" ? 16 : height - 24, size: fontSize, font: regular, color: rgb(.18, .22, .2), opacity: .9 });
          }
        });
        downloadBlob(await document.save({ useObjectStreams: true }), `${baseName(files[0].name)}_DANH_DAU.pdf`);
        setNotice(`Đã cập nhật ${document.getPageCount()} trang${addWatermark ? " · dấu bản quyền" : ""}${addNumbers ? " · số trang" : ""}.`);
      } else if (mode === "optimize") {
        const before = files[0].size;
        const document = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false, updateMetadata: false });
        const bytes = await document.save({ useObjectStreams: true, addDefaultPage: false, objectsPerTick: 50 });
        downloadBlob(bytes, `${baseName(files[0].name)}_TOI_UU.pdf`);
        const change = Math.round((1 - bytes.length / before) * 1000) / 10;
        setNotice(change > 0 ? `Đã tối ưu cấu trúc: giảm ${change}% (${(before / 1024 / 1024).toFixed(1)} MB → ${(bytes.length / 1024 / 1024).toFixed(1)} MB).` : "Đã tối ưu cấu trúc. Tệp không giảm đáng kể vì dữ liệu ảnh đã được nén sẵn.");
      } else if (mode === "sanitize") {
        const document = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false, updateMetadata: false });
        document.setTitle(""); document.setAuthor(""); document.setSubject(""); document.setKeywords([]); document.setProducer("SY LAND"); document.setCreator("SY LAND");
        downloadBlob(await document.save({ useObjectStreams: true }), `${baseName(files[0].name)}_DA_LAM_SACH.pdf`);
        setNotice(`Đã làm sạch thông tin mô tả của ${document.getPageCount()} trang. Nội dung hiển thị không thay đổi.`);
      } else if (mode === "compress") {
        const pdfjs = await import("pdfjs-dist");
        const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
        pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
        const source = await pdfjs.getDocument({ data: new Uint8Array(await files[0].arrayBuffer()) }).promise;
        if (source.numPages > 80) throw new Error("Nén ảnh hỗ trợ tối đa 80 trang mỗi lượt.");
        const profiles = { small: { scale: 1.15, quality: .52 }, balanced: { scale: 1.55, quality: .7 }, clear: { scale: 2, quality: .84 } } as const;
        const profile = profiles[compressQuality]; const output = await PDFDocument.create();
        for (let pageNumber = 1; pageNumber <= source.numPages; pageNumber++) {
          setNotice(`Đang nén trang ${pageNumber}/${source.numPages}…`);
          const page = await source.getPage(pageNumber); const base = page.getViewport({ scale: 1 }); const viewport = page.getViewport({ scale: profile.scale });
          const canvas = document.createElement("canvas"); canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
          await page.render({ canvas, canvasContext: canvas.getContext("2d")!, viewport, background: "white" }).promise;
          const jpeg = await new Promise<Blob>((resolve, reject) => canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("Không tạo được ảnh nén.")), "image/jpeg", profile.quality));
          const image = await output.embedJpg(await jpeg.arrayBuffer()); const target = output.addPage([base.width, base.height]);
          target.drawImage(image, { x: 0, y: 0, width: base.width, height: base.height });
        }
        const bytes = await output.save({ useObjectStreams: true }); const before = files[0].size; const change = Math.round((1 - bytes.length / before) * 1000) / 10;
        downloadBlob(bytes, `${baseName(files[0].name)}_NEN_${compressQuality.toUpperCase()}.pdf`);
        setNotice(change > 0 ? `Đã nén ${source.numPages} trang: giảm ${change}% (${(before / 1024 / 1024).toFixed(1)} MB → ${(bytes.length / 1024 / 1024).toFixed(1)} MB).` : "Đã tạo bản nén ảnh nhưng dung lượng không giảm; nên giữ tệp gốc.");
      } else {
        const source = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false });
        const total = source.getPageCount();
        const groups = splitKind === "each" ? Array.from({ length: total }, (_, index) => [index]) : parseRanges(ranges, total);
        if (splitKind === "selected") {
          const output = await PDFDocument.create();
          const selected = [...new Set(groups.flat())];
          (await output.copyPages(source, selected)).forEach((page) => output.addPage(page));
          downloadBlob(await output.save(), `${baseName(files[0].name)}_TRANG_DA_CHON.pdf`);
          setNotice(`Đã xuất ${selected.length}/${total} trang vào một tệp.`);
        } else {
          const { zipSync } = await import("fflate");
          const archive: Record<string, Uint8Array> = {};
          for (let index = 0; index < groups.length; index++) {
            const output = await PDFDocument.create();
            (await output.copyPages(source, groups[index])).forEach((page) => output.addPage(page));
            const label = splitKind === "each" ? `TRANG_${String(index + 1).padStart(3, "0")}` : `PHAN_${String(index + 1).padStart(2, "0")}_TRANG_${groups[index][0] + 1}-${groups[index].at(-1)! + 1}`;
            archive[`${baseName(files[0].name)}_${label}.pdf`] = await output.save();
          }
          downloadBlob(zipSync(archive, { level: 0 }), `${baseName(files[0].name)}_DA_TACH.zip`, "application/zip");
          setNotice(`Đã tạo ${groups.length} tệp PDF trong một gói ZIP.`);
        }
      }
    } catch (reason) {
      console.error(reason);
      setNotice("Không xử lý được. PDF có thể bị khóa, hỏng hoặc dùng cấu trúc chưa được hỗ trợ.");
    } finally { setBusy(false); }
  }

  return (
    <section className="pdf-toolkit" id="cong-cu-pdf" aria-labelledby="pdf-toolkit-title">
      <div className="section-heading split-heading"><div><p className="section-kicker">Bộ công cụ PDF địa chính</p><h2 id="pdf-toolkit-title">Tách, chọn trang, nối và xoay PDF<br />ngay trên thiết bị.</h2></div><p>Thiết kế độc lập cho SỸ LAND. Tệp không được tải lên máy chủ; bản gốc không bị sửa hoặc ghi đè.</p></div>
      <div className="pdf-tool-shell">
        <div className="pdf-tool-tabs" role="tablist" aria-label="Chọn công cụ PDF">
          <button className={mode === "split" ? "active" : ""} type="button" onClick={() => changeMode("split")}>Tách và xuất trang</button>
          <button className={mode === "merge" ? "active" : ""} type="button" onClick={() => changeMode("merge")}>Nối PDF</button>
          <button className={mode === "rotate" ? "active" : ""} type="button" onClick={() => changeMode("rotate")}>Xoay PDF</button>
          <button className={mode === "organize" ? "active" : ""} type="button" onClick={() => changeMode("organize")}>Sắp xếp · Xóa trang</button>
          <button className={mode === "resize" ? "active" : ""} type="button" onClick={() => changeMode("resize")}>Chuyển sang A4</button>
          <button className={mode === "annotate" ? "active" : ""} type="button" onClick={() => changeMode("annotate")}>Số trang · Dấu bản quyền</button>
          <button className={mode === "optimize" ? "active" : ""} type="button" onClick={() => changeMode("optimize")}>Tối ưu PDF</button>
          <button className={mode === "compare" ? "active" : ""} type="button" onClick={() => changeMode("compare")}>So sánh PDF</button>
          <button className={mode === "sanitize" ? "active" : ""} type="button" onClick={() => changeMode("sanitize")}>Làm sạch metadata</button>
          <button className={mode === "ocr" ? "active" : ""} type="button" onClick={() => changeMode("ocr")}>OCR PDF scan</button>
          <button className={mode === "compress" ? "active" : ""} type="button" onClick={() => changeMode("compress")}>Nén ảnh PDF</button>
        </div>
        <div className="pdf-tool-body">
          <div className="pdf-pick">
            <input ref={inputRef} type="file" accept="application/pdf,.pdf" multiple={mode === "merge" || mode === "compare"} onChange={chooseFiles} />
            <span aria-hidden="true">PDF</span><h3>{mode === "merge" ? "Chọn nhiều PDF theo thứ tự cần nối" : mode === "compare" ? "Chọn hai phiên bản PDF" : "Chọn một tệp PDF"}</h3><p>{mode === "merge" ? "Tối đa 30 tệp. Có thể di chuyển thứ tự sau khi chọn." : mode === "compare" ? "So sánh nội dung văn bản theo từng trang; không tải tệp lên máy chủ." : "Xử lý cục bộ; khuyến nghị tệp không quá 100 MB."}</p>
            <button type="button" onClick={() => inputRef.current?.click()}>Chọn PDF từ thiết bị</button>
          </div>
          <div className="pdf-options">
            {mode === "split" && <><h3>Phương thức tách</h3><div className="pdf-choice-grid"><label><input type="radio" checked={splitKind === "ranges"} onChange={() => setSplitKind("ranges")} /><span><b>Theo khoảng</b><small>Tạo một PDF cho mỗi khoảng</small></span></label><label><input type="radio" checked={splitKind === "selected"} onChange={() => setSplitKind("selected")} /><span><b>Xuất trang đã chọn</b><small>Gộp trang chọn vào một PDF</small></span></label><label><input type="radio" checked={splitKind === "each"} onChange={() => setSplitKind("each")} /><span><b>Mỗi trang một tệp</b><small>Tải kết quả dạng ZIP</small></span></label></div>{splitKind !== "each" && <label className="range-input">Khoảng/trang cần xử lý<input value={ranges} onChange={(event) => setRanges(event.target.value)} placeholder="Ví dụ: 1-3, 5, 8-10" /><small>Dùng dấu phẩy để ngăn cách nhiều khoảng.</small></label>}</>}
            {mode === "merge" && <><h3>Thứ tự nối</h3><div className="merge-list">{files.length ? files.map((file, index) => <div key={`${file.name}-${file.lastModified}`}><span>{index + 1}</span><p><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => move(index, -1)} disabled={index === 0}>↑</button><button type="button" onClick={() => move(index, 1)} disabled={index === files.length - 1}>↓</button><button type="button" onClick={() => setFiles((current) => current.filter((_, itemIndex) => itemIndex !== index))}>×</button></div>) : <p className="pdf-empty">Chưa chọn tệp.</p>}</div></>}
            {mode === "rotate" && <><h3>Góc xoay toàn bộ trang</h3><div className="angle-options">{[90, 180, 270].map((value) => <button className={angle === value ? "active" : ""} type="button" key={value} onClick={() => setAngle(value)}>↻ {value}°</button>)}</div></>}
            {mode === "organize" && <><h3>Xem trước và sắp xếp trang</h3><div className="page-organizer">{previews.map((item, index) => <article className={removedPages.has(item.page) ? "removed" : ""} key={item.page}><div><img src={item.image} alt={`Xem trước trang ${item.page}`} /><span>{index + 1}</span></div><small>Trang gốc {item.page}</small><div><button type="button" onClick={() => movePage(index, -1)} disabled={index === 0}>←</button><button type="button" onClick={() => movePage(index, 1)} disabled={index === previews.length - 1}>→</button><button type="button" className="remove-page" onClick={() => toggleRemoved(item.page)}>{removedPages.has(item.page) ? "Giữ" : "Xóa"}</button></div></article>)}</div></>}
            {mode === "resize" && <div className="resize-info"><span>A3</span><b>→</b><span>A4</span><div><h3>Chuyển toàn bộ trang về A4</h3><p>Giữ tỷ lệ nội dung, tự nhận hướng dọc/ngang và căn giữa để thuận tiện in hoặc ký số.</p></div></div>}
            {mode === "annotate" && <><h3>Đánh dấu tài liệu</h3><div className="annotate-options"><label className="annotate-check"><input type="checkbox" checked={addWatermark} onChange={(event) => setAddWatermark(event.target.checked)} /><span>Thêm dấu bản quyền</span></label><label>Nội dung dấu<input value={watermark} maxLength={60} disabled={!addWatermark} onChange={(event) => setWatermark(event.target.value)} placeholder="Ví dụ: SỸ LAND - BẢN KIỂM TRA" /><small>Chữ tiếng Việt được tự chuyển sang không dấu để tương thích PDF.</small></label><label className="annotate-check"><input type="checkbox" checked={addNumbers} onChange={(event) => setAddNumbers(event.target.checked)} /><span>Thêm số trang</span></label><div className="number-config"><label>Bắt đầu từ<input type="number" min="1" value={numberStart} disabled={!addNumbers} onChange={(event) => setNumberStart(Math.max(1, Number(event.target.value) || 1))} /></label><label>Vị trí<select value={numberPosition} disabled={!addNumbers} onChange={(event) => setNumberPosition(event.target.value as "bottom" | "top")}><option value="bottom">Giữa chân trang</option><option value="top">Giữa đầu trang</option></select></label></div></div></>}
            {mode === "optimize" && <div className="optimize-info"><span>⇣</span><div><h3>Tối ưu cấu trúc PDF</h3><p>Sắp xếp lại đối tượng và luồng dữ liệu để giảm dung lượng khi có thể. Công cụ không hạ độ phân giải ảnh nên giữ nguyên chất lượng hồ sơ scan.</p><ul><li>Không xóa nội dung</li><li>Không giảm chất lượng ảnh</li><li>Kết quả phụ thuộc cấu trúc tệp gốc</li></ul></div></div>}
            {mode === "compare" && <><div className="compare-files">{files.map((file, index) => <article key={`${file.name}-${file.lastModified}`}><span>{index ? "B" : "A"}</span><div><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></div></article>)}</div>{compareResult && <div className="compare-result"><div><strong>{compareResult.pagesA}</strong><span>Trang bản A</span></div><div><strong>{compareResult.pagesB}</strong><span>Trang bản B</span></div><div><strong>{compareResult.samePages}</strong><span>Trang trùng khớp</span></div><div><strong>{compareResult.changedPages.length}</strong><span>Trang thay đổi</span></div><p><b>Trang cần kiểm tra:</b> {compareResult.changedPages.length ? compareResult.changedPages.slice(0, 100).join(", ") : "Không có"}{compareResult.changedPages.length > 100 ? "…" : ""}</p><small>Độ phủ văn bản: {compareResult.textCoverage}%. PDF scan không có lớp chữ cần OCR trước để so sánh chính xác.</small></div>}</>}
            {mode === "sanitize" && <div className="optimize-info"><span>⌫</span><div><h3>Làm sạch metadata</h3><p>Xóa tiêu đề, tác giả, chủ đề và từ khóa ẩn trong PDF trước khi gửi cho người khác. Nội dung và hình ảnh hiển thị được giữ nguyên.</p><ul><li>Không xóa chữ ký hoặc chú thích hiển thị</li><li>Không thay thế việc kiểm tra dữ liệu cá nhân trong nội dung</li><li>Tạo tệp mới, không ghi đè bản gốc</li></ul></div></div>}
            {mode === "ocr" && <div className="optimize-info"><span>OCR</span><div><h3>Nhận dạng chữ tiếng Việt</h3><p>Chuyển từng trang scan thành văn bản TXT có phân cách trang. Xử lý cục bộ, tối đa 40 trang mỗi lượt.</p><ul><li>Kết quả OCR cần được người dùng kiểm tra lại</li><li>Không tự động thay thế dữ liệu hồ sơ gốc</li><li>Độ chính xác phụ thuộc chất lượng bản scan</li></ul></div></div>}
            {mode === "compress" && <><h3>Mức nén hình ảnh</h3><div className="compression-options">{([{ key: "small", title: "Dung lượng nhỏ", note: "Phù hợp gửi xem nhanh" }, { key: "balanced", title: "Cân bằng", note: "Khuyến nghị cho hồ sơ thông thường" }, { key: "clear", title: "Rõ nét", note: "Ưu tiên chữ nhỏ và bản đồ" }] as const).map((item) => <label key={item.key}><input type="radio" checked={compressQuality === item.key} onChange={() => setCompressQuality(item.key)} /><span><b>{item.title}</b><small>{item.note}</small></span></label>)}</div><p className="compression-warning"><b>Lưu ý:</b> Bản nén được dựng lại từ ảnh từng trang nên có thể mất lớp chữ tìm kiếm, liên kết, biểu mẫu và trạng thái chữ ký số. Luôn giữ tệp gốc.</p></>}
            {mode !== "merge" && mode !== "organize" && mode !== "compare" && files[0] && <div className="single-pdf"><span>✓</span><p><b>{files[0].name}</b><small>{(files[0].size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => setFiles([])}>×</button></div>}
            <div className="pdf-run"><p className={notice.includes("Không") ? "error" : ""}>{notice || (files.length ? `${files.length} tệp · ${(totalSize / 1024 / 1024).toFixed(1)} MB` : "Kết quả sẽ được tải về thiết bị.")}</p><div className="pdf-run-actions">{mode === "compare" && compareResult && <button className="secondary-pdf-action" type="button" onClick={() => { const rows = ["Chi_tieu,Ban_A,Ban_B", `Ten_tep,\"${files[0].name.replace(/\"/g, '\"\"')}\",\"${files[1].name.replace(/\"/g, '\"\"')}\"`, `So_trang,${compareResult.pagesA},${compareResult.pagesB}`, `Trang_trung_khop,${compareResult.samePages},`, `Trang_thay_doi,\"${compareResult.changedPages.join("; ")}\",`, `Do_phu_van_ban,${compareResult.textCoverage}%,`]; downloadText("\uFEFF" + rows.join("\r\n"), `SYLAND_BAO_CAO_SO_SANH_${Date.now()}.csv`, "text/csv;charset=utf-8"); }}>Tải báo cáo CSV</button>}<button type="button" disabled={busy || !files.length || (mode === "compare" && files.length !== 2) || (mode === "organize" && !previews.length)} onClick={() => void run()}>{busy ? "Đang xử lý…" : mode === "merge" ? "Nối PDF" : mode === "rotate" ? "Xoay và tải PDF" : mode === "organize" ? "Xuất PDF đã sắp xếp" : mode === "resize" ? "Chuyển và tải PDF A4" : mode === "annotate" ? "Thêm dấu và tải PDF" : mode === "optimize" ? "Tối ưu và tải PDF" : mode === "compare" ? "So sánh hai PDF" : mode === "sanitize" ? "Làm sạch và tải PDF" : mode === "ocr" ? "OCR và tải TXT" : mode === "compress" ? "Nén và tải PDF" : "Tách và tải kết quả"}</button></div></div>
          </div>
        </div>
      </div>
    </section>
  );
}
