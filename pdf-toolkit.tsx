"use client";

import { ChangeEvent, useMemo, useRef, useState } from "react";

type Mode = "split" | "merge" | "rotate" | "organize" | "resize";
type PagePreview = { page: number; image: string };

function downloadBlob(bytes: Uint8Array, name: string, type = "application/pdf") {
  const blob = new Blob([bytes as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  anchor.click();
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
    setFiles(mode === "merge" ? selected.slice(0, 30) : selected.slice(0, 1));
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

  function changeMode(next: Mode) { setMode(next); setFiles([]); setPreviews([]); setRemovedPages(new Set()); setNotice(""); }

  function movePage(index: number, direction: -1 | 1) { setPreviews((current) => { const target = index + direction; if (target < 0 || target >= current.length) return current; const next = [...current]; [next[index], next[target]] = [next[target], next[index]]; return next; }); }
  function toggleRemoved(page: number) { setRemovedPages((current) => { const next = new Set(current); if (next.has(page)) next.delete(page); else next.add(page); return next; }); }

  async function run() {
    if (!files.length) { setNotice("Hãy chọn tệp PDF trước khi xử lý."); return; }
    setBusy(true); setNotice("Đang xử lý trên thiết bị…");
    try {
      const { PDFDocument, degrees } = await import("pdf-lib");
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
        </div>
        <div className="pdf-tool-body">
          <div className="pdf-pick">
            <input ref={inputRef} type="file" accept="application/pdf,.pdf" multiple={mode === "merge"} onChange={chooseFiles} />
            <span aria-hidden="true">PDF</span><h3>{mode === "merge" ? "Chọn nhiều PDF theo thứ tự cần nối" : "Chọn một tệp PDF"}</h3><p>{mode === "merge" ? "Tối đa 30 tệp. Có thể di chuyển thứ tự sau khi chọn." : "Xử lý cục bộ; khuyến nghị tệp không quá 100 MB."}</p>
            <button type="button" onClick={() => inputRef.current?.click()}>Chọn PDF từ thiết bị</button>
          </div>
          <div className="pdf-options">
            {mode === "split" && <><h3>Phương thức tách</h3><div className="pdf-choice-grid"><label><input type="radio" checked={splitKind === "ranges"} onChange={() => setSplitKind("ranges")} /><span><b>Theo khoảng</b><small>Tạo một PDF cho mỗi khoảng</small></span></label><label><input type="radio" checked={splitKind === "selected"} onChange={() => setSplitKind("selected")} /><span><b>Xuất trang đã chọn</b><small>Gộp trang chọn vào một PDF</small></span></label><label><input type="radio" checked={splitKind === "each"} onChange={() => setSplitKind("each")} /><span><b>Mỗi trang một tệp</b><small>Tải kết quả dạng ZIP</small></span></label></div>{splitKind !== "each" && <label className="range-input">Khoảng/trang cần xử lý<input value={ranges} onChange={(event) => setRanges(event.target.value)} placeholder="Ví dụ: 1-3, 5, 8-10" /><small>Dùng dấu phẩy để ngăn cách nhiều khoảng.</small></label>}</>}
            {mode === "merge" && <><h3>Thứ tự nối</h3><div className="merge-list">{files.length ? files.map((file, index) => <div key={`${file.name}-${file.lastModified}`}><span>{index + 1}</span><p><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => move(index, -1)} disabled={index === 0}>↑</button><button type="button" onClick={() => move(index, 1)} disabled={index === files.length - 1}>↓</button><button type="button" onClick={() => setFiles((current) => current.filter((_, itemIndex) => itemIndex !== index))}>×</button></div>) : <p className="pdf-empty">Chưa chọn tệp.</p>}</div></>}
            {mode === "rotate" && <><h3>Góc xoay toàn bộ trang</h3><div className="angle-options">{[90, 180, 270].map((value) => <button className={angle === value ? "active" : ""} type="button" key={value} onClick={() => setAngle(value)}>↻ {value}°</button>)}</div></>}
            {mode === "organize" && <><h3>Xem trước và sắp xếp trang</h3><div className="page-organizer">{previews.map((item, index) => <article className={removedPages.has(item.page) ? "removed" : ""} key={item.page}><div><img src={item.image} alt={`Xem trước trang ${item.page}`} /><span>{index + 1}</span></div><small>Trang gốc {item.page}</small><div><button type="button" onClick={() => movePage(index, -1)} disabled={index === 0}>←</button><button type="button" onClick={() => movePage(index, 1)} disabled={index === previews.length - 1}>→</button><button type="button" className="remove-page" onClick={() => toggleRemoved(item.page)}>{removedPages.has(item.page) ? "Giữ" : "Xóa"}</button></div></article>)}</div></>}
            {mode === "resize" && <div className="resize-info"><span>A3</span><b>→</b><span>A4</span><div><h3>Chuyển toàn bộ trang về A4</h3><p>Giữ tỷ lệ nội dung, tự nhận hướng dọc/ngang và căn giữa để thuận tiện in hoặc ký số.</p></div></div>}
            {mode !== "merge" && mode !== "organize" && files[0] && <div className="single-pdf"><span>✓</span><p><b>{files[0].name}</b><small>{(files[0].size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => setFiles([])}>×</button></div>}
            <div className="pdf-run"><p className={notice.includes("Không") ? "error" : ""}>{notice || (files.length ? `${files.length} tệp · ${(totalSize / 1024 / 1024).toFixed(1)} MB` : "Kết quả sẽ được tải về thiết bị.")}</p><button type="button" disabled={busy || !files.length || (mode === "organize" && !previews.length)} onClick={() => void run()}>{busy ? "Đang xử lý…" : mode === "merge" ? "Nối PDF" : mode === "rotate" ? "Xoay và tải PDF" : mode === "organize" ? "Xuất PDF đã sắp xếp" : mode === "resize" ? "Chuyển và tải PDF A4" : "Tách và tải kết quả"}</button></div>
          </div>
        </div>
      </div>
    </section>
  );
}
