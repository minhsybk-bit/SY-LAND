"use client";

import { ChangeEvent, useMemo, useRef, useState } from "react";
import { consumeDailyUsage, EntitlementSnapshot, useEntitlements } from "./entitlements";

type Mode = "split" | "merge" | "rotate" | "organize" | "resize" | "annotate" | "optimize" | "compare" | "sanitize" | "ocr" | "compress" | "dedupe" | "toimage" | "imagetopdf" | "wordtopdf" | "crop" | "batch";
type PagePreview = {
  page: number;
  image: string;
  status?: "ok" | "review" | "error";
  decision?: "unreviewed" | "blank" | "content";
  inkRatio?: number;
  textChars?: number;
  extractedText?: string;
  ocrText?: string;
  note?: string;
};
type PageComparison = { page: number; similarity: number | null; status: "same" | "review" | "changed" | "missing" | "scan"; note: string };
type CompareResult = { pagesA: number; pagesB: number; changedPages: number[]; samePages: number; reviewPages: number; scanPages: number; textCoverage: number; details: PageComparison[] };
type ToolCategory = "all" | "split" | "edit" | "convert" | "review" | "optimize";
const MAX_BATCH_TOTAL_BYTES = 300 * 1024 * 1024;
const MAX_IMAGE_TOTAL_BYTES = 500 * 1024 * 1024;
const MODE_LEVEL: Record<Mode, number> = {
  split: 25, merge: 25, rotate: 25, organize: 25,
  crop: 40, resize: 40, sanitize: 40,
  annotate: 70, optimize: 70, compress: 70, dedupe: 70,
  toimage: 70, imagetopdf: 70, wordtopdf: 70, batch: 70,
  compare: 100, ocr: 100,
};
function canUseMode(entitlements: EntitlementSnapshot, mode: Mode) {
  return entitlements.role === "admin" || entitlements.fullTools || entitlements.featurePercent >= MODE_LEVEL[mode];
}

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

function normalizeCompareText(input: string) {
  return input.normalize("NFKC").toLocaleLowerCase("vi").replace(/[\u00ad\u200b-\u200d\ufeff]/g, "").replace(/\s+/g, " ").replace(/\s+([,.;:!?%)])/g, "$1").trim();
}

function textSimilarity(left: string, right: string) {
  if (left === right) return 100;
  const leftTokens = left.split(/\s+/).filter(Boolean); const rightTokens = right.split(/\s+/).filter(Boolean);
  if (!leftTokens.length || !rightTokens.length) return 0;
  const counts = new Map<string, number>(); leftTokens.forEach((token) => counts.set(token, (counts.get(token) || 0) + 1));
  let overlap = 0; rightTokens.forEach((token) => { const count = counts.get(token) || 0; if (count > 0) { overlap++; counts.set(token, count - 1); } });
  return Math.round((2 * overlap / (leftTokens.length + rightTokens.length)) * 1000) / 10;
}

export default function PdfToolkit() {
  const entitlements = useEntitlements();
  const [mode, setMode] = useState<Mode>("split");
  const [files, setFiles] = useState<File[]>([]);
  const [ranges, setRanges] = useState("1");
  const [splitKind, setSplitKind] = useState<"ranges" | "each" | "selected" | "visual" | "fixed" | "cuts" | "parity">("ranges");
  const [pagesPerPart, setPagesPerPart] = useState(2);
  const [outputPrefix, setOutputPrefix] = useState("SYLAND");
  const [cutPoints, setCutPoints] = useState("3, 7");
  const [angle, setAngle] = useState(90);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [previews, setPreviews] = useState<PagePreview[]>([]);
  const [inspectedPage, setInspectedPage] = useState<number | null>(null);
  const [ocrPage, setOcrPage] = useState<number | null>(null);
  const [removedPages, setRemovedPages] = useState<Set<number>>(new Set());
  const [selectedPages, setSelectedPages] = useState<Set<number>>(new Set());
  const [watermark, setWatermark] = useState("SY LAND");
  const [addWatermark, setAddWatermark] = useState(true);
  const [addNumbers, setAddNumbers] = useState(true);
  const [numberStart, setNumberStart] = useState(1);
  const [numberPosition, setNumberPosition] = useState<"bottom" | "top">("bottom");
  const [compareResult, setCompareResult] = useState<CompareResult | null>(null);
  const [compressQuality, setCompressQuality] = useState<"small" | "balanced" | "clear">("balanced");
  const [duplicateGroups, setDuplicateGroups] = useState<number[][]>([]);
  const [duplicateScanned, setDuplicateScanned] = useState(false);
  const [imageFormat, setImageFormat] = useState<"jpg" | "png">("jpg");
  const [imageResolution, setImageResolution] = useState<"screen" | "standard" | "high">("standard");
  const [imageRanges, setImageRanges] = useState("1");
  const [toolSearch, setToolSearch] = useState("");
  const [toolCategory, setToolCategory] = useState<ToolCategory>("edit");
  const [imagePageMode, setImagePageMode] = useState<"a4" | "original">("a4");
  const [cropMargins, setCropMargins] = useState({ top: 0, right: 0, bottom: 0, left: 0 });
  const [batchAction, setBatchAction] = useState<"optimize" | "sanitize" | "rotate">("optimize");
  const [wordPdfQuality, setWordPdfQuality] = useState<"standard" | "high">("standard");
  const [wordPreviewing, setWordPreviewing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0, label: "Sẵn sàng" });
  const [paused, setPaused] = useState(false);
  const [imagePreviews, setImagePreviews] = useState<string[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);
  const wordPreviewRef = useRef<HTMLDivElement>(null);
  const imagePreviewUrls = useRef<string[]>([]);
  const taskControl = useRef({ cancelled: false, paused: false });
  const totalSize = useMemo(() => files.reduce((sum, file) => sum + file.size, 0), [files]);
  const maxBatchFiles = entitlements.maxParcelsPerRun;
  const batchLimitLabel = maxBatchFiles == null ? "không giới hạn số tệp" : `tối đa ${maxBatchFiles} tệp`;
  const toolName = useMemo(() => ({ split: "Tách và xuất trang", merge: "Nối PDF", rotate: "Xoay PDF", organize: "Sắp xếp và xóa trang", resize: "Chuyển sang A4", annotate: "Số trang và dấu bản quyền", optimize: "Tối ưu PDF", compare: "So sánh PDF", sanitize: "Làm sạch metadata", ocr: "OCR PDF scan", compress: "Nén ảnh PDF", dedupe: "Tìm trang trùng", toimage: "PDF sang ảnh", imagetopdf: "Ảnh sang PDF", wordtopdf: "Word sang PDF hàng loạt", crop: "Cắt lề PDF", batch: "Xử lý PDF hàng loạt" }[mode]), [mode]);

  function toolVisible(category: Exclude<ToolCategory, "all">, keywords: string) {
    const query = toolSearch.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    const normalized = keywords.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    return (toolCategory === "all" || toolCategory === category) && (!query || normalized.includes(query));
  }

  async function checkpoint() {
    if (taskControl.current.cancelled) throw new Error("SYLAND_TASK_CANCELLED");
    while (taskControl.current.paused) {
      await new Promise((resolve) => window.setTimeout(resolve, 120));
      if (taskControl.current.cancelled) throw new Error("SYLAND_TASK_CANCELLED");
    }
  }

  function togglePause() {
    const next = !taskControl.current.paused;
    taskControl.current.paused = next; setPaused(next);
    setNotice(next ? "Đã tạm dừng tại điểm an toàn. Nhấn Tiếp tục để chạy tiếp." : "Đang tiếp tục xử lý…");
  }

  function cancelTask() {
    taskControl.current.cancelled = true; taskControl.current.paused = false; setPaused(false);
    setNotice("Đang hủy tác vụ tại điểm an toàn gần nhất…");
  }

  async function renderPreviews(file: File) {
    taskControl.current = { cancelled: false, paused: false }; setPaused(false); setBusy(true);
    setProgress({ current: 0, total: 0, label: "Đang mở và kiểm tra cấu trúc PDF…" });
    setNotice("Đang đọc nội dung và tạo bản xem trước để người dùng xác nhận…");
    try {
      const pdfjs = await import("pdfjs-dist");
      const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
      pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
      const pdfDoc = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
      const next: PagePreview[] = [];
      let pagesToReview = 0;
      const limit = Math.min(pdfDoc.numPages, mode === "organize" || mode === "split" ? 80 : 8);
      for (let pageNumber = 1; pageNumber <= limit; pageNumber++) {
        await checkpoint();
        setProgress({ current: pageNumber, total: pdfDoc.numPages, label: `Đang kiểm tra trang ${pageNumber}/${pdfDoc.numPages}` });
        try {
          const page = await pdfDoc.getPage(pageNumber);
          const base = page.getViewport({ scale: 1 });
          if (base.width < 10 || base.height < 10) throw new Error("Kích thước trang không hợp lệ");
          const viewport = page.getViewport({ scale: Math.min(.55, 260 / base.width) });
          const canvas = window.document.createElement("canvas");
          canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
          const context = canvas.getContext("2d", { willReadFrequently: true })!;
          await page.render({ canvas, canvasContext: context, viewport, background: "white" }).promise;
          const text = await page.getTextContent();
          const extractedText = text.items.map((item) => "str" in item ? item.str : "").join(" ").replace(/\s+/g, " ").trim();
          const textChars = extractedText.replace(/\s/g, "").length;
          const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
          let dark = 0; let sampled = 0;
          for (let index = 0; index < pixels.length; index += 16) {
            sampled++; if ((pixels[index] + pixels[index + 1] + pixels[index + 2]) / 3 < 245) dark++;
          }
          const inkRatio = sampled ? dark / sampled : 0;
          const needsReview = inkRatio < .0025 && textChars < 5;
          if (needsReview) pagesToReview++;
          next.push({
            page: pageNumber,
            image: canvas.toDataURL("image/jpeg", .82),
            status: needsReview ? "review" : "ok",
            decision: needsReview ? "unreviewed" : "content",
            inkRatio,
            textChars,
            extractedText,
            note: needsReview
              ? "Chưa đủ dữ liệu để kết luận — mở trang, xem nội dung hoặc chạy OCR rồi xác nhận"
              : `Đã đọc được ${textChars} ký tự hoặc phát hiện nội dung hình ảnh`,
          });
        } catch (reason) {
          if (reason instanceof Error && reason.message === "SYLAND_TASK_CANCELLED") throw reason;
          next.push({ page: pageNumber, image: "", status: "error", note: reason instanceof Error ? reason.message : "Không render được trang" });
        }
        setPreviews([...next]);
      }
      setPreviews(next); setRemovedPages(new Set()); setSelectedPages(new Set());
      const errors = next.filter((item) => item.status === "error").length;
      const missing = pdfDoc.numPages < 2 ? ` Cảnh báo: tệp chỉ có ${pdfDoc.numPages} trang, thiếu trang ${pdfDoc.numPages < 1 ? "1–2" : "2"}.` : "";
      setProgress({ current: limit, total: pdfDoc.numPages, label: `Hoàn tất: ${pagesToReview} trang cần xác minh · ${errors} trang lỗi` });
      setNotice(pdfDoc.numPages > 80 ? `Đã đọc 80/${pdfDoc.numPages} trang đầu.${missing}` : `Đã đọc ${pdfDoc.numPages} trang: ${pagesToReview} trang cần người dùng xác minh, ${errors} lỗi.${missing}`);
    } catch (reason) {
      console.error(reason);
      setNotice(reason instanceof Error && reason.message === "SYLAND_TASK_CANCELLED" ? "Đã hủy kiểm tra. Kết quả một phần vẫn được giữ để xem lại." : "Không tạo được ảnh xem trước của PDF này.");
    } finally { if (!taskControl.current.cancelled) setProgress((current) => ({ ...current, current: current.total, label: current.current < current.total ? "Hoàn tất tác vụ" : current.label })); setBusy(false); }
  }

  async function renderWordPreview(file: File) {
    if (!wordPreviewRef.current) return;
    setWordPreviewing(true);
    setNotice("Đang dựng bản xem trước DOCX đầu tiên ngay trên thiết bị…");
    try {
      const { renderAsync } = await import("docx-preview");
      wordPreviewRef.current.innerHTML = "";
      await renderAsync(await file.arrayBuffer(), wordPreviewRef.current, wordPreviewRef.current, {
        className: "syland-docx",
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        breakPages: true,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        useBase64URL: true,
      });
      setNotice("Đã tạo bản xem trước tệp đầu tiên. Hãy kiểm tra bố cục trước khi chuyển đổi.");
    } catch (reason) {
      console.error(reason);
      wordPreviewRef.current.innerHTML = "";
      setNotice("Không dựng được bản xem trước DOCX. Tệp có thể bị khóa, hỏng hoặc chứa thành phần chưa được hỗ trợ.");
    } finally {
      setWordPreviewing(false);
    }
  }

  async function chooseFiles(event: ChangeEvent<HTMLInputElement>) {
    const selected = Array.from(event.target.files || []).filter((file) =>
      mode === "imagetopdf"
        ? /image\/(jpeg|png|webp)/i.test(file.type) || /\.(jpe?g|png|webp)$/i.test(file.name)
        : mode === "wordtopdf"
          ? file.name.toLowerCase().endsWith(".docx") || file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          : file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"));
    imagePreviewUrls.current.forEach((url) => URL.revokeObjectURL(url)); imagePreviewUrls.current = [];
    setDuplicateGroups([]); setDuplicateScanned(false); setCompareResult(null);
    const countLimited = mode === "merge" || mode === "imagetopdf" || mode === "wordtopdf" || mode === "batch"
      ? (maxBatchFiles == null ? selected : selected.slice(0, maxBatchFiles))
      : mode === "compare" ? selected.slice(0, 2) : selected.slice(0, 1);
    const hardTotalLimit = mode === "imagetopdf" ? MAX_IMAGE_TOTAL_BYTES : MAX_BATCH_TOTAL_BYTES;
    const totalLimit = Math.min(hardTotalLimit, entitlements.maxTotalUploadMB * 1024 * 1024);
    const fileLimit = entitlements.maxFileSizeMB * 1024 * 1024;
    const accepted: File[] = [];
    let acceptedBytes = 0;
    for (const file of countLimited) {
      if (file.size > fileLimit) continue;
      if (acceptedBytes + file.size > totalLimit) break;
      accepted.push(file); acceptedBytes += file.size;
    }
    setFiles(accepted);
    if (mode === "imagetopdf") { imagePreviewUrls.current = accepted.slice(0, 12).map((file) => URL.createObjectURL(file)); setImagePreviews([...imagePreviewUrls.current]); } else setImagePreviews([]);
    setNotice(!selected.length
      ? mode === "wordtopdf" ? "Hãy chọn tệp Word định dạng DOCX hợp lệ." : "Hãy chọn tệp PDF hợp lệ."
      : accepted.length < selected.length
        ? `Đã nhận ${accepted.length}/${selected.length} tệp theo quyền tài khoản: ${batchLimitLabel}, ${entitlements.maxFileSizeMB} MB/tệp và ${Math.round(totalLimit / 1024 / 1024)} MB/lượt.`
        : "");
    event.target.value = "";
    if (mode === "wordtopdf" && accepted[0]) await renderWordPreview(accepted[0]);
    else if (mode !== "imagetopdf" && accepted[0]) await renderPreviews(accepted[0]);
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

  function changeMode(next: Mode) {
    if (!canUseMode(entitlements, next)) {
      setNotice(`Công cụ này cần gói ${MODE_LEVEL[next] >= 100 ? "Pro" : "Plus"} hoặc quyền quản trị. Tài khoản hiện có ${entitlements.featurePercent}% công cụ.`);
      return;
    }
    taskControl.current.cancelled = true; imagePreviewUrls.current.forEach((url) => URL.revokeObjectURL(url)); imagePreviewUrls.current = []; setImagePreviews([]); if (wordPreviewRef.current) wordPreviewRef.current.innerHTML = ""; setMode(next); setFiles([]); setPreviews([]); setInspectedPage(null); setRemovedPages(new Set()); setSelectedPages(new Set()); setCompareResult(null); setDuplicateGroups([]); setDuplicateScanned(false); setProgress({ current: 0, total: 0, label: "Sẵn sàng" }); setPaused(false); setNotice("");
  }

  function movePage(index: number, direction: -1 | 1) { setPreviews((current) => { const target = index + direction; if (target < 0 || target >= current.length) return current; const next = [...current]; [next[index], next[target]] = [next[target], next[index]]; return next; }); }
  function toggleRemoved(page: number) { setRemovedPages((current) => { const next = new Set(current); if (next.has(page)) next.delete(page); else next.add(page); return next; }); }
  function toggleSelected(page: number) { setSelectedPages((current) => { const next = new Set(current); if (next.has(page)) next.delete(page); else next.add(page); return next; }); }

  function decidePage(page: number, decision: "blank" | "content") {
    setPreviews((current) => current.map((item) => item.page === page
      ? { ...item, decision, note: decision === "blank" ? "Người dùng đã xác nhận đây là trang trắng" : "Người dùng đã xác nhận trang có nội dung và được giữ lại" }
      : item));
    setRemovedPages((current) => {
      const next = new Set(current);
      if (decision === "content") next.delete(page);
      return next;
    });
  }

  async function readPageWithOcr(pageNumber: number) {
    if (!files[0]) return;
    setOcrPage(pageNumber);
    try {
      const pdfjs = await import("pdfjs-dist");
      const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
      const { createWorker } = await import("tesseract.js");
      pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
      const source = await pdfjs.getDocument({ data: new Uint8Array(await files[0].arrayBuffer()) }).promise;
      const page = await source.getPage(pageNumber);
      const base = page.getViewport({ scale: 1 });
      const viewport = page.getViewport({ scale: Math.min(2.2, 2200 / Math.max(base.width, base.height)) });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
      await page.render({ canvas, canvasContext: canvas.getContext("2d", { willReadFrequently: true })!, viewport, background: "white" }).promise;
      const worker = await createWorker("vie");
      try {
        const result = await worker.recognize(canvas);
        const ocrText = result.data.text.replace(/\s+/g, " ").trim();
        setPreviews((current) => current.map((item) => item.page === pageNumber
          ? { ...item, ocrText, decision: ocrText ? "content" : item.decision, note: ocrText ? `OCR đọc được ${ocrText.length} ký tự — trang có nội dung` : "OCR không đọc được chữ; vẫn phải xem hình ảnh trước khi xác nhận" }
          : item));
      } finally { await worker.terminate(); }
    } catch (reason) {
      console.error(reason);
      setNotice("Không OCR được trang này. Hãy kiểm tra trực tiếp bằng bản xem trước.");
    } finally { setOcrPage(null); }
  }

  async function run() {
    if (!canUseMode(entitlements, mode)) {
      setNotice(`Tài khoản chưa có quyền sử dụng ${toolName}. Hãy nâng cấp gói phù hợp.`);
      return;
    }
    if (!files.length) { setNotice(mode === "wordtopdf" ? "Hãy chọn tệp DOCX trước khi xử lý." : "Hãy chọn tệp PDF trước khi xử lý."); return; }
    const usage = consumeDailyUsage(entitlements);
    if (!usage.allowed) {
      setNotice("Đã hết lượt trải nghiệm hôm nay. Hãy đăng ký/đăng nhập hoặc nâng cấp gói để tiếp tục.");
      return;
    }
    taskControl.current = { cancelled: false, paused: false }; setPaused(false);
    setBusy(true); setProgress({ current: 0, total: 1, label: "Đang chuẩn bị tác vụ…" }); setNotice(usage.remaining == null ? "Đang xử lý trên thiết bị…" : `Đang xử lý trên thiết bị · còn ${usage.remaining} lượt hôm nay.`);
    try {
      if (mode === "compare") {
        if (files.length !== 2) throw new Error("Hãy chọn đúng hai tệp PDF để so sánh.");
        const pdfjs = await import("pdfjs-dist");
        const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
        pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
        const documents = await Promise.all(files.map(async (file) => pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise));
        const texts: string[][] = [];
        for (let documentIndex = 0; documentIndex < documents.length; documentIndex++) {
          const document = documents[documentIndex];
          const pages: string[] = [];
          for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber++) {
            await checkpoint(); setProgress({ current: pageNumber + documentIndex * document.numPages, total: documents.reduce((sum, item) => sum + item.numPages, 0), label: `Đang đọc bản ${documentIndex ? "B" : "A"} · trang ${pageNumber}/${document.numPages}` });
            const content = await (await document.getPage(pageNumber)).getTextContent();
            pages.push(normalizeCompareText(content.items.map((item) => "str" in item ? item.str : "").join(" ")));
          }
          texts.push(pages);
        }
        const maximum = Math.max(texts[0].length, texts[1].length);
        const changedPages: number[] = []; const details: PageComparison[] = [];
        let samePages = 0; let reviewPages = 0; let scanPages = 0; let pagesWithText = 0;
        for (let index = 0; index < maximum; index++) {
          const left = texts[0][index] || ""; const right = texts[1][index] || "";
          if (left || right) pagesWithText++;
          const page = index + 1;
          if (index >= texts[0].length || index >= texts[1].length) { changedPages.push(page); details.push({ page, similarity: null, status: "missing", note: index >= texts[0].length ? "Chỉ có trong bản B" : "Chỉ có trong bản A" }); continue; }
          if (!left && !right) { scanPages++; details.push({ page, similarity: null, status: "scan", note: "Cả hai trang không có lớp chữ; cần OCR" }); continue; }
          if (!left || !right) { scanPages++; details.push({ page, similarity: null, status: "scan", note: "Một bản không có lớp chữ; cần OCR" }); continue; }
          const similarity = textSimilarity(left, right);
          if (similarity >= 99.5) { samePages++; details.push({ page, similarity, status: "same", note: "Nội dung trùng khớp" }); }
          else if (similarity >= 92) { reviewPages++; details.push({ page, similarity, status: "review", note: "Khác biệt nhỏ; nên kiểm tra" }); }
          else { changedPages.push(page); details.push({ page, similarity, status: "changed", note: "Nội dung thay đổi đáng kể" }); }
        }
        setCompareResult({ pagesA: texts[0].length, pagesB: texts[1].length, changedPages, samePages, reviewPages, scanPages, textCoverage: maximum ? Math.round(pagesWithText / maximum * 100) : 0, details });
        setNotice(changedPages.length ? `Phát hiện ${changedPages.length} trang thay đổi đáng kể hoặc bị thêm/bớt; ${reviewPages} trang có khác biệt nhỏ.` : scanPages ? `Không thấy thay đổi rõ ràng; còn ${scanPages} trang scan cần OCR để kết luận.` : reviewPages ? `Có ${reviewPages} trang khác biệt nhỏ cần kiểm tra.` : "Hai tệp có nội dung văn bản trùng khớp theo từng trang.");
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
            await checkpoint(); setProgress({ current: pageNumber, total: source.numPages, label: `Đang OCR trang ${pageNumber}/${source.numPages}` });
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
      if (mode === "dedupe") {
        const pdfjs = await import("pdfjs-dist");
        const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
        pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
        const source = await pdfjs.getDocument({ data: new Uint8Array(await files[0].arrayBuffer()) }).promise;
        if (source.numPages > 150) throw new Error("Quét trang trùng hỗ trợ tối đa 150 trang mỗi lượt.");
        const hashes = new Map<string, number[]>();
        for (let pageNumber = 1; pageNumber <= source.numPages; pageNumber++) {
          await checkpoint(); setProgress({ current: pageNumber, total: source.numPages, label: `Đang tìm trang trùng ${pageNumber}/${source.numPages}` });
          setNotice(`Đang đối chiếu trang ${pageNumber}/${source.numPages}…`);
          const page = await source.getPage(pageNumber); const base = page.getViewport({ scale: 1 });
          const viewport = page.getViewport({ scale: Math.min(.75, 700 / Math.max(base.width, base.height)) });
          const canvas = document.createElement("canvas"); canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
          await page.render({ canvas, canvasContext: canvas.getContext("2d", { willReadFrequently: true })!, viewport, background: "white" }).promise;
          const data = canvas.getContext("2d", { willReadFrequently: true })!.getImageData(0, 0, canvas.width, canvas.height).data;
          const digest = await crypto.subtle.digest("SHA-256", data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength));
          const hash = Array.from(new Uint8Array(digest)).map((value) => value.toString(16).padStart(2, "0")).join("");
          hashes.set(hash, [...(hashes.get(hash) || []), pageNumber]);
        }
        const groups = [...hashes.values()].filter((group) => group.length > 1);
        setDuplicateGroups(groups); setDuplicateScanned(true);
        const repeated = groups.reduce((sum, group) => sum + group.length - 1, 0);
        setNotice(repeated ? `Phát hiện ${repeated} trang lặp trong ${groups.length} nhóm. Trang đầu mỗi nhóm sẽ được giữ lại.` : `Đã kiểm tra ${source.numPages} trang; không phát hiện trang giống hệt nhau.`);
        return;
      }
      if (mode === "toimage") {
        const pdfjs = await import("pdfjs-dist");
        const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
        const { zipSync } = await import("fflate");
        pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
        const source = await pdfjs.getDocument({ data: new Uint8Array(await files[0].arrayBuffer()) }).promise;
        const pages = [...new Set(parseRanges(imageRanges, source.numPages).flat())];
        if (pages.length > 80) throw new Error("Chuyển ảnh hỗ trợ tối đa 80 trang mỗi lượt.");
        const scales = { screen: 1.15, standard: 1.7, high: 2.35 } as const; const archive: Record<string, Uint8Array> = {};
        for (let index = 0; index < pages.length; index++) {
          await checkpoint(); setProgress({ current: index + 1, total: pages.length, label: `Đang xuất ảnh ${index + 1}/${pages.length}` });
          const pageNumber = pages[index] + 1; setNotice(`Đang xuất ảnh trang ${pageNumber} (${index + 1}/${pages.length})…`);
          const page = await source.getPage(pageNumber); const viewport = page.getViewport({ scale: scales[imageResolution] });
          const canvas = document.createElement("canvas"); canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
          await page.render({ canvas, canvasContext: canvas.getContext("2d")!, viewport, background: "white" }).promise;
          const mime = imageFormat === "png" ? "image/png" : "image/jpeg"; const quality = imageFormat === "jpg" ? .88 : undefined;
          const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob((value) => value ? resolve(value) : reject(new Error("Không tạo được ảnh.")), mime, quality));
          archive[`${baseName(files[0].name)}_TRANG_${String(pageNumber).padStart(3, "0")}.${imageFormat}`] = new Uint8Array(await blob.arrayBuffer());
        }
        downloadBlob(zipSync(archive, { level: 0 }), `${baseName(files[0].name)}_${imageFormat.toUpperCase()}.zip`, "application/zip");
        setNotice(`Đã xuất ${pages.length} trang thành ${imageFormat.toUpperCase()} và đóng gói ZIP.`);
        return;
      }
      if (mode === "wordtopdf") {
        const [{ renderAsync }, html2canvasModule, jspdfModule, { zipSync, strToU8 }] = await Promise.all([
          import("docx-preview"),
          import("html2canvas"),
          import("jspdf"),
          import("fflate"),
        ]);
        const html2canvas = html2canvasModule.default;
        const { jsPDF } = jspdfModule;
        const archive: Record<string, Uint8Array> = {};
        const report = ["Tep_Word,So_trang,Dung_luong_goc,Dung_luong_PDF,Trang_thai,Ghi_chu"];
        const renderHost = document.createElement("div");
        renderHost.className = "word-pdf-render-host";
        document.body.appendChild(renderHost);
        let success = 0;
        try {
          for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
            await checkpoint();
            const file = files[fileIndex];
            setNotice(`Đang chuyển ${fileIndex + 1}/${files.length}: ${file.name}`);
            setProgress({ current: fileIndex, total: files.length, label: `Đang dựng ${file.name}` });
            try {
              renderHost.innerHTML = "";
              await renderAsync(await file.arrayBuffer(), renderHost, renderHost, {
                className: "syland-export-docx",
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                breakPages: true,
                renderHeaders: true,
                renderFooters: true,
                renderFootnotes: true,
                useBase64URL: true,
              });
              await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
              const pageNodes = Array.from(renderHost.querySelectorAll<HTMLElement>("section.syland-export-docx"));
              const pages = pageNodes.length ? pageNodes : Array.from(renderHost.querySelectorAll<HTMLElement>("section.docx"));
              const outputPages = pages.length ? pages : [renderHost];
              let pdf: InstanceType<typeof jsPDF> | null = null;
              for (let pageIndex = 0; pageIndex < outputPages.length; pageIndex++) {
                await checkpoint();
                setProgress({
                  current: fileIndex + pageIndex / Math.max(1, outputPages.length),
                  total: files.length,
                  label: `Đang tạo PDF ${fileIndex + 1}/${files.length} · trang ${pageIndex + 1}/${outputPages.length}`,
                });
                const canvas = await html2canvas(outputPages[pageIndex], {
                  scale: wordPdfQuality === "high" ? 2 : 1.45,
                  backgroundColor: "#ffffff",
                  useCORS: true,
                  logging: false,
                  imageTimeout: 15000,
                });
                const orientation = canvas.width > canvas.height ? "landscape" : "portrait";
                if (!pdf) {
                  pdf = new jsPDF({ orientation, unit: "px", format: [canvas.width, canvas.height], hotfixes: ["px_scaling"], compress: true });
                } else {
                  pdf.addPage([canvas.width, canvas.height], orientation);
                }
                pdf.addImage(canvas.toDataURL("image/jpeg", wordPdfQuality === "high" ? .94 : .88), "JPEG", 0, 0, canvas.width, canvas.height, undefined, "FAST");
                canvas.width = 1; canvas.height = 1;
              }
              if (!pdf) throw new Error("Không tìm thấy trang Word để chuyển đổi.");
              const bytes = new Uint8Array(pdf.output("arraybuffer"));
              const outputName = `${file.name.replace(/\.docx$/i, "").replace(/[^a-zA-Z0-9À-ỹ_-]+/g, "_")}.pdf`;
              archive[outputName] = bytes;
              report.push(`"${file.name.replace(/"/g, '""')}",${outputPages.length},${file.size},${bytes.length},Thành công,`);
              success++;
            } catch (reason) {
              if (reason instanceof Error && reason.message === "SYLAND_TASK_CANCELLED") throw reason;
              console.error(reason);
              report.push(`"${file.name.replace(/"/g, '""')}",,${file.size},,Lỗi,"${(reason instanceof Error ? reason.message : "Không chuyển đổi được").replace(/"/g, '""')}"`);
            }
          }
        } finally {
          renderHost.remove();
        }
        archive["SYLAND_NHAT_KY_WORD_SANG_PDF.csv"] = strToU8("\uFEFF" + report.join("\r\n"));
        downloadBlob(zipSync(archive, { level: 0 }), `SYLAND_WORD_SANG_PDF_${files.length}_TEP.zip`, "application/zip");
        setProgress({ current: files.length, total: files.length, label: `Hoàn tất ${success}/${files.length} tệp` });
        setNotice(`Đã chuyển ${success}/${files.length} tệp Word. ZIP gồm các PDF thành công và nhật ký CSV.`);
        return;
      }
      const { PDFDocument, StandardFonts, degrees, rgb } = await import("pdf-lib");
      if (mode === "batch") {
        const { zipSync, strToU8 } = await import("fflate"); const archive: Record<string, Uint8Array> = {}; const report = ["Tep,Thao_tac,Trang,Dung_luong_goc,Dung_luong_moi,Trang_thai,Ghi_chu"];
        let success = 0;
        for (let index = 0; index < files.length; index++) {
          await checkpoint(); setProgress({ current: index + 1, total: files.length, label: `Đang xử lý tệp ${index + 1}/${files.length}` });
          const file = files[index]; setNotice(`Đang xử lý ${index + 1}/${files.length}: ${file.name}`);
          try {
            const document = await PDFDocument.load(await file.arrayBuffer(), { ignoreEncryption: false, updateMetadata: batchAction !== "sanitize" });
            if (batchAction === "rotate") document.getPages().forEach((page) => page.setRotation(degrees((page.getRotation().angle + angle) % 360)));
            if (batchAction === "sanitize") { document.setTitle(""); document.setAuthor(""); document.setSubject(""); document.setKeywords([]); document.setProducer("SY LAND"); document.setCreator("SY LAND"); }
            const bytes = await document.save({ useObjectStreams: true, addDefaultPage: false, objectsPerTick: 50 }); const suffix = batchAction === "rotate" ? `XOAY_${angle}` : batchAction === "sanitize" ? "LAM_SACH" : "TOI_UU";
            archive[`${baseName(file.name)}_${suffix}.pdf`] = bytes; report.push(`"${file.name.replace(/"/g, '""')}",${batchAction},${document.getPageCount()},${file.size},${bytes.length},Thành công,`); success++;
          } catch (reason) { console.error(reason); report.push(`"${file.name.replace(/"/g, '""')}",${batchAction},,,,Lỗi,"PDF bị khóa, hỏng hoặc chưa được hỗ trợ"`); }
        }
        archive["SYLAND_NHAT_KY_XU_LY.csv"] = strToU8("\uFEFF" + report.join("\r\n")); downloadBlob(zipSync(archive, { level: 0 }), `SYLAND_HANG_LOAT_${batchAction.toUpperCase()}.zip`, "application/zip");
        setNotice(`Hoàn tất ${success}/${files.length} tệp. ZIP gồm kết quả thành công và nhật ký CSV.`);
      } else if (mode === "imagetopdf") {
        const output = await PDFDocument.create();
        for (let index = 0; index < files.length; index++) {
          await checkpoint(); setProgress({ current: index + 1, total: files.length, label: `Đang ghép ảnh ${index + 1}/${files.length}` });
          setNotice(`Đang đưa ảnh ${index + 1}/${files.length} vào PDF…`);
          const file = files[index]; let bytes = await file.arrayBuffer(); let isPng = /png$/i.test(file.name) || file.type === "image/png";
          if (/webp$/i.test(file.name) || file.type === "image/webp") {
            const bitmap = await createImageBitmap(file); const canvas = document.createElement("canvas"); canvas.width = bitmap.width; canvas.height = bitmap.height; canvas.getContext("2d")!.drawImage(bitmap, 0, 0); bitmap.close();
            const jpeg = await new Promise<Blob>((resolve, reject) => canvas.toBlob((value) => value ? resolve(value) : reject(new Error("Không chuyển được WebP.")), "image/jpeg", .92)); bytes = await jpeg.arrayBuffer(); isPng = false;
          }
          const image = isPng ? await output.embedPng(bytes) : await output.embedJpg(bytes);
          if (imagePageMode === "original") {
            const scale = 72 / 96; const page = output.addPage([image.width * scale, image.height * scale]);
            page.drawImage(image, { x: 0, y: 0, width: page.getWidth(), height: page.getHeight() });
          } else {
            const landscape = image.width > image.height; const width = landscape ? 841.89 : 595.28; const height = landscape ? 595.28 : 841.89; const margin = 24;
            const scale = Math.min((width - margin * 2) / image.width, (height - margin * 2) / image.height); const drawWidth = image.width * scale; const drawHeight = image.height * scale;
            const page = output.addPage([width, height]); page.drawImage(image, { x: (width - drawWidth) / 2, y: (height - drawHeight) / 2, width: drawWidth, height: drawHeight });
          }
        }
        downloadBlob(await output.save({ useObjectStreams: true }), `SYLAND_ANH_THANH_PDF_${files.length}_TRANG.pdf`);
        setNotice(`Đã ghép ${files.length} ảnh thành PDF ${files.length} trang.`);
      } else if (mode === "merge") {
        const output = await PDFDocument.create();
        for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
          await checkpoint(); setProgress({ current: fileIndex + 1, total: files.length, label: `Đang nối PDF ${fileIndex + 1}/${files.length}` });
          const file = files[fileIndex];
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
      } else if (mode === "crop") {
        const document = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false }); const pointPerMm = 72 / 25.4;
        document.getPages().forEach((page) => {
          const left = cropMargins.left * pointPerMm; const right = cropMargins.right * pointPerMm; const top = cropMargins.top * pointPerMm; const bottom = cropMargins.bottom * pointPerMm;
          const width = page.getWidth() - left - right; const height = page.getHeight() - top - bottom;
          if (width <= 20 || height <= 20) throw new Error("Lề cắt lớn hơn kích thước trang.");
          page.setCropBox(left, bottom, width, height);
        });
        downloadBlob(await document.save({ useObjectStreams: true }), `${baseName(files[0].name)}_CAT_LE.pdf`);
        setNotice(`Đã cắt lề hiển thị cho ${document.getPageCount()} trang. Nội dung ẩn ngoài vùng cắt có thể vẫn tồn tại trong cấu trúc PDF.`);
      } else if (mode === "organize") {
        const source = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false });
        const order = previews.filter((item) => !removedPages.has(item.page)).map((item) => item.page - 1);
        if (!order.length) throw new Error("Không còn trang để xuất.");
        const output = await PDFDocument.create();
        for (let index = 0; index < order.length; index++) {
          await checkpoint();
          setProgress({ current: index + 1, total: order.length, label: `Đang tạo tệp kết quả ${index + 1}/${order.length}` });
          const [page] = await output.copyPages(source, [order[index]]); output.addPage(page);
        }
        downloadBlob(await output.save(), `${baseName(files[0].name)}_SAP_XEP.pdf`);
        setProgress({ current: order.length, total: order.length, label: `Hoàn tất · giữ ${order.length} trang · xóa ${removedPages.size} trang` });
        setNotice(`Đã xuất ${order.length} trang; loại bỏ ${removedPages.size} trang.`);
      } else if (mode === "resize") {
        const source = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false });
        const output = await PDFDocument.create();
        for (let pageIndex = 0; pageIndex < source.getPageCount(); pageIndex++) {
          await checkpoint(); setProgress({ current: pageIndex + 1, total: source.getPageCount(), label: `Đang chuyển A4 ${pageIndex + 1}/${source.getPageCount()}` });
          const sourcePage = source.getPages()[pageIndex];
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
          await checkpoint(); setProgress({ current: pageNumber, total: source.numPages, label: `Đang nén trang ${pageNumber}/${source.numPages}` });
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
        const cuts = splitKind === "cuts" ? [...new Set(cutPoints.split(",").map((value) => Number(value.trim())).filter((value) => Number.isInteger(value) && value > 0 && value < total))].sort((a, b) => a - b) : [];
        if (splitKind === "cuts" && !cuts.length) throw new Error(`Hãy nhập ít nhất một mốc từ 1 đến ${Math.max(1, total - 1)}.`);
        const cutGroups = splitKind === "cuts" ? [...cuts, total].map((end, index) => { const start = index ? cuts[index - 1] : 0; return Array.from({ length: end - start }, (_, page) => start + page); }) : [];
        const parityGroups = [Array.from({ length: total }, (_, index) => index).filter((index) => (index + 1) % 2 === 1), Array.from({ length: total }, (_, index) => index).filter((index) => (index + 1) % 2 === 0)].filter((group) => group.length);
        const groups = splitKind === "each" ? Array.from({ length: total }, (_, index) => [index]) : splitKind === "parity" ? parityGroups : splitKind === "fixed" ? Array.from({ length: Math.ceil(total / pagesPerPart) }, (_, part) => Array.from({ length: Math.min(pagesPerPart, total - part * pagesPerPart) }, (__, index) => part * pagesPerPart + index)) : splitKind === "cuts" ? cutGroups : splitKind === "visual" ? [[...selectedPages].sort((a, b) => a - b).map((page) => page - 1)].filter((group) => group.length) : parseRanges(ranges, total);
        if (splitKind === "visual" && !groups.length) throw new Error("Hãy chọn ít nhất một trang từ ảnh xem trước.");
        if (splitKind === "selected" || splitKind === "visual") {
          const output = await PDFDocument.create();
          const selected = [...new Set(groups.flat())];
          (await output.copyPages(source, selected)).forEach((page) => output.addPage(page));
          downloadBlob(await output.save(), `${baseName(files[0].name)}_TRANG_DA_CHON.pdf`);
          setNotice(`Đã xuất ${selected.length}/${total} trang vào một tệp.`);
        } else {
          const { zipSync } = await import("fflate");
          const archive: Record<string, Uint8Array> = {};
          for (let index = 0; index < groups.length; index++) {
            await checkpoint(); setProgress({ current: index + 1, total: groups.length, label: `Đang tạo phần ${index + 1}/${groups.length}` });
            const output = await PDFDocument.create();
            (await output.copyPages(source, groups[index])).forEach((page) => output.addPage(page));
            const label = splitKind === "each" ? `TRANG_${String(index + 1).padStart(3, "0")}` : splitKind === "parity" ? (index === 0 ? "TRANG_LE" : "TRANG_CHAN") : `PHAN_${String(index + 1).padStart(2, "0")}_TRANG_${groups[index][0] + 1}-${groups[index].at(-1)! + 1}`;
            const prefix = baseName(outputPrefix.trim() || baseName(files[0].name));
            archive[`${prefix}_${label}.pdf`] = await output.save();
          }
          downloadBlob(zipSync(archive, { level: 0 }), `${baseName(files[0].name)}_DA_TACH.zip`, "application/zip");
          setNotice(`Đã tạo ${groups.length} tệp PDF trong một gói ZIP.`);
        }
      }
    } catch (reason) {
      console.error(reason);
      setNotice(reason instanceof Error && reason.message === "SYLAND_TASK_CANCELLED" ? "Đã hủy tác vụ. Tệp gốc không bị thay đổi." : reason instanceof Error ? reason.message : "Không xử lý được. PDF có thể bị khóa, hỏng hoặc dùng cấu trúc chưa được hỗ trợ.");
    } finally { if (!taskControl.current.cancelled) setProgress((current) => ({ ...current, current: current.total, label: current.current < current.total ? "Hoàn tất tác vụ" : current.label })); setBusy(false); }
  }

  return (
    <section className="pdf-toolkit" id="cong-cu-pdf" aria-labelledby="pdf-toolkit-title">
      <div className="section-heading split-heading"><div><p className="section-kicker">Bộ công cụ PDF địa chính</p><h2 id="pdf-toolkit-title">Tách, chọn trang, nối và xoay PDF<br />ngay trên thiết bị.</h2></div><p>Thiết kế độc lập cho SỸ LAND. Tệp không được tải lên máy chủ; bản gốc không bị sửa hoặc ghi đè.<br /><b>{entitlements.reason}</b></p></div>
      <div className="pdf-tool-shell">
        <aside className="pdf-tool-sidebar" aria-label="Danh mục Bộ công cụ PDF">
        <div className="pdf-tool-finder"><label><span aria-hidden="true">⌕</span><input value={toolSearch} onChange={(event) => setToolSearch(event.target.value)} placeholder="Tìm theo công việc: xóa trang, OCR, nén ảnh…" /></label><div aria-label="Nhóm công cụ">{([{ key: "all", label: "Tổng quan", icon: "▦" }, { key: "split", label: "Tách trang", icon: "⑂" }, { key: "edit", label: "Chỉnh sửa", icon: "✎" }, { key: "convert", label: "Chuyển đổi", icon: "⇄" }, { key: "review", label: "Kiểm tra", icon: "✓" }, { key: "optimize", label: "Tối ưu", icon: "↯" }] as const).map((item) => <button className={toolCategory === item.key ? "active" : ""} type="button" key={item.key} onClick={() => setToolCategory(item.key)}><span aria-hidden="true">{item.icon}</span>{item.label}</button>)}</div></div>
        <div className="pdf-tool-list-heading"><div><b>Chọn công cụ</b><small>Danh sách được lọc theo nhóm nghiệp vụ ở trên</small></div><span>{toolCategory === "all" ? "Tất cả công cụ" : toolCategory === "split" ? "Tách trang" : toolCategory === "edit" ? "Chỉnh sửa" : toolCategory === "convert" ? "Chuyển đổi" : toolCategory === "review" ? "Kiểm tra" : "Tối ưu"}</span></div>
        <div className="pdf-tool-tabs" role="tablist" aria-label="Chọn công cụ PDF">
          {toolVisible("split", "tách xuất trang chia đều chẵn lẻ khoảng mốc") && <button className={mode === "split" ? "active" : ""} type="button" onClick={() => changeMode("split")}>Tách và xuất trang</button>}
          {toolVisible("edit", "nối ghép gộp pdf") && <button className={mode === "merge" ? "active" : ""} type="button" onClick={() => changeMode("merge")}>Nối PDF</button>}
          {toolVisible("edit", "xoay trang pdf") && <button className={mode === "rotate" ? "active" : ""} type="button" onClick={() => changeMode("rotate")}>Xoay PDF</button>}
          {toolVisible("edit", "sắp xếp xóa trang") && <button className={mode === "organize" ? "active" : ""} type="button" onClick={() => changeMode("organize")}>Sắp xếp · Xóa trang</button>}
          {toolVisible("edit", "cắt lề crop mép trắng pdf") && <button className={mode === "crop" ? "active" : ""} type="button" onClick={() => changeMode("crop")}>Cắt lề PDF</button>}
          {toolVisible("convert", "chuyển khổ a3 a4 in ký số") && <button className={mode === "resize" ? "active" : ""} type="button" onClick={() => changeMode("resize")}>Chuyển sang A4</button>}
          {toolVisible("edit", "số trang dấu bản quyền watermark") && <button className={mode === "annotate" ? "active" : ""} type="button" onClick={() => changeMode("annotate")}>Số trang · Dấu bản quyền</button>}
          {toolVisible("optimize", "tối ưu cấu trúc dung lượng") && <button className={mode === "optimize" ? "active" : ""} type="button" onClick={() => changeMode("optimize")}>Tối ưu PDF</button>}
          {toolVisible("review", "so sánh đối chiếu sai khác") && <button className={mode === "compare" ? "active" : ""} type="button" onClick={() => changeMode("compare")}>So sánh PDF</button>}
          {toolVisible("optimize", "làm sạch metadata riêng tư") && <button className={mode === "sanitize" ? "active" : ""} type="button" onClick={() => changeMode("sanitize")}>Làm sạch metadata</button>}
          {toolVisible("review", "ocr nhận dạng chữ scan tiếng việt") && <button className={mode === "ocr" ? "active" : ""} type="button" onClick={() => changeMode("ocr")}>OCR PDF scan</button>}
          {toolVisible("optimize", "nén ảnh giảm dung lượng pdf") && <button className={mode === "compress" ? "active" : ""} type="button" onClick={() => changeMode("compress")}>Nén ảnh PDF</button>}
          {toolVisible("optimize", "xử lý hàng loạt nhiều pdf nhật ký") && <button className={mode === "batch" ? "active" : ""} type="button" onClick={() => changeMode("batch")}>Xử lý PDF hàng loạt</button>}
          {toolVisible("review", "tìm trang trùng lặp kiểm tra") && <button className={mode === "dedupe" ? "active" : ""} type="button" onClick={() => changeMode("dedupe")}>Tìm trang trùng</button>}
          {toolVisible("convert", "pdf sang ảnh jpg png chuyển đổi") && <button className={mode === "toimage" ? "active" : ""} type="button" onClick={() => changeMode("toimage")}>PDF sang ảnh</button>}
          {toolVisible("convert", "ảnh jpg png sang pdf ghép ảnh chụp") && <button className={mode === "imagetopdf" ? "active" : ""} type="button" onClick={() => changeMode("imagetopdf")}>Ảnh sang PDF</button>}
          {toolVisible("convert", "word docx sang pdf hàng loạt chuyển đổi văn bản") && <button className={mode === "wordtopdf" ? "active" : ""} type="button" onClick={() => changeMode("wordtopdf")}>Word → PDF hàng loạt</button>}
          {!(["split", "edit", "convert", "review", "optimize"] as const).some((category) => toolVisible(category, category === "split" ? "tách xuất trang chia đều chẵn lẻ khoảng mốc" : category === "edit" ? "nối ghép gộp xoay sắp xếp xóa cắt lề crop số trang dấu bản quyền watermark" : category === "convert" ? "chuyển khổ a3 a4 pdf sang ảnh jpg png webp ảnh sang pdf ghép ảnh chụp word docx hàng loạt" : category === "review" ? "so sánh đối chiếu ocr scan trang trùng" : "tối ưu cấu trúc metadata nén dung lượng xử lý hàng loạt nhật ký")) && <p className="tool-not-found">Không tìm thấy công cụ phù hợp.</p>}
        </div>
        </aside>
        <div className="pdf-tool-body">
          <div className="pdf-workspace-summary"><div><span>KHÔNG GIAN XỬ LÝ</span><strong>{toolName}</strong></div><p><b>{batchLimitLabel[0].toUpperCase() + batchLimitLabel.slice(1)}/lượt</b><small>Xử lý tuần tự để giảm tải bộ nhớ · tệp luôn ở trên thiết bị</small></p></div>
          <div className="pdf-pick">
            <input ref={inputRef} type="file" accept={mode === "imagetopdf" ? "image/jpeg,image/png,image/webp,.jpg,.jpeg,.png,.webp" : mode === "wordtopdf" ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document,.docx" : "application/pdf,.pdf"} multiple={mode === "merge" || mode === "compare" || mode === "imagetopdf" || mode === "wordtopdf" || mode === "batch"} onChange={chooseFiles} />
            <span aria-hidden="true">{mode === "imagetopdf" ? "IMG" : mode === "wordtopdf" ? "DOCX" : "PDF"}</span><h3>{mode === "merge" ? "Chọn nhiều PDF theo thứ tự cần nối" : mode === "batch" ? "Chọn nhiều PDF để xử lý hàng loạt" : mode === "wordtopdf" ? "Chọn nhiều tệp Word DOCX" : mode === "imagetopdf" ? "Chọn ảnh JPG/PNG/WebP theo thứ tự" : mode === "compare" ? "Chọn hai phiên bản PDF" : "Chọn một tệp PDF"}</h3><p>{mode === "merge" ? `${batchLimitLabel} · 300 MB/lượt. Có thể di chuyển thứ tự sau khi chọn.` : mode === "batch" ? `${batchLimitLabel} · 300 MB/lượt. Kết quả và nhật ký CSV được đóng gói ZIP.` : mode === "wordtopdf" ? `${batchLimitLabel} · 300 MB/lượt. Chuyển cục bộ, không tải tài liệu lên máy chủ.` : mode === "imagetopdf" ? `${batchLimitLabel} · 500 MB/lượt. Có thể đổi thứ tự trước khi tạo PDF.` : mode === "compare" ? "So sánh nội dung văn bản theo từng trang; không tải tệp lên máy chủ." : "Xử lý cục bộ; khuyến nghị tệp không quá 100 MB."}</p>
            <button type="button" onClick={() => inputRef.current?.click()}>{mode === "imagetopdf" ? "Chọn ảnh từ thiết bị" : mode === "wordtopdf" ? "Chọn DOCX từ thiết bị" : "Chọn PDF từ thiết bị"}</button>
          </div>
          <div className="pdf-options">
            <div className="pdf-operation-head"><div><span>CÔNG CỤ ĐANG CHỌN</span><h3>{toolName}</h3></div><b>PHIÊN BẢN CHÍNH THỨC</b></div>
            <div className="pdf-universal-workflow"><div className={files.length ? "done" : "active"}><b>01</b><span>Chọn tệp</span><small>Kiểm tra định dạng và dung lượng</small></div><div className={files.length && (previews.length || imagePreviews.length) ? "done" : files.length ? "active" : ""}><b>02</b><span>Xem trước</span><small>Rà soát nguồn trước khi xử lý</small></div><div className={files.length ? "active" : ""}><b>03</b><span>Thiết lập</span><small>Chọn phạm vi và phương án</small></div><div className={busy ? "active" : notice.startsWith("Đã") ? "done" : ""}><b>04</b><span>Xuất kết quả</span><small>Tạo tệp mới, giữ nguyên bản gốc</small></div></div>
            {files.length > 0 && mode !== "wordtopdf" && <div className="pdf-source-preview"><header><div><b>Xem trước tệp nguồn</b><small>{mode === "imagetopdf" ? `${files.length} ảnh đã chọn` : previews.length ? `Đã đọc ${previews.length} trang đầu của ${files[0].name}` : files[0].name}</small></div>{mode !== "imagetopdf" && <button type="button" disabled={busy} onClick={() => void renderPreviews(files[0])}>Tải lại xem trước</button>}</header><div>{mode === "imagetopdf" ? imagePreviews.slice(0, 6).map((image, index) => <figure key={image}><img src={image} alt={`Ảnh nguồn ${index + 1}`} /><figcaption>Ảnh {index + 1}</figcaption></figure>) : previews.slice(0, 6).map((item) => <figure className={item.status === "error" ? "error" : item.status === "blank" ? "blank" : ""} key={item.page}>{item.image ? <img src={item.image} alt={`Trang ${item.page}`} /> : <span>Không xem được</span>}<figcaption>Trang {item.page}{item.status === "blank" ? " · Nghi trắng" : item.status === "error" ? " · Lỗi" : ""}</figcaption></figure>)}</div>{(previews.length > 6 || imagePreviews.length > 6) && <p>Đang hiển thị 6 mục đầu. Danh sách đầy đủ nằm trong phần thiết lập của công cụ.</p>}</div>}
            {mode === "split" && <><h3>Phương thức tách</h3><div className="pdf-choice-grid split-choice-grid"><label><input type="radio" checked={splitKind === "ranges"} onChange={() => setSplitKind("ranges")} /><span><b>Theo khoảng</b><small>Tạo một PDF cho mỗi khoảng</small></span></label><label><input type="radio" checked={splitKind === "selected"} onChange={() => setSplitKind("selected")} /><span><b>Nhập trang cần xuất</b><small>Gộp trang chọn vào một PDF</small></span></label><label><input type="radio" checked={splitKind === "visual"} onChange={() => setSplitKind("visual")} /><span><b>Chọn trực quan</b><small>Bấm vào ảnh từng trang</small></span></label><label><input type="radio" checked={splitKind === "cuts"} onChange={() => setSplitKind("cuts")} /><span><b>Tách tại mốc</b><small>Chia sau các trang chỉ định</small></span></label><label><input type="radio" checked={splitKind === "fixed"} onChange={() => setSplitKind("fixed")} /><span><b>Chia đều</b><small>Tự tách mỗi N trang</small></span></label><label><input type="radio" checked={splitKind === "parity"} onChange={() => setSplitKind("parity")} /><span><b>Trang chẵn · lẻ</b><small>Tạo hai PDF xen kẽ</small></span></label><label><input type="radio" checked={splitKind === "each"} onChange={() => setSplitKind("each")} /><span><b>Mỗi trang một tệp</b><small>Tải kết quả dạng ZIP</small></span></label></div>{splitKind !== "each" && splitKind !== "parity" && splitKind !== "visual" && splitKind !== "fixed" && splitKind !== "cuts" && <label className="range-input">Khoảng/trang cần xử lý<input value={ranges} onChange={(event) => setRanges(event.target.value)} placeholder="Ví dụ: 1-3, 5, 8-10" /><small>Dùng dấu phẩy để ngăn cách nhiều khoảng.</small></label>}{splitKind === "cuts" && <label className="range-input">Tách sau các trang<input value={cutPoints} onChange={(event) => setCutPoints(event.target.value)} placeholder="Ví dụ: 3, 7" /><small>Ví dụ 3, 7 sẽ tạo các phần 1–3, 4–7 và 8–hết.</small></label>}{splitKind === "fixed" && <div className="fixed-split-options"><label>Số trang mỗi phần<input type="number" min="1" max="100" value={pagesPerPart} onChange={(event) => setPagesPerPart(Math.min(100, Math.max(1, Number(event.target.value) || 1)))} /></label><p>{previews.length ? `Dự kiến tạo ${Math.ceil(previews.length / pagesPerPart)} phần từ ${previews.length} trang.` : "Chọn PDF để tính số phần dự kiến."}</p></div>}{splitKind === "parity" && <p className="parity-note">Tệp thứ nhất gồm trang 1, 3, 5…; tệp thứ hai gồm trang 2, 4, 6… Trang được đánh số theo thứ tự hiển thị của PDF.</p>}{splitKind === "visual" && <><div className="visual-select-toolbar"><b>Đã chọn {selectedPages.size} trang</b><button type="button" onClick={() => setSelectedPages(new Set(previews.map((item) => item.page)))}>Chọn tất cả</button><button type="button" onClick={() => setSelectedPages(new Set())}>Bỏ chọn</button></div><div className="visual-page-picker">{previews.map((item) => <button className={selectedPages.has(item.page) ? "selected" : ""} type="button" key={item.page} onClick={() => toggleSelected(item.page)}><span>{selectedPages.has(item.page) ? "✓" : item.page}</span><img src={item.image} alt={`Trang ${item.page}`} /><small>Trang {item.page}</small></button>)}</div>{!previews.length && <p className="pdf-empty">Chọn PDF để tạo ảnh xem trước.</p>}</>}{(splitKind === "ranges" || splitKind === "cuts" || splitKind === "fixed" || splitKind === "parity" || splitKind === "each") && <label className="range-input">Tiền tố tên tệp kết quả<input value={outputPrefix} maxLength={60} onChange={(event) => setOutputPrefix(event.target.value)} placeholder="Ví dụ: HOSO_DANGKY" /><small>Tên kết quả được tự loại bỏ ký tự không hợp lệ.</small></label>}</>}
            {mode === "merge" && <><h3>Thứ tự nối</h3><div className="merge-list">{files.length ? files.map((file, index) => <div key={`${file.name}-${file.lastModified}`}><span>{index + 1}</span><p><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => move(index, -1)} disabled={index === 0}>↑</button><button type="button" onClick={() => move(index, 1)} disabled={index === files.length - 1}>↓</button><button type="button" onClick={() => setFiles((current) => current.filter((_, itemIndex) => itemIndex !== index))}>×</button></div>) : <p className="pdf-empty">Chưa chọn tệp.</p>}</div></>}
            {mode === "rotate" && <><h3>Góc xoay toàn bộ trang</h3><div className="angle-options">{[90, 180, 270].map((value) => <button className={angle === value ? "active" : ""} type="button" key={value} onClick={() => setAngle(value)}>↻ {value}°</button>)}</div></>}
            {mode === "crop" && <><h3>Cắt lề hiển thị toàn bộ trang</h3><div className="crop-options">{(["top", "right", "bottom", "left"] as const).map((side) => <label key={side}>{side === "top" ? "Lề trên" : side === "right" ? "Lề phải" : side === "bottom" ? "Lề dưới" : "Lề trái"}<input type="number" min="0" max="100" step="1" value={cropMargins[side]} onChange={(event) => setCropMargins((current) => ({ ...current, [side]: Math.min(100, Math.max(0, Number(event.target.value) || 0)) }))} /><span>mm</span></label>)}</div><p className="compression-warning"><b>Lưu ý:</b> Cắt lề chỉ thay đổi vùng hiển thị/in. Nội dung nằm ngoài vùng cắt có thể vẫn còn trong cấu trúc PDF; không dùng chức năng này để che dữ liệu mật.</p></>}
            {mode === "organize" && <><h3>Đọc nội dung, xác nhận và chỉnh sửa trang</h3><div className="pdf-quality-workflow"><ol><li><b>1. Chọn PDF</b><span>Tệp được đọc cục bộ trên thiết bị</span></li><li><b>2. Đọc nội dung</b><span>Kiểm tra lớp chữ và hình ảnh từng trang</span></li><li><b>3. Xác nhận</b><span>Người dùng quyết định trắng, giữ lại hoặc xóa</span></li><li><b>4. Xuất bản mới</b><span>Không ghi đè tệp gốc</span></li></ol><div className="pdf-quality-actions"><p>Hệ thống không tự kết luận hoặc tự xóa trang. Hãy mở trang để xem toàn bộ nội dung; dùng OCR khi tài liệu là bản scan.</p><button type="button" disabled={!files[0] || busy} onClick={() => files[0] && void renderPreviews(files[0])}>Đọc lại toàn bộ</button><button type="button" disabled={!previews.length || busy} onClick={() => setRemovedPages(new Set())}>Giữ lại tất cả trang</button></div></div><div className="page-organizer">{previews.map((item, index) => <article className={`${removedPages.has(item.page) ? "removed" : ""} ${item.status === "review" ? "needs-review" : item.status === "error" ? "page-error" : ""} ${item.decision === "blank" ? "confirmed-blank" : ""}`} key={item.page}><button className="page-preview-open" type="button" onClick={() => setInspectedPage(item.page)} aria-label={`Mở xem toàn bộ trang ${item.page}`}>{item.image ? <img src={item.image} alt={`Xem trước trang ${item.page}`} /> : <b className="preview-error">Không xem được</b>}<span>{index + 1}</span>{item.status === "error" ? <em>TRANG LỖI</em> : item.decision === "blank" ? <em>ĐÃ XÁC NHẬN TRẮNG</em> : item.status === "review" ? <em>CẦN XÁC MINH</em> : <em className="has-content">CÓ NỘI DUNG</em>}</button><small>Trang gốc {item.page}{typeof item.inkRatio === "number" ? ` · hình/ảnh ${(item.inkRatio * 100).toFixed(2)}% · ${item.textChars || 0} ký tự` : ""}</small><p>{item.note}</p><div><button type="button" onClick={() => movePage(index, -1)} disabled={index === 0 || busy}>←</button><button type="button" onClick={() => movePage(index, 1)} disabled={index === previews.length - 1 || busy}>→</button><button type="button" className="inspect-page" onClick={() => setInspectedPage(item.page)}>Xem & xác nhận</button></div></article>)}</div>{!previews.length && <p className="pdf-empty">Chọn PDF để đọc nội dung và tạo bản xem trước từng trang.</p>}{inspectedPage !== null && (() => { const item = previews.find((preview) => preview.page === inspectedPage); if (!item) return null; const readableText = item.ocrText || item.extractedText || ""; return <div className="page-inspector" role="dialog" aria-modal="true" aria-label={`Kiểm tra trang ${item.page}`}><div className="page-inspector-panel"><header><div><b>Kiểm tra trang {item.page}</b><small>Chỉ xác nhận trang trắng sau khi đã xem hình ảnh và nội dung đọc được</small></div><button type="button" onClick={() => setInspectedPage(null)} aria-label="Đóng">×</button></header><div className="page-inspector-body"><figure>{item.image ? <img src={item.image} alt={`Toàn bộ trang ${item.page}`} /> : <span>Không dựng được hình ảnh trang</span>}</figure><section><h4>Nội dung đọc được</h4><div className="page-text-content">{readableText || "Chưa đọc được lớp chữ. Nếu đây là PDF scan, hãy nhấn “Đọc OCR trang này” rồi kiểm tra lại hình ảnh."}</div><dl><div><dt>Ký tự lớp chữ</dt><dd>{item.textChars || 0}</dd></div><div><dt>Tỷ lệ hình/mực</dt><dd>{typeof item.inkRatio === "number" ? `${(item.inkRatio * 100).toFixed(3)}%` : "Không xác định"}</dd></div><div><dt>Kết luận</dt><dd>{item.decision === "blank" ? "Đã xác nhận trang trắng" : item.decision === "content" ? "Trang có nội dung/giữ lại" : "Chưa xác nhận"}</dd></div></dl><button className="ocr-page-action" type="button" disabled={ocrPage === item.page} onClick={() => void readPageWithOcr(item.page)}>{ocrPage === item.page ? "Đang OCR…" : "Đọc OCR trang này"}</button></section></div><footer><button type="button" onClick={() => { decidePage(item.page, "content"); setInspectedPage(null); }}>Có nội dung · Giữ lại</button><button type="button" className="confirm-blank" onClick={() => decidePage(item.page, "blank")}>Xác nhận trang trắng</button><button type="button" className="delete-confirmed-page" disabled={item.decision !== "blank"} onClick={() => { setRemovedPages((current) => new Set(current).add(item.page)); setInspectedPage(null); }}>Xóa trang đã xác nhận</button></footer></div></div>; })()}</>}
            {mode === "resize" && <div className="resize-info"><span>A3</span><b>→</b><span>A4</span><div><h3>Chuyển toàn bộ trang về A4</h3><p>Giữ tỷ lệ nội dung, tự nhận hướng dọc/ngang và căn giữa để thuận tiện in hoặc ký số.</p></div></div>}
            {mode === "annotate" && <><h3>Đánh dấu tài liệu</h3><div className="annotate-options"><label className="annotate-check"><input type="checkbox" checked={addWatermark} onChange={(event) => setAddWatermark(event.target.checked)} /><span>Thêm dấu bản quyền</span></label><label>Nội dung dấu<input value={watermark} maxLength={60} disabled={!addWatermark} onChange={(event) => setWatermark(event.target.value)} placeholder="Ví dụ: SỸ LAND - BẢN KIỂM TRA" /><small>Chữ tiếng Việt được tự chuyển sang không dấu để tương thích PDF.</small></label><label className="annotate-check"><input type="checkbox" checked={addNumbers} onChange={(event) => setAddNumbers(event.target.checked)} /><span>Thêm số trang</span></label><div className="number-config"><label>Bắt đầu từ<input type="number" min="1" value={numberStart} disabled={!addNumbers} onChange={(event) => setNumberStart(Math.max(1, Number(event.target.value) || 1))} /></label><label>Vị trí<select value={numberPosition} disabled={!addNumbers} onChange={(event) => setNumberPosition(event.target.value as "bottom" | "top")}><option value="bottom">Giữa chân trang</option><option value="top">Giữa đầu trang</option></select></label></div></div></>}
            {mode === "optimize" && <div className="optimize-info"><span>⇣</span><div><h3>Tối ưu cấu trúc PDF</h3><p>Sắp xếp lại đối tượng và luồng dữ liệu để giảm dung lượng khi có thể. Công cụ không hạ độ phân giải ảnh nên giữ nguyên chất lượng hồ sơ scan.</p><ul><li>Không xóa nội dung</li><li>Không giảm chất lượng ảnh</li><li>Kết quả phụ thuộc cấu trúc tệp gốc</li></ul></div></div>}
            {mode === "compare" && <><div className="compare-files">{files.map((file, index) => <article key={`${file.name}-${file.lastModified}`}><span>{index ? "B" : "A"}</span><div><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></div></article>)}</div>{compareResult && <div className="compare-result"><div><strong>{compareResult.pagesA}</strong><span>Trang bản A</span></div><div><strong>{compareResult.pagesB}</strong><span>Trang bản B</span></div><div><strong>{compareResult.samePages}</strong><span>Trùng khớp</span></div><div><strong>{compareResult.reviewPages}</strong><span>Khác biệt nhỏ</span></div><div><strong>{compareResult.changedPages.length}</strong><span>Thay đổi lớn</span></div><div><strong>{compareResult.scanPages}</strong><span>Cần OCR</span></div><p><b>Trang thay đổi lớn/thêm/bớt:</b> {compareResult.changedPages.length ? compareResult.changedPages.slice(0, 100).join(", ") : "Không có"}{compareResult.changedPages.length > 100 ? "…" : ""}</p><div className="compare-detail-list">{compareResult.details.filter((item) => item.status !== "same").slice(0, 40).map((item) => <span className={`compare-${item.status}`} key={item.page}>Trang {item.page}: {item.similarity === null ? "—" : `${item.similarity}%`} · {item.note}</span>)}</div><small>Độ phủ văn bản: {compareResult.textCoverage}%. Ngưỡng: ≥99,5% trùng khớp; 92–99,4% cần kiểm tra; dưới 92% thay đổi đáng kể.</small></div>}</>}
            {mode === "sanitize" && <div className="optimize-info"><span>⌫</span><div><h3>Làm sạch metadata</h3><p>Xóa tiêu đề, tác giả, chủ đề và từ khóa ẩn trong PDF trước khi gửi cho người khác. Nội dung và hình ảnh hiển thị được giữ nguyên.</p><ul><li>Không xóa chữ ký hoặc chú thích hiển thị</li><li>Không thay thế việc kiểm tra dữ liệu cá nhân trong nội dung</li><li>Tạo tệp mới, không ghi đè bản gốc</li></ul></div></div>}
            {mode === "ocr" && <div className="optimize-info"><span>OCR</span><div><h3>Nhận dạng chữ tiếng Việt</h3><p>Chuyển từng trang scan thành văn bản TXT có phân cách trang. Xử lý cục bộ, tối đa 40 trang mỗi lượt.</p><ul><li>Kết quả OCR cần được người dùng kiểm tra lại</li><li>Không tự động thay thế dữ liệu hồ sơ gốc</li><li>Độ chính xác phụ thuộc chất lượng bản scan</li></ul></div></div>}
            {mode === "compress" && <><h3>Mức nén hình ảnh</h3><div className="compression-options">{([{ key: "small", title: "Dung lượng nhỏ", note: "Phù hợp gửi xem nhanh" }, { key: "balanced", title: "Cân bằng", note: "Khuyến nghị cho hồ sơ thông thường" }, { key: "clear", title: "Rõ nét", note: "Ưu tiên chữ nhỏ và bản đồ" }] as const).map((item) => <label key={item.key}><input type="radio" checked={compressQuality === item.key} onChange={() => setCompressQuality(item.key)} /><span><b>{item.title}</b><small>{item.note}</small></span></label>)}</div><p className="compression-warning"><b>Lưu ý:</b> Bản nén được dựng lại từ ảnh từng trang nên có thể mất lớp chữ tìm kiếm, liên kết, biểu mẫu và trạng thái chữ ký số. Luôn giữ tệp gốc.</p></>}
            {mode === "dedupe" && <div className="dedupe-panel"><div className="optimize-info"><span>≡</span><div><h3>Phát hiện trang giống hệt nhau</h3><p>Tạo ảnh kiểm tra cục bộ rồi so sánh SHA-256 của từng trang. Chỉ đánh dấu khi hình ảnh trang khớp hoàn toàn ở cùng độ phân giải kiểm tra.</p><ul><li>Giữ trang xuất hiện đầu tiên</li><li>Không tự xóa trước khi người dùng xác nhận tải</li><li>Tối đa 150 trang mỗi lượt</li></ul></div></div>{duplicateScanned && <div className="duplicate-results">{duplicateGroups.length ? duplicateGroups.map((group, index) => <p key={group.join("-")}><b>Nhóm {index + 1}</b><span>Giữ trang {group[0]}</span><span>Đề xuất bỏ: {group.slice(1).join(", ")}</span></p>) : <p className="no-duplicate"><b>Không có trang trùng</b><span>Không phát hiện trang có hình ảnh giống hệt nhau.</span></p>}</div>}</div>}
            {mode === "toimage" && <><h3>Xuất trang PDF thành ảnh</h3><div className="image-export-options"><div><b>Định dạng</b><label><input type="radio" checked={imageFormat === "jpg"} onChange={() => setImageFormat("jpg")} /> JPG</label><label><input type="radio" checked={imageFormat === "png"} onChange={() => setImageFormat("png")} /> PNG</label></div><div><b>Độ phân giải</b>{([{ key: "screen", label: "Xem nhanh" }, { key: "standard", label: "Tiêu chuẩn" }, { key: "high", label: "Rõ nét" }] as const).map((item) => <label key={item.key}><input type="radio" checked={imageResolution === item.key} onChange={() => setImageResolution(item.key)} /> {item.label}</label>)}</div></div><label className="range-input">Trang cần xuất<input value={imageRanges} onChange={(event) => setImageRanges(event.target.value)} placeholder="Ví dụ: 1-3, 5, 8-10" /><small>Tối đa 80 trang mỗi lượt. PNG rõ hơn nhưng thường có dung lượng lớn hơn JPG.</small></label></>}
            {mode === "imagetopdf" && <><h3>Sắp xếp ảnh và chọn khổ trang</h3><div className="image-page-mode"><label><input type="radio" checked={imagePageMode === "a4"} onChange={() => setImagePageMode("a4")} /><span><b>Căn vừa A4</b><small>Tự nhận hướng dọc/ngang, có lề an toàn</small></span></label><label><input type="radio" checked={imagePageMode === "original"} onChange={() => setImagePageMode("original")} /><span><b>Theo tỷ lệ ảnh</b><small>Giữ toàn bộ ảnh, không thêm lề A4</small></span></label></div><div className="merge-list image-merge-list">{files.length ? files.map((file, index) => <div key={`${file.name}-${file.lastModified}`}><span>{index + 1}</span><p><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => move(index, -1)} disabled={index === 0}>↑</button><button type="button" onClick={() => move(index, 1)} disabled={index === files.length - 1}>↓</button><button type="button" onClick={() => setFiles((current) => current.filter((_, itemIndex) => itemIndex !== index))}>×</button></div>) : <p className="pdf-empty">Chưa chọn ảnh.</p>}</div></>}
            {mode === "wordtopdf" && <div className="word-pdf-tool"><div className="pdf-quality-workflow"><ol><li><b>1. Chọn DOCX</b><span>{batchLimitLabel} mỗi lượt</span></li><li><b>2. Xem trước</b><span>Kiểm tra tệp đầu tiên và bố cục</span></li><li><b>3. Chuyển tuần tự</b><span>Có tiến độ, tạm dừng và hủy</span></li><li><b>4. Tải ZIP</b><span>PDF và nhật ký CSV</span></li></ol></div><h3>Chất lượng PDF</h3><div className="word-pdf-quality"><label><input type="radio" checked={wordPdfQuality === "standard"} onChange={() => setWordPdfQuality("standard")} /><span><b>Tiêu chuẩn</b><small>Nhanh hơn, phù hợp văn bản thông thường</small></span></label><label><input type="radio" checked={wordPdfQuality === "high"} onChange={() => setWordPdfQuality("high")} /><span><b>Rõ nét</b><small>Chữ và hình rõ hơn, xử lý lâu hơn</small></span></label></div><div className="merge-list word-pdf-list">{files.length ? files.map((file, index) => <div key={`${file.name}-${file.lastModified}`}><span>{index + 1}</span><p><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" disabled={busy || wordPreviewing} onClick={() => { setFiles((current) => current.filter((_, itemIndex) => itemIndex !== index)); if (index === 0 && wordPreviewRef.current) wordPreviewRef.current.innerHTML = ""; }}>×</button></div>) : <p className="pdf-empty">Chưa chọn tệp DOCX.</p>}</div><section className="word-pdf-preview"><header><div><b>Xem trước tệp Word đầu tiên</b><small>{files[0]?.name || "Chọn DOCX để xem bố cục trước khi chuyển"}</small></div>{files[0] && <button type="button" disabled={wordPreviewing || busy} onClick={() => void renderWordPreview(files[0])}>{wordPreviewing ? "Đang dựng…" : "Xem lại"}</button>}</header><div ref={wordPreviewRef} className="word-preview-surface" /></section><p className="compression-warning"><b>Lưu ý:</b> Tệp DOCX được dựng lại bằng trình duyệt. Font chưa cài, macro, theo dõi thay đổi, biểu đồ hoặc bố cục Word phức tạp có thể khác bản gốc; hãy kiểm tra bản xem trước và PDF kết quả trước khi phát hành.</p><p className="batch-log-note">Mỗi DOCX tạo một PDF. ZIP kèm SYLAND_NHAT_KY_WORD_SANG_PDF.csv để đối chiếu tệp thành công và tệp lỗi.</p></div>}
            {mode === "batch" && <><h3>Xử lý đồng loạt · {batchLimitLabel}</h3><div className="batch-pdf-actions"><label><input type="radio" checked={batchAction === "optimize"} onChange={() => setBatchAction("optimize")} /><span><b>Tối ưu cấu trúc</b><small>Giữ nguyên chất lượng ảnh</small></span></label><label><input type="radio" checked={batchAction === "sanitize"} onChange={() => setBatchAction("sanitize")} /><span><b>Làm sạch metadata</b><small>Xóa tác giả, tiêu đề và từ khóa</small></span></label><label><input type="radio" checked={batchAction === "rotate"} onChange={() => setBatchAction("rotate")} /><span><b>Xoay đồng loạt</b><small>Dùng góc xoay bên dưới</small></span></label></div>{batchAction === "rotate" && <div className="angle-options batch-angle-options">{[90, 180, 270].map((value) => <button className={angle === value ? "active" : ""} type="button" key={value} onClick={() => setAngle(value)}>↻ {value}°</button>)}</div>}<div className="merge-list batch-pdf-list">{files.length ? files.map((file, index) => <div key={`${file.name}-${file.lastModified}`}><span>{index + 1}</span><p><b>{file.name}</b><small>{(file.size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => setFiles((current) => current.filter((_, itemIndex) => itemIndex !== index))}>×</button></div>) : <p className="pdf-empty">Chưa chọn PDF.</p>}</div><p className="batch-log-note">ZIP luôn kèm SYLAND_NHAT_KY_XU_LY.csv, ghi tệp thành công, tệp lỗi, số trang và dung lượng trước/sau.</p></>}
            {mode !== "merge" && mode !== "imagetopdf" && mode !== "wordtopdf" && mode !== "batch" && mode !== "organize" && mode !== "compare" && files[0] && <div className="single-pdf"><span>✓</span><p><b>{files[0].name}</b><small>{(files[0].size / 1024 / 1024).toFixed(1)} MB</small></p><button type="button" onClick={() => setFiles([])}>×</button></div>}
            {progress.total > 0 && <div className="pdf-task-progress" role="status" aria-live="polite"><div><span style={{ width: `${Math.min(100, Math.round(progress.current / progress.total * 100))}%` }} /></div><p>{progress.label} · {Math.min(100, Math.round(progress.current / progress.total * 100))}%</p></div>}
            <div className="pdf-run"><p className={notice.includes("Không") ? "error" : ""}>{notice || (files.length ? `${files.length} tệp · ${(totalSize / 1024 / 1024).toFixed(1)} MB` : "Kết quả sẽ được tải về thiết bị.")}</p><div className="pdf-run-actions">{busy && <><button className="secondary-pdf-action" type="button" onClick={togglePause}>{paused ? "Tiếp tục" : "Tạm dừng"}</button><button className="danger-pdf-action" type="button" onClick={cancelTask}>Hủy tác vụ</button></>}{mode === "compare" && compareResult && <button className="secondary-pdf-action" type="button" onClick={() => { const rows = ["Trang,Ty_le_tuong_dong,Phan_loai,Ghi_chu", ...compareResult.details.map((item) => `${item.page},${item.similarity === null ? "" : item.similarity + "%"},${item.status},\"${item.note.replace(/\"/g, '\"\"')}\"`)]; rows.unshift(`Tep_B,\"${files[1].name.replace(/\"/g, '\"\"')}\"`, `Tep_A,\"${files[0].name.replace(/\"/g, '\"\"')}\"`); downloadText("\uFEFF" + rows.join("\r\n"), `SYLAND_BAO_CAO_SO_SANH_${Date.now()}.csv`, "text/csv;charset=utf-8"); }}>Tải báo cáo CSV</button>}{mode === "dedupe" && duplicateGroups.length > 0 && <button className="secondary-pdf-action" type="button" onClick={() => { void (async () => { setBusy(true); try { const { PDFDocument } = await import("pdf-lib"); const source = await PDFDocument.load(await files[0].arrayBuffer(), { ignoreEncryption: false }); const remove = new Set(duplicateGroups.flatMap((group) => group.slice(1))); const keep = source.getPageIndices().filter((index) => !remove.has(index + 1)); const output = await PDFDocument.create(); (await output.copyPages(source, keep)).forEach((page) => output.addPage(page)); downloadBlob(await output.save({ useObjectStreams: true }), `${baseName(files[0].name)}_BO_TRANG_TRUNG.pdf`); setNotice(`Đã tạo PDF mới gồm ${keep.length} trang, loại ${remove.size} trang lặp. Tệp gốc không thay đổi.`); } catch (reason) { console.error(reason); setNotice("Không tạo được PDF đã loại trang trùng."); } finally { setBusy(false); } })(); }}>Tải PDF đã loại trùng</button>}<button type="button" disabled={busy || !files.length || (mode === "compare" && files.length !== 2) || (mode === "organize" && !previews.length)} onClick={() => void run()}>{busy ? "Đang xử lý…" : mode === "merge" ? "Nối PDF" : mode === "rotate" ? "Xoay và tải PDF" : mode === "organize" ? `Xuất PDF mới · xóa ${removedPages.size} trang` : mode === "resize" ? "Chuyển và tải PDF A4" : mode === "annotate" ? "Thêm dấu và tải PDF" : mode === "optimize" ? "Tối ưu và tải PDF" : mode === "compare" ? "So sánh hai PDF" : mode === "sanitize" ? "Làm sạch và tải PDF" : mode === "ocr" ? "OCR và tải TXT" : mode === "compress" ? "Nén và tải PDF" : mode === "dedupe" ? "Quét trang trùng" : mode === "toimage" ? "Xuất ảnh và tải ZIP" : mode === "imagetopdf" ? "Ghép ảnh và tải PDF" : mode === "wordtopdf" ? "Chuyển Word và tải ZIP" : mode === "crop" ? "Cắt lề và tải PDF" : mode === "batch" ? "Xử lý và tải ZIP" : "Tách và tải kết quả"}</button></div></div>
          </div>
        </div>
      </div>
    </section>
  );
}
