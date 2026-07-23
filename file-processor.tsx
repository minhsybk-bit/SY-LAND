"use client";

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";
import { useEntitlements } from "./entitlements";

type CellValue = string | number | boolean | null;

type ProcessedResult = {
  kind: "word" | "pdf" | "excel";
  title: string;
  description: string;
  metrics: Array<{ label: string; value: string }>;
  text?: string;
  rows?: CellValue[][];
  sheetName?: string;
  warning?: string;
  ocr?: boolean;
};

type ProcessingProgress = {
  label: string;
  percent: number;
};

type LandFields = {
  communeCode: string;
  mapSheet: string;
  parcel: string;
  area: string;
  prefix: "CHUACOGIAY" | "COGCN";
  suffix: string;
};

type BatchItem = {
  id: string;
  file: File;
  status: "waiting" | "processing" | "done" | "review" | "error";
  fields: LandFields;
  message: string;
};

type MasterSheet = {
  fileName: string;
  sheetName: string;
  sheets: Array<{ name: string; rows: CellValue[][] }>;
  rows: CellValue[][];
  headers: string[];
  headerRow: number;
  codeColumn: number;
  sheetColumn: number;
  parcelColumn: number;
};

type HistoryRecord = {
  id: string;
  fileName: string;
  fileType: string;
  processedAt: string;
  result: string;
  suggestedName: string;
};

type LocationProfile = {
  provinceNew: string;
  provinceOld: string;
  communeNew: string;
  communeOld: string;
  communeCode: string;
  villageNew: string;
  villageOld: string;
  aliases: string;
};

const EMPTY_LOCATION: LocationProfile = { provinceNew: "", provinceOld: "", communeNew: "", communeOld: "", communeCode: "", villageNew: "", villageOld: "", aliases: "" };

const ACCEPTED = ".docx,.pdf,.xlsx,.xls,.csv";
const MAX_FILE_SIZE = 20 * 1024 * 1024;
const MAX_BATCH_TOTAL_SIZE = 500 * 1024 * 1024;

function formatBytes(bytes: number) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function safeBaseName(name: string) {
  return name.replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9À-ỹ_-]+/g, "_");
}

function escapeCsv(value: CellValue) {
  const text = value == null ? "" : String(value);
  return `"${text.replaceAll('"', '""')}"`;
}

function normalizeHeader(value: CellValue) {
  return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase("vi-VN").replace(/đ/g, "d").replace(/[^a-z0-9]/g, "");
}

function normalizeKeyNumber(value: CellValue) {
  return String(value ?? "").replace(/\D/g, "").replace(/^0+(?=\d)/, "");
}

function buildParcelKey(code: CellValue, sheet: CellValue, parcel: CellValue) {
  const codeDigits = String(code ?? "").replace(/\D/g, "");
  const normalizedSheet = normalizeKeyNumber(sheet);
  const normalizedParcel = normalizeKeyNumber(parcel);
  return codeDigits && normalizedSheet && normalizedParcel ? `${codeDigits.padStart(5, "0")}|${normalizedSheet}|${normalizedParcel}` : "";
}

function findColumn(headers: string[], kind: "code" | "sheet" | "parcel") {
  const aliases = { code: ["maxa", "maxaphuong", "maxadonvihanhchinh", "maxacuthe"], sheet: ["soto", "sotobando", "sotobandodiachinh", "sohieutrenbandodiachinh"], parcel: ["sothua", "sothuadat", "sothututhua"] }[kind];
  return headers.findIndex((header) => aliases.some((alias) => header === alias || header.includes(alias)));
}

function detectHeaderRow(rows: CellValue[][]) {
  let bestRow = 0;
  let bestScore = -1;
  rows.slice(0, 20).forEach((row, rowIndex) => {
    const headers = row.map(normalizeHeader);
    const score = (["code", "sheet", "parcel"] as const).reduce((total, kind) => total + (findColumn(headers, kind) >= 0 ? 3 : 0), 0) + row.filter((cell) => String(cell).trim()).length * 0.02;
    if (score > bestScore) { bestScore = score; bestRow = rowIndex; }
  });
  return bestRow;
}

const VILLAGE_CODE_MAP: Array<{ code: string; names: string[] }> = [
  { code: "02146", names: ["khuổi phầy", "kim vân", "quốc tuấn", "nà mỏ", "khuổi hát", "kim hỷ"] },
  { code: "02143", names: ["nà làng", "bản giang", "pàn xả", "vằng khít", "khuổi nộc", "lương thượng"] },
  { code: "02140", names: ["bản kén", "nặm cà", "chợ mới", "tân an", "nà diệc", "bản sảng", "nà lẹng", "cốc phia", "nà dường", "văn lang"] },
];

function normalizeSearchText(value: string) {
  return value.normalize("NFC").toLocaleLowerCase("vi-VN").replace(/\s+/g, " ").trim();
}

function parseConfiguredVillageCodes(profile?: LocationProfile) {
  if (!profile) return [] as Array<{ code: string; names: string[] }>;
  const mappings: Array<{ code: string; names: string[] }> = [];
  profile.aliases.split(/\n|\|/).forEach((line) => {
    const match = line.trim().match(/^([0-9]{4,5})\s*[:=]\s*(.+)$/);
    if (!match) return;
    mappings.push({
      code: match[1].padStart(5, "0"),
      names: match[2].split(/[,;]/).map(normalizeSearchText).filter((name) => name.length >= 2),
    });
  });
  return mappings;
}

function findVillageCode(text: string, profile?: LocationProfile) {
  const normalized = normalizeSearchText(text);
  const configuredMappings = parseConfiguredVillageCodes(profile);
  const configuredMatch = configuredMappings.find((entry) => entry.names.some((name) => normalized.includes(name)));
  if (configuredMatch) return configuredMatch.code;

  if (profile?.communeCode) {
    // Ưu tiên tên thôn/bản/xóm trước tên xã theo yêu cầu nghiệp vụ.
    const villageNames = [profile.villageOld, profile.villageNew]
      .map(normalizeSearchText).filter((name) => name.length >= 2);
    if (villageNames.some((name) => normalized.includes(name))) return profile.communeCode.padStart(5, "0");
    const configuredNames = [profile.communeOld, profile.communeNew, profile.provinceOld, profile.provinceNew, ...profile.aliases.split(/[,;\n]/).filter((name) => !/^\s*[0-9]{4,5}\s*[:=]/.test(name))]
      .map(normalizeSearchText).filter((name) => name.length >= 2);
    if (!configuredNames.length || configuredNames.some((name) => normalized.includes(name))) return profile.communeCode.padStart(5, "0");
    return "";
  }
  return VILLAGE_CODE_MAP.find((entry) => entry.names.some((name) => normalized.includes(name)))?.code || "";
}

function extractNumberedSection(text: string, section: "a" | "b") {
  const normalized = text.replace(/\r/g, "\n");
  const marker = section === "a" ? "a" : "b";
  const nextMarker = section === "a" ? "b" : "c";
  const pattern = new RegExp(
    `(?:^|\\n)\\s*2\\s*[.\\-)]?\\s*${marker}\\s*[.\\-):]?\\s*([\\s\\S]*?)(?=(?:\\n\\s*2\\s*[.\\-)]?\\s*${nextMarker}\\s*[.\\-):]?)|$)`,
    "iu",
  );
  return normalized.match(pattern)?.[1]?.trim() || "";
}

function sanitizeSuffix(value: string) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/gi, "D")
    .toUpperCase().replace(/[^A-Z0-9_-]/g, "").replace(/^_+|_+$/g, "").slice(0, 24);
}

function cleanNumber(value?: string) {
  return (value || "").replace(/\D/g, "").replace(/^0+(?=\d)/, "").trim();
}

function normalizeCommuneCode(value?: string) {
  const digits = (value || "").replace(/\D/g, "");
  return digits && digits.length <= 5 ? digits.padStart(5, "0") : "";
}

function extractLandFields(text: string, fileName: string, profile?: LocationProfile): LandFields {
  const fileMatch = fileName.match(/(?:CHUACOGIAY|CHUACAPGIAY|COGCN)[_-]([0-9]{4,5})[_-]([0-9]{1,6})[_-]([0-9]{1,7})(?=[_.-]|$)/i);
  const explicitCode = text.match(/m[aã]\s*x[aã]\s*[:\-]?\s*([0-9]{4,5})/iu)?.[1];
  const section2a = extractNumberedSection(text, "a");
  const section2b = extractNumberedSection(text, "b");
  const parcelAddress = text.match(/(?:địa\s*chỉ\s*thửa\s*đất|thửa\s*đất\s*tại)([\s\S]{0,280})/iu)?.[1] || "";
  // Mục 2.b là nguồn xác định địa bàn; tên thôn trong cấu hình được dò trước tên xã.
  const communeCode = normalizeCommuneCode(fileMatch?.[1] || explicitCode || findVillageCode(section2b, profile) || findVillageCode(parcelAddress, profile) || findVillageCode(text, profile) || profile?.communeCode);
  // Mục 2.a là nguồn chính của số tờ/số thửa; toàn văn chỉ là phương án dự phòng cho mẫu khác.
  const mapSheetFromText = (section2a || text).match(/tờ\s*bản\s*đồ(?:\s*địa\s*chính)?\s*(?:số)?\s*[:\-]?\s*([0-9]{1,6})/iu)?.[1]
    || text.match(/tờ\s*bản\s*đồ(?:\s*địa\s*chính)?\s*(?:số)?\s*[:\-]?\s*([0-9]{1,6})/iu)?.[1];
  const parcelFromText = (section2a || text).match(/(?:thửa\s*đất|thửa)\s*(?:số)?\s*[:\-]?\s*([0-9]{1,7})/iu)?.[1]
    || text.match(/(?:thửa\s*đất|thửa)\s*(?:số)?\s*[:\-]?\s*([0-9]{1,7})/iu)?.[1];
  const area = text.match(/diện\s*tích(?:\s*thửa\s*đất)?\s*[:\-]?\s*([0-9][0-9.,\s]*)\s*m(?:2|²)/iu)?.[1];
  const forestry = /\b(?:RSX|RPH|RDD)\b/i.test(`${text} ${fileName}`);
  let mapSheet = cleanNumber(fileMatch?.[2] || mapSheetFromText);
  if (forestry && /^[123]$/.test(mapSheet)) mapSheet = `${mapSheet}10000`;

  return {
    communeCode,
    mapSheet,
    parcel: cleanNumber(fileMatch?.[3] || parcelFromText),
    area: cleanNumber(area),
    prefix: /^COGCN[_-]/i.test(fileName) ? "COGCN" : "CHUACOGIAY",
    suffix: /[_-]TBXN(?:[_.-]|$)/i.test(fileName) ? "TBXN" : /[_-]DDK(?:[_.-]|$)/i.test(fileName) ? "DDK" : "GT",
  };
}

function reviewLandFields(fields: LandFields, text: string, fileName: string) {
  const notes: string[] = [];
  if (!/^\d{5}$/.test(fields.communeCode)) notes.push("mã xã phải có đúng 5 chữ số");
  if (!fields.mapSheet) notes.push("thiếu số tờ bản đồ");
  if (!fields.parcel) notes.push("thiếu số thửa");
  if (/CHUACAPGIAY/i.test(fileName)) notes.push("đã chuẩn hóa tiền tố cũ CHUACAPGIAY thành CHUACOGIAY");
  if (/\b(?:RSX|RPH|RDD)\b/i.test(`${text} ${fileName}`) && /^[123]10000$/.test(fields.mapSheet)) notes.push(`đã áp dụng quy tắc tờ đất rừng ${fields.mapSheet.slice(0, 1)} → ${fields.mapSheet}`);
  if (fields.mapSheet && fields.parcel && fields.mapSheet === fields.parcel) notes.push("số tờ trùng số thửa, nên kiểm tra lại");
  return notes;
}

function buildSuggestedName(file: File, fields: LandFields) {
  const suffix = sanitizeSuffix(fields.suffix);
  if (!/^\d{5}$/.test(fields.communeCode) || !/^\d{1,6}$/.test(fields.mapSheet) || !/^\d{1,7}$/.test(fields.parcel) || !suffix) return "";
  const extension = file.name.split(".").pop()?.toLowerCase() || "pdf";
  return `${fields.prefix}_${fields.communeCode}_${fields.mapSheet}_${fields.parcel}_${suffix}.${extension}`;
}

async function processWord(file: File): Promise<ProcessedResult> {
  const mammoth = await import("mammoth/mammoth.browser");
  const output = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
  const text = output.value.trim();
  const words = text ? text.split(/\s+/).length : 0;
  const paragraphs = text ? text.split(/\n+/).filter(Boolean).length : 0;

  return {
    kind: "word",
    title: "Đã đọc nội dung Word",
    description: "Văn bản được trích xuất để kiểm tra, tìm kiếm và tiếp tục xây dựng nội dung nghiệp vụ.",
    metrics: [
      { label: "Số từ", value: words.toLocaleString("vi-VN") },
      { label: "Đoạn văn", value: paragraphs.toLocaleString("vi-VN") },
      { label: "Cảnh báo", value: String(output.messages.length) },
    ],
    text: text || "Không tìm thấy nội dung chữ trong tệp Word này.",
    warning: output.messages.length ? "Một số thành phần định dạng phức tạp có thể chưa được thể hiện trong phần xem trước." : undefined,
  };
}

async function processPdf(
  file: File,
  onProgress: (progress: ProcessingProgress) => void,
  enableOcr = true,
): Promise<ProcessedResult> {
  const pdfjs = await import("pdfjs-dist");
  const workerUrl = await import("pdfjs-dist/build/pdf.worker.min.mjs?url");
  pdfjs.GlobalWorkerOptions.workerSrc = workerUrl.default;
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pages: string[] = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    onProgress({ label: `Đang kiểm tra lớp chữ · Trang ${pageNumber}/${pdf.numPages}`, percent: Math.round((pageNumber / pdf.numPages) * 25) });
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const text = content.items
      .map((item) => ("str" in item ? item.str : ""))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
    pages.push(text);
  }

  const missingPageIndexes = pages
    .map((pageText, index) => (pageText ? -1 : index))
    .filter((index) => index >= 0);
  const nativeTextPages = pages.length - missingPageIndexes.length;
  const ocrLimit = Math.min(missingPageIndexes.length, 20);
  let ocrCompleted = 0;
  let ocrFailed = false;

  if (ocrLimit > 0 && enableOcr) {
    onProgress({ label: "Đang tải bộ nhận dạng tiếng Việt…", percent: 28 });
    const { createWorker } = await import("tesseract.js");
    let currentOcrPage = 0;
    const worker = await createWorker("vie", 1, {
      logger: (message) => {
        if (message.status === "recognizing text") {
          const pageProgress = typeof message.progress === "number" ? message.progress : 0;
          const completedShare = currentOcrPage / Math.max(ocrLimit, 1);
          const currentShare = pageProgress / Math.max(ocrLimit, 1);
          onProgress({
            label: `OCR tiếng Việt · Trang ${currentOcrPage + 1}/${ocrLimit}`,
            percent: Math.min(98, Math.round(30 + (completedShare + currentShare) * 68)),
          });
        }
      },
    });

    try {
      for (let index = 0; index < ocrLimit; index += 1) {
        currentOcrPage = index;
        const pageIndex = missingPageIndexes[index];
        const page = await pdf.getPage(pageIndex + 1);
        const originalViewport = page.getViewport({ scale: 1 });
        const scale = Math.min(2, 1800 / originalViewport.width);
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        const context = canvas.getContext("2d", { willReadFrequently: true });
        if (!context) throw new Error("Không thể tạo vùng ảnh cho OCR.");
        await page.render({ canvas, canvasContext: context, viewport }).promise;
        const recognition = await worker.recognize(canvas);
        pages[pageIndex] = recognition.data.text.replace(/\s+\n/g, "\n").trim();
        ocrCompleted += 1;
        canvas.width = 1;
        canvas.height = 1;
      }
    } catch (reason) {
      console.error(reason);
      ocrFailed = true;
    } finally {
      await worker.terminate();
    }
  }

  onProgress({ label: "Đang hoàn thiện kết quả…", percent: 100 });
  const text = pages
    .map((pageText, index) => pageText ? `--- Trang ${index + 1} ---\n${pageText}` : "")
    .filter(Boolean)
    .join("\n\n");
  const remainingMissing = pages.filter((pageText) => !pageText).length;
  const warnings = [
    missingPageIndexes.length > 20 ? `OCR giới hạn 20 trang mỗi lượt; còn ${missingPageIndexes.length - 20} trang chưa nhận dạng.` : "",
    missingPageIndexes.length && !enableOcr ? `${missingPageIndexes.length} trang scan được bỏ qua trong chế độ hàng loạt; hãy xử lý riêng để chạy OCR.` : "",
    ocrFailed ? "OCR đã dừng do không tải được mô hình hoặc trang PDF quá phức tạp." : "",
    remainingMissing && !ocrFailed ? `Còn ${remainingMissing} trang chưa nhận dạng được chữ.` : "",
  ].filter(Boolean);

  return {
    kind: "pdf",
    title: ocrCompleted ? "Đã OCR tiếng Việt cho PDF scan" : text ? "Đã trích xuất chữ từ PDF" : "Chưa nhận dạng được nội dung",
    description: ocrCompleted
      ? "Các trang scan đã được chuyển thành chữ ngay trên thiết bị và ghép cùng lớp chữ sẵn có."
      : text
        ? "Nội dung chữ đã được đọc trực tiếp từ PDF và sẵn sàng để rà soát."
        : "PDF chưa có lớp chữ và OCR chưa hoàn thành.",
    metrics: [
      { label: "Số trang", value: String(pdf.numPages) },
      { label: "Chữ có sẵn", value: String(nativeTextPages) },
      { label: "Trang đã OCR", value: String(ocrCompleted) },
      { label: "Ký tự", value: text.length.toLocaleString("vi-VN") },
    ],
    text: text || "Không tìm thấy nội dung chữ có thể trích xuất trong tệp PDF này.",
    warning: warnings.length ? warnings.join(" ") : undefined,
    ocr: ocrCompleted > 0,
  };
}

async function processExcel(file: File): Promise<ProcessedResult> {
  const XLSX = await import("xlsx");
  let workbook;

  if (file.name.toLowerCase().endsWith(".csv")) {
    workbook = XLSX.read(await file.text(), { type: "string", cellDates: true });
  } else {
    workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  }

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<CellValue[]>(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  const columns = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const filled = rows.reduce(
    (total, row) => total + row.filter((cell) => String(cell).trim() !== "").length,
    0,
  );
  const blankRows = rows.filter((row) => row.every((cell) => !String(cell ?? "").trim())).length;
  const rowKeys = rows.slice(1).map((row) => row.map((cell) => String(cell ?? "").trim()).join("\u001f")).filter((key) => key.replaceAll("\u001f", ""));
  const duplicateRows = rowKeys.length - new Set(rowKeys).size;

  return {
    kind: "excel",
    title: "Đã đọc bảng dữ liệu",
    description: `Đang hiển thị sheet “${sheetName}”. Có thể xem trước dữ liệu và xuất lại thành CSV chuẩn hóa.`,
    metrics: [
      { label: "Số sheet", value: String(workbook.SheetNames.length) },
      { label: "Số dòng", value: rows.length.toLocaleString("vi-VN") },
      { label: "Số cột", value: columns.toLocaleString("vi-VN") },
      { label: "Ô có dữ liệu", value: filled.toLocaleString("vi-VN") },
      { label: "Dòng trống", value: blankRows.toLocaleString("vi-VN") },
      { label: "Dòng trùng", value: duplicateRows.toLocaleString("vi-VN") },
    ],
    rows,
    sheetName,
    warning: [rows.length > 200 ? "Bảng xem trước chỉ hiển thị 200 dòng đầu." : "", blankRows ? `Phát hiện ${blankRows} dòng trống.` : "", duplicateRows ? `Phát hiện ${duplicateRows} dòng trùng nội dung.` : ""].filter(Boolean).join(" ") || undefined,
  };
}

function BatchProcessor({ onProcessed, locationProfile }: { onProcessed: (record: HistoryRecord) => void; locationProfile: LocationProfile }) {
  const entitlements = useEntitlements();
  const maxBatchFiles = entitlements.maxParcelsPerRun;
  const inputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);
  const masterInputRef = useRef<HTMLInputElement>(null);
  const stopRequestedRef = useRef(false);
  const [items, setItems] = useState<BatchItem[]>([]);
  const [running, setRunning] = useState(false);
  const [master, setMaster] = useState<MasterSheet | null>(null);
  const [masterError, setMasterError] = useState("");
  const [comparisonFilter, setComparisonFilter] = useState<"all" | "matched" | "unmatched" | "missing">("all");
  const [batchSearch, setBatchSearch] = useState("");
  const [batchSort, setBatchSort] = useState<"added" | "name" | "parcel" | "comparison" | "status">("added");
  const [batchPrefix, setBatchPrefix] = useState<LandFields["prefix"]>("CHUACOGIAY");
  const [batchSuffix, setBatchSuffix] = useState("GT");
  const [archiveMode, setArchiveMode] = useState<"flat" | "folders">("folders");
  const [archiveRoot, setArchiveRoot] = useState("HO_SO_DAT_DAI");
  const [batchOcr, setBatchOcr] = useState(false);
  const [autoClassify, setAutoClassify] = useState(true);
  const [batchProgress, setBatchProgress] = useState({ current: 0, total: 0, label: "" });
  const [batchNotice, setBatchNotice] = useState("");
  const [contentHashes, setContentHashes] = useState<Record<string, string>>({});
  const [contentDuplicateIds, setContentDuplicateIds] = useState<Set<string>>(new Set());

  const suggestedNames = useMemo(
    () => items.map((item) => buildSuggestedName(item.file, item.fields)),
    [items],
  );
  const duplicateNames = useMemo(() => {
    const counts = new Map<string, number>();
    suggestedNames.filter(Boolean).forEach((name) => counts.set(name.toLocaleLowerCase(), (counts.get(name.toLocaleLowerCase()) || 0) + 1));
    return new Set([...counts.entries()].filter(([, count]) => count > 1).map(([name]) => name));
  }, [suggestedNames]);
  const readyCount = items.filter((item, index) => suggestedNames[index] && !duplicateNames.has(suggestedNames[index].toLocaleLowerCase()) && item.status !== "error").length;
  const cmdReadyCount = items.filter((item, index) => /\.pdf$/i.test(item.file.name) && suggestedNames[index] && !duplicateNames.has(suggestedNames[index].toLocaleLowerCase()) && (item.status === "done" || item.status === "review")).length;
  const masterKeyCounts = useMemo(() => {
    const counts = new Map<string, number>();
    if (!master || master.codeColumn < 0 || master.sheetColumn < 0 || master.parcelColumn < 0) return counts;
    master.rows.slice(master.headerRow + 1).forEach((row) => {
      const key = buildParcelKey(row[master.codeColumn], row[master.sheetColumn], row[master.parcelColumn]);
      if (key) counts.set(key, (counts.get(key) || 0) + 1);
    });
    return counts;
  }, [master]);
  const masterKeySources = useMemo(() => {
    const sources = new Map<string, Set<string>>();
    if (!master || master.codeColumn < 0 || master.sheetColumn < 0 || master.parcelColumn < 0) return sources;
    master.rows.slice(master.headerRow + 1).forEach((row) => {
      const key = buildParcelKey(row[master.codeColumn], row[master.sheetColumn], row[master.parcelColumn]);
      if (!key) return;
      const source = master.sheetName === "__ALL__" ? String(row[3] || "Không rõ sheet") : master.sheetName;
      if (!sources.has(key)) sources.set(key, new Set());
      sources.get(key)?.add(source);
    });
    return sources;
  }, [master]);
  const comparison = useMemo(() => items.map((item) => {
    if (!master) return "none" as const;
    const key = buildParcelKey(item.fields.communeCode, item.fields.mapSheet, item.fields.parcel);
    if (!key) return "missing" as const;
    return masterKeyCounts.has(key) ? "matched" as const : "unmatched" as const;
  }), [items, master, masterKeyCounts]);
  const matchedCount = comparison.filter((status) => status === "matched").length;
  const unmatchedCount = comparison.filter((status) => status === "unmatched").length;
  const duplicateMasterCount = [...masterKeyCounts.values()].filter((count) => count > 1).length;
  const crossSheetDuplicateCount = [...masterKeySources.values()].filter((sources) => sources.size > 1).length;
  const matchSources = items.map((item) => {
    const key = buildParcelKey(item.fields.communeCode, item.fields.mapSheet, item.fields.parcel);
    return key ? [...(masterKeySources.get(key) || [])] : [];
  });
  const visibleItems = items.map((item, index) => ({ item, index })).filter(({ item, index }) => {
    const matchesFilter = comparisonFilter === "all" || comparison[index] === comparisonFilter;
    const query = normalizeHeader(batchSearch);
    const searchable = normalizeHeader(`${item.file.name} ${item.fields.communeCode} ${item.fields.mapSheet} ${item.fields.parcel} ${suggestedNames[index]}`);
    return matchesFilter && (!query || searchable.includes(query));
  }).sort((a, b) => {
    if (batchSort === "name") return a.item.file.name.localeCompare(b.item.file.name, "vi", { numeric: true });
    if (batchSort === "parcel") return buildParcelKey(a.item.fields.communeCode, a.item.fields.mapSheet, a.item.fields.parcel).localeCompare(buildParcelKey(b.item.fields.communeCode, b.item.fields.mapSheet, b.item.fields.parcel), "vi", { numeric: true });
    if (batchSort === "comparison") return comparison[a.index].localeCompare(comparison[b.index]);
    if (batchSort === "status") return a.item.status.localeCompare(b.item.status);
    return a.index - b.index;
  });

  async function loadMasterFile(file?: File) {
    if (!file) return;
    setMasterError("");
    if (file.size > MAX_FILE_SIZE || !/\.(xlsx|xls|csv)$/i.test(file.name)) {
      setMasterError("File tổng phải là XLSX, XLS hoặc CSV và không vượt quá 20 MB.");
      return;
    }
    try {
      const XLSX = await import("xlsx");
      const workbook = file.name.toLowerCase().endsWith(".csv") ? XLSX.read(await file.text(), { type: "string", cellDates: true }) : XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
      const sheets = workbook.SheetNames.map((name) => ({ name, rows: XLSX.utils.sheet_to_json<CellValue[]>(workbook.Sheets[name], { header: 1, defval: "", raw: false }) }));
      const sheetName = sheets[0].name;
      const rows = sheets[0].rows;
      if (!rows.length) throw new Error("empty");
      const headerRow = detectHeaderRow(rows);
      const headers = rows[headerRow].map((value, index) => String(value).trim() || `Cột ${index + 1}`);
      const normalized = headers.map(normalizeHeader);
      setMaster({ fileName: file.name, sheetName, sheets, rows, headers, headerRow, codeColumn: findColumn(normalized, "code"), sheetColumn: findColumn(normalized, "sheet"), parcelColumn: findColumn(normalized, "parcel") });
    } catch (reason) {
      console.error(reason);
      setMaster(null);
      setMasterError("Không đọc được file tổng. Hãy kiểm tra sheet đầu tiên và hàng tiêu đề.");
    } finally {
      if (masterInputRef.current) masterInputRef.current.value = "";
    }
  }

  function changeMasterHeaderRow(headerRow: number) {
    setMaster((current) => {
      if (!current || !current.rows[headerRow]) return current;
      const headers = current.rows[headerRow].map((value, index) => String(value).trim() || `Cột ${index + 1}`);
      const normalized = headers.map(normalizeHeader);
      return { ...current, headerRow, headers, codeColumn: findColumn(normalized, "code"), sheetColumn: findColumn(normalized, "sheet"), parcelColumn: findColumn(normalized, "parcel") };
    });
  }

  function changeMasterSheet(sheetName: string) {
    setMaster((current) => {
      if (!current) return current;
      if (sheetName === "__ALL__") {
        const combinedRows: CellValue[][] = [["Mã xã", "Số tờ", "Số thửa", "Sheet nguồn"]];
        current.sheets.forEach((sheet) => {
          if (!sheet.rows.length) return;
          const detectedHeaderRow = detectHeaderRow(sheet.rows);
          const normalized = sheet.rows[detectedHeaderRow].map(normalizeHeader);
          const codeColumn = findColumn(normalized, "code");
          const sheetColumn = findColumn(normalized, "sheet");
          const parcelColumn = findColumn(normalized, "parcel");
          if (codeColumn < 0 || sheetColumn < 0 || parcelColumn < 0) return;
          sheet.rows.slice(detectedHeaderRow + 1).forEach((row) => {
            const key = buildParcelKey(row[codeColumn], row[sheetColumn], row[parcelColumn]);
            if (key) combinedRows.push([row[codeColumn], row[sheetColumn], row[parcelColumn], sheet.name]);
          });
        });
        return { ...current, sheetName: "__ALL__", rows: combinedRows, headerRow: 0, headers: ["Mã xã", "Số tờ", "Số thửa", "Sheet nguồn"], codeColumn: 0, sheetColumn: 1, parcelColumn: 2 };
      }
      const selected = current.sheets.find((sheet) => sheet.name === sheetName);
      if (!selected?.rows.length) return current;
      const headerRow = detectHeaderRow(selected.rows);
      const headers = selected.rows[headerRow].map((value, index) => String(value).trim() || `Cột ${index + 1}`);
      const normalized = headers.map(normalizeHeader);
      return { ...current, sheetName, rows: selected.rows, headerRow, headers, codeColumn: findColumn(normalized, "code"), sheetColumn: findColumn(normalized, "sheet"), parcelColumn: findColumn(normalized, "parcel") };
    });
  }

  function addFiles(fileList?: FileList | null) {
    if (!fileList) return;
    setBatchNotice("");
    const availableSlots = maxBatchFiles == null ? Number.POSITIVE_INFINITY : Math.max(0, maxBatchFiles - items.length);
    const selected = Array.from(fileList).slice(0, availableSlots);
    const existingKeys = new Set(items.map((item) => `${item.file.name}-${item.file.size}-${item.file.lastModified}`));
    const existingBytes = items.reduce((total, item) => total + item.file.size, 0);
    let acceptedBytes = existingBytes;
    const uniqueFiles = selected.filter((file) => !existingKeys.has(`${file.name}-${file.size}-${file.lastModified}`));
    const capacityFiles = uniqueFiles.filter((file) => {
      if (acceptedBytes + file.size > MAX_BATCH_TOTAL_SIZE) return false;
      acceptedBytes += file.size;
      return true;
    });
    const next = capacityFiles
      .map<BatchItem>((file, index) => {
        const extension = file.name.split(".").pop()?.toLowerCase() || "";
        const valid = ["docx", "pdf", "xlsx", "xls", "csv"].includes(extension) && file.size <= MAX_FILE_SIZE;
        const fields = extractLandFields("", file.name, locationProfile);
        if (!autoClassify) {
          fields.prefix = batchPrefix;
          fields.suffix = batchSuffix;
        }
        return {
          id: `${file.name}-${file.size}-${file.lastModified}-${index}`,
          file,
          status: valid ? "waiting" : "error",
          fields,
          message: valid ? "Chờ xử lý" : "Sai định dạng hoặc vượt quá 20 MB",
        };
      });
    setItems((current) => maxBatchFiles == null ? [...current, ...next] : [...current, ...next].slice(0, maxBatchFiles));
    const duplicateCount = selected.length - uniqueFiles.length;
    const omittedCount = Math.max(0, Array.from(fileList).length - selected.length) + (uniqueFiles.length - capacityFiles.length);
    if (omittedCount || duplicateCount) {
      setBatchNotice([
        omittedCount ? `${omittedCount} tệp chưa được thêm do vượt ${maxBatchFiles == null ? "giới hạn kỹ thuật tổng 500 MB" : `hạn mức ${maxBatchFiles} thửa/lần hoặc tổng 500 MB`}.` : "",
        duplicateCount ? `${duplicateCount} tệp trùng đã được bỏ qua.` : "",
      ].filter(Boolean).join(" "));
    } else {
      setBatchNotice(`Đã thêm ${next.length} tệp. ${maxBatchFiles == null ? "Tài khoản không giới hạn số thửa/lần." : `Còn có thể thêm ${Math.max(0, maxBatchFiles - items.length - next.length)} tệp.`}`);
    }
    if (inputRef.current) inputRef.current.value = "";
  }

  function updateItem(id: string, patch: Omit<Partial<BatchItem>, "fields"> & { fields?: Partial<LandFields> }) {
    setItems((current) => current.map((item) => item.id === id ? {
      ...item,
      ...patch,
      fields: patch.fields ? { ...item.fields, ...patch.fields } : item.fields,
    } : item));
  }

  function applyDefaults() {
    setItems((current) => current.map((item) => ({ ...item, fields: { ...item.fields, prefix: batchPrefix, suffix: batchSuffix } })));
  }

  async function scanContentDuplicates() {
    if (!items.length || running) return;
    setRunning(true);
    const hashes: Record<string, string> = {};
    const firstByHash = new Map<string, string>();
    const duplicates = new Set<string>();
    setBatchProgress({ current: 0, total: items.length, label: "Đang kiểm tra tệp trùng…" });
    try {
      for (let index = 0; index < items.length; index += 1) {
        const item = items[index];
        setBatchProgress({ current: index, total: items.length, label: `SHA-256 · ${item.file.name}` });
        const digest = await crypto.subtle.digest("SHA-256", await item.file.arrayBuffer());
        const hash = Array.from(new Uint8Array(digest)).map((value) => value.toString(16).padStart(2, "0")).join("");
        hashes[item.id] = hash;
        if (firstByHash.has(hash)) duplicates.add(item.id);
        else firstByHash.set(hash, item.id);
        setBatchProgress({ current: index + 1, total: items.length, label: `SHA-256 · ${item.file.name}` });
      }
      setContentHashes(hashes);
      setContentDuplicateIds(duplicates);
      setItems((current) => current.map((item) => duplicates.has(item.id) ? { ...item, status: "review", message: [item.message, "Trùng nội dung với tệp đã thêm trước đó"].filter(Boolean).join(" · ") } : item));
      setBatchProgress({ current: items.length, total: items.length, label: duplicates.size ? `Phát hiện ${duplicates.size} tệp trùng nội dung` : "Không phát hiện tệp trùng nội dung" });
    } catch (reason) {
      console.error(reason);
      setBatchProgress({ current: 0, total: items.length, label: "Không thể tính SHA-256 trên trình duyệt này" });
    } finally {
      setRunning(false);
    }
  }

  function removeContentDuplicates() {
    if (!contentDuplicateIds.size || running) return;
    setItems((current) => current.filter((item) => !contentDuplicateIds.has(item.id)));
    setContentDuplicateIds(new Set());
  }

  async function processAll() {
    if (!items.length || running) return;
    setRunning(true);
    stopRequestedRef.current = false;
    const processable = items.filter((item) => item.status !== "error");
    setBatchProgress({ current: 0, total: processable.length, label: "Đang chuẩn bị…" });
    let ocrPdfCount = 0;
    let completedCount = 0;
    for (const item of items) {
      if (stopRequestedRef.current) break;
      if (item.status === "error") continue;
      setBatchProgress({ current: completedCount, total: processable.length, label: item.file.name });
      updateItem(item.id, { status: "processing", message: "Đang đọc nội dung…" });
      try {
        const extension = item.file.name.split(".").pop()?.toLowerCase();
        let text = "";
        let scanNote = "";
        if (extension === "docx") {
          text = (await processWord(item.file)).text || "";
        } else if (extension === "pdf") {
          const runOcr = batchOcr && ocrPdfCount < 5;
          if (runOcr) ocrPdfCount += 1;
          const output = await processPdf(item.file, () => undefined, runOcr);
          text = output.text || "";
          scanNote = output.warning || "";
          if (batchOcr && !runOcr) scanNote = [scanNote, "Đã đạt giới hạn OCR 5 PDF trong một lượt; hãy xử lý riêng tệp scan này."].filter(Boolean).join(" ");
        } else {
          await processExcel(item.file);
        }
        const extracted = extractLandFields(text, item.file.name, locationProfile);
        extracted.prefix = item.fields.prefix;
        extracted.suffix = item.fields.suffix;
        const complete = Boolean(extracted.communeCode && extracted.mapSheet && extracted.parcel);
        const reviewNotes = reviewLandFields(extracted, text, item.file.name);
        const needsReview = !complete || reviewNotes.some((note) => note.includes("nên kiểm tra"));
        const detailMessage = [scanNote, ...reviewNotes].filter(Boolean).join(" · ");
        updateItem(item.id, {
          fields: extracted,
          status: needsReview ? "review" : "done",
          message: detailMessage || "Đã nhận diện đủ trường",
        });
        onProcessed({ id: `${Date.now()}-${item.id}`, fileName: item.file.name, fileType: extension?.toUpperCase() || "TỆP", processedAt: new Date().toISOString(), result: needsReview ? "Cần rà soát" : "Hoàn thành", suggestedName: buildSuggestedName(item.file, extracted) });
      } catch (reason) {
        console.error(reason);
        updateItem(item.id, { status: "error", message: "Không đọc được tệp" });
      }
      completedCount += 1;
      setBatchProgress({ current: completedCount, total: processable.length, label: item.file.name });
    }
    if (stopRequestedRef.current) setItems((current) => current.map((item) => item.status === "processing" ? { ...item, status: "review", message: "Đã dừng; cần xử lý lại tệp này" } : item));
    setRunning(false);
    setBatchProgress((current) => ({ ...current, label: stopRequestedRef.current ? "Đã dừng theo yêu cầu" : "Đã hoàn thành lượt xử lý" }));
  }

  async function downloadZip() {
    const validItems = items.filter((item, index) => suggestedNames[index] && !duplicateNames.has(suggestedNames[index].toLocaleLowerCase()) && item.status !== "error");
    if (!validItems.length) return;
    const { zipSync } = await import("fflate");
    const archive: Record<string, Uint8Array> = {};
    for (const item of validItems) {
      const fileName = buildSuggestedName(item.file, item.fields);
      const safeRoot = safeBaseName(archiveRoot.trim() || "HO_SO_DAT_DAI");
      const path = archiveMode === "folders" ? `${safeRoot}/${item.fields.prefix}/${item.fields.suffix}/${fileName}` : fileName;
      archive[path] = new Uint8Array(await item.file.arrayBuffer());
    }
    const manifestRows: CellValue[][] = [["STT", "Tệp nguồn", "Tên chuẩn", "Mã xã", "Số tờ", "Số thửa", "SHA-256", "Trạng thái", "Ghi chú"], ...validItems.map((item, index) => [index + 1, item.file.name, buildSuggestedName(item.file, item.fields), item.fields.communeCode, item.fields.mapSheet, item.fields.parcel, contentHashes[item.id] || "Chưa tính", item.status, item.message])];
    archive[`${safeBaseName(archiveRoot.trim() || "HO_SO_DAT_DAI")}_NHAT_KY.csv`] = new TextEncoder().encode(`\uFEFF${manifestRows.map((row) => row.map(escapeCsv).join(",")).join("\r\n")}`);
    const zipped = zipSync(archive, { level: 6 });
    const blob = new Blob([zipped], { type: "application/zip" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `HO_SO_DA_CHUAN_HOA_${new Date().toISOString().slice(0, 10)}.zip`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function downloadRenameCmd() {
    const pdfItems = items.filter((item, index) =>
      /\.pdf$/i.test(item.file.name)
      && Boolean(suggestedNames[index])
      && !duplicateNames.has(suggestedNames[index].toLocaleLowerCase())
      && (item.status === "done" || item.status === "review"),
    );
    if (!pdfItems.length) return;
    const escapeBatchName = (name: string) => name.replaceAll("%", "%%");
    const commands = pdfItems.map((item) =>
      `ren "${escapeBatchName(item.file.name)}" "${escapeBatchName(buildSuggestedName(item.file, item.fields))}"`,
    );
    const content = [
      "@echo off",
      "chcp 65001 >nul",
      "setlocal",
      "echo SỸ LAND - DOI TEN PDF HANG LOAT",
      "echo Thu muc: %CD%",
      "",
      ...commands,
      "",
      `echo Da thuc hien ${commands.length} lenh doi ten.`,
      "pause",
    ].join("\r\n");
    const blob = new Blob(["\uFEFF", content], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `SY_LAND_DOI_TEN_PDF_${new Date().toISOString().slice(0, 10)}.bat`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function downloadComparisonReport() {
    if (!master || !items.length) return;
    const labels = { none: "Chưa đối chiếu", missing: "Thiếu khóa", matched: "Trùng khớp", unmatched: "Không có trong file tổng" };
    const rows: CellValue[][] = [["STT", "Tệp nguồn", "Mã xã", "Số tờ", "Số thửa", "Tên đề xuất", "Kết quả đối chiếu", "Sheet tìm thấy", "Ghi chú"], ...items.map((item, index) => [index + 1, item.file.name, item.fields.communeCode, item.fields.mapSheet, item.fields.parcel, suggestedNames[index], labels[comparison[index]], matchSources[index].join("; "), item.message])];
    const blob = new Blob(["\uFEFF", rows.map((row) => row.map(escapeCsv).join(",")).join("\r\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `BAO_CAO_DOI_CHIEU_${new Date().toISOString().slice(0, 10)}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  async function downloadReviewWorkbook() {
    if (!master) return;
    const reviewRows = items.flatMap((item, index) => comparison[index] === "unmatched" || comparison[index] === "missing" ? [{
      "Tệp nguồn": item.file.name,
      "Mã xã": item.fields.communeCode,
      "Số tờ": item.fields.mapSheet,
      "Số thửa": item.fields.parcel,
      "Tên đề xuất": suggestedNames[index],
      "Kết quả": comparison[index] === "unmatched" ? "Không có trong file tổng" : "Thiếu khóa đối chiếu",
      "Ghi chú": item.message,
    }] : []);
    if (!reviewRows.length) return;
    const XLSX = await import("xlsx");
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(reviewRows);
    sheet["!cols"] = [{ wch: 32 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 42 }, { wch: 24 }, { wch: 34 }];
    XLSX.utils.book_append_sheet(workbook, sheet, "CAN_RA_SOAT");
    XLSX.writeFile(workbook, `HO_SO_CAN_RA_SOAT_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  async function downloadBatchWorkbook() {
    if (!items.length) return;
    const XLSX = await import("xlsx");
    const rows = items.map((item, index) => ({
      "STT": index + 1,
      "Tệp nguồn": item.file.name,
      "Tỉnh/thành phố mới": locationProfile.provinceNew,
      "Tỉnh/thành phố cũ": locationProfile.provinceOld,
      "Xã/phường mới": locationProfile.communeNew,
      "Xã/phường cũ": locationProfile.communeOld,
      "Thôn/bản/xóm mới": locationProfile.villageNew,
      "Thôn/bản/xóm cũ": locationProfile.villageOld,
      "Mã xã": item.fields.communeCode,
      "Số tờ": item.fields.mapSheet,
      "Số thửa": item.fields.parcel,
      "Diện tích (m²)": item.fields.area,
      "Nhóm hồ sơ": item.fields.prefix,
      "Hậu tố": item.fields.suffix,
      "Tên tệp đề xuất": suggestedNames[index],
      "SHA-256": contentHashes[item.id] || "Chưa tính",
      "Đối chiếu file tổng": comparison[index] === "matched" ? "Trùng khớp" : comparison[index] === "unmatched" ? "Không có trong file tổng" : comparison[index] === "missing" ? "Thiếu khóa" : "Chưa đối chiếu",
      "Sheet tìm thấy": matchSources[index].join("; "),
      "Trạng thái": item.status === "done" ? "Hoàn thành" : item.status === "review" ? "Cần rà soát" : item.status === "error" ? "Có lỗi" : item.status === "processing" ? "Đang xử lý" : "Chờ xử lý",
      "Ghi chú": item.message,
    }));
    const summary = [
      { "Chỉ tiêu": "Tổng số tệp", "Giá trị": items.length },
      { "Chỉ tiêu": "Đã hoàn thành", "Giá trị": items.filter((item) => item.status === "done").length },
      { "Chỉ tiêu": "Cần rà soát", "Giá trị": items.filter((item) => item.status === "review").length },
      { "Chỉ tiêu": "Có lỗi", "Giá trị": items.filter((item) => item.status === "error").length },
      { "Chỉ tiêu": "Trùng khớp file tổng", "Giá trị": matchedCount },
      { "Chỉ tiêu": "Không có trong file tổng", "Giá trị": unmatchedCount },
      { "Chỉ tiêu": "Tên tệp bị trùng", "Giá trị": duplicateNames.size },
      { "Chỉ tiêu": "Thời điểm xuất", "Giá trị": new Date().toLocaleString("vi-VN") },
    ];
    const workbook = XLSX.utils.book_new();
    const dataSheet = XLSX.utils.json_to_sheet(rows);
    dataSheet["!cols"] = [{ wch: 7 }, { wch: 34 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 11 }, { wch: 11 }, { wch: 11 }, { wch: 16 }, { wch: 18 }, { wch: 10 }, { wch: 48 }, { wch: 66 }, { wch: 24 }, { wch: 28 }, { wch: 16 }, { wch: 48 }];
    const summarySheet = XLSX.utils.json_to_sheet(summary);
    summarySheet["!cols"] = [{ wch: 30 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(workbook, dataSheet, "TONG_HOP_HO_SO");
    XLSX.utils.book_append_sheet(workbook, summarySheet, "TONG_QUAN");
    XLSX.writeFile(workbook, `TONG_HOP_HO_SO_DAT_DAI_${new Date().toISOString().slice(0, 10)}.xlsx`, { compression: true });
  }

  async function downloadMasterDuplicates() {
    if (!master || !duplicateMasterCount) return;
    const duplicateRows = [...masterKeyCounts.entries()].filter(([, count]) => count > 1).map(([key, count]) => {
      const [communeCode, mapSheet, parcel] = key.split("|");
      const sources = [...(masterKeySources.get(key) || [])];
      return { "Mã xã": communeCode, "Số tờ": mapSheet, "Số thửa": parcel, "Số lần xuất hiện": count, "Số sheet": sources.length, "Sheet liên quan": sources.join("; "), "Phân loại": sources.length > 1 ? "Trùng giữa nhiều sheet" : "Trùng dòng trong một sheet" };
    });
    const XLSX = await import("xlsx");
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(duplicateRows);
    sheet["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 17 }, { wch: 12 }, { wch: 38 }, { wch: 28 }];
    XLSX.utils.book_append_sheet(workbook, sheet, "KHOA_TRUNG");
    XLSX.writeFile(workbook, `BAO_CAO_KHOA_TRUNG_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <section className="batch-shell" id="xu-ly-hang-loat" aria-labelledby="batch-title">
      <div className="batch-heading">
        <div><span className="demo-label">04 · XỬ LÝ HÀNG LOẠT</span><h3 id="batch-title">Chuẩn hóa nhiều hồ sơ trong một lượt</h3><p>{maxBatchFiles == null ? "Không giới hạn số thửa theo tài khoản" : `Hạn mức ${maxBatchFiles} thửa/lần`}, 20 MB mỗi tệp và 500 MB mỗi lượt. Hệ thống xử lý tuần tự để hạn chế treo trình duyệt.</p><small>{entitlements.reason}</small></div>
        <div className="batch-summary"><strong>{items.length}</strong><span>Tệp đã chọn</span><strong>{readyCount}</strong><span>Sẵn sàng tải</span><strong>{maxBatchFiles == null ? "∞" : Math.max(0, maxBatchFiles - items.length)}</strong><span>Còn có thể thêm</span></div>
      </div>
      <div className="batch-toolbar">
        <label className="batch-ocr-toggle"><input type="checkbox" checked={autoClassify} onChange={(event) => setAutoClassify(event.target.checked)} /><span><b>Tự phân loại tên tệp</b><small>Nhận COGCN/CHUACOGIAY và GT/TBXN/DDK</small></span></label>
        <label>Nhóm hồ sơ<select value={batchPrefix} onChange={(event) => setBatchPrefix(event.target.value as LandFields["prefix"])}><option value="CHUACOGIAY">Chưa có giấy</option><option value="COGCN">Có GCN</option></select></label>
        <label>Hậu tố tự đặt<input value={batchSuffix} maxLength={24} onChange={(event) => setBatchSuffix(sanitizeSuffix(event.target.value))} placeholder="GT hoặc TBXN" list="syland-suffixes" /><datalist id="syland-suffixes"><option value="GT" /><option value="TBXN" /><option value="DDK" /></datalist></label>
        <label>Cách đóng gói<select value={archiveMode} onChange={(event) => setArchiveMode(event.target.value as "flat" | "folders")}><option value="folders">Chia thư mục tự động</option><option value="flat">Một thư mục</option></select></label>
        {archiveMode === "folders" && <label className="archive-root-label">Tên thư mục gốc<input value={archiveRoot} maxLength={50} onChange={(event) => setArchiveRoot(event.target.value)} placeholder="HO_SO_DAT_DAI" /></label>}
        <label className="batch-ocr-toggle"><input type="checkbox" checked={batchOcr} onChange={(event) => setBatchOcr(event.target.checked)} /><span><b>OCR PDF scan</b><small>Tối đa 5 PDF/lượt; xử lý lâu hơn</small></span></label>
        <button type="button" className="secondary-action" onClick={applyDefaults} disabled={!items.length}>Áp dụng cho tất cả</button>
        <input ref={inputRef} type="file" accept={ACCEPTED} multiple onChange={(event) => addFiles(event.target.files)} aria-label="Chọn nhiều tệp hồ sơ" />
        <input ref={folderInputRef} type="file" accept={ACCEPTED} multiple {...({ webkitdirectory: "", directory: "" } as Record<string, string>)} onChange={(event) => addFiles(event.target.files)} aria-label="Chọn thư mục hồ sơ" />
        <button type="button" className="batch-add" onClick={() => inputRef.current?.click()}>+ Chọn nhiều tệp</button>
        <button type="button" className="batch-add folder-add" onClick={() => folderInputRef.current?.click()}>+ Chọn cả thư mục</button>
      </div>
      {batchNotice && <p className="batch-capacity-notice" role="status">{batchNotice}</p>}

      <div className="rename-workflow-note">
        <strong>Đổi tên PDF theo nội dung hồ sơ</strong>
        <span>1. Đọc “Thửa đất số” và “Tờ bản đồ số” tại mục 2.a</span>
        <span>2. Đọc mục 2.b, ưu tiên tên thôn để xác định xã cũ/mã xã trong cấu hình</span>
        <span>3. Xem trước cú pháp CHUACOGIAY_MÃXÃ_SỐTỜ_SỐTHỬA_{sanitizeSuffix(batchSuffix) || "HẬUTỐ"}.pdf</span>
        <span>4. Tải tệp CMD và chạy trong thư mục chứa PDF gốc</span>
      </div>

      {archiveMode === "folders" && <div className="folder-preview" aria-label="Cấu trúc thư mục ZIP"><span>ZIP sẽ được sắp xếp:</span><code>{safeBaseName(archiveRoot.trim() || "HO_SO_DAT_DAI")} / COGCN hoặc CHUACOGIAY / GT, TBXN hoặc DDK / tệp</code></div>}

      <div className="master-panel" id="doi-chieu-du-lieu">
        <div className="master-intro"><span className="demo-label">05 · ĐỐI CHIẾU FILE TỔNG</span><strong>So khớp theo Mã xã + Số tờ + Số thửa</strong><small>Chỉ đọc dữ liệu trên thiết bị, không thay đổi file Excel gốc.</small></div>
        <input ref={masterInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={(event) => void loadMasterFile(event.target.files?.[0])} aria-label="Chọn file Excel tổng" />
        <button type="button" className="master-file-button" onClick={() => masterInputRef.current?.click()}>{master ? "Đổi file tổng" : "+ Chọn file tổng"}</button>
        {master && <div className="master-config">
          <div className="master-file"><strong>{master.fileName}</strong><small>{master.sheetName === "__ALL__" ? `Tất cả ${master.sheets.length} sheet có cấu trúc phù hợp` : `Sheet “${master.sheetName}”`} · {Math.max(0, master.rows.length - master.headerRow - 1).toLocaleString("vi-VN")} dòng dữ liệu · Tiêu đề dòng {master.headerRow + 1}</small></div>
          <div className="master-selects"><label>Sheet dữ liệu<select value={master.sheetName} onChange={(event) => changeMasterSheet(event.target.value)}><option value="__ALL__">Tất cả sheet phù hợp</option>{master.sheets.map((sheet) => <option key={sheet.name} value={sheet.name}>{sheet.name}</option>)}</select></label><label>Hàng tiêu đề<select value={master.headerRow} disabled={master.sheetName === "__ALL__"} onChange={(event) => changeMasterHeaderRow(Number(event.target.value))}>{master.rows.slice(0, 20).map((_, rowIndex) => <option key={rowIndex} value={rowIndex}>Dòng {rowIndex + 1}</option>)}</select></label>{(["codeColumn", "sheetColumn", "parcelColumn"] as const).map((field, index) => <label key={field}>{["Cột mã xã", "Cột số tờ", "Cột số thửa"][index]}<select value={master[field]} disabled={master.sheetName === "__ALL__"} onChange={(event) => setMaster((current) => current ? { ...current, [field]: Number(event.target.value) } : current)}><option value={-1}>Chọn cột…</option>{master.headers.map((header, column) => <option key={`${header}-${column}`} value={column}>{header}</option>)}</select></label>)}</div>
          <div className="comparison-stats"><span><b>{matchedCount}</b> Trùng khớp</span><span><b>{unmatchedCount}</b> Không khớp</span><span><b>{duplicateMasterCount}</b> Khóa trùng</span>{master.sheetName === "__ALL__" && <span><b>{crossSheetDuplicateCount}</b> Trùng nhiều sheet</span>}</div>
        </div>}
        {masterError && <p className="master-error">{masterError}</p>}
        {master && (master.codeColumn < 0 || master.sheetColumn < 0 || master.parcelColumn < 0) && <p className="master-error">Chưa nhận diện đủ 3 cột. Hãy chọn thủ công trong các danh sách trên.</p>}
      </div>

      {items.length > 0 && <div className="batch-list-controls">
        <div className="batch-search"><span aria-hidden="true">⌕</span><input value={batchSearch} onChange={(event) => setBatchSearch(event.target.value)} placeholder="Tìm tên tệp, mã xã, số tờ, số thửa…" aria-label="Tìm trong danh sách hồ sơ" />{batchSearch && <button type="button" onClick={() => setBatchSearch("")} aria-label="Xóa tìm kiếm">×</button>}</div>
        <label>Sắp xếp<select value={batchSort} onChange={(event) => setBatchSort(event.target.value as typeof batchSort)}><option value="added">Thứ tự thêm</option><option value="name">Tên tệp</option><option value="parcel">Mã xã · Tờ · Thửa</option><option value="comparison">Kết quả đối chiếu</option><option value="status">Trạng thái xử lý</option></select></label>
        {master && <div className="comparison-filter" aria-label="Lọc kết quả đối chiếu"><span>Hiển thị</span>{(["all", "matched", "unmatched", "missing"] as const).map((filter) => <button type="button" key={filter} className={comparisonFilter === filter ? "active" : ""} onClick={() => setComparisonFilter(filter)}>{filter === "all" ? `Tất cả (${items.length})` : filter === "matched" ? `Trùng khớp (${matchedCount})` : filter === "unmatched" ? `Không khớp (${unmatchedCount})` : `Thiếu khóa (${comparison.filter((status) => status === "missing").length})`}</button>)}</div>}
        <small>{visibleItems.length} hồ sơ đang hiển thị</small>
      </div>}

      {!items.length ? (
        <button type="button" className="batch-empty" onClick={() => inputRef.current?.click()} onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); addFiles(event.dataTransfer.files); }}><span aria-hidden="true">＋</span><strong>Chọn hoặc kéo thả nhiều Word, PDF, Excel</strong><small>Có thể chọn cả thư mục; hệ thống xử lý lần lượt và không ghi đè tệp gốc.</small></button>
      ) : !visibleItems.length ? (
        <div className="batch-no-results"><strong>Không tìm thấy hồ sơ phù hợp</strong><button type="button" onClick={() => { setBatchSearch(""); setComparisonFilter("all"); }}>Xóa bộ lọc</button></div>
      ) : (
        <div className="batch-table-wrap" tabIndex={0}>
          <table className="batch-table">
            <thead><tr><th>Tệp nguồn</th><th>Mã xã</th><th>Số tờ</th><th>Số thửa</th><th>Tên đề xuất</th><th>Đối chiếu</th><th>Trạng thái</th><th /></tr></thead>
            <tbody>{visibleItems.map(({ item, index }) => {
              const suggested = suggestedNames[index];
              const duplicate = Boolean(suggested && duplicateNames.has(suggested.toLocaleLowerCase()));
              return <tr key={item.id} className={duplicate ? "duplicate-row" : ""}>
                <td><strong>{item.file.name}</strong><small>{formatBytes(item.file.size)}</small>{contentHashes[item.id] && <small className="file-hash" title={contentHashes[item.id]}>SHA-256: {contentHashes[item.id].slice(0, 12)}…</small>}{contentDuplicateIds.has(item.id) && <b className="duplicate-note">Trùng nội dung</b>}</td>
                <td><input value={item.fields.communeCode} aria-label={`Mã xã của ${item.file.name}`} maxLength={5} onChange={(event) => updateItem(item.id, { fields: { communeCode: event.target.value.replace(/\D/g, "").slice(0, 5) } })} /></td>
                <td><input value={item.fields.mapSheet} aria-label={`Số tờ của ${item.file.name}`} maxLength={6} onChange={(event) => updateItem(item.id, { fields: { mapSheet: event.target.value.replace(/\D/g, "").slice(0, 6) } })} /></td>
                <td><input value={item.fields.parcel} aria-label={`Số thửa của ${item.file.name}`} maxLength={7} onChange={(event) => updateItem(item.id, { fields: { parcel: event.target.value.replace(/\D/g, "").slice(0, 7) } })} /></td>
                <td><code>{suggested || "Chưa đủ dữ liệu"}</code>{duplicate && <b className="duplicate-note">Trùng tên</b>}</td>
                <td><span className={`match-chip match-${comparison[index]}`}>{comparison[index] === "matched" ? "Trùng khớp" : comparison[index] === "unmatched" ? "Không có trong tổng" : comparison[index] === "missing" ? "Thiếu khóa" : "Chưa có file tổng"}</span>{matchSources[index].length > 0 && <small className="match-source">Sheet: {matchSources[index].join(", ")}</small>}{matchSources[index].length > 1 && <b className="multi-sheet-note">Có ở nhiều sheet</b>}</td>
                <td><span className={`batch-status status-${item.status}`}>{item.status === "processing" ? "Đang xử lý" : item.status === "done" ? "Hoàn thành" : item.status === "review" ? "Cần rà soát" : item.status === "error" ? "Có lỗi" : "Chờ xử lý"}</span><small>{item.message}</small></td>
                <td><button type="button" onClick={() => setItems((current) => current.filter((entry) => entry.id !== item.id))} aria-label={`Bỏ ${item.file.name}`}>×</button></td>
              </tr>;
            })}</tbody>
          </table>
        </div>
      )}

      <div className="batch-actions">
        <button type="button" className="button button-primary" onClick={processAll} disabled={!items.length || running}>{running ? "Đang xử lý…" : "Bắt đầu nhận diện"}<span aria-hidden="true">→</span></button>
        {running && <button type="button" className="button batch-stop" onClick={() => { stopRequestedRef.current = true; setBatchProgress((current) => ({ ...current, label: "Sẽ dừng sau tệp hiện tại…" })); }}>Dừng an toàn</button>}
        <button type="button" className="button report-download" onClick={() => void downloadBatchWorkbook()} disabled={!items.length || running}>Xuất Excel tổng hợp <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={() => void scanContentDuplicates()} disabled={!items.length || running}>Kiểm tra tệp trùng</button>
        {contentDuplicateIds.size > 0 && <button type="button" className="button batch-stop" onClick={removeContentDuplicates} disabled={running}>Bỏ {contentDuplicateIds.size} bản trùng</button>}
        <button type="button" className="button batch-download" onClick={downloadZip} disabled={!readyCount || running}>Tải {readyCount} tệp dạng ZIP <span aria-hidden="true">↓</span></button>
        <button type="button" className="button cmd-download" onClick={downloadRenameCmd} disabled={!cmdReadyCount || running}>Tải {cmdReadyCount} lệnh CMD đổi tên PDF <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={downloadComparisonReport} disabled={!master || !items.length}>Tải báo cáo CSV <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={() => void downloadReviewWorkbook()} disabled={!master || (!unmatchedCount && !comparison.includes("missing"))}>Xuất hồ sơ cần rà soát (.xlsx) <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={() => void downloadMasterDuplicates()} disabled={!master || !duplicateMasterCount}>Xuất {duplicateMasterCount} khóa trùng (.xlsx) <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={() => { setItems([]); setBatchProgress({ current: 0, total: 0, label: "" }); }} disabled={!items.length || running}>Xóa danh sách</button>
        {duplicateNames.size > 0 && <p><span aria-hidden="true">!</span>Có {duplicateNames.size} tên bị trùng. Hãy sửa mã xã, số tờ hoặc số thửa trước khi tải.</p>}
      </div>
      {batchProgress.total > 0 && <div className="batch-progress" aria-live="polite"><div><strong>{batchProgress.label}</strong><span>{batchProgress.current}/{batchProgress.total} tệp · {Math.round((batchProgress.current / Math.max(1, batchProgress.total)) * 100)}%</span></div><i><b style={{ width: `${(batchProgress.current / Math.max(1, batchProgress.total)) * 100}%` }} /></i></div>}
    </section>
  );
}

export default function FileProcessor() {
  const inputRef = useRef<HTMLInputElement>(null);
  const locationImportRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<ProcessedResult | null>(null);
  const [status, setStatus] = useState<"idle" | "processing" | "done" | "error">("idle");
  const [error, setError] = useState("");
  const [dragActive, setDragActive] = useState(false);
  const [progress, setProgress] = useState<ProcessingProgress>({ label: "Đang chuẩn bị…", percent: 0 });
  const [landFields, setLandFields] = useState<LandFields>({ communeCode: "", mapSheet: "", parcel: "", area: "", prefix: "CHUACOGIAY", suffix: "GT" });
  const [history, setHistory] = useState<HistoryRecord[]>([]);
  const [locationProfile, setLocationProfile] = useState<LocationProfile>(EMPTY_LOCATION);
  const [locationLoaded, setLocationLoaded] = useState(false);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      try {
        const stored = localStorage.getItem("sy-land-processing-history");
        if (stored) setHistory(JSON.parse(stored));
      } catch { /* Bỏ qua lịch sử không hợp lệ. */ }
    }, 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    try {
      const stored = localStorage.getItem("sy-land-location-profile");
      if (stored) setLocationProfile({ ...EMPTY_LOCATION, ...JSON.parse(stored) });
    } catch { /* Bỏ qua cấu hình địa bàn cũ không hợp lệ. */ }
    setLocationLoaded(true);
  }, []);

  useEffect(() => {
    if (locationLoaded) localStorage.setItem("sy-land-location-profile", JSON.stringify(locationProfile));
  }, [locationProfile, locationLoaded]);

  function addHistory(record: HistoryRecord) {
    setHistory((current) => {
      const next = [record, ...current].slice(0, 30);
      localStorage.setItem("sy-land-processing-history", JSON.stringify(next));
      return next;
    });
  }

  function exportLocationProfile() {
    const blob = new Blob([JSON.stringify({ product: "SY LAND", version: 1, location: locationProfile }, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `CAU_HINH_DIA_BAN_${locationProfile.communeCode || "SY_LAND"}.json`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  async function importLocationProfile(file?: File) {
    if (!file || file.size > 1024 * 1024) return;
    try {
      const parsed = JSON.parse(await file.text());
      const source = parsed?.location || parsed;
      const next = { ...EMPTY_LOCATION, ...source, communeCode: normalizeCommuneCode(String(source?.communeCode || "")) };
      setLocationProfile(next);
    } catch {
      setError("Không đọc được cấu hình địa bàn. Hãy chọn đúng tệp JSON đã xuất từ SỸ LAND.");
      setStatus("error");
    } finally {
      if (locationImportRef.current) locationImportRef.current.value = "";
    }
  }

  const suggestedFileName = useMemo(() => {
    return file ? buildSuggestedName(file, landFields) : "";
  }, [file, landFields]);
  const landReviewNotes = useMemo(() => file && result?.kind !== "excel" ? reviewLandFields(landFields, result?.text || "", file.name) : [], [file, result, landFields]);

  async function handleFile(nextFile?: File) {
    if (!nextFile) return;
    setError("");
    setResult(null);
    setFile(nextFile);

    const extension = nextFile.name.split(".").pop()?.toLowerCase();
    if (!extension || !["docx", "pdf", "xlsx", "xls", "csv"].includes(extension)) {
      setStatus("error");
      setError("Định dạng chưa được hỗ trợ. Vui lòng chọn DOCX, PDF, XLSX, XLS hoặc CSV.");
      return;
    }
    if (nextFile.size > MAX_FILE_SIZE) {
      setStatus("error");
      setError("Tệp vượt quá 20 MB. Hãy chọn tệp nhỏ hơn để xử lý an toàn trên thiết bị.");
      return;
    }

    setStatus("processing");
    setProgress({ label: "Đang mở tệp…", percent: 3 });
    try {
      let processed: ProcessedResult;
      if (extension === "docx") processed = await processWord(nextFile);
      else if (extension === "pdf") processed = await processPdf(nextFile, setProgress);
      else processed = await processExcel(nextFile);
      if (processed.text && processed.kind !== "excel") {
        setLandFields(extractLandFields(processed.text, nextFile.name, locationProfile));
      }
      setResult(processed);
      setStatus("done");
      addHistory({ id: `${Date.now()}-${nextFile.name}`, fileName: nextFile.name, fileType: extension.toUpperCase(), processedAt: new Date().toISOString(), result: processed.ocr ? "Đã OCR" : "Hoàn thành", suggestedName: "" });
    } catch (reason) {
      console.error(reason);
      setStatus("error");
      setError("Không thể đọc tệp này. Tệp có thể bị khóa, hỏng hoặc dùng cấu trúc chưa được hỗ trợ.");
    }
  }

  function onInputChange(event: ChangeEvent<HTMLInputElement>) {
    void handleFile(event.target.files?.[0]);
    event.target.value = "";
  }

  function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragActive(false);
    void handleFile(event.dataTransfer.files?.[0]);
  }

  function exportResult() {
    if (!result || !file) return;
    const isTable = result.kind === "excel" && result.rows;
    const content = isTable
      ? result.rows!.map((row) => row.map(escapeCsv).join(",")).join("\r\n")
      : [
          `Tệp nguồn: ${file.name}`,
          `Loại: ${result.kind.toUpperCase()}`,
          `Kích thước: ${formatBytes(file.size)}`,
          `Mã xã: ${landFields.communeCode || "Chưa xác định"}`,
          `Số tờ: ${landFields.mapSheet || "Chưa xác định"}`,
          `Số thửa: ${landFields.parcel || "Chưa xác định"}`,
          `Diện tích: ${landFields.area || "Chưa xác định"}`,
          `Tên tệp đề xuất: ${suggestedFileName || "Chưa đủ dữ liệu"}`,
          "",
          result.text || "",
        ].join("\n");
    const extension = isTable ? "csv" : "txt";
    const blob = new Blob(["\uFEFF", content], { type: isTable ? "text/csv;charset=utf-8" : "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${safeBaseName(file.name)}_KET_QUA.${extension}`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  async function exportCleanedWorkbook() {
    if (!result?.rows || !file) return;
    const XLSX = await import("xlsx");
    const seen = new Set<string>();
    const cleaned = result.rows
      .map((row) => row.map((cell) => typeof cell === "string" ? cell.trim().replace(/\s+/g, " ") : cell))
      .filter((row, index) => {
        if (row.every((cell) => !String(cell ?? "").trim())) return false;
        if (index === 0) return true;
        const key = row.map((cell) => String(cell ?? "")).join("\u001f");
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    const workbook = XLSX.utils.book_new();
    const cleanSheet = XLSX.utils.aoa_to_sheet(cleaned);
    cleanSheet["!cols"] = Array.from({ length: Math.min(30, Math.max(1, ...cleaned.map((row) => row.length))) }, () => ({ wch: 18 }));
    XLSX.utils.book_append_sheet(workbook, cleanSheet, "DU_LIEU_DA_LAM_SACH");

    const headerRow = detectHeaderRow(cleaned);
    const normalizedHeaders = (cleaned[headerRow] || []).map(normalizeHeader);
    const codeColumn = findColumn(normalizedHeaders, "code");
    const sheetColumn = findColumn(normalizedHeaders, "sheet");
    const parcelColumn = findColumn(normalizedHeaders, "parcel");
    const issues: Array<Record<string, CellValue>> = [];
    const keyRows = new Map<string, number[]>();
    if (codeColumn >= 0 && sheetColumn >= 0 && parcelColumn >= 0) {
      cleaned.slice(headerRow + 1).forEach((row, offset) => {
        const excelRow = headerRow + offset + 2;
        const code = String(row[codeColumn] ?? "").replace(/\D/g, "");
        const mapSheet = normalizeKeyNumber(row[sheetColumn]);
        const parcel = normalizeKeyNumber(row[parcelColumn]);
        const key = buildParcelKey(code, mapSheet, parcel);
        if (!/^\d{5}$/.test(code)) issues.push({ "Dòng": excelRow, "Loại lỗi": "Mã xã không hợp lệ", "Mã xã": code, "Số tờ": mapSheet, "Số thửa": parcel, "Khuyến nghị": "Nhập đúng mã xã 5 chữ số" });
        if (!mapSheet) issues.push({ "Dòng": excelRow, "Loại lỗi": "Thiếu số tờ", "Mã xã": code, "Số tờ": mapSheet, "Số thửa": parcel, "Khuyến nghị": "Bổ sung số tờ bản đồ" });
        if (!parcel) issues.push({ "Dòng": excelRow, "Loại lỗi": "Thiếu số thửa", "Mã xã": code, "Số tờ": mapSheet, "Số thửa": parcel, "Khuyến nghị": "Bổ sung số thửa" });
        if (key) keyRows.set(key, [...(keyRows.get(key) || []), excelRow]);
      });
      keyRows.forEach((rowNumbers, key) => {
        if (rowNumbers.length < 2) return;
        const [code, mapSheet, parcel] = key.split("|");
        issues.push({ "Dòng": rowNumbers.join(", "), "Loại lỗi": "Trùng khóa thửa đất", "Mã xã": code, "Số tờ": mapSheet, "Số thửa": parcel, "Khuyến nghị": "Kiểm tra các dòng có cùng Mã xã + Số tờ + Số thửa" });
      });
    } else {
      issues.push({ "Dòng": "—", "Loại lỗi": "Không nhận diện đủ cột", "Mã xã": "", "Số tờ": "", "Số thửa": "", "Khuyến nghị": "Đặt tiêu đề cột rõ ràng: Mã xã, Số tờ, Số thửa" });
    }
    const reportSheet = XLSX.utils.json_to_sheet(issues.length ? issues : [{ "Dòng": "—", "Loại lỗi": "Không phát hiện lỗi khóa địa chính", "Mã xã": "", "Số tờ": "", "Số thửa": "", "Khuyến nghị": "Vẫn cần chuyên viên kiểm tra trước khi sử dụng" }]);
    reportSheet["!cols"] = [{ wch: 15 }, { wch: 28 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 48 }];
    XLSX.utils.book_append_sheet(workbook, reportSheet, "BAO_CAO_KIEM_TRA");
    XLSX.writeFile(workbook, `${safeBaseName(file.name)}_DA_LAM_SACH.xlsx`, { compression: true });
  }

  function reset() {
    setFile(null);
    setResult(null);
    setStatus("idle");
    setError("");
    setProgress({ label: "Đang chuẩn bị…", percent: 0 });
    setLandFields({ communeCode: "", mapSheet: "", parcel: "", area: "", prefix: "CHUACOGIAY", suffix: "GT" });
  }

  function updateLandField<Key extends keyof LandFields>(key: Key, value: LandFields[Key]) {
    setLandFields((current) => ({ ...current, [key]: value }));
  }

  function downloadRenamedCopy() {
    if (!file || !suggestedFileName) return;
    const url = URL.createObjectURL(file);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = suggestedFileName;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function exportHistory() {
    const rows: CellValue[][] = [["STT", "Thời gian", "Tệp nguồn", "Loại", "Kết quả", "Tên đề xuất"], ...history.map((entry, index) => [index + 1, new Date(entry.processedAt).toLocaleString("vi-VN"), entry.fileName, entry.fileType, entry.result, entry.suggestedName])];
    const url = URL.createObjectURL(new Blob(["\uFEFF", rows.map((row) => row.map(escapeCsv).join(",")).join("\r\n")], { type: "text/csv;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `NHAT_KY_XU_LY_${new Date().toISOString().slice(0, 10)}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function clearHistory() {
    localStorage.removeItem("sy-land-processing-history");
    setHistory([]);
  }

  return (
    <>
    <div className="file-workspace-head">
      <div><span className="demo-label">TRUNG TÂM XỬ LÝ TỆP</span><h3>Xử lý hồ sơ theo quy trình rõ ràng</h3><p>Chọn đúng khu vực công việc; dữ liệu được xử lý cục bộ và không ghi đè tệp gốc.</p></div>
      <nav aria-label="Điều hướng Xử lý tệp"><a href="#xu-ly-don">01 · Một tệp</a><a href="#xu-ly-hang-loat">02 · Hàng loạt</a><a href="#doi-chieu-du-lieu">03 · Đối chiếu</a><a href="#nhat-ky-xu-ly">04 · Nhật ký</a></nav>
    </div>
    <div className="processor-shell" id="xu-ly-don">
      <details className="location-profile">
        <summary><div><span className="location-glyph" aria-hidden="true">⌖</span><div><strong>Địa bàn xử lý hồ sơ</strong><small>Khai báo một lần để nhận diện địa danh và mã xã trên toàn quốc.</small></div></div><span>{locationProfile.communeCode ? `Mã xã ${locationProfile.communeCode}` : "Chưa khai báo"}</span></summary>
        <div className="location-form">
          <label>Tỉnh/thành phố hiện nay<input value={locationProfile.provinceNew} placeholder="Ví dụ: Thái Nguyên" onChange={(event) => setLocationProfile((current) => ({ ...current, provinceNew: event.target.value }))} /></label>
          <label>Tỉnh/thành phố trước sắp xếp<input value={locationProfile.provinceOld} placeholder="Nếu có" onChange={(event) => setLocationProfile((current) => ({ ...current, provinceOld: event.target.value }))} /></label>
          <label>Xã/phường hiện nay<input value={locationProfile.communeNew} placeholder="Tên xã/phường mới" onChange={(event) => setLocationProfile((current) => ({ ...current, communeNew: event.target.value }))} /></label>
          <label>Xã/phường trước sắp xếp<input value={locationProfile.communeOld} placeholder="Tên xã/phường cũ" onChange={(event) => setLocationProfile((current) => ({ ...current, communeOld: event.target.value }))} /></label>
          <label>Mã xã dùng đặt tên tệp<input value={locationProfile.communeCode} inputMode="numeric" maxLength={5} placeholder="Mã 5 chữ số" onChange={(event) => setLocationProfile((current) => ({ ...current, communeCode: event.target.value.replace(/\D/g, "").slice(0, 5) }))} /></label>
          <label>Thôn/bản/xóm hiện nay<input value={locationProfile.villageNew} placeholder="Tên hiện nay" onChange={(event) => setLocationProfile((current) => ({ ...current, villageNew: event.target.value }))} /></label>
          <label>Thôn/bản/xóm trước sắp xếp<input value={locationProfile.villageOld} placeholder="Tên cũ" onChange={(event) => setLocationProfile((current) => ({ ...current, villageOld: event.target.value }))} /></label>
          <label className="location-aliases">Tên gọi khác hoặc bảng thôn → mã xã<textarea value={locationProfile.aliases} placeholder={"Tên gọi khác, cách nhau bằng dấu phẩy; hoặc mỗi xã một dòng:\n02146 = Khuổi Phầy, Kim Vân, Quốc Tuấn\n02143 = Bản Giang, Pàn Xả, Khuổi Nộc"} onChange={(event) => setLocationProfile((current) => ({ ...current, aliases: event.target.value }))} /></label>
          <div className="location-actions"><p><span aria-hidden="true">✓</span>Cấu hình được lưu cục bộ trên thiết bị và áp dụng cho cả xử lý một tệp, hàng loạt và đối chiếu.</p><input ref={locationImportRef} type="file" accept="application/json,.json" onChange={(event) => void importLocationProfile(event.target.files?.[0])} aria-label="Nhập cấu hình địa bàn" /><button type="button" onClick={exportLocationProfile}>Xuất cấu hình</button><button type="button" onClick={() => locationImportRef.current?.click()}>Nhập cấu hình</button><button type="button" onClick={() => setLocationProfile(EMPTY_LOCATION)}>Xóa cấu hình</button></div>
        </div>
      </details>
      <div className="privacy-strip">
        <span className="privacy-icon" aria-hidden="true">✓</span>
        <div><strong>Tệp được xử lý ngay trên thiết bị</strong><small>Không tải lên hoặc lưu trữ trên máy chủ trong phiên bản này.</small></div>
        <span className="supported-types">DOCX · PDF · XLSX · XLS · CSV</span>
      </div>

      <div className="processor-grid">
        <section className="upload-panel" aria-labelledby="upload-title">
          <span className="demo-label">01 · CHỌN TỆP</span>
          <h3 id="upload-title">Đưa hồ sơ vào khu vực xử lý</h3>
          <p>Chọn một tệp Word, PDF hoặc Excel. Dung lượng tối đa 20 MB.</p>
          <div
            className={`drop-zone ${dragActive ? "drag-active" : ""}`}
            onDragEnter={(event) => { event.preventDefault(); setDragActive(true); }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={() => setDragActive(false)}
            onDrop={onDrop}
          >
            <input ref={inputRef} type="file" accept={ACCEPTED} onChange={onInputChange} aria-label="Chọn tệp Word, PDF hoặc Excel" />
            <span className="upload-glyph" aria-hidden="true">↑</span>
            <strong>Kéo thả tệp vào đây</strong>
            <span>hoặc</span>
            <button type="button" onClick={() => inputRef.current?.click()}>Chọn tệp từ thiết bị</button>
          </div>
          {file && (
            <div className="selected-file">
              <span className={`type-badge type-${file.name.split(".").pop()?.toLowerCase()}`}>{file.name.split(".").pop()?.toUpperCase()}</span>
              <div><strong>{file.name}</strong><small>{formatBytes(file.size)}</small></div>
              <button type="button" onClick={reset} aria-label="Bỏ tệp đã chọn">×</button>
            </div>
          )}
        </section>

        <section className="processing-panel" aria-labelledby="processing-title">
          <span className="demo-label">02 · ĐỌC VÀ KIỂM TRA</span>
          <h3 id="processing-title">Kết quả phân tích</h3>

          {status === "idle" && <div className="empty-result"><span aria-hidden="true">▤</span><p>Kết quả sẽ xuất hiện tại đây sau khi chọn tệp.</p></div>}
          {status === "processing" && <div className="processing-state" aria-live="polite"><span className="spinner" aria-hidden="true" /><strong>{progress.label}</strong><div className="ocr-progress"><i style={{ width: `${progress.percent}%` }} /></div><b>{progress.percent}%</b><p>PDF scan có thể mất vài phút. Không đóng trang trong khi OCR đang chạy.</p></div>}
          {status === "error" && <div className="error-state" role="alert"><strong>Chưa xử lý được tệp</strong><p>{error}</p><button type="button" onClick={() => inputRef.current?.click()}>Chọn tệp khác</button></div>}

          {status === "done" && result && (
            <div className="processed-result" aria-live="polite">
              <div className="result-title-row"><span className="result-check">✓</span><div><h4>{result.title}</h4><p>{result.description}</p>{result.ocr && <span className="ocr-badge">OCR · Tiếng Việt</span>}</div></div>
              <div className="metric-grid">{result.metrics.map((metric) => <div key={metric.label}><strong>{metric.value}</strong><span>{metric.label}</span></div>)}</div>
              {result.warning && <p className="result-warning"><span aria-hidden="true">!</span>{result.warning}</p>}

              {result.rows ? (
                <div className="table-preview" tabIndex={0} aria-label={`Xem trước sheet ${result.sheetName}`}>
                  <table><tbody>{result.rows.slice(0, 200).map((row, rowIndex) => <tr key={rowIndex}>{row.slice(0, 20).map((cell, cellIndex) => rowIndex === 0 ? <th key={cellIndex}>{String(cell)}</th> : <td key={cellIndex}>{String(cell)}</td>)}</tr>)}</tbody></table>
                </div>
              ) : (
                <div className="text-preview" tabIndex={0}>{result.text}</div>
              )}

              {result.kind !== "excel" && (
                <div className="land-extraction">
                  <div className="extraction-heading"><div><span className="demo-label">03 · THÔNG TIN ĐỊA CHÍNH</span><h4>Kiểm tra trường dữ liệu đã nhận diện</h4></div><span className="review-chip">Cần chuyên viên xác nhận</span></div>
                  <div className="land-field-grid">
                    <label>Mã xã<input value={landFields.communeCode} inputMode="numeric" maxLength={5} placeholder="02140" onChange={(event) => updateLandField("communeCode", event.target.value.replace(/\D/g, "").slice(0, 5))} /></label>
                    <label>Số tờ bản đồ<input value={landFields.mapSheet} inputMode="numeric" maxLength={6} placeholder="29 hoặc 110000" onChange={(event) => updateLandField("mapSheet", event.target.value.replace(/\D/g, "").slice(0, 6))} /></label>
                    <label>Số thửa<input value={landFields.parcel} inputMode="numeric" maxLength={7} placeholder="199" onChange={(event) => updateLandField("parcel", event.target.value.replace(/\D/g, "").slice(0, 7))} /></label>
                    <label>Diện tích (m²)<input value={landFields.area} inputMode="decimal" placeholder="1.253,6" onChange={(event) => updateLandField("area", event.target.value.replace(/[^0-9.,]/g, ""))} /></label>
                    <label>Nhóm hồ sơ<select value={landFields.prefix} onChange={(event) => updateLandField("prefix", event.target.value as LandFields["prefix"])}><option value="CHUACOGIAY">Chưa có giấy</option><option value="COGCN">Có GCN</option></select></label>
                    <label>Hậu tố<select value={landFields.suffix} onChange={(event) => updateLandField("suffix", event.target.value as LandFields["suffix"])}><option value="GT">GT</option><option value="TBXN">TBXN</option><option value="DDK">DDK</option></select></label>
                  </div>
                  {landReviewNotes.length > 0 && <div className="land-review-list" role="status"><strong>Kiểm tra trước khi xuất:</strong><ul>{landReviewNotes.map((note) => <li key={note}>{note}</li>)}</ul></div>}
                  <div className={`filename-suggestion ${suggestedFileName ? "complete" : "incomplete"}`}>
                    <span>Tên tệp đề xuất</span>
                    <strong>{suggestedFileName || "Cần nhập đủ mã xã, số tờ và số thửa"}</strong>
                    {suggestedFileName && <button type="button" onClick={downloadRenamedCopy}>Tải bản sao theo tên chuẩn <span aria-hidden="true">↓</span></button>}
                  </div>
                  <p className="field-rule"><span aria-hidden="true">i</span>Hồ sơ RSX/RPH/RDD dùng tờ 1, 2, 3 được chuẩn hóa thành 110000, 210000, 310000. Tệp gốc không bị đổi tên hay ghi đè.</p>
                </div>
              )}

              <div className="processor-actions"><button type="button" className="download-button" onClick={exportResult}>Tải kết quả <span aria-hidden="true">↓</span></button>{result.kind === "excel" && <button type="button" className="secondary-action" onClick={() => void exportCleanedWorkbook()}>Tải Excel đã làm sạch</button>}<button type="button" className="secondary-action" onClick={() => inputRef.current?.click()}>Xử lý tệp khác</button></div>
            </div>
          )}
        </section>
      </div>
    </div>
    <BatchProcessor onProcessed={addHistory} locationProfile={locationProfile} />
    <section className="local-history" id="nhat-ky-xu-ly" aria-labelledby="history-title">
      <div className="history-heading"><div><span className="demo-label">06 · NHẬT KÝ CỤC BỘ</span><h3 id="history-title">Lịch sử xử lý gần đây</h3><p>Chỉ lưu tên tệp và trạng thái trên thiết bị này; không lưu nội dung hồ sơ.</p></div><div><button type="button" onClick={exportHistory} disabled={!history.length}>Tải CSV</button><button type="button" onClick={clearHistory} disabled={!history.length}>Xóa lịch sử</button></div></div>
      {!history.length ? <p className="history-empty">Chưa có hoạt động. Nhật ký sẽ xuất hiện sau khi xử lý tệp.</p> : <div className="history-table-wrap"><table><thead><tr><th>Thời gian</th><th>Tệp nguồn</th><th>Loại</th><th>Kết quả</th><th>Tên đề xuất</th></tr></thead><tbody>{history.map((entry) => <tr key={entry.id}><td>{new Date(entry.processedAt).toLocaleString("vi-VN")}</td><td>{entry.fileName}</td><td>{entry.fileType}</td><td><span>{entry.result}</span></td><td><code>{entry.suggestedName || "—"}</code></td></tr>)}</tbody></table></div>}
    </section>
    </>
  );
}
