"use client";

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";

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
  suffix: "GT" | "TBXN" | "DDK";
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

const ACCEPTED = ".docx,.pdf,.xlsx,.xls,.csv";
const MAX_FILE_SIZE = 20 * 1024 * 1024;

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

function findVillageCode(text: string) {
  const normalized = text.toLocaleLowerCase("vi-VN");
  return VILLAGE_CODE_MAP.find((entry) => entry.names.some((name) => normalized.includes(name)))?.code || "";
}

function cleanNumber(value?: string) {
  return (value || "").replace(/\s+/g, "").trim();
}

function extractLandFields(text: string, fileName: string): LandFields {
  const fileMatch = fileName.match(/(?:CHUACOGIAY|COGCN)_([0-9]{4,5})_([0-9]{1,6})_([0-9]{1,7})/i);
  const explicitCode = text.match(/m[aã]\s*x[aã]\s*[:\-]?\s*([0-9]{4,5})/iu)?.[1];
  const parcelAddress = text.match(/(?:địa\s*chỉ\s*thửa\s*đất|thửa\s*đất\s*tại)([\s\S]{0,280})/iu)?.[1] || "";
  const communeCode = (fileMatch?.[1] || explicitCode || findVillageCode(parcelAddress) || findVillageCode(text)).padStart(5, "0");
  const mapSheetFromText = text.match(/tờ\s*bản\s*đồ(?:\s*địa\s*chính)?\s*(?:số)?\s*[:\-]?\s*([0-9]{1,6})/iu)?.[1];
  const parcelFromText = text.match(/(?:thửa\s*đất|thửa)\s*(?:số)?\s*[:\-]?\s*([0-9]{1,7})/iu)?.[1];
  const area = text.match(/diện\s*tích(?:\s*thửa\s*đất)?\s*[:\-]?\s*([0-9][0-9.,\s]*)\s*m(?:2|²)/iu)?.[1];
  const forestry = /\b(?:RSX|RPH|RDD)\b/i.test(text);
  let mapSheet = cleanNumber(fileMatch?.[2] || mapSheetFromText);
  if (forestry && mapSheet === "1") mapSheet = "110000";
  if (forestry && mapSheet === "2") mapSheet = "210000";

  return {
    communeCode,
    mapSheet,
    parcel: cleanNumber(fileMatch?.[3] || parcelFromText),
    area: cleanNumber(area),
    prefix: fileName.toUpperCase().startsWith("COGCN_") ? "COGCN" : "CHUACOGIAY",
    suffix: fileName.toUpperCase().includes("_TBXN") ? "TBXN" : fileName.toUpperCase().includes("_DDK") ? "DDK" : "GT",
  };
}

function buildSuggestedName(file: File, fields: LandFields) {
  if (!fields.communeCode || !fields.mapSheet || !fields.parcel) return "";
  const extension = file.name.split(".").pop()?.toLowerCase() || "pdf";
  return `${fields.prefix}_${fields.communeCode}_${fields.mapSheet}_${fields.parcel}_${fields.suffix}.${extension}`;
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

  return {
    kind: "excel",
    title: "Đã đọc bảng dữ liệu",
    description: `Đang hiển thị sheet “${sheetName}”. Có thể xem trước dữ liệu và xuất lại thành CSV chuẩn hóa.`,
    metrics: [
      { label: "Số sheet", value: String(workbook.SheetNames.length) },
      { label: "Số dòng", value: rows.length.toLocaleString("vi-VN") },
      { label: "Số cột", value: columns.toLocaleString("vi-VN") },
      { label: "Ô có dữ liệu", value: filled.toLocaleString("vi-VN") },
    ],
    rows,
    sheetName,
    warning: rows.length > 200 ? "Bảng xem trước chỉ hiển thị 200 dòng đầu; tệp xuất vẫn giữ toàn bộ dữ liệu đã đọc." : undefined,
  };
}

function BatchProcessor({ onProcessed }: { onProcessed: (record: HistoryRecord) => void }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const masterInputRef = useRef<HTMLInputElement>(null);
  const [items, setItems] = useState<BatchItem[]>([]);
  const [running, setRunning] = useState(false);
  const [master, setMaster] = useState<MasterSheet | null>(null);
  const [masterError, setMasterError] = useState("");
  const [comparisonFilter, setComparisonFilter] = useState<"all" | "matched" | "unmatched" | "missing">("all");
  const [batchSearch, setBatchSearch] = useState("");
  const [batchSort, setBatchSort] = useState<"added" | "name" | "parcel" | "comparison" | "status">("added");
  const [batchPrefix, setBatchPrefix] = useState<LandFields["prefix"]>("CHUACOGIAY");
  const [batchSuffix, setBatchSuffix] = useState<LandFields["suffix"]>("GT");
  const [archiveMode, setArchiveMode] = useState<"flat" | "folders">("folders");
  const [archiveRoot, setArchiveRoot] = useState("HO_SO_DAT_DAI");

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
    const selected = Array.from(fileList).slice(0, 50);
    const existingKeys = new Set(items.map((item) => `${item.file.name}-${item.file.size}-${item.file.lastModified}`));
    const next = selected
      .filter((file) => !existingKeys.has(`${file.name}-${file.size}-${file.lastModified}`))
      .map<BatchItem>((file, index) => {
        const extension = file.name.split(".").pop()?.toLowerCase() || "";
        const valid = ["docx", "pdf", "xlsx", "xls", "csv"].includes(extension) && file.size <= MAX_FILE_SIZE;
        const fields = extractLandFields("", file.name);
        fields.prefix = batchPrefix;
        fields.suffix = batchSuffix;
        return {
          id: `${file.name}-${file.size}-${file.lastModified}-${index}`,
          file,
          status: valid ? "waiting" : "error",
          fields,
          message: valid ? "Chờ xử lý" : "Sai định dạng hoặc vượt quá 20 MB",
        };
      });
    setItems((current) => [...current, ...next].slice(0, 50));
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

  async function processAll() {
    if (!items.length || running) return;
    setRunning(true);
    for (const item of items) {
      if (item.status === "error") continue;
      updateItem(item.id, { status: "processing", message: "Đang đọc nội dung…" });
      try {
        const extension = item.file.name.split(".").pop()?.toLowerCase();
        let text = "";
        let scanNote = "";
        if (extension === "docx") {
          text = (await processWord(item.file)).text || "";
        } else if (extension === "pdf") {
          const output = await processPdf(item.file, () => undefined, false);
          text = output.text || "";
          scanNote = output.warning || "";
        } else {
          await processExcel(item.file);
        }
        const extracted = extractLandFields(text, item.file.name);
        extracted.prefix = item.fields.prefix;
        extracted.suffix = item.fields.suffix;
        const complete = Boolean(extracted.communeCode && extracted.mapSheet && extracted.parcel);
        updateItem(item.id, {
          fields: extracted,
          status: complete ? "done" : "review",
          message: complete ? (scanNote || "Đã nhận diện đủ trường") : (scanNote || "Cần bổ sung trường còn thiếu"),
        });
        onProcessed({ id: `${Date.now()}-${item.id}`, fileName: item.file.name, fileType: extension?.toUpperCase() || "TỆP", processedAt: new Date().toISOString(), result: complete ? "Hoàn thành" : "Cần rà soát", suggestedName: buildSuggestedName(item.file, extracted) });
      } catch (reason) {
        console.error(reason);
        updateItem(item.id, { status: "error", message: "Không đọc được tệp" });
      }
    }
    setRunning(false);
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
    const zipped = zipSync(archive, { level: 6 });
    const blob = new Blob([zipped], { type: "application/zip" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `HO_SO_DA_CHUAN_HOA_${new Date().toISOString().slice(0, 10)}.zip`;
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
    <section className="batch-shell" aria-labelledby="batch-title">
      <div className="batch-heading">
        <div><span className="demo-label">04 · XỬ LÝ HÀNG LOẠT</span><h3 id="batch-title">Chuẩn hóa nhiều hồ sơ trong một lượt</h3><p>Tối đa 50 tệp, 20 MB mỗi tệp. PDF scan được đánh dấu để OCR riêng nhằm tránh quá tải.</p></div>
        <div className="batch-summary"><strong>{items.length}</strong><span>Tệp đã chọn</span><strong>{readyCount}</strong><span>Sẵn sàng tải</span></div>
      </div>
      <div className="batch-toolbar">
        <label>Nhóm hồ sơ<select value={batchPrefix} onChange={(event) => setBatchPrefix(event.target.value as LandFields["prefix"])}><option value="CHUACOGIAY">Chưa có giấy</option><option value="COGCN">Có GCN</option></select></label>
        <label>Hậu tố<select value={batchSuffix} onChange={(event) => setBatchSuffix(event.target.value as LandFields["suffix"])}><option value="GT">GT</option><option value="TBXN">TBXN</option><option value="DDK">DDK</option></select></label>
        <label>Cách đóng gói<select value={archiveMode} onChange={(event) => setArchiveMode(event.target.value as "flat" | "folders")}><option value="folders">Chia thư mục tự động</option><option value="flat">Một thư mục</option></select></label>
        {archiveMode === "folders" && <label className="archive-root-label">Tên thư mục gốc<input value={archiveRoot} maxLength={50} onChange={(event) => setArchiveRoot(event.target.value)} placeholder="HO_SO_DAT_DAI" /></label>}
        <button type="button" className="secondary-action" onClick={applyDefaults} disabled={!items.length}>Áp dụng cho tất cả</button>
        <input ref={inputRef} type="file" accept={ACCEPTED} multiple onChange={(event) => addFiles(event.target.files)} aria-label="Chọn nhiều tệp hồ sơ" />
        <button type="button" className="batch-add" onClick={() => inputRef.current?.click()}>+ Chọn nhiều tệp</button>
      </div>

      {archiveMode === "folders" && <div className="folder-preview" aria-label="Cấu trúc thư mục ZIP"><span>ZIP sẽ được sắp xếp:</span><code>{safeBaseName(archiveRoot.trim() || "HO_SO_DAT_DAI")} / COGCN hoặc CHUACOGIAY / GT, TBXN hoặc DDK / tệp</code></div>}

      <div className="master-panel">
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
        <button type="button" className="batch-empty" onClick={() => inputRef.current?.click()}><span aria-hidden="true">＋</span><strong>Chọn nhiều Word, PDF hoặc Excel</strong><small>Hệ thống sẽ xử lý lần lượt và không ghi đè tệp gốc.</small></button>
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
                <td><strong>{item.file.name}</strong><small>{formatBytes(item.file.size)}</small></td>
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
        <button type="button" className="button batch-download" onClick={downloadZip} disabled={!readyCount || running}>Tải {readyCount} tệp dạng ZIP <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={downloadComparisonReport} disabled={!master || !items.length}>Tải báo cáo CSV <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={() => void downloadReviewWorkbook()} disabled={!master || (!unmatchedCount && !comparison.includes("missing"))}>Xuất hồ sơ cần rà soát (.xlsx) <span aria-hidden="true">↓</span></button>
        <button type="button" className="button report-download" onClick={() => void downloadMasterDuplicates()} disabled={!master || !duplicateMasterCount}>Xuất {duplicateMasterCount} khóa trùng (.xlsx) <span aria-hidden="true">↓</span></button>
        {duplicateNames.size > 0 && <p><span aria-hidden="true">!</span>Có {duplicateNames.size} tên bị trùng. Hãy sửa mã xã, số tờ hoặc số thửa trước khi tải.</p>}
      </div>
    </section>
  );
}

export default function FileProcessor() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<ProcessedResult | null>(null);
  const [status, setStatus] = useState<"idle" | "processing" | "done" | "error">("idle");
  const [error, setError] = useState("");
  const [dragActive, setDragActive] = useState(false);
  const [progress, setProgress] = useState<ProcessingProgress>({ label: "Đang chuẩn bị…", percent: 0 });
  const [landFields, setLandFields] = useState<LandFields>({ communeCode: "", mapSheet: "", parcel: "", area: "", prefix: "CHUACOGIAY", suffix: "GT" });
  const [history, setHistory] = useState<HistoryRecord[]>([]);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      try {
        const stored = localStorage.getItem("sy-land-processing-history");
        if (stored) setHistory(JSON.parse(stored));
      } catch { /* Bỏ qua lịch sử không hợp lệ. */ }
    }, 0);
    return () => window.clearTimeout(timer);
  }, []);

  function addHistory(record: HistoryRecord) {
    setHistory((current) => {
      const next = [record, ...current].slice(0, 30);
      localStorage.setItem("sy-land-processing-history", JSON.stringify(next));
      return next;
    });
  }

  const suggestedFileName = useMemo(() => {
    if (!file || !landFields.communeCode || !landFields.mapSheet || !landFields.parcel) return "";
    const extension = file.name.split(".").pop()?.toLowerCase() || "pdf";
    return `${landFields.prefix}_${landFields.communeCode}_${landFields.mapSheet}_${landFields.parcel}_${landFields.suffix}.${extension}`;
  }, [file, landFields]);

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
        setLandFields(extractLandFields(processed.text, nextFile.name));
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
    <div className="processor-shell">
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
                  <div className={`filename-suggestion ${suggestedFileName ? "complete" : "incomplete"}`}>
                    <span>Tên tệp đề xuất</span>
                    <strong>{suggestedFileName || "Cần nhập đủ mã xã, số tờ và số thửa"}</strong>
                    {suggestedFileName && <button type="button" onClick={downloadRenamedCopy}>Tải bản sao theo tên chuẩn <span aria-hidden="true">↓</span></button>}
                  </div>
                  <p className="field-rule"><span aria-hidden="true">i</span>Số tờ lâm nghiệp 110000 hoặc 210000 được giữ nguyên. Tệp gốc không bị đổi tên hay ghi đè.</p>
                </div>
              )}

              <div className="processor-actions"><button type="button" className="download-button" onClick={exportResult}>Tải kết quả <span aria-hidden="true">↓</span></button><button type="button" className="secondary-action" onClick={() => inputRef.current?.click()}>Xử lý tệp khác</button></div>
            </div>
          )}
        </section>
      </div>
    </div>
    <BatchProcessor onProcessed={addHistory} />
    <section className="local-history" aria-labelledby="history-title">
      <div className="history-heading"><div><span className="demo-label">06 · NHẬT KÝ CỤC BỘ</span><h3 id="history-title">Lịch sử xử lý gần đây</h3><p>Chỉ lưu tên tệp và trạng thái trên thiết bị này; không lưu nội dung hồ sơ.</p></div><div><button type="button" onClick={exportHistory} disabled={!history.length}>Tải CSV</button><button type="button" onClick={clearHistory} disabled={!history.length}>Xóa lịch sử</button></div></div>
      {!history.length ? <p className="history-empty">Chưa có hoạt động. Nhật ký sẽ xuất hiện sau khi xử lý tệp.</p> : <div className="history-table-wrap"><table><thead><tr><th>Thời gian</th><th>Tệp nguồn</th><th>Loại</th><th>Kết quả</th><th>Tên đề xuất</th></tr></thead><tbody>{history.map((entry) => <tr key={entry.id}><td>{new Date(entry.processedAt).toLocaleString("vi-VN")}</td><td>{entry.fileName}</td><td>{entry.fileType}</td><td><span>{entry.result}</span></td><td><code>{entry.suggestedName || "—"}</code></td></tr>)}</tbody></table></div>}
    </section>
    </>
  );
}
