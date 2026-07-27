"use client";

import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import { useEntitlements } from "./entitlements";
import "./creator-studio.css";

type SourceInspection = {
  platform: string;
  visibility: "public" | "unlisted" | "private" | "restricted" | "unknown";
  ownership: "verified" | "declared" | "unverified";
  license: string;
  risk: "low" | "medium" | "high" | "blocked";
  canProcess: boolean;
  reasons: string[];
  sourceTitle?: string;
  sourceCreator?: string;
};

type CreatorJob = {
  id: string;
  status: "queued" | "downloading" | "transcribing" | "translating" | "dubbing" | "rendering" | "completed" | "failed" | "cancelled";
  progress: number;
  message: string;
  outputUrl?: string;
};

const API_BASE = String(import.meta.env.VITE_SYLAND_CREATOR_API_URL || "").trim().replace(/\/$/, "");
const REMOTE_SESSION_KEY = "sy-land-auth-session";
const MAX_UPLOAD_BYTES = 500 * 1024 * 1024;
const SUPPORTED_FILE_TYPES = ["video/mp4", "video/quicktime", "video/webm", "video/x-matroska"];

const progressLabels: Record<CreatorJob["status"], string> = {
  queued: "Đang chờ máy chủ",
  downloading: "Đang tiếp nhận video",
  transcribing: "Whisper đang nhận diện lời thoại",
  translating: "AI đang dịch sang tiếng Việt",
  dubbing: "Đang tạo giọng đọc tiếng Việt",
  rendering: "Đang dựng video và phụ đề",
  completed: "Đã hoàn thành",
  failed: "Xử lý thất bại",
  cancelled: "Đã hủy tác vụ",
};

function readAccessToken() {
  try {
    return String(JSON.parse(localStorage.getItem(REMOTE_SESSION_KEY) || "null")?.accessToken || "");
  } catch {
    return "";
  }
}

function platformFromUrl(value: string) {
  try {
    const host = new URL(value).hostname.toLowerCase().replace(/^www\./, "");
    if (host.includes("youtube.com") || host === "youtu.be") return "YouTube";
    if (host.includes("tiktok.com")) return "TikTok";
    if (host.includes("douyin.com")) return "Douyin";
    if (host.includes("bilibili.com") || host === "b23.tv") return "Bilibili";
    if (host.includes("instagram.com")) return "Instagram";
    if (host.includes("facebook.com") || host === "fb.watch") return "Facebook";
    return host;
  } catch {
    return "";
  }
}

function visibilityLabel(value: SourceInspection["visibility"]) {
  if (value === "public") return "Công khai";
  if (value === "unlisted") return "Không công khai";
  if (value === "private") return "Riêng tư";
  if (value === "restricted") return "Bị giới hạn";
  return "Chưa xác định";
}

function ownershipLabel(value: SourceInspection["ownership"]) {
  if (value === "verified") return "Đã xác minh";
  if (value === "declared") return "Người dùng xác nhận";
  return "Chưa xác minh";
}

function riskLabel(value: SourceInspection["risk"]) {
  if (value === "low") return "Rủi ro thấp";
  if (value === "medium") return "Cần rà soát";
  if (value === "high") return "Rủi ro cao";
  return "Không được xử lý";
}

async function apiRequest(path: string, init: RequestInit = {}) {
  if (!API_BASE) throw new Error("Dịch vụ xử lý video chưa được cấu hình trên website.");
  const token = readAccessToken();
  if (!token) throw new Error("Hãy đăng nhập tài khoản SỸ LAND trước khi sử dụng Creator.");
  const response = await fetch(`${API_BASE}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init.body instanceof FormData ? {} : { "Content-Type": "application/json" }),
      ...(init.headers || {}),
    },
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.message || data.detail || "Máy chủ Creator từ chối yêu cầu.");
  return data;
}

function CreatorStudio() {
  const entitlements = useEntitlements();
  const fileInput = useRef<HTMLInputElement>(null);
  const pollTimer = useRef<number | null>(null);
  const [sourceMode, setSourceMode] = useState<"link" | "upload">("link");
  const [sourceUrl, setSourceUrl] = useState("");
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [provider, setProvider] = useState("openai");
  const [voice, setVoice] = useState("female-north");
  const [whisperModel, setWhisperModel] = useState("small");
  const [backgroundVolume, setBackgroundVolume] = useState(10);
  const [subtitleSize, setSubtitleSize] = useState("medium");
  const [rightsConfirmed, setRightsConfirmed] = useState(false);
  const [inspection, setInspection] = useState<SourceInspection | null>(null);
  const [inspectionState, setInspectionState] = useState<"idle" | "checking" | "done" | "error">("idle");
  const [job, setJob] = useState<CreatorJob | null>(null);
  const [message, setMessage] = useState("");

  const preliminaryPlatform = useMemo(() => platformFromUrl(sourceUrl), [sourceUrl]);
  const accountReady = entitlements.role !== "guest";
  const jobActive = Boolean(job && !["completed", "failed", "cancelled"].includes(job.status));
  const canCreate = Boolean(
    accountReady &&
    rightsConfirmed &&
    inspection?.canProcess &&
    (sourceMode === "link" ? sourceUrl.trim() : sourceFile) &&
    !job?.id,
  );

  useEffect(() => () => {
    if (pollTimer.current) window.clearTimeout(pollTimer.current);
  }, []);

  const safeOutputUrl = useMemo(() => {
    if (!job?.outputUrl) return "";
    try {
      const value = new URL(job.outputUrl);
      return value.protocol === "https:" ? value.toString() : "";
    } catch {
      return "";
    }
  }, [job?.outputUrl]);

  function resetResult() {
    setInspection(null);
    setInspectionState("idle");
    setMessage("");
    setJob(null);
    if (pollTimer.current) window.clearTimeout(pollTimer.current);
  }

  function changeFile(event: ChangeEvent<HTMLInputElement>) {
    const next = event.target.files?.[0] || null;
    if (!next) return;
    if (next.size > MAX_UPLOAD_BYTES) {
      setMessage("Video vượt quá 500 MB. Hãy chọn video ngắn hơn hoặc nén video trước.");
      event.target.value = "";
      return;
    }
    if (!SUPPORTED_FILE_TYPES.includes(next.type) && !/\.(mp4|mov|webm|mkv)$/i.test(next.name)) {
      setMessage("Chỉ hỗ trợ MP4, MOV, WebM hoặc MKV.");
      event.target.value = "";
      return;
    }
    resetResult();
    setSourceFile(next);
  }

  async function inspectSource() {
    setMessage("");
    setJob(null);
    if (!accountReady) {
      setMessage("Hãy đăng nhập tài khoản SỸ LAND tại mục Tài khoản trước khi kiểm tra nguồn video.");
      return;
    }
    if (sourceMode === "link" && !sourceUrl.trim()) {
      setMessage("Hãy dán đường link video.");
      return;
    }
    if (sourceMode === "upload" && !sourceFile) {
      setMessage("Hãy chọn video từ máy.");
      return;
    }
    setInspectionState("checking");
    try {
      if (sourceMode === "upload") {
        setInspection({
          platform: "Tệp tải lên",
          visibility: "unknown",
          ownership: "declared",
          license: "Do người dùng khai báo",
          risk: "medium",
          canProcess: true,
          sourceTitle: sourceFile?.name,
          reasons: [
            "Không phát hiện nguồn nền tảng từ tệp cục bộ.",
            "Người dùng phải sở hữu hoặc được phép sử dụng video.",
            "Máy chủ sẽ kiểm tra định dạng và thời lượng; quyền sở hữu vẫn cần bằng chứng của người dùng.",
          ],
        });
      } else {
        const result = await apiRequest("/v1/video/inspect", {
          method: "POST",
          body: JSON.stringify({ url: sourceUrl.trim() }),
        });
        setInspection(result as SourceInspection);
      }
      setInspectionState("done");
    } catch (error) {
      setInspectionState("error");
      setMessage(error instanceof Error ? error.message : "Không thể kiểm tra nguồn video.");
    }
  }

  async function pollJob(jobId: string) {
    try {
      const next = await apiRequest(`/v1/video/jobs/${encodeURIComponent(jobId)}`) as CreatorJob;
      setJob(next);
      if (!["completed", "failed", "cancelled"].includes(next.status)) {
        pollTimer.current = window.setTimeout(() => pollJob(jobId), 3000);
      }
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Không thể cập nhật tiến độ.");
    }
  }

  async function createVideo(event: FormEvent) {
    event.preventDefault();
    if (!canCreate) return;
    setMessage("");
    try {
      const settings = {
        provider,
        voice,
        whisperModel,
        backgroundVolume,
        subtitleSize,
        rightsConfirmed,
      };
      let created: CreatorJob;
      if (sourceMode === "upload" && sourceFile) {
        const body = new FormData();
        body.append("video", sourceFile);
        body.append("settings", JSON.stringify(settings));
        created = await apiRequest("/v1/video/jobs/upload", { method: "POST", body }) as CreatorJob;
      } else {
        created = await apiRequest("/v1/video/jobs", {
          method: "POST",
          body: JSON.stringify({ url: sourceUrl.trim(), settings, inspection }),
        }) as CreatorJob;
      }
      setJob(created);
      pollJob(created.id);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Không thể tạo tác vụ video.");
    }
  }

  async function cancelJob() {
    if (!job || !jobActive) return;
    if (pollTimer.current) window.clearTimeout(pollTimer.current);
    setMessage("");
    try {
      const next = await apiRequest(`/v1/video/jobs/${encodeURIComponent(job.id)}/cancel`, {
        method: "POST",
      }) as CreatorJob;
      setJob(next);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Không thể hủy tác vụ.");
      pollJob(job.id);
    }
  }

  function startAnotherVideo() {
    if (pollTimer.current) window.clearTimeout(pollTimer.current);
    setJob(null);
    setMessage("");
  }

  return (
    <section className="creator-studio" aria-labelledby="creator-title">
      <div className="creator-heading">
        <div>
          <p className="section-kicker">SỸ LAND Creator</p>
          <h2 id="creator-title">Việt hóa Video bằng AI.</h2>
          <p>Dịch lời thoại, lồng tiếng Việt, ghép phụ đề và kiểm tra nguồn sử dụng trong một quy trình có kiểm soát.</p>
        </div>
        <div className="creator-account-badge">
          <span>{entitlements.role === "guest" ? "Chưa đăng nhập" : `Gói ${entitlements.plan}`}</span>
          <strong>{entitlements.fullTools ? "Creator nâng cao" : "Creator tiêu chuẩn"}</strong>
          <small>{entitlements.reason}</small>
        </div>
      </div>

      <div className="creator-shell">
        <aside className="creator-steps" aria-label="Quy trình Creator">
          <span className="creator-product-mark">C</span>
          <ol>
            <li className={inspectionState === "idle" ? "active" : ""}><span>1</span><div><b>Chọn nguồn</b><small>Link công khai hoặc tệp của bạn</small></div></li>
            <li className={inspectionState === "checking" ? "active" : inspection ? "complete" : ""}><span>2</span><div><b>Kiểm tra quyền</b><small>Riêng tư, giấy phép và rủi ro</small></div></li>
            <li className={jobActive ? "active" : job?.status === "completed" ? "complete" : ""}><span>3</span><div><b>AI Việt hóa</b><small>Whisper, dịch và lồng tiếng</small></div></li>
            <li className={job?.status === "completed" ? "complete" : ""}><span>4</span><div><b>Xuất MP4</b><small>Xem trước và tải kết quả</small></div></li>
          </ol>
          <div className="creator-safety-note"><b>Nguyên tắc xử lý</b><p>Không vượt đăng nhập, không tự mở video riêng tư và không khẳng định tuyệt đối về bản quyền.</p></div>
        </aside>

        <form className="creator-workspace" onSubmit={createVideo}>
          <div className="creator-mode-tabs" role="tablist" aria-label="Chọn nguồn video">
            <button type="button" disabled={jobActive} className={sourceMode === "link" ? "active" : ""} onClick={() => { setSourceMode("link"); resetResult(); }}>Dán đường link</button>
            <button type="button" disabled={jobActive} className={sourceMode === "upload" ? "active" : ""} onClick={() => { setSourceMode("upload"); resetResult(); }}>Tải video của bạn</button>
          </div>

          {sourceMode === "link" ? (
            <div className="creator-link-input">
              <label htmlFor="creator-url">Đường link video</label>
              <div><span aria-hidden="true">↗</span><input id="creator-url" disabled={jobActive} type="url" value={sourceUrl} onChange={(event) => { setSourceUrl(event.target.value); resetResult(); }} placeholder="https://youtube.com/... hoặc TikTok, Douyin, Bilibili" /></div>
              <small>{preliminaryPlatform ? `Đã nhận diện liên kết ${preliminaryPlatform}.` : "Hỗ trợ nguồn công khai; nguồn cần đăng nhập phải kết nối tài khoản chính chủ."}</small>
            </div>
          ) : (
            <div className="creator-upload">
              <input ref={fileInput} disabled={jobActive} type="file" accept="video/mp4,video/quicktime,video/webm,video/x-matroska,.mkv" onChange={changeFile} />
              <span aria-hidden="true">▶</span>
              <div><b>{sourceFile?.name || "Chọn video từ máy"}</b><small>{sourceFile ? `${(sourceFile.size / 1024 / 1024).toFixed(1)} MB` : "MP4, MOV, WebM hoặc MKV · tối đa 500 MB"}</small></div>
              <button type="button" disabled={jobActive} onClick={() => fileInput.current?.click()}>{sourceFile ? "Chọn lại" : "Chọn tệp"}</button>
            </div>
          )}

          <div className="creator-settings">
            <label>AI dịch<select disabled={jobActive} value={provider} onChange={(event) => setProvider(event.target.value)}><option value="openai">OpenAI</option><option value="gemini">Google Gemini</option></select></label>
            <label>Giọng đọc<select disabled={jobActive} value={voice} onChange={(event) => setVoice(event.target.value)}><option value="female-north">Nữ · Bắc</option><option value="male-north">Nam · Bắc</option><option value="female-south">Nữ · Nam</option><option value="male-south">Nam · Nam</option></select></label>
            <label>Whisper<select disabled={jobActive} value={whisperModel} onChange={(event) => setWhisperModel(event.target.value)}><option value="base">Base · nhanh</option><option value="small">Small · cân bằng</option><option value="medium">Medium · chính xác</option></select></label>
            <label>Cỡ phụ đề<select disabled={jobActive} value={subtitleSize} onChange={(event) => setSubtitleSize(event.target.value)}><option value="small">Nhỏ</option><option value="medium">Vừa</option><option value="large">Lớn</option></select></label>
          </div>

          <label className="creator-volume">Âm lượng tiếng gốc/nhạc nền <b>{backgroundVolume}%</b><input disabled={jobActive} type="range" min="0" max="30" step="1" value={backgroundVolume} onChange={(event) => setBackgroundVolume(Number(event.target.value))} /></label>

          <div className="creator-inspection">
            <div className="creator-inspection-head">
              <div><b>Kiểm tra quyền sử dụng</b><small>Đây là đánh giá rủi ro, không phải chứng nhận pháp lý.</small></div>
              <button type="button" disabled={inspectionState === "checking" || jobActive} onClick={inspectSource}>{inspectionState === "checking" ? "Đang kiểm tra…" : "Kiểm tra nguồn video"}</button>
            </div>

            {inspection && (
              <div className={`creator-risk risk-${inspection.risk}`}>
                <div className="creator-risk-title"><span aria-hidden="true">{inspection.risk === "low" ? "✓" : inspection.risk === "blocked" ? "×" : "!"}</span><div><strong>{riskLabel(inspection.risk)}</strong><small>{inspection.sourceTitle || inspection.platform}{inspection.sourceCreator ? ` · ${inspection.sourceCreator}` : ""}</small></div></div>
                <dl>
                  <div><dt>Nền tảng</dt><dd>{inspection.platform}</dd></div>
                  <div><dt>Quyền riêng tư</dt><dd>{visibilityLabel(inspection.visibility)}</dd></div>
                  <div><dt>Chủ sở hữu</dt><dd>{ownershipLabel(inspection.ownership)}</dd></div>
                  <div><dt>Giấy phép</dt><dd>{inspection.license || "Chưa xác định"}</dd></div>
                </dl>
                <ul>{inspection.reasons.map((reason) => <li key={reason}>{reason}</li>)}</ul>
              </div>
            )}
          </div>

          <label className="creator-rights-confirm">
            <input type="checkbox" disabled={jobActive} checked={rightsConfirmed} onChange={(event) => setRightsConfirmed(event.target.checked)} />
            <span>Tôi xác nhận sở hữu video hoặc đã được chủ sở hữu cho phép dịch, lồng tiếng và xuất bản nội dung này.</span>
          </label>

          {message && <p className="creator-message" role="alert">{message}</p>}

          {job && (
            <div className={`creator-job ${job.status}`}>
              <div><b>{progressLabels[job.status]}</b><span>{Math.max(0, Math.min(100, Math.round(job.progress)))}%</span></div>
              <progress max="100" value={job.progress} />
              <p>{job.message}</p>
              <div className="creator-job-controls">
                {job.status === "completed" && safeOutputUrl && <a className="button button-primary" href={safeOutputUrl} download>Tải video MP4 <span aria-hidden="true">↓</span></a>}
                {jobActive && <button type="button" className="creator-cancel" onClick={cancelJob}>Hủy tác vụ</button>}
                {!jobActive && <button type="button" className="creator-restart" onClick={startAnotherVideo}>Tạo video khác</button>}
              </div>
            </div>
          )}

          <div className="creator-actions">
            <button className="button button-primary" type="submit" disabled={!canCreate}>Tạo Video <span aria-hidden="true">→</span></button>
            <p>{!accountReady ? "Cần đăng nhập tài khoản SỸ LAND." : !inspection ? "Kiểm tra nguồn video trước khi xử lý." : !rightsConfirmed ? "Cần xác nhận quyền sử dụng." : inspection.canProcess ? "Sẵn sàng gửi tới máy chủ xử lý." : "Nguồn video này không được phép xử lý."}</p>
          </div>
        </form>
      </div>
    </section>
  );
}

export default CreatorStudio;
