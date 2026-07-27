import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";

const root = new URL("..", import.meta.url).pathname;
const dist = join(root, "pages-dist");
const failures = [];
const pass = [];

function assert(condition, message) {
  if (condition) pass.push(message);
  else failures.push(message);
}

assert(existsSync(join(dist, "github-pages.html")), "Có tệp HTML phát hành");
const html = readFileSync(join(dist, "github-pages.html"), "utf8");
const assetPaths = [...html.matchAll(/(?:src|href)="\/SY-LAND\/([^"]+)"/g)].map((match) => match[1]);
assert(assetPaths.length >= 3, "HTML tham chiếu đủ JS, CSS và favicon");
for (const assetPath of assetPaths) {
  assert(existsSync(join(dist, assetPath)), `Tài nguyên tồn tại: ${assetPath}`);
}

const assetsDir = join(dist, "assets");
const assets = readdirSync(assetsDir);
const entryJs = assets.find((name) => /^github-pages-.*\.js$/.test(name));
const entryCss = assets.find((name) => /^github-pages-.*\.css$/.test(name));
assert(Boolean(entryJs), "Có JavaScript khởi tạo");
assert(Boolean(entryCss), "Có CSS phát hành");
if (entryJs) assert(statSync(join(assetsDir, entryJs)).size < 300 * 1024, "JavaScript khởi tạo dưới 300 KB");
if (entryCss) assert(statSync(join(assetsDir, entryCss)).size < 140 * 1024, "CSS phát hành dưới 140 KB");

const sourceFiles = readdirSync(root).filter((name) => /\.(tsx|ts|css)$/.test(name));
const source = sourceFiles.map((name) => readFileSync(join(root, name), "utf8")).join("\n");
const page = readFileSync(join(root, "page.tsx"), "utf8");
const accountPortal = readFileSync(join(root, "account-portal.tsx"), "utf8");
const paymentCenter = readFileSync(join(root, "payment-center.tsx"), "utf8");
for (const id of ["minh-hoa", "cong-cu-pdf", "syland-creator", "tai-phan-mem", "tai-khoan", "thanh-toan", "phap-ly"]) {
  assert(page.includes(`id="${id}"`), `Có điểm điều hướng #${id}`);
}
assert(page.includes("lazy(() => import(\"./file-processor\"))"), "Xử lý tệp được tải theo nhu cầu");
assert(page.includes("lazy(() => import(\"./pdf-toolkit\"))"), "Công cụ PDF được tải theo nhu cầu");
assert(page.includes("lazy(() => import(\"./creator-studio\"))"), "SỸ LAND Creator được tải theo nhu cầu");
assert(!source.includes("@import \"tailwindcss\""), "Không nạp Tailwind không sử dụng");
assert(!source.includes("VITE_OPENAI_API_KEY"), "Không đưa OpenAI API Key vào trình duyệt");
assert(!source.includes("VITE_GEMINI_API_KEY"), "Không đưa Gemini API Key vào trình duyệt");
assert(accountPortal.includes('remoteAuth("/settings"'), "Google OAuth kiểm tra trạng thái nhà cung cấp trước khi chuyển hướng");
assert(accountPortal.includes("unable to exchange external code"), "Google OAuth giải thích lỗi Client Secret");
assert(!accountPortal.includes('url.searchParams.set("oauth"'), "OAuth dùng URL callback GitHub Pages chính xác, không thêm query thừa");
assert(paymentCenter.includes("SUPABASE_REPAIR_AUTH_PAYMENTS.sql"), "Thông báo thanh toán trỏ đúng bản phục hồi hợp nhất");

const browserOutput = assets
  .filter((name) => name.endsWith(".js"))
  .map((name) => readFileSync(join(assetsDir, name), "utf8"))
  .join("\n");
assert(!browserOutput.includes("service_role"), "Không đưa service_role vào website");
assert(!browserOutput.includes("3950549732"), "Không đưa số tài khoản ngân hàng vào mã website");
assert(!browserOutput.includes("admin/123"), "Không chứa tài khoản quản trị mặc định không an toàn");

const schemaSql = readFileSync(join(root, "SUPABASE_SCHEMA.sql"), "utf8");
const paymentSql = readFileSync(join(root, "SUPABASE_PAYMENTS.sql"), "utf8");
const paymentFix = readFileSync(join(root, "SUPABASE_PAYMENT_PLANS_FIX.sql"), "utf8");
const repairSql = readFileSync(join(root, "SUPABASE_REPAIR_AUTH_PAYMENTS.sql"), "utf8");
const stableAccountSql = readFileSync(join(root, "SUPABASE_STABLE_ACCOUNT_SYNC.sql"), "utf8");
const creatorSql = readFileSync(join(root, "SUPABASE_CREATOR.sql"), "utf8");
const creatorApi = readFileSync(join(root, "creator_api", "main.py"), "utf8");
const creatorSources = readFileSync(join(root, "creator_api", "sources.py"), "utf8");
const creatorPipeline = readFileSync(join(root, "creator_api", "pipeline.py"), "utf8");
const creatorTasks = readFileSync(join(root, "creator_api", "tasks.py"), "utf8");
const creatorCelery = readFileSync(join(root, "creator_api", "celery_app.py"), "utf8");
const creatorStudio = readFileSync(join(root, "creator-studio.tsx"), "utf8");
assert(
  schemaSql.includes('create policy "profile_self_read"') &&
    schemaSql.includes("id = auth.uid() or public.is_syland_admin()"),
  "Hồ sơ tài khoản được cô lập bằng RLS"
);
assert(
  schemaSql.includes('create policy "license_owner_read"') &&
    schemaSql.includes("auth.jwt() ->> 'email'"),
  "Bản quyền chỉ hiển thị cho chủ tài khoản hoặc quản trị viên"
);
assert(
  paymentSql.includes('create policy "payment_owner_read"') &&
    paymentSql.includes("user_id = auth.uid() or public.is_syland_admin()"),
  "Đơn thanh toán được cô lập theo tài khoản"
);
assert(
  paymentSql.includes("create or replace function public.admin_confirm_payment") &&
    paymentSql.includes("if not public.is_syland_admin()"),
  "Chỉ quản trị viên được xác nhận thanh toán"
);
assert(
  paymentSql.includes("new.license_code := v_code") &&
    paymentSql.includes("new.confirmed_by := auth.uid()"),
  "Xác nhận thanh toán tự động cấp và lưu mã bản quyền"
);
assert(paymentSql.includes("disable trigger user"), "Bản cài đặt thanh toán tạm dừng trigger khi nâng cấp dữ liệu");
assert(paymentSql.includes("enable trigger user"), "Bản cài đặt thanh toán bật lại trigger sau nâng cấp");
assert(paymentFix.includes("begin;") && paymentFix.includes("commit;"), "Bản vá thanh toán chạy trong transaction");
assert(!paymentFix.includes("update public.payment_orders"), "Bản vá gói không kích hoạt trigger của đơn cũ");
assert(paymentFix.includes("not valid"), "Ràng buộc mới không quét dữ liệu lịch sử");
assert(repairSql.includes("drop trigger if exists issue_paid_license_trigger"), "Bản phục hồi gỡ trigger cũ trước khi sửa");
assert(!repairSql.includes("update public.payment_orders\nset plan"), "Bản phục hồi không chuẩn hóa/cập nhật đơn lịch sử");
assert(repairSql.includes("create trigger issue_paid_license_trigger"), "Bản phục hồi tạo lại trigger cấp mã");
assert(repairSql.includes("admin_ready"), "Bản phục hồi tự kiểm tra quyền quản trị");
assert(
  stableAccountSql.includes("get_my_syland_entitlements") &&
    stableAccountSql.includes("link_license_to_auth_user_trigger"),
  "Website và Windows dùng cùng nguồn quyền theo auth.users.id"
);
assert(
  accountPortal.includes('remoteData("/rpc/get_my_syland_entitlements"') &&
    accountPortal.includes("serverEntitlements || resolveEntitlements"),
  "Website ưu tiên quyền chuẩn từ máy chủ"
);
assert(
  accountPortal.includes('await remoteAuth("/logout"') &&
    accountPortal.includes("localStorage.removeItem(REMOTE_SESSION_KEY)"),
  "Đăng xuất thu hồi phiên máy chủ và xóa phiên cục bộ"
);
assert(
  stableAccountSql.includes("'schema_version', 2") &&
    stableAccountSql.includes("'recommended_upgrade'"),
  "RPC quyền chuẩn có phiên bản chính sách và gợi ý nâng cấp"
);
assert(
  !accountPortal.includes("code_challenge: challenge") &&
    accountPortal.includes('sessionStorage.removeItem(OAUTH_VERIFIER_KEY)'),
  "Google OAuth dùng phiên do Supabase quản lý"
);
assert(
  accountPortal.includes('currentEntitlements.plan !== "Dùng thử"') &&
    accountPortal.includes('currentEntitlements.role === "admin" ? "Không giới hạn"'),
  "Gói quản trị hiển thị hoạt động và không giới hạn"
);
assert(
  creatorSql.includes("alter table public.creator_projects enable row level security") &&
    creatorSql.includes("user_id = auth.uid() or public.is_syland_admin()"),
  "Dự án Creator được cô lập theo tài khoản bằng RLS"
);
assert(
  creatorSql.includes("Tác vụ chỉ được tạo qua máy chủ Creator") &&
    !creatorSql.includes('create policy "creator_job_owner_insert"'),
  "Trình duyệt không thể tự tạo tác vụ Creator trong cơ sở dữ liệu"
);
assert(
  creatorSql.includes("creator_reserve_minutes") &&
    creatorSql.includes("pg_advisory_xact_lock") &&
    creatorSql.includes("CREATOR_QUOTA_EXCEEDED"),
  "Hạn mức Creator được giữ nguyên tử theo tài khoản"
);
assert(
  creatorSql.includes("creator_usage_summary") &&
    creatorSql.includes("'remainingMinutes'") &&
    creatorApi.includes('@app.get("/v1/video/usage"'),
  "API Creator trả gói và số phút còn lại từ máy chủ"
);
assert(
  creatorStudio.includes('apiRequest("/v1/video/usage")') &&
    creatorStudio.includes("phút còn lại") &&
    creatorStudio.includes("làm mới"),
  "Giao diện Creator hiển thị hạn mức phút và ngày làm mới"
);
assert(
  creatorSql.includes("revoke all on function public.creator_reserve_minutes") &&
    creatorSql.includes("grant execute on function public.creator_reserve_minutes") &&
    creatorSql.includes("to service_role"),
  "Chỉ máy chủ service_role được ghi hạn mức Creator"
);
assert(
  creatorApi.includes("Depends(authenticated_user)") &&
    creatorApi.includes("valid_download_signature") &&
    creatorApi.includes("get_owned_job"),
  "Creator API xác thực tài khoản và ký liên kết tải riêng"
);
assert(
  creatorSources.includes('"cookiefile": None') &&
    creatorSources.includes('"geo_bypass": False') &&
    creatorSources.includes('risk="blocked"'),
  "Quét nguồn Creator không dùng cookie hoặc vượt video riêng tư"
);
assert(
  creatorPipeline.includes("client.responses.create(") &&
    creatorPipeline.includes("store=False") &&
    !creatorPipeline.includes("VITE_OPENAI_API_KEY"),
  "Dịch OpenAI dùng Responses API và không lưu khóa ở client"
);
assert(
  creatorApi.includes('@app.get("/ready")') &&
    creatorApi.includes('"storage": False') &&
    creatorApi.includes('"database": False') &&
    creatorApi.includes('"queue": False'),
  "Creator API có readiness check cho lưu trữ, Supabase và Redis"
);
assert(
  creatorApi.includes('@app.post("/v1/video/jobs/{job_id}/cancel"') &&
    creatorTasks.includes("class JobCancelled") &&
    creatorTasks.includes("refund_minutes"),
  "Chủ tài khoản có thể hủy tác vụ và được hoàn hạn mức"
);
assert(
  creatorTasks.includes('@celery_app.task(name="creator.cleanup_expired_outputs")') &&
    creatorCelery.includes('"cleanup-expired-creator-outputs"') &&
    creatorCelery.includes('include=["creator_api.tasks"]'),
  "Worker đăng ký tác vụ và tự dọn MP4 hết thời hạn"
);

console.log(`SỸ LAND verification: ${pass.length} kiểm tra đạt.`);
if (failures.length) {
  console.error(failures.map((message) => `- ${message}`).join("\n"));
  process.exit(1);
}
