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
for (const id of ["minh-hoa", "cong-cu-pdf", "tai-phan-mem", "tai-khoan", "thanh-toan", "phap-ly"]) {
  assert(page.includes(`id="${id}"`), `Có điểm điều hướng #${id}`);
}
assert(page.includes("lazy(() => import(\"./file-processor\"))"), "Xử lý tệp được tải theo nhu cầu");
assert(page.includes("lazy(() => import(\"./pdf-toolkit\"))"), "Công cụ PDF được tải theo nhu cầu");
assert(!source.includes("@import \"tailwindcss\""), "Không nạp Tailwind không sử dụng");

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
for (const sql of [paymentSql, paymentFix]) {
  assert(sql.includes("disable trigger user"), "Bản vá thanh toán tạm dừng trigger an toàn");
  assert(sql.includes("enable trigger user"), "Bản vá thanh toán bật lại trigger");
}
assert(paymentFix.includes("begin;") && paymentFix.includes("commit;"), "Bản vá thanh toán chạy trong transaction");

console.log(`SỸ LAND verification: ${pass.length} kiểm tra đạt.`);
if (failures.length) {
  console.error(failures.map((message) => `- ${message}`).join("\n"));
  process.exit(1);
}
