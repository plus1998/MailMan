#!/usr/bin/env bun
/**
 * 用授权码换 token，并（可选）导入 Mail Man + 验证取件。
 * code 从文件读，不走命令行参数（里面含 * ! $）。
 *
 * 用法：
 *   bun scripts/exchange.ts --code-file /tmp/mm_code.txt --key <pkce key> \
 *       --email <目标邮箱> [--password <pw>] [--import] [--verify] [--api http://localhost:3000]
 */
const DEFAULT_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753";
const REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPE = "https://outlook.office.com/mail.read offline_access";
const DIR = `${import.meta.dir}/../../../..`.replace(/\/+$/, "") + "/.mailman-oauth";

function arg(name: string, fallback?: string): string | undefined {
  const i = process.argv.indexOf(`--${name}`);
  if (i >= 0 && process.argv[i + 1] && !process.argv[i + 1].startsWith("--")) return process.argv[i + 1];
  const eq = process.argv.find((a) => a.startsWith(`--${name}=`));
  if (eq) return eq.slice(name.length + 3);
  return process.argv.includes(`--${name}`) ? "true" : fallback;
}

const flag = (name: string) => process.argv.includes(`--${name}`);

const codeFile = arg("code-file");
const key = arg("key");
const email = arg("email");
const clientId = arg("client-id", DEFAULT_CLIENT_ID)!;
const api = arg("api", "http://localhost:3000")!;

if (!codeFile) {
  console.error("缺 --code-file <path>（文件里只放 code）");
  process.exit(2);
}
const code = (await Bun.file(codeFile).text()).trim();
if (!code) {
  console.error(`${codeFile} 是空的`);
  process.exit(2);
}
let verifier = "";
if (key) {
  const f = `${DIR}/${key}.verifier.txt`;
  if (Bun.file(f).size > 0) verifier = (await Bun.file(f).text()).trim();
  else console.error(`警告：${f} 不存在，PKCE 校验会失败（authorize 与换 token 必须是同一次）`);
}

const form = new URLSearchParams({
  client_id: clientId,
  grant_type: "authorization_code",
  code,
  redirect_uri: REDIRECT_URI,
  scope: SCOPE,
});
if (verifier) form.set("code_verifier", verifier);

const res = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
  method: "POST",
  headers: { "Content-Type": "application/x-www-form-urlencoded" },
  body: form,
});
const body = (await res.json()) as Record<string, string>;

if (!res.ok || body.error) {
  console.error(`换 token 失败 HTTP ${res.status}`);
  console.error(`${body.error}: ${body.error_description ?? "(无 error_description)"}`);
  console.error("→ 查 .agents/skills/outlook-token-import/references/aadsts.md");
  process.exit(1);
}

const rt = body.refresh_token ?? "";
console.log("换 token 成功");
console.log(`  scope      ${body.scope}`);
console.log(`  expires_in ${body.expires_in}s  refresh_token ${rt.length} 字符 ${rt.slice(0, 12)}…${rt.slice(-8)}`);
if (!rt) {
  console.error("响应里没有 refresh_token（scope 缺 offline_access？），无法长期保存");
  process.exit(1);
}

if (!flag("import")) {
  console.log("\n（未加 --import，只换不导入）");
} else {
  if (!email) {
    console.error("导入需要 --email <目标邮箱>（这个 client 不返回 client_info，无法自动解出 email）");
    process.exit(2);
  }
  const payload = {
    format: "json",
    text: JSON.stringify({
      refresh_token: rt,
      client_id: clientId,
      email,
      ...(process.env.MAILMAN_PW ? { password: process.env.MAILMAN_PW } : {}),
    }),
  };
  const r = await fetch(`${api}/api/mailboxes/import`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const d = (await r.json()) as { total: number; ok: number; failed: number; items: { status: string; email: string | null; action?: string; error?: string }[] };
  console.log(`导入：total=${d.total} ok=${d.ok} failed=${d.failed}`);
  for (const i of d.items) console.log(`  ${i.status} ${i.email ?? "-"} ${i.action ?? ""} ${i.error ?? ""}`.trimEnd());
  if (d.failed > 0) process.exit(1);
}

if (flag("verify") && email) {
  const r = await fetch(`${api}/api/messages?${new URLSearchParams({ email, top: "5" })}`);
  const d = (await r.json()) as { messages?: { receivedAt: string; subject: string | null; from: { address: string } }[]; error?: string };
  if (d.error || !d.messages) {
    console.error(`取件验证失败：${d.error}`);
    process.exit(1);
  }
  console.log(`取件验证：拿到 ${d.messages.length} 封`);
  for (const m of d.messages) console.log(`  ${m.receivedAt} | ${m.from?.address} | ${(m.subject ?? "").slice(0, 60)}`);
}

console.log("\n收尾：rm -f <code-file> && rm -rf .mailman-oauth");
