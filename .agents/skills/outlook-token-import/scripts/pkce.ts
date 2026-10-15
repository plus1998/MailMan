#!/usr/bin/env bun
/**
 * 生成 PKCE 并打印 authorize URL。
 * verifier 存到 .mailman-oauth/<key>.verifier.txt（gitignore），换 token 时读同一个文件。
 *
 * 用法：bun scripts/pkce.ts --key aaronfox [--client-id <id>] [--scope <scope>] [--login-hint <email>]
 */
import { mkdirSync } from "node:fs";

const DEFAULT_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"; // Mozilla Thunderbird（native 公共客户端）
const REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPE = "https://outlook.office.com/mail.read offline_access";
const DIR = `${import.meta.dir}/../../../..`.replace(/\/+$/, "") + "/.mailman-oauth";

function arg(name: string, fallback?: string): string | undefined {
  const i = process.argv.indexOf(`--${name}`);
  if (i >= 0 && process.argv[i + 1]) return process.argv[i + 1];
  const eq = process.argv.find((a) => a.startsWith(`--${name}=`));
  return eq ? eq.slice(name.length + 3) : fallback;
}

const key = arg("key") ?? new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const clientId = arg("client-id", DEFAULT_CLIENT_ID)!;
const scope = arg("scope", SCOPE)!;
const loginHint = arg("login-hint");

const verifier = Buffer.from(crypto.getRandomValues(new Uint8Array(48)))
  .toString("base64url");
const challenge = Buffer.from(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)))
  .toString("base64url");

mkdirSync(DIR, { recursive: true });
await Bun.write(`${DIR}/${key}.verifier.txt`, verifier);
await Bun.write(`${DIR}/${key}.challenge.txt`, challenge);

const url = new URL("https://login.microsoftonline.com/common/oauth2/v2.0/authorize");
url.searchParams.set("client_id", clientId);
url.searchParams.set("response_type", "code");
url.searchParams.set("redirect_uri", REDIRECT_URI);
url.searchParams.set("scope", scope);
url.searchParams.set("state", `mailman-${key}`);
url.searchParams.set("client_info", "1");
url.searchParams.set("code_challenge", challenge);
url.searchParams.set("code_challenge_method", "S256");
if (loginHint) url.searchParams.set("login_hint", loginHint);

console.log(`key:        ${key}`);
console.log(`client_id:  ${clientId}`);
console.log(`verifier:   ${DIR}/${key}.verifier.txt`);
console.log(`\n在浏览器打开：\n${url.toString()}\n`);
console.log("登录后从跳转地址取 code，再跑 scripts/exchange.ts");
