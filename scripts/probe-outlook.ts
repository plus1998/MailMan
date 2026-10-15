/**
 * 命令行探针：不走数据库，直接验证 refresh token 与邮件接口。
 *   bun run scripts/probe-outlook.ts <client_id> <refresh_token> [message_id]
 */
import { getMessage, listMessages, refreshAccessToken } from "../src/outlook";

const [clientId, refreshToken, messageId] = process.argv.slice(2);
if (!clientId || !refreshToken) {
  console.error("用法: bun run scripts/probe-outlook.ts <client_id> <refresh_token> [message_id]");
  process.exit(1);
}

const tokens = await refreshAccessToken({ clientId, refreshToken });
console.log("access_token:", tokens.accessToken.slice(0, 40) + "…");
console.log("scope:", tokens.scope);
console.log("expires_in:", tokens.expiresIn);
console.log("新 refresh_token（已轮换）:", tokens.refreshToken.slice(0, 40) + "…");

const list = await listMessages(tokens.accessToken, { top: 5 });
console.log(`\nInbox 前 ${list.messages.length} 封:`);
for (const m of list.messages) {
  console.log(`- [${m.receivedAt}] ${m.subject} <${m.from.address}>`);
}

if (messageId) {
  const detail = await getMessage(tokens.accessToken, messageId);
  console.log(`\n邮件 ${messageId}`);
  console.log("subject:", detail.subject);
  console.log("body:", detail.bodyContentType, `${detail.bodyContent.length} chars`);
}
