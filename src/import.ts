import { saveTokens, setMailboxError, upsertMailbox } from "./db";
import { OutlookError, refreshAccessToken } from "./outlook";
import { parseImportInput, type ImportFormat, type ParsedLineOk } from "./parse";

export { DELIMITER, detectFormat, parseImportInput, parseImportText, parseJsonText, parseLine } from "./parse";
export type { ImportFormat, ParsedLine, ParsedLineErr, ParsedLineOk } from "./parse";

export type ImportItemResult = {
  lineNo: number;
  email: string | null;
  status: "ok" | "error";
  action?: "created" | "updated";
  ignoredParams?: number;
  error?: string;
  errorCode?: string;
};

export type ImportResult = {
  total: number;
  ok: number;
  failed: number;
  ignoredParams: number;
  format: "lines" | "json";
  items: ImportItemResult[];
};

const DEFAULT_CONCURRENCY = 4;

/**
 * 批量导入：逐行刷新 access token。
 * 刷新失败 -> 该行报错并跳过（不写入邮箱，已有记录则标记为 error）。
 * 刷新成功 -> 保存 access token + 轮换后的新 refresh token。
 */
export async function importMailboxes(text: string, concurrency = DEFAULT_CONCURRENCY, format: ImportFormat = "auto"): Promise<ImportResult> {
  const { parsed, format: resolvedFormat } = parseImportInput(text, format);
  const items: ImportItemResult[] = [];
  const tasks: ParsedLineOk[] = [];
  /** email -> items 下标 */
  const slot = new Map<string, number>();
  let ignoredParams = 0;

  for (const p of parsed) {
    if (!p.ok) {
      items.push({ lineNo: p.lineNo, email: null, status: "error", error: p.reason });
      continue;
    }
    const prev = slot.get(p.email);
    if (prev !== undefined) {
      items[prev] = {
        lineNo: items[prev]!.lineNo,
        email: p.email,
        status: "error",
        error: `同一批次内 email 重复，已改用第 ${p.lineNo} 行的数据`,
      };
    }
    slot.set(p.email, items.length);
    items.push({ lineNo: p.lineNo, email: p.email, status: "error", error: "pending" });
    tasks.push(p);
    ignoredParams += p.ignored.length;
  }

  let cursor = 0;
  const worker = async () => {
    for (;;) {
      const idx = cursor++;
      if (idx >= tasks.length) return;
      const t = tasks[idx]!;
      const out = items[slot.get(t.email)!]!;
      try {
        const tokens = await refreshAccessToken({ clientId: t.clientId, refreshToken: t.refreshToken });
        const action = upsertMailbox({
          email: t.email,
          provider: "outlook",
          password: t.password,
          clientId: t.clientId,
          refreshToken: tokens.refreshToken,
        });
        saveTokens(t.email, {
          accessToken: tokens.accessToken,
          expiresAt: tokens.expiresAt,
          refreshToken: tokens.refreshToken,
        });
        out.status = "ok";
        out.action = action;
        out.ignoredParams = t.ignored.length;
        delete out.error;
      } catch (e) {
        const isOutlook = e instanceof OutlookError;
        const message = isOutlook ? e.message : `未知错误：${(e as Error).message}`;
        // 拿不到 token：报错跳过，不新增记录；若之前已入库则标记异常
        setMailboxError(t.email, message);
        out.status = "error";
        out.error = message;
        out.errorCode = isOutlook ? e.code : undefined;
      }
    }
  };

  const pool = Math.max(1, Math.min(concurrency, 16));
  await Promise.all(Array.from({ length: Math.min(pool, tasks.length || 1) }, worker));

  return {
    total: parsed.length,
    ok: items.filter((i) => i.status === "ok").length,
    failed: items.filter((i) => i.status === "error").length,
    ignoredParams,
    format: resolvedFormat,
    items,
  };
}
