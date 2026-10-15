/** 导入文本解析：email----password----clientid----refreshtoken（多余字段忽略） */

export const DELIMITER = "----";

export type ParsedLineOk = {
  ok: true;
  lineNo: number;
  raw: string;
  email: string;
  password: string;
  clientId: string;
  refreshToken: string;
  ignored: string[];
};

export type ParsedLineErr = { ok: false; lineNo: number; raw: string; reason: string };
export type ParsedLine = ParsedLineOk | ParsedLineErr;

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function parseLine(line: string, lineNo: number): ParsedLine {
  const raw = line.trim();
  const parts = raw.split(DELIMITER).map((p) => p.trim());
  const [email, password, clientId, refreshToken] = parts;
  const ignored = parts.slice(4);

  if (parts.length < 4) {
    return {
      ok: false,
      lineNo,
      raw,
      reason: `字段不足：需要 4 段（email${DELIMITER}password${DELIMITER}clientid${DELIMITER}refreshtoken），实际 ${parts.length} 段`,
    };
  }
  if (!email) return { ok: false, lineNo, raw, reason: "email 为空" };
  if (!EMAIL_RE.test(email)) return { ok: false, lineNo, raw, reason: `email 格式非法：${email}` };
  if (!clientId) return { ok: false, lineNo, raw, reason: "clientid 为空" };
  if (!refreshToken) return { ok: false, lineNo, raw, reason: "refreshtoken 为空" };

  return {
    ok: true,
    lineNo,
    raw,
    email: email.toLowerCase(),
    password: password ?? "",
    clientId,
    refreshToken,
    ignored,
  };
}

/* ---------------- JSON 导入 ----------------
 * 兼容直接粘贴 OAuth token 响应（单个对象 / 对象数组 / 每行一个 JSON）。
 * email   <- client_info.email | client_info.preferred_username | id_token.email | 显式 email
 * client_id <- 显式 client_id | id_token.aud（client_info 里没有 client_id）
 * refresh_token <- refresh_token
 */

export type ImportFormat = "auto" | "lines" | "json";

function decodeB64Json(seg: string): Record<string, unknown> | null {
  if (!seg) return null;
  try {
    const value = JSON.parse(Buffer.from(seg.trim(), "base64url").toString("utf8"));
    return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;
  } catch {
    return null;
  }
}

function jwtPayload(idToken: string): Record<string, unknown> | null {
  const parts = idToken.split(".");
  if (parts.length < 2) return null;
  return decodeB64Json(parts[1]!);
}

function str(v: unknown): string {
  return typeof v === "string" ? v.trim() : typeof v === "number" ? String(v) : "";
}

function entryToLine(entry: unknown, lineNo: number, raw: string): ParsedLine {
  if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
    return { ok: false, lineNo, raw, reason: "不是 JSON 对象" };
  }
  const o = entry as Record<string, unknown>;
  const clientInfo = decodeB64Json(str(o.client_info));
  const idp = jwtPayload(str(o.id_token));

  const refreshToken = str(o.refresh_token);
  if (!refreshToken) return { ok: false, lineNo, raw, reason: "缺少 refresh_token" };

  const email =
    str(o.email) || str(clientInfo?.email) || str(clientInfo?.preferred_username) ||
    str(idp?.email) || str(idp?.preferred_username);
  if (!email) {
    return { ok: false, lineNo, raw, reason: "无法确定 email：client_info / id_token 里没有 email，也没有显式 email 字段" };
  }
  if (!EMAIL_RE.test(email)) return { ok: false, lineNo, raw, reason: `email 格式非法：${email}` };

  const clientId = str(o.client_id) || str(idp?.aud);
  if (!clientId) {
    return { ok: false, lineNo, raw, reason: "无法确定 client_id：JSON 里没有 client_id，id_token 里也没有 aud" };
  }

  return {
    ok: true,
    lineNo,
    raw,
    email: email.toLowerCase(),
    password: str(o.password) || str(o.account),
    clientId,
    refreshToken,
    ignored: [],
  };
}

export function parseJsonText(text: string): ParsedLine[] {
  const trimmed = text.trim();
  if (!trimmed) return [];

  // 1) 整体当作 JSON（单个对象或对象数组）
  try {
    const value: unknown = JSON.parse(trimmed);
    const entries = Array.isArray(value) ? value : [value];
    return entries.map((e, i) => entryToLine(e, i + 1, `json#${i + 1}`));
  } catch (e) {
    if (trimmed.startsWith("[")) {
      return [{ ok: false, lineNo: 1, raw: trimmed.slice(0, 60), reason: `JSON 解析失败：${(e as Error).message}` }];
    }
  }

  // 2) JSON Lines：每行一个 JSON 对象
  const lines = text.split(/\r?\n/);
  const out: ParsedLine[] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]!;
    const lineNo = i + 1;
    if (!line.trim() || line.trim().startsWith("#")) continue;
    let value: unknown;
    try {
      value = JSON.parse(line.trim());
    } catch (e) {
      out.push({ ok: false, lineNo, raw: line.trim().slice(0, 60), reason: `JSON 解析失败：${(e as Error).message}` });
      continue;
    }
    out.push(entryToLine(value, lineNo, `json#${lineNo}`));
  }
  return out;
}

export function detectFormat(text: string): Exclude<ImportFormat, "auto"> {
  const t = text.trim();
  if (t.startsWith("{") || t.startsWith("[")) return "json";
  const first = (t.split(/\r?\n/)[0] ?? "").trim();
  return first.startsWith("{") ? "json" : "lines";
}

export function parseImportInput(text: string, format: ImportFormat = "auto"): { parsed: ParsedLine[]; format: Exclude<ImportFormat, "auto"> } {
  const resolved = format === "auto" ? detectFormat(text) : format;
  return { parsed: resolved === "json" ? parseJsonText(text) : parseImportText(text), format: resolved };
}

export function parseImportText(text: string): ParsedLine[] {
  const out: ParsedLine[] = [];
  const lines = text.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]!;
    if (!line.trim() || line.trim().startsWith("#")) continue;
    out.push(parseLine(line, i + 1));
  }
  return out;
}
