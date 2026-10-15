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
