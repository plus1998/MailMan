import { describe, expect, test } from "bun:test";
import { DELIMITER, detectFormat, parseImportInput, parseImportText, parseJsonText, parseLine } from "./parse";

const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString("base64url");
const CI = b64({ email: "Reader@Example.com", preferred_username: "reader@example.com", name: "Test User" });
const IDT = `${b64({ alg: "none" })}.${b64({ aud: "cid-from-aud", email: "reader@example.com" })}.sig`;

const OK = `a@b.com----pwd----cid----rt`;

describe("parseJsonText", () => {
  test("单个对象：email 从 client_info 解出，client_id 取 id_token.aud", () => {
    const rows = parseJsonText(JSON.stringify({ refresh_token: "RT-1", client_info: CI, id_token: IDT }));
    expect(rows.length).toBe(1);
    const r = rows[0]!;
    if (!r.ok) throw new Error(`应解析成功：${r.reason}`);
    expect(r.email).toBe("reader@example.com");
    expect(r.clientId).toBe("cid-from-aud");
    expect(r.refreshToken).toBe("RT-1");
  });

  test("对象数组批量，序号从 1 开始", () => {
    const rows = parseJsonText(JSON.stringify([
      { refresh_token: "RT-1", client_info: CI, id_token: IDT },
      { refresh_token: "RT-2", client_info: CI, id_token: IDT },
    ]));
    expect(rows.length).toBe(2);
    expect(rows.every((r) => r.ok)).toBe(true);
    expect(rows.map((r) => r.lineNo)).toEqual([1, 2]);
  });

  test("显式 client_id / email 优先于 client_info / id_token", () => {
    const r = parseJsonText(JSON.stringify({ refresh_token: "RT", client_id: "cid-explicit", email: "Explicit@Example.com", client_info: CI, id_token: IDT }))[0]!;
    if (!r.ok) throw new Error(`应解析成功：${r.reason}`);
    expect(r.clientId).toBe("cid-explicit");
    expect(r.email).toBe("explicit@example.com");
  });

  test("缺 refresh_token / 解不出 email / 解不出 client_id 分别报错", () => {
    const noRt = parseJsonText(JSON.stringify({ client_info: CI }))[0]!;
    if (noRt.ok) throw new Error("缺 refresh_token 应失败");
    expect(noRt.reason).toContain("refresh_token");

    const noEmail = parseJsonText(JSON.stringify({ refresh_token: "RT", client_id: "cid", id_token: `${b64({ alg: "none" })}.${b64({ aud: "cid-from-aud" })}.sig` }))[0]!;
    if (noEmail.ok) throw new Error("解不出 email 应失败");
    expect(noEmail.reason).toContain("无法确定 email");

    const noCid = parseJsonText(JSON.stringify({ refresh_token: "RT", client_info: CI }))[0]!;
    if (noCid.ok) throw new Error("解不出 client_id 应失败");
    expect(noCid.reason).toContain("无法确定 client_id");
  });

  test("JSON Lines：每行一个对象，行号保留", () => {
    const one = JSON.stringify({ refresh_token: "RT-1", client_info: CI, id_token: IDT });
    const rows = parseJsonText(`# 注释\n\n${one}\nnot-json`);
    expect(rows.length).toBe(2);
    expect(rows[0]!.ok).toBe(true);
    expect(rows[0]!.lineNo).toBe(3);
    if (rows[1]!.ok) throw new Error("非法 JSON 行应失败");
    expect(rows[1]!.reason).toContain("JSON 解析失败");
  });

  test("数组残缺时不抛异常，返回错误条目", () => {
    const rows = parseJsonText(`[{"refresh_token":"RT"`);
    expect(rows.length).toBe(1);
    if (rows[0]!.ok) throw new Error("残缺数组应返回错误条目");
    expect(rows[0]!.reason).toContain("JSON 解析失败");
  });

  test("client_info 不是合法 base64 时降级到 id_token", () => {
    const r = parseJsonText(JSON.stringify({ refresh_token: "RT", client_info: "@@@", id_token: IDT }))[0]!;
    if (!r.ok) throw new Error(`应解析成功：${r.reason}`);
    expect(r.email).toBe("reader@example.com");
  });
});

describe("detectFormat / parseImportInput", () => {
  test("以 { 或 [ 开头识别为 json，否则 lines", () => {
    expect(detectFormat('{"refresh_token":"RT"}')).toBe("json");
    expect(detectFormat(`\n  [{"a":1}]`)).toBe("json");
    expect(detectFormat(OK)).toBe("lines");
    expect(detectFormat(`# 注释\n${OK}`)).toBe("lines");
  });

  test("auto 自动分流，显式 format 强制按指定格式", () => {
    expect(parseImportInput(OK, "auto").format).toBe("lines");
    expect(parseImportInput(JSON.stringify({ refresh_token: "RT", client_info: CI, id_token: IDT }, null, 2), "auto").format).toBe("json");
    const forced = parseImportInput(JSON.stringify({ refresh_token: "RT", client_info: CI, id_token: IDT }), "lines");
    expect(forced.format).toBe("lines");
    expect(forced.parsed[0]!.ok).toBe(false);
  });
});

describe("parseLine", () => {
  test("标准 4 段", () => {
    const r = parseLine(OK, 1);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.email).toBe("a@b.com");
    expect(r.password).toBe("pwd");
    expect(r.clientId).toBe("cid");
    expect(r.refreshToken).toBe("rt");
    expect(r.ignored).toEqual([]);
  });

  test("超出 4 段只取前 4 段，其余忽略", () => {
    const r = parseLine(`${OK}----extra1----extra2`, 1);
    if (!r.ok) throw new Error("应解析成功");
    expect(r.clientId).toBe("cid");
    expect(r.refreshToken).toBe("rt");
    expect(r.ignored).toEqual(["extra1", "extra2"]);
  });

  test("少于 4 段报错", () => {
    const r = parseLine(`a@b.com${DELIMITER}pwd${DELIMITER}cid`, 3);
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.lineNo).toBe(3);
    expect(r.reason).toContain("字段不足");
  });

  test("email 非法 / 关键字段为空", () => {
    expect(parseLine(`notanemail----p----c----r`, 1).ok).toBe(false);
    expect(parseLine(`a@b.com----p----${DELIMITER}r`, 1).ok).toBe(false);
    expect(parseLine(`a@b.com----p----c----`, 1).ok).toBe(false);
  });

  test("email 统一小写，两端空白去掉", () => {
    const r = parseLine(`  A@B.COM ---- pwd ---- cid ---- rt `, 1);
    if (!r.ok) throw new Error("应解析成功");
    expect(r.email).toBe("a@b.com");
    expect(r.password).toBe("pwd");
  });

  test("refresh token 中含 ---- 以外的特殊字符不受影响", () => {
    const rt = "M.C507_FAKE.0.U.MsaArtifacts.-Cj2q!0YBir!gLkSre*pJkBK7in!x";
    const r = parseLine(`a@b.com----p----cid----${rt}`, 1);
    if (!r.ok) throw new Error("应解析成功");
    expect(r.refreshToken).toBe(rt);
  });
});

describe("parseImportText", () => {
  test("跳过空行与 # 注释行，行号保留原始行号", () => {
    const text = `# 注释\n\n${OK}\n   \n${OK}----x`;
    const rows = parseImportText(text);
    expect(rows.length).toBe(2);
    expect(rows[0]!.lineNo).toBe(3);
    expect(rows[1]!.lineNo).toBe(5);
  });
});
