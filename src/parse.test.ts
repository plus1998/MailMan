import { describe, expect, test } from "bun:test";
import { DELIMITER, parseImportText, parseLine } from "./parse";

const OK = `a@b.com----pwd----cid----rt`;

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
