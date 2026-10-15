import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { OutlookError, getMessage, listMessages, refreshAccessToken } from "./outlook";

const realFetch = globalThis.fetch;
let calls: { url: string; init: RequestInit }[] = [];

const mockFetch = (handler: (url: string, init: RequestInit) => Response | Promise<Response>) => {
  globalThis.fetch = ((url: string | URL, init?: RequestInit) =>
    Promise.resolve(handler(String(url), init ?? {}))) as unknown as typeof fetch;
};

beforeEach(() => {
  calls = [];
  mockFetch((url, init) => {
    calls.push({ url, init });
    throw new Error("test 未 mock fetch");
  });
});

afterEach(() => {
  globalThis.fetch = realFetch;
});

const jsonResponse = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });

const TOKEN_OK = {
  token_type: "Bearer",
  scope: "https://outlook.office.com/Mail.Read",
  expires_in: 3599,
  ext_expires_in: 3599,
  access_token: "FAKE-ACCESS-TOKEN-000000",
  refresh_token: "FAKE-REFRESH-TOKEN-ROTATED-000000",
};

describe("refreshAccessToken", () => {
  test("成功：解析 access token、过期时间、轮换后的 refresh token", async () => {
    mockFetch(() => jsonResponse(TOKEN_OK));
    const before = Date.now();
    const t = await refreshAccessToken({ clientId: "cid", refreshToken: "rt" });
    expect(t.accessToken).toBe(TOKEN_OK.access_token);
    expect(t.refreshToken).toBe(TOKEN_OK.refresh_token);
    expect(t.expiresIn).toBe(3599);
    expect(t.expiresAt).toBeGreaterThanOrEqual(before + 3599 * 1000);
  });

  test("请求参数与 curl 一致", async () => {
    mockFetch((url, init) => {
      calls.push({ url: String(url), init: init ?? {} });
      return jsonResponse(TOKEN_OK);
    });
    await refreshAccessToken({ clientId: "9e5f94bc-e8a4-4e73-b8be-63364c29d753", refreshToken: "FAKE-RT-1" });
    const c = calls[0]!;
    expect(c.url).toBe("https://login.microsoftonline.com/common/oauth2/v2.0/token");
    const form = new URLSearchParams(String(c.init.body));
    expect(form.get("grant_type")).toBe("refresh_token");
    expect(form.get("client_id")).toBe("9e5f94bc-e8a4-4e73-b8be-63364c29d753");
    expect(form.get("refresh_token")).toBe("FAKE-RT-1");
    expect(form.get("redirect_uri")).toBe("https://login.microsoftonline.com/common/oauth2/token");
    expect(form.get("scope")).toBe("https://outlook.office.com/mail.read offline_access");
  });

  test("失败：抛出微软的 error + error_description，不兜底", async () => {
    mockFetch(() =>
      jsonResponse(
        { error: "invalid_grant", error_description: "AADSTS70043: The refresh token has expired.", error_codes: [70043] },
        400,
      ),
    );
    let err: unknown;
    try {
      await refreshAccessToken({ clientId: "cid", refreshToken: "rt" });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(OutlookError);
    const oe = err as OutlookError;
    expect(oe.code).toBe("invalid_grant");
    expect(oe.message).toContain("AADSTS70043");
    expect(oe.status).toBe(400);
  });

  test("失败：响应缺少 refresh_token 时报错（scope 漏了 offline_access）", async () => {
    mockFetch(() => jsonResponse({ access_token: "x", expires_in: 3599 }));
    let msg = "";
    try {
      await refreshAccessToken({ clientId: "cid", refreshToken: "rt" });
    } catch (e) {
      msg = (e as Error).message;
    }
    expect(msg).toContain("refresh_token");
  });
});

const LIST_OK = {
  "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Me/MailFolders('Inbox')/Messages(Subject,From,ReceivedDateTime,BodyPreview)",
  value: [
    {
      Id: "AQMkADAwATNiZmYAZS1iNzU1AC04MGJjLTAwAi0wMAoA1",
      ReceivedDateTime: "2026-09-30T15:35:07Z",
      Subject: "Get real help, not just answers",
      BodyPreview: "Study, explore, and grow with ChatGPT.",
      From: { EmailAddress: { Name: "ChatGPT", Address: "noreply@email.openai.com" } },
    },
    {
      Id: "AQMkADAwATNiZmYAZS1iNzU1AC04MGJjLTAwAi0wMAoA2",
      ReceivedDateTime: "2026-09-20T17:17:02Z",
      Subject: "你的临时 ChatGPT 登录代码",
      BodyPreview: "输入此临时验证码以继续：\r\n\r\n302505",
      From: { EmailAddress: { Name: "ChatGPT", Address: "noreply@tm.openai.com" } },
    },
  ],
  "@odata.nextLink": "https://outlook.office.com/api/v2.0/me/mailFolders/Inbox/messages?$skip=10",
};

describe("listMessages", () => {
  test("归一化字段 + nextLink", async () => {
    mockFetch(() => jsonResponse(LIST_OK));
    const r = await listMessages("token", { top: 10, skip: 20 });
    expect(r.messages.length).toBe(2);
    expect(r.messages[0]).toEqual({
      id: "AQMkADAwATNiZmYAZS1iNzU1AC04MGJjLTAwAi0wMAoA1",
      subject: "Get real help, not just answers",
      from: { name: "ChatGPT", address: "noreply@email.openai.com" },
      receivedAt: "2026-09-30T15:35:07Z",
      bodyPreview: "Study, explore, and grow with ChatGPT.",
      isRead: false,
    });
    expect(r.messages[1]!.subject).toBe("你的临时 ChatGPT 登录代码");
    expect(r.nextLink).toContain("$skip=10");
  });

  test("URL：folder / $select / $top / $skip / $orderby", async () => {
    mockFetch((url) => {
      calls.push({ url: String(url), init: {} });
      return jsonResponse(LIST_OK);
    });
    await listMessages("token", { folder: "Junk", top: 50, skip: 30 });
    const u = new URL(calls[0]!.url);
    expect(u.pathname).toBe("/api/v2.0/me/mailFolders/Junk/messages");
    expect(u.searchParams.get("$top")).toBe("50");
    expect(u.searchParams.get("$skip")).toBe("30");
    expect(u.searchParams.get("$orderby")).toBe("receivedDateTime desc");
    expect(u.searchParams.get("$select")).toContain("bodyPreview");
  });

  test("401 时抛出上游错误", async () => {
    mockFetch(() => jsonResponse({ "odata.error": { message: { value: "Invalid security token." } }, error: "invalid_token" }, 401));
    let err: unknown;
    try {
      await listMessages("bad", {});
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(OutlookError);
    expect((err as OutlookError).status).toBe(401);
  });
});

describe("getMessage", () => {
  test("取回 HTML 正文，id 做 URL 编码", async () => {
    mockFetch((url) => {
      calls.push({ url: String(url), init: {} });
      return jsonResponse({
        Id: "AQMk ADw/x+=",
        Subject: "你的临时 ChatGPT 登录代码",
        ReceivedDateTime: "2026-09-20T17:17:02Z",
        From: { EmailAddress: { Name: "ChatGPT", Address: "noreply@tm.openai.com" } },
        Body: { ContentType: "HTML", Content: "<html><body>302505</body></html>" },
      });
    });
    const d = await getMessage("token", "AQMk ADw/x+=");
    expect(d.bodyContentType).toBe("HTML");
    expect(d.bodyContent).toContain("302505");
    expect(d.subject).toBe("你的临时 ChatGPT 登录代码");
    expect(calls[0]!.url).toContain("/me/messages/AQMk%20ADw%2Fx%2B%3D?");
  });
});
