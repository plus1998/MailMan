/**
 * Outlook REST API v2.0 + Microsoft identity v2.0 token endpoint.
 * 错误一律抛出 OutlookError（带 MS 返回的 error / error_description），不做兜底。
 */

const TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
const API_BASE = "https://outlook.office.com/api/v2.0";

export const DEFAULT_SCOPE = "https://outlook.office.com/mail.read offline_access";
export const DEFAULT_REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/token";

export class OutlookError extends Error {
  readonly status: number;
  readonly code?: string;
  readonly detail?: string;

  constructor(message: string, opts: { status?: number; code?: string; detail?: string } = {}) {
    super(message);
    this.name = "OutlookError";
    this.status = opts.status ?? 0;
    this.code = opts.code;
    this.detail = opts.detail;
  }
}

export type TokenSet = {
  accessToken: string;
  /** epoch ms */
  expiresAt: number;
  /** MS 会轮换 refresh token，必须回写 */
  refreshToken: string;
  scope: string;
  expiresIn: number;
};

function describeMsError(status: number, body: unknown): OutlookError {
  const b = (body ?? {}) as {
    error?: string;
    error_description?: string;
    error_codes?: number[];
    timestamp?: string;
    message?: string;
  };
  const code = b.error ?? (b.error_codes?.length ? String(b.error_codes[0]) : undefined);
  const detail = b.error_description ?? b.message;
  const msg = detail ?? `HTTP ${status}`;
  return new OutlookError(
    code ? `${code}: ${msg}` : msg,
    { status, code, detail },
  );
}

export async function refreshAccessToken(input: {
  clientId: string;
  refreshToken: string;
  scope?: string;
  redirectUri?: string;
}): Promise<TokenSet> {
  const form = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: input.refreshToken,
    client_id: input.clientId,
    redirect_uri: input.redirectUri ?? DEFAULT_REDIRECT_URI,
    scope: input.scope ?? DEFAULT_SCOPE,
  });

  let res: Response;
  try {
    res = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
      body: form.toString(),
    });
  } catch (e) {
    throw new OutlookError(`无法连接 token 端点：${(e as Error).message}`);
  }

  const text = await res.text();
  let body: unknown = null;
  try {
    body = text ? JSON.parse(text) : null;
  } catch {
    /* 非 JSON 响应，保留原文 */
  }

  if (!res.ok) {
    if (body) throw describeMsError(res.status, body);
    throw new OutlookError(`HTTP ${res.status}: ${text.slice(0, 300)}`, { status: res.status });
  }

  const data = body as {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
    scope?: string;
  } | null;

  if (!data?.access_token) {
    throw new OutlookError(`token 响应缺少 access_token：${text.slice(0, 300)}`, { status: res.status });
  }
  if (!data.refresh_token) {
    throw new OutlookError(`token 响应缺少 refresh_token（scope 未包含 offline_access？）：${text.slice(0, 300)}`, {
      status: res.status,
    });
  }

  const expiresIn = Number(data.expires_in ?? 3599);
  return {
    accessToken: data.access_token,
    expiresAt: Date.now() + expiresIn * 1000,
    refreshToken: data.refresh_token,
    scope: data.scope ?? "",
    expiresIn,
  };
}

async function apiGet<T>(accessToken: string, url: string): Promise<T> {
  let res: Response;
  try {
    res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" },
    });
  } catch (e) {
    throw new OutlookError(`无法连接 Outlook API：${(e as Error).message}`);
  }

  const text = await res.text();
  let body: unknown = null;
  try {
    body = text ? JSON.parse(text) : null;
  } catch {
    /* ignore */
  }

  if (!res.ok) {
    if (body) throw describeMsError(res.status, body);
    throw new OutlookError(`HTTP ${res.status}: ${text.slice(0, 300)}`, { status: res.status });
  }
  if (!body) throw new OutlookError(`Outlook API 返回空响应（HTTP ${res.status}）`, { status: res.status });
  return body as T;
}

/* ---------- 响应归一化 ---------- */

type RawAddress = { Name?: string; Address?: string };
type RawFrom = { EmailAddress?: RawAddress };
type RawMessage = {
  Id?: string;
  ReceivedDateTime?: string;
  Subject?: string;
  BodyPreview?: string;
  IsRead?: boolean;
  From?: RawFrom;
  Body?: { ContentType?: string; Content?: string };
};

export type MailAddress = { name: string; address: string };
export type MailSummary = {
  id: string;
  subject: string;
  from: MailAddress;
  receivedAt: string; // ISO
  bodyPreview: string;
  isRead: boolean;
};
export type MailDetail = MailSummary & {
  bodyContentType: string;
  bodyContent: string;
};

function addr(raw: RawFrom | undefined): MailAddress {
  const a = raw?.EmailAddress;
  return { name: a?.Name ?? "", address: a?.Address ?? "" };
}

function summary(raw: RawMessage): MailSummary {
  if (!raw.Id) throw new OutlookError("邮件响应缺少 Id");
  return {
    id: raw.Id,
    subject: raw.Subject ?? "(无主题)",
    from: addr(raw.From),
    receivedAt: raw.ReceivedDateTime ?? "",
    bodyPreview: raw.BodyPreview ?? "",
    isRead: raw.IsRead ?? false,
  };
}

export async function listMessages(
  accessToken: string,
  opts: { folder?: string; top?: number; skip?: number } = {},
): Promise<{ messages: MailSummary[]; nextLink: string | null; total: number | null }> {
  const folder = encodeURIComponent(opts.folder ?? "Inbox");
  const params = new URLSearchParams({
    $select: "id,subject,from,receivedDateTime,bodyPreview,isRead",
    $top: String(Math.min(Math.max(opts.top ?? 20, 1), 100)),
    $orderby: "receivedDateTime desc",
  });
  if (opts.skip && opts.skip > 0) params.set("$skip", String(opts.skip));

  const url = `${API_BASE}/me/mailFolders/${folder}/messages?${params.toString()}`;
  const body = await apiGet<{ value?: RawMessage[]; "@odata.nextLink"?: string; "@odata.count"?: number }>(
    accessToken,
    url,
  );
  const value = body.value;
  if (!Array.isArray(value)) throw new OutlookError("邮件列表响应缺少 value 数组");
  return {
    messages: value.map(summary),
    nextLink: body["@odata.nextLink"] ?? null,
    total: body["@odata.count"] ?? null,
  };
}

export async function getMessage(accessToken: string, id: string): Promise<MailDetail> {
  const params = new URLSearchParams({ $select: "id,subject,from,receivedDateTime,bodyPreview,isRead,body" });
  const url = `${API_BASE}/me/messages/${encodeURIComponent(id)}?${params.toString()}`;
  const raw = await apiGet<RawMessage>(accessToken, url);
  return {
    ...summary(raw),
    bodyContentType: raw.Body?.ContentType ?? "text",
    bodyContent: raw.Body?.Content ?? "",
  };
}
