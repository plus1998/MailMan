import { getMailbox, saveTokens, setMailboxError, type MailboxRow } from "./db";
import {
  OutlookError,
  getMessage as msGetMessage,
  listMessages as msListMessages,
  refreshAccessToken,
  type MailDetail,
  type MailSummary,
} from "./outlook";

export class MailboxNotFoundError extends Error {
  constructor(email: string) {
    super(`邮箱未导入：${email}`);
    this.name = "MailboxNotFoundError";
  }
}

/** access token 提前 5 分钟视为过期，避免边界 401 */
const REFRESH_BUFFER_MS = 5 * 60 * 1000;

/** 同一邮箱并发请求时只刷新一次（refresh token 会被 MS 轮换，重复刷新会互相打废） */
const inflightRefresh = new Map<string, Promise<string>>();

export function requireMailbox(email: string): MailboxRow {
  const row = getMailbox(email);
  if (!row) throw new MailboxNotFoundError(email);
  return row;
}

async function doRefresh(row: MailboxRow): Promise<string> {
  try {
    const tokens = await refreshAccessToken({ clientId: row.client_id, refreshToken: row.refresh_token });
    saveTokens(row.email, {
      accessToken: tokens.accessToken,
      expiresAt: tokens.expiresAt,
      refreshToken: tokens.refreshToken,
    });
    return tokens.accessToken;
  } catch (e) {
    const message = e instanceof OutlookError ? e.message : `未知错误：${(e as Error).message}`;
    setMailboxError(row.email, message);
    throw e;
  }
}

export async function getAccessToken(email: string): Promise<string> {
  const row = requireMailbox(email);
  if (row.access_token && row.access_token_expires_at && row.access_token_expires_at - Date.now() > REFRESH_BUFFER_MS) {
    return row.access_token;
  }
  const pending = inflightRefresh.get(row.email);
  if (pending) return pending;
  const p = doRefresh(row).finally(() => inflightRefresh.delete(row.email));
  inflightRefresh.set(row.email, p);
  return p;
}

export async function forceRefresh(email: string): Promise<{ expiresAt: number }> {
  const row = requireMailbox(email);
  const pending = inflightRefresh.get(row.email);
  if (pending) {
    await pending;
    const after = getMailbox(row.email);
    return { expiresAt: after?.access_token_expires_at ?? Date.now() };
  }
  const p = doRefresh(row).finally(() => inflightRefresh.delete(row.email));
  inflightRefresh.set(row.email, p);
  await p;
  return { expiresAt: getMailbox(row.email)?.access_token_expires_at ?? Date.now() };
}

export type MessageListResult = {
  messages: MailSummary[];
  nextLink: string | null;
  folder: string;
  top: number;
  skip: number;
};

export async function fetchMessages(
  email: string,
  opts: { folder?: string; top?: number; skip?: number } = {},
): Promise<MessageListResult> {
  const token = await getAccessToken(email);
  const res = await msListMessages(token, opts);
  return {
    messages: res.messages,
    nextLink: res.nextLink,
    folder: opts.folder ?? "Inbox",
    top: opts.top ?? 20,
    skip: opts.skip ?? 0,
  };
}

export async function fetchMessage(email: string, id: string): Promise<MailDetail> {
  const token = await getAccessToken(email);
  return msGetMessage(token, id);
}
