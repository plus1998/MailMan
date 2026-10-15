export type PublicMailbox = {
  email: string;
  provider: string;
  clientId: string;
  hasPassword: boolean;
  hasRefreshToken: boolean;
  hasAccessToken: boolean;
  accessTokenExpiresAt: number | null;
  status: "active" | "error";
  lastError: string | null;
  createdAt: number;
  updatedAt: number;
};

export type MailAddress = { name: string; address: string };

export type MailSummary = {
  id: string;
  subject: string;
  from: MailAddress;
  receivedAt: string;
  bodyPreview: string;
  isRead: boolean;
};

export type MailDetail = MailSummary & { bodyContentType: string; bodyContent: string };

export type MessageListResult = {
  messages: MailSummary[];
  nextLink: string | null;
  folder: string;
  top: number;
  skip: number;
};

export type ImportItem = {
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
  items: ImportItem[];
};

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(path, init);
  const text = await res.text();
  let data: unknown = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error(`响应不是 JSON（HTTP ${res.status}）：${text.slice(0, 300)}`);
    }
  }
  if (!res.ok) {
    const msg = (data as { error?: string } | null)?.error ?? `HTTP ${res.status}`;
    throw new Error(msg);
  }
  return data as T;
}

export const api = {
  listMailboxes: () => request<{ mailboxes: PublicMailbox[] }>("/api/mailboxes"),

  importMailboxes: (text: string) =>
    request<ImportResult>("/api/mailboxes/import", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    }),

  refresh: (email: string) =>
    request<{ ok: true; email: string; expiresAt: number }>(`/api/mailboxes/refresh?email=${encodeURIComponent(email)}`, {
      method: "POST",
    }),

  remove: (email: string) =>
    request<{ ok: true; deleted: boolean }>(`/api/mailboxes?email=${encodeURIComponent(email)}`, { method: "DELETE" }),

  messages: (email: string, folder: string, top: number, skip: number) =>
    request<MessageListResult>(
      `/api/messages?email=${encodeURIComponent(email)}&folder=${encodeURIComponent(folder)}&top=${top}&skip=${skip}`,
    ),

  message: (email: string, id: string) =>
    request<MailDetail>(`/api/message?email=${encodeURIComponent(email)}&id=${encodeURIComponent(id)}`),
};
