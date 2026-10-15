import { countMailboxes, deleteMailbox, listMailboxes, toPublicMailbox } from "./db";
import { importMailboxes } from "./import";
import { MailboxNotFoundError, fetchMessage, fetchMessages, forceRefresh } from "./mail-service";
import { OutlookError } from "./outlook";

const json = (data: unknown, status = 200) =>
  Response.json(data, { status, headers: { "Cache-Control": "no-store" } });

function errorToResponse(e: unknown, ctx: string): Response {
  if (e instanceof MailboxNotFoundError) return json({ error: e.message }, 404);
  if (e instanceof OutlookError) {
    const status = e.status === 401 || e.status === 403 ? 401 : e.status >= 400 && e.status < 500 ? 502 : 502;
    return json({ error: `${ctx}失败：${e.message}`, code: e.code, upstreamStatus: e.status }, status);
  }
  return json({ error: `${ctx}失败：${(e as Error).message}` }, 500);
}

function badRequest(message: string): Response {
  return json({ error: message }, 400);
}

function intParam(url: URL, name: string, fallback: number, min: number, max: number): number {
  const raw = url.searchParams.get(name);
  if (raw === null || raw === "") return fallback;
  const n = Number(raw);
  if (!Number.isFinite(n)) throw new Error(`参数 ${name} 不是数字：${raw}`);
  return Math.min(Math.max(Math.trunc(n), min), max);
}

function requireEmail(url: URL, body?: Record<string, unknown> | null): string {
  const fromQuery = url.searchParams.get("email");
  const fromBody = typeof body?.email === "string" ? body.email : null;
  const email = (fromQuery || fromBody || "").trim();
  if (!email) throw new Error("缺少参数 email");
  return email.toLowerCase();
}

export async function handleApi(req: Request, url: URL): Promise<Response> {
  const route = `${req.method} ${url.pathname}`;

  try {
    if (route === "GET /api/health") {
      return json({ ok: true, mailboxes: countMailboxes() });
    }

    if (route === "GET /api/mailboxes") {
      return json({ mailboxes: listMailboxes().map(toPublicMailbox) });
    }

    if (route === "POST /api/mailboxes/import") {
      const body = (await req.json().catch(() => null)) as { text?: unknown; concurrency?: unknown } | null;
      const text = typeof body?.text === "string" ? body.text : null;
      if (!text || !text.trim()) return badRequest("请求体缺少 text（多行 email----password----clientid----refreshtoken）");
      const concurrency = typeof body?.concurrency === "number" ? body.concurrency : undefined;
      const result = await importMailboxes(text, concurrency);
      return json(result);
    }

    if (route === "POST /api/mailboxes/refresh") {
      const body = (await req.json().catch(() => ({}))) as Record<string, unknown>;
      const email = requireEmail(url, body);
      const { expiresAt } = await forceRefresh(email);
      return json({ ok: true, email, expiresAt });
    }

    if (route === "DELETE /api/mailboxes") {
      const email = requireEmail(url);
      const deleted = deleteMailbox(email);
      if (!deleted) return json({ error: `邮箱不存在：${email}` }, 404);
      return json({ ok: true, email, deleted: true });
    }

    if (route === "GET /api/messages") {
      const email = requireEmail(url);
      const folder = url.searchParams.get("folder")?.trim() || "Inbox";
      const top = intParam(url, "top", 20, 1, 100);
      const skip = intParam(url, "skip", 0, 0, 100_000);
      const result = await fetchMessages(email, { folder, top, skip });
      return json(result);
    }

    if (route === "GET /api/message") {
      const email = requireEmail(url);
      const id = url.searchParams.get("id")?.trim();
      if (!id) return badRequest("缺少参数 id");
      return json(await fetchMessage(email, id));
    }

    return json({ error: `未知接口：${route}` }, 404);
  } catch (e) {
    return errorToResponse(e, req.method === "GET" ? "读取" : "操作");
  }
}
