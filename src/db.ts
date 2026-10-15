import { Database } from "bun:sqlite";
import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";

export type Provider = "outlook";
export type MailboxStatus = "active" | "error";

export type MailboxRow = {
  id: number;
  email: string;
  provider: Provider;
  password: string | null;
  client_id: string;
  refresh_token: string;
  access_token: string | null;
  access_token_expires_at: number | null; // epoch ms
  status: MailboxStatus;
  last_error: string | null;
  created_at: number;
  updated_at: number;
};

export type PublicMailbox = {
  email: string;
  provider: Provider;
  clientId: string;
  hasPassword: boolean;
  hasRefreshToken: boolean;
  hasAccessToken: boolean;
  accessTokenExpiresAt: number | null;
  status: MailboxStatus;
  lastError: string | null;
  createdAt: number;
  updatedAt: number;
};

const DB_PATH = resolve(process.env.MAILMAN_DB ?? "data/mailman.sqlite");
mkdirSync(dirname(DB_PATH), { recursive: true });

export const db = new Database(DB_PATH, { create: true });
db.exec("PRAGMA journal_mode = WAL;");
db.exec(`
  CREATE TABLE IF NOT EXISTS mailboxes (
    id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    email                  TEXT    NOT NULL UNIQUE,
    provider               TEXT    NOT NULL DEFAULT 'outlook',
    password               TEXT,
    client_id              TEXT    NOT NULL,
    refresh_token          TEXT    NOT NULL,
    access_token           TEXT,
    access_token_expires_at INTEGER,
    status                 TEXT    NOT NULL DEFAULT 'active',
    last_error             TEXT,
    created_at             INTEGER NOT NULL,
    updated_at             INTEGER NOT NULL
  );
`);

const now = () => Date.now();

type Named = Record<string, string | number | null>;

const qInsert = db.query<unknown, [Named]>(
  `INSERT INTO mailboxes (email, provider, password, client_id, refresh_token, status, last_error, created_at, updated_at)
   VALUES ($email, $provider, $password, $client_id, $refresh_token, 'active', NULL, $now, $now)
   ON CONFLICT(email) DO UPDATE SET
     provider     = excluded.provider,
     password     = excluded.password,
     client_id    = excluded.client_id,
     refresh_token = excluded.refresh_token,
     access_token = NULL,
     access_token_expires_at = NULL,
     status       = 'active',
     last_error   = NULL,
     updated_at   = excluded.updated_at`,
);

const qSelectByEmail = db.query<MailboxRow, [Named]>(
  "SELECT * FROM mailboxes WHERE email = $email",
);
const qSelectAll = db.query<MailboxRow, []>(
  "SELECT * FROM mailboxes ORDER BY created_at DESC, id DESC",
);
const qDeleteByEmail = db.query<unknown, [string]>(
  "DELETE FROM mailboxes WHERE email = $email",
);
const qSaveTokens = db.query<unknown, [Named]>(
  `UPDATE mailboxes
     SET access_token = $access_token,
         access_token_expires_at = $expires_at,
         refresh_token = $refresh_token,
         status = 'active',
         last_error = NULL,
         updated_at = $now
   WHERE email = $email`,
);
const qSetError = db.query<unknown, [Named]>(
  `UPDATE mailboxes SET status = 'error', last_error = $last_error, updated_at = $now WHERE email = $email`,
);
const qCount = db.query<{ n: number }, []>("SELECT COUNT(*) AS n FROM mailboxes");

export function listMailboxes(): MailboxRow[] {
  return qSelectAll.all();
}

export function getMailbox(email: string): MailboxRow | null {
  return qSelectByEmail.get({ $email: email.toLowerCase() }) ?? null;
}

export function countMailboxes(): number {
  return qCount.get()?.n ?? 0;
}

/** 返回 "created" | "updated" */
export function upsertMailbox(input: {
  email: string;
  provider: Provider;
  password: string | null;
  clientId: string;
  refreshToken: string;
}): "created" | "updated" {
  const email = input.email.toLowerCase();
  const existed = getMailbox(email) !== null;
  qInsert.run({
    $email: email,
    $provider: input.provider,
    $password: input.password,
    $client_id: input.clientId,
    $refresh_token: input.refreshToken,
    $now: now(),
  });
  return existed ? "updated" : "created";
}

export function saveTokens(
  email: string,
  tokens: { accessToken: string; expiresAt: number; refreshToken: string },
): void {
  qSaveTokens.run({
    $email: email.toLowerCase(),
    $access_token: tokens.accessToken,
    $expires_at: tokens.expiresAt,
    $refresh_token: tokens.refreshToken,
    $now: now(),
  });
}

export function setMailboxError(email: string, message: string): void {
  qSetError.run({ $email: email.toLowerCase(), $last_error: message, $now: now() });
}

export function deleteMailbox(email: string): boolean {
  return qDeleteByEmail.run(email.toLowerCase()).changes > 0;
}

export function toPublicMailbox(row: MailboxRow): PublicMailbox {
  return {
    email: row.email,
    provider: row.provider,
    clientId: row.client_id,
    hasPassword: !!row.password,
    hasRefreshToken: !!row.refresh_token,
    hasAccessToken: !!row.access_token,
    accessTokenExpiresAt: row.access_token_expires_at,
    status: row.status,
    lastError: row.last_error,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}
