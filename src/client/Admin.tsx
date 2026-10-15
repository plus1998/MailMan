import { useCallback, useEffect, useState } from "react";
import { api, type ImportResult, type PublicMailbox } from "./api";
import { navigate } from "./App";

const PLACEHOLDER = `support@example.com----P@ssw0rd----9e5f94bc-e8a4-4e73-b8be-63364c29d753----M.C507_BAY.0.U.MsaArtifacts.-xxxx
// 以 # 开头的行和空行会被忽略；每行只取前 4 段，多余参数忽略`;

function fmtTime(ms: number | null): string {
  if (!ms) return "—";
  const d = new Date(ms);
  const diff = ms - Date.now();
  const suffix = diff > 0 ? `（${Math.round(diff / 60000)} 分钟后过期）` : "（已过期）";
  return `${d.toLocaleString()} ${suffix}`;
}

export function Admin() {
  const [text, setText] = useState("");
  const [boxes, setBoxes] = useState<PublicMailbox[]>([]);
  const [result, setResult] = useState<ImportResult | null>(null);
  const [importing, setImporting] = useState(false);
  const [busy, setBusy] = useState<Record<string, string>>({});
  const [error, setError] = useState<string | null>(null);

  const reload = useCallback(async () => {
    const r = await api.listMailboxes();
    setBoxes(r.mailboxes);
  }, []);

  useEffect(() => {
    reload().catch((e: Error) => setError(e.message));
  }, [reload]);

  const doImport = async () => {
    if (!text.trim()) {
      setError("请先粘贴要导入的邮箱");
      return;
    }
    setImporting(true);
    setError(null);
    try {
      const r = await api.importMailboxes(text);
      setResult(r);
      await reload();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setImporting(false);
    }
  };

  const doRefresh = async (email: string) => {
    setBusy((b) => ({ ...b, [email]: "刷新中…" }));
    setError(null);
    try {
      await api.refresh(email);
      await reload();
    } catch (e) {
      setError(`${email}: ${(e as Error).message}`);
      await reload();
    } finally {
      setBusy((b) => {
        const next = { ...b };
        delete next[email];
        return next;
      });
    }
  };

  const doDelete = async (email: string) => {
    if (!confirm(`删除 ${email}？`)) return;
    setBusy((b) => ({ ...b, [email]: "删除中…" }));
    try {
      await api.remove(email);
      await reload();
    } catch (e) {
      setError(`${email}: ${(e as Error).message}`);
    } finally {
      setBusy((b) => {
        const next = { ...b };
        delete next[email];
        return next;
      });
    }
  };

  return (
    <>
      <section className="card">
        <h2>批量导入邮箱</h2>
        <p className="muted">
          每行一条：<code>email----password----clientid----refreshtoken</code>。导入时会用 refresh token 换取 access
          token，换取失败该行报错并跳过（不入库）；成功则保存 access token 与轮换后的新 refresh token。
        </p>
        <textarea
          className="ta"
          value={text}
          spellCheck={false}
          placeholder={PLACEHOLDER}
          onChange={(e) => setText(e.target.value)}
        />
        <div className="row">
          <button className="btn primary" disabled={importing} onClick={doImport}>
            {importing ? "导入中…" : "导入"}
          </button>
          <button className="btn" disabled={importing} onClick={() => { setText(""); setResult(null); setError(null); }}>
            清空
          </button>
          <span className="muted">共 {text.split(/\r?\n/).filter((l) => l.trim()).length} 行</span>
        </div>

        {error && <div className="banner error">{error}</div>}

        {result && (
          <>
            <div className="banner">
              总计 {result.total} 行 · 成功 <b className="ok">{result.ok}</b> · 失败{" "}
              <b className="bad">{result.failed}</b> · 忽略多余参数 {result.ignoredParams} 个
            </div>
            <div className="table-wrap">
            <table className="table">
              <thead>
                <tr><th>行</th><th>email</th><th>结果</th><th>说明</th></tr>
              </thead>
              <tbody>
                {result.items.map((it) => (
                  <tr key={it.lineNo}>
                    <td>{it.lineNo}</td>
                    <td>{it.email ?? "—"}</td>
                    <td>
                      <span className={`tag ${it.status === "ok" ? "active" : "error"}`}>
                        {it.status === "ok" ? (it.action === "updated" ? "已更新" : "已导入") : "失败"}
                      </span>
                    </td>
                    <td className="wrap">
                      {it.status === "ok"
                        ? [it.ignoredParams ? `忽略多余参数 ${it.ignoredParams} 个` : null].filter(Boolean).join(" ") || "token 有效"
                        : `${it.errorCode ? `[${it.errorCode}] ` : ""}${it.error ?? ""}`}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            </div>
          </>
        )}
      </section>

      <section className="card">
        <h2>已导入邮箱（{boxes.length}）</h2>
        {boxes.length === 0 && <div className="muted">暂无邮箱</div>}
        {boxes.length > 0 && (
          <div className="table-wrap">
          <table className="table mailbox-table">
            <thead>
              <tr>
                <th>邮箱</th><th>client_id</th><th>access token</th><th>状态 / 最后错误</th><th>操作</th>
              </tr>
            </thead>
            <tbody>
              {boxes.map((b) => (
                <tr key={b.email} className={b.lastError ? "has-error" : ""}>
                  <td data-label="邮箱">{b.email}</td>
                  <td data-label="client_id" className="meta mono">{b.clientId}</td>
                  <td data-label="access token" className="meta">{b.hasAccessToken ? fmtTime(b.accessTokenExpiresAt) : "未获取"}</td>
                  <td data-label="状态" className="wrap">
                    <div className="cell">
                      <span className={`tag ${b.status}`}>{b.status === "active" ? "正常" : "异常"}</span>
                      {b.lastError && <div className="err-text">{b.lastError}</div>}
                    </div>
                  </td>
                  <td data-label="操作" className="acts">
                    {busy[b.email] ? (
                      <span className="muted">{busy[b.email]}</span>
                    ) : (
                      <>
                        <a href={`/inbox?email=${encodeURIComponent(b.email)}`}
                           onClick={(e) => { e.preventDefault(); navigate(`/inbox?email=${encodeURIComponent(b.email)}`); }}>
                          查看邮件
                        </a>
                        {" · "}
                        <a href="#" onClick={(e) => { e.preventDefault(); doRefresh(b.email); }}>刷新 token</a>
                        {" · "}
                        <a href="#" className="bad" onClick={(e) => { e.preventDefault(); doDelete(b.email); }}>删除</a>
                      </>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
        )}
      </section>
    </>
  );
}
