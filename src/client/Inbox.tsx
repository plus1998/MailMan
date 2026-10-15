import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { api, type MailDetail, type MailSummary, type PublicMailbox } from "./api";
import { navigate } from "./App";

const FOLDERS = [
  { value: "Inbox", label: "收件箱" },
  { value: "Junk", label: "垃圾邮件" },
  { value: "SentItems", label: "已发送" },
  { value: "Drafts", label: "草稿" },
];

function fmtDate(iso: string): string {
  if (!iso) return "—";
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? iso : d.toLocaleString();
}

/** 邮件正文：不执行脚本，只允许链接开新窗口 */
function MailBody({ detail }: { detail: MailDetail }) {
  const srcDoc = useMemo(
    () =>
      `<!doctype html><html><head><meta charset="utf-8"><base target="_blank">` +
      `<style>body{margin:0;padding:16px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;font-size:14px;color:#111;word-break:break-word}img{max-width:100%;height:auto}a{color:#2563eb}</style>` +
      `</head><body>${detail.bodyContent}</body></html>`,
    [detail.bodyContent],
  );
  if (detail.bodyContentType.toUpperCase() !== "HTML") {
    return <pre className="plain-text">{detail.bodyContent}</pre>;
  }
  return <iframe className="mail-frame" sandbox="allow-popups allow-popups-to-escape-sandbox" srcDoc={srcDoc} title="邮件正文" />;
}

function Picker({ onPick }: { onPick: (email: string) => void }) {
  const [boxes, setBoxes] = useState<PublicMailbox[]>([]);
  const [value, setValue] = useState("");
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    api.listMailboxes().then((r) => setBoxes(r.mailboxes)).catch((e: Error) => setError(e.message));
  }, []);

  return (
    <section className="card">
      <h2>缺少 email 参数</h2>
      <p className="muted">
        用法：<code>/inbox?email=xxx@yyy.com</code>
      </p>
      {error && <div className="banner error">{error}</div>}
      <div className="row">
        <input
          className="input"
          placeholder="输入已导入的邮箱地址"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && value.trim()) onPick(value.trim().toLowerCase());
          }}
        />
        <button className="btn primary" onClick={() => value.trim() && onPick(value.trim().toLowerCase())}>查看</button>
      </div>
      {boxes.length > 0 && (
        <div className="table-wrap">
        <table className="table">
          <thead><tr><th>已导入邮箱</th><th>状态</th></tr></thead>
          <tbody>
            {boxes.map((b) => (
              <tr key={b.email}>
                <td>
                  <a href={`/inbox?email=${encodeURIComponent(b.email)}`}
                     onClick={(e) => { e.preventDefault(); onPick(b.email); }}>
                    {b.email}
                  </a>
                </td>
                <td><span className={`tag ${b.status}`}>{b.status === "active" ? "正常" : "异常"}</span></td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      )}
    </section>
  );
}

export function Inbox({ search }: { search: string }) {
  const email = useMemo(() => (new URLSearchParams(search).get("email") ?? "").trim().toLowerCase(), [search]);
  const [folder, setFolder] = useState("Inbox");
  const [top, setTop] = useState(20);
  const [messages, setMessages] = useState<MailSummary[]>([]);
  const [skip, setSkip] = useState(0);
  const [hasMore, setHasMore] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [detail, setDetail] = useState<MailDetail | null>(null);
  const [detailLoading, setDetailLoading] = useState(false);
  const [detailError, setDetailError] = useState<string | null>(null);
  const listSeq = useRef(0);

  const load = useCallback(
    async (nextSkip: number, append: boolean) => {
      if (!email) return;
      const seq = ++listSeq.current;
      setLoading(true);
      setError(null);
      try {
        const r = await api.messages(email, folder, top, nextSkip);
        if (seq !== listSeq.current) return; // 有更新的请求已发出，丢弃旧结果
        setMessages((prev) => (append ? [...prev, ...r.messages] : r.messages));
        setSkip(r.skip + r.messages.length);
        setHasMore(r.nextLink !== null && r.messages.length > 0);
      } catch (e) {
        if (seq === listSeq.current) setError((e as Error).message);
      } finally {
        if (seq === listSeq.current) setLoading(false);
      }
    },
    [email, folder, top],
  );

  useEffect(() => {
    setMessages([]);
    setSkip(0);
    setDetail(null);
    load(0, false);
  }, [load]);

  const openMessage = async (m: MailSummary) => {
    setDetailLoading(true);
    setDetailError(null);
    try {
      setDetail(await api.message(email, m.id));
    } catch (e) {
      setDetailError((e as Error).message);
    } finally {
      setDetailLoading(false);
    }
  };

  const refreshToken = async () => {
    setLoading(true);
    setError(null);
    try {
      await api.refresh(email);
      await load(0, false);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const detailOpen = detail !== null || detailLoading || detailError !== null;

  // 窄屏弹窗：Esc 关闭 + 锁页面滚动
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") { setDetail(null); setDetailError(null); }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  useEffect(() => {
    document.body.classList.toggle("modal-open", detailOpen);
    return () => document.body.classList.remove("modal-open");
  }, [detailOpen]);

  const closeDetail = () => { setDetail(null); setDetailError(null); };

  if (!email) return <Picker onPick={(e) => navigate(`/inbox?email=${encodeURIComponent(e)}`)} />;

  return (
    <div className="inbox">
      <section className="card">
        <div className="row between">
          <h2 className="nowrap">{email}</h2>
          <div className="row">
            <label className="muted">
              文件夹
              <select className="input" value={folder} onChange={(e) => setFolder(e.target.value)}>
                {FOLDERS.map((f) => <option key={f.value} value={f.value}>{f.label}</option>)}
              </select>
            </label>
            <label className="muted">
              每页
              <select className="input" value={top} onChange={(e) => setTop(Number(e.target.value))}>
                {[10, 20, 50, 100].map((n) => <option key={n} value={n}>{n}</option>)}
              </select>
            </label>
            <button className="btn" disabled={loading} onClick={() => load(0, false)}>{loading ? "加载中…" : "刷新列表"}</button>
            <button className="btn" disabled={loading} onClick={refreshToken}>重新取 token</button>
          </div>
        </div>
        {error && <div className="banner error">{error}</div>}
        {!error && !loading && messages.length === 0 && <div className="muted">没有邮件</div>}
      </section>

      <div className="inbox-body">
        <section className="card pane list-pane">
          <div className="table-wrap">
          <table className="table mail-list">
            <thead>
              <tr><th>主题</th><th>发件人</th><th>时间</th></tr>
            </thead>
            <tbody>
              {messages.map((m) => (
                <tr key={m.id} className={detail?.id === m.id ? "selected" : ""} onClick={() => openMessage(m)}>
                  <td className="wrap">
                    <div className={m.isRead ? "subject" : "subject unread"}>{m.subject}</div>
                    <div className="muted preview">{m.bodyPreview.slice(0, 160)}</div>
                  </td>
                  <td className="nowrap">
                    {m.from.name || m.from.address}
                    <div className="muted small">{m.from.address}</div>
                  </td>
                  <td className="nowrap">{fmtDate(m.receivedAt)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
          <div className="row">
            <button className="btn" disabled={loading || !hasMore} onClick={() => load(skip, true)}>
              {hasMore ? "加载更多" : "没有更多了"}
            </button>
            <span className="muted">已加载 {messages.length} 封</span>
          </div>
        </section>

        <section className={`card pane detail-pane${detailOpen ? " open" : ""}`}>
          <div className="detail-bar">
            <h3>邮件正文</h3>
            <button className="close" onClick={closeDetail} aria-label="关闭">×</button>
          </div>
          {detailLoading && <div className="muted">加载中…</div>}
          {detailError && <div className="banner error">{detailError}</div>}
          {!detail && !detailLoading && !detailError && <div className="muted">点击邮件行查看正文</div>}
          {detail && (
            <>
              <div className="detail-head">
                <div className="subject">{detail.subject}</div>
                <div className="muted small">
                  发件人：{detail.from.name} &lt;{detail.from.address}&gt; · {fmtDate(detail.receivedAt)} ·{" "}
                  {detail.bodyContentType.toUpperCase()}
                </div>
              </div>
              <MailBody detail={detail} />
            </>
          )}
        </section>
      </div>
    </div>
  );
}
