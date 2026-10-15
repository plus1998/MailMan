import { useEffect, useState } from "react";
import { Admin } from "./Admin";
import { Inbox } from "./Inbox";
import { api, type PublicMailbox } from "./api";

const NAV_EVENT = "mailman:navigate";

export function navigate(path: string) {
  history.pushState(null, "", path);
  window.dispatchEvent(new Event(NAV_EVENT));
}

function useLocation() {
  const [loc, setLoc] = useState({ path: location.pathname, search: location.search });
  useEffect(() => {
    const update = () => setLoc({ path: location.pathname, search: location.search });
    window.addEventListener("popstate", update);
    window.addEventListener(NAV_EVENT, update);
    return () => {
      window.removeEventListener("popstate", update);
      window.removeEventListener(NAV_EVENT, update);
    };
  }, []);
  return loc;
}

function Home() {
  const [boxes, setBoxes] = useState<PublicMailbox[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    api.listMailboxes().then((r) => setBoxes(r.mailboxes)).catch((e: Error) => setError(e.message));
  }, []);

  return (
    <section className="card">
      <h2>邮箱取件系统 · v1（仅 Outlook）</h2>
      <ul className="plain">
        <li>
          <a href="/admin" onClick={(e) => { e.preventDefault(); navigate("/admin"); }}>
            /admin
          </a>
          {" "}— 批量导入邮箱（<code>email----password----clientid----refreshtoken</code>）
        </li>
        <li>
          <a href="/inbox" onClick={(e) => { e.preventDefault(); navigate("/inbox"); }}>
            /inbox?email=xxx
          </a>
          {" "}— 查看指定邮箱的邮件列表
        </li>
      </ul>

      {error && <div className="banner error">{error}</div>}
      {boxes === null && !error && <div className="muted">加载中…</div>}
      {boxes !== null && boxes.length === 0 && <div className="muted">还没有导入任何邮箱，先去 /admin 导入。</div>}
      {boxes !== null && boxes.length > 0 && (
        <div className="table-wrap">
        <table className="table">
          <thead>
            <tr><th>邮箱</th><th>状态</th><th>操作</th></tr>
          </thead>
          <tbody>
            {boxes.map((b) => (
              <tr key={b.email}>
                <td>{b.email}</td>
                <td><span className={`tag ${b.status}`}>{b.status === "active" ? "正常" : "异常"}</span></td>
                <td>
                  <a href={`/inbox?email=${encodeURIComponent(b.email)}`}
                     onClick={(e) => { e.preventDefault(); navigate(`/inbox?email=${encodeURIComponent(b.email)}`); }}>
                    查看邮件
                  </a>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      )}
    </section>
  );
}

export function App() {
  const { path, search } = useLocation();

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <a href="/" onClick={(e) => { e.preventDefault(); navigate("/"); }}>Mail Man</a>
        </div>
        <nav>
          <a href="/admin" className={path === "/admin" ? "active" : ""}
             onClick={(e) => { e.preventDefault(); navigate("/admin"); }}>邮箱管理</a>
          <a href="/inbox" className={path === "/inbox" ? "active" : ""}
             onClick={(e) => { e.preventDefault(); navigate("/inbox"); }}>收件箱</a>
        </nav>
      </header>

      <main>
        {path === "/admin" ? <Admin /> : path === "/inbox" ? <Inbox search={search} /> : <Home />}
      </main>

      <footer className="foot">bun + react · Outlook REST API v2.0 · 无鉴权（仅限本地/内网使用）</footer>
    </div>
  );
}
