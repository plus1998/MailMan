import { Glob } from "bun";
import { handleApi } from "./src/api";
import { countMailboxes } from "./src/db";

const DEV = process.env.NODE_ENV !== "production";
const PORT = Number(process.env.PORT ?? 3000);

/* ---------------- 前端构建（内存产物 + 源码 mtime 缓存失效） ---------------- */

type Bundle = { js: string; css: string; stamp: number };
let bundle: Bundle | null = null;

async function clientStamp(): Promise<number> {
  const glob = new Glob("src/client/**/*.{ts,tsx,css}");
  let max = 0;
  for await (const file of glob.scan()) max = Math.max(max, Bun.file(file).lastModified);
  return max;
}

async function getBundle(): Promise<Bundle> {
  const stamp = DEV ? await clientStamp() : 0;
  if (bundle && (!DEV || bundle.stamp === stamp)) return bundle;

  const result = await Bun.build({
    entrypoints: ["src/client/main.tsx"],
    target: "browser",
    minify: !DEV,
  });
  if (!result.success) {
    throw new Error(`前端构建失败：\n${result.logs.map((l) => JSON.stringify(l)).join("\n")}`);
  }

  let js = "";
  let css = "";
  for (const out of result.outputs) {
    const text = await out.text();
    if (out.path.endsWith(".css")) css += text;
    else js += text;
  }
  bundle = { js, css, stamp };
  return bundle;
}

const textResponse = (body: string, type: string) =>
  new Response(body, { headers: { "Content-Type": type, "Cache-Control": DEV ? "no-store" : "public, max-age=3600" } });

function shell(title: string): string {
  return `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title === SITE ? SITE : `${title} · ${SITE}`}</title>
<link rel="stylesheet" href="/assets/app.css">
</head>
<body>
<div id="root"></div>
<script src="/assets/app.js"></script>
</body>
</html>`;
}

/* ---------------- 路由 ---------------- */

const PAGES: Record<string, string> = {
  "/": "Mail Man",
  "/admin": "邮箱管理",
  "/inbox": "收件箱",
};
const SITE = "Mail Man";

async function fetchHandler(req: Request): Promise<Response> {
  const url = new URL(req.url);

  if (url.pathname.startsWith("/api/")) return handleApi(req, url);

  if (url.pathname === "/assets/app.js") return textResponse((await getBundle()).js, "text/javascript; charset=utf-8");
  if (url.pathname === "/assets/app.css") return textResponse((await getBundle()).css, "text/css; charset=utf-8");

  if (PAGES[url.pathname] !== undefined && req.method === "GET") {
    await getBundle(); // 构建失败时直接抛错，不返回半截页面
    return textResponse(shell(PAGES[url.pathname]!), "text/html; charset=utf-8");
  }

  if (url.pathname === "/favicon.ico") return new Response(null, { status: 204 });

  return new Response(`404 Not Found: ${url.pathname}`, {
    status: 404,
    headers: { "Content-Type": "text/plain; charset=utf-8" },
  });
}

const server = Bun.serve({
  port: PORT,
  fetch: fetchHandler,
  error(err) {
    console.error("[error]", err);
    return Response.json({ error: `服务端异常：${(err as Error).message}` }, { status: 500 });
  },
});

// 启动即构建一次，构建错误直接暴露
await getBundle();

console.log(`Mail Man 已启动: ${server.url} (mode=${DEV ? "dev" : "production"})`);
console.log(`  管理页  ${server.url}admin`);
console.log(`  收件箱  ${server.url}inbox?email=<邮箱>`);
console.log(`  邮箱数  ${countMailboxes()}`);
