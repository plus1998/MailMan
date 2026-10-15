---
name: outlook-token-import
description: 用账号密码走浏览器 OAuth 流程，为 Outlook/Hotmail 邮箱取到可服务端刷新的 refresh token 并导入 Mail Man。当需要新增邮箱、邮箱 token 失效需要重取、或遇到 AADSTS90023 / AADSTS70043 / AADSTS9002313 时使用。
compatibility: bun 1.x；内置浏览器 MCP（eco_agent_browser）或 Playwright + Chrome；Mail Man 服务已启动
---

# Outlook 账号密码 → refresh token → 导入 Mail Man

**先读 `references/aadsts.md`**：里面是错误码 → 原因 → 下一步动作的对照表，遇到任何 `AADSTSxxxx` 先查它。

## 0. 最重要的一条：client_id 决定 token 能不能用

取 token 时用的 `client_id` 注册类型，决定这个 refresh token 能否被 Mail Man 服务端刷新：

| client_id | 类型 | 服务端刷新 |
| --- | --- | --- |
| `9199bf20-a13f-4107-85dc-02114787ef48`（outlook.live.com 网页自用） | SPA + 第一方 | ✗ `AADSTS90023`。加 `Origin: https://outlook.office.com` 后 SPA 那关过了，但会报「第一方应用不允许用户 consent」，**死路，别再试** |
| `9e5f94bc-e8a4-4e73-b8be-63364c29d753`（Mozilla Thunderbird） | Native 公共客户端 | ✓ 实测可用 |

**永远用 Thunderbird 的 client_id 走授权码 + PKCE 流程**，不要图省事直接抓 outlook.live.com 网页自己的 token。

## 1. 生成 PKCE 与 authorize URL

```bash
bun .agents/skills/outlook-token-import/scripts/pkce.ts --key <邮箱或别名>
```

verifier 存到 `.mailman-oauth/<key>.verifier.txt`（已 gitignore），并打印 authorize URL。

## 2. 在浏览器里走完登录

打开上一步打印的 URL。**优先用内置浏览器 MCP**（`mcp__eco_agent_browser__*`），它没有抓包接口，但能用 `eval` 读 `location.href` 和页面 DOM，够用。

内置浏览器 MCP 的坑（实测）：

- `agent_browser_open` 打在 **UI 焦点 tab**，而 `eval/snapshot/click` 只认**第一个 tab**。两者不一致时会打不中元素（`Element not found`）。
- 解法：让用户关掉其它 tab，只留 OAuth 那个；或先 `agent_browser_tab_switch` 再确认 `eval` 里 `location.href` 已是目标页再操作。
- `tab_switch` 的 `index` 参数有 off-by-one（`index:3` 会找 t4），用 `tab:"t3"`。

按页面状态依次处理（每步之后用 `eval` 读 `location.href` + 可见 input/button 再决定）：

| 页面 | 特征 | 动作 |
| --- | --- | --- |
| 账户输入 | `#i0116`（name=loginfmt） | `fill` 邮箱 → `click #idSIButton9` |
| 密码 | `#i0118` 可见 | `fill` 密码 → `click #idSIButton9` |
| 邮箱验证（陌生设备） | `#proofConfirmationText`，文案「我们会将验证码发送到 xx*****@…」 | 填**完整**恢复邮箱地址 → `click #primaryButton`（发送验证码）→ 见第 3 步取码 → 填码提交 |
| 保持登录 | `#acceptButton` / `#declineButton` | `click #acceptButton` |
| 授权 consent | 文案「是否允许此应用访问你的信息? / Mozilla Thunderbird / 读取你的邮件」 | 「接受」按钮**没有 id**，用 eval 按文本找并 click（见下） |

按文本点按钮：

```js
(()=>{const b=[...document.querySelectorAll('button,input[type=submit]')]
  .filter(x=>x.offsetParent!==null).find(x=>(x.value||x.textContent||'').trim()==='接受');
  if(!b) return 'not found'; b.click(); return 'clicked';})()
```

## 3. 验证码：从 Mail Man 自己读

微软把单用码发到恢复邮箱时，那个邮箱通常**已经在 Mail Man 里**（导入过）。直接调 API 读，不用问用户：

```bash
# 列出候选（发件人 account-security-noreply@accountprotection.microsoft.com）
curl -s "localhost:3000/api/messages?email=<恢复邮箱>&top=5"
# 取正文里的数字码
curl -s "localhost:3000/api/message?email=<恢复邮箱>&id=<消息id>"
```

正文里是 4–8 位数字码（`/api/message` 返回 `bodyContent`）。注意列表字段是 `receivedAt`（不是 `received`）。

## 4. 取 code 并换 token

consent 之后浏览器会跳到 `https://login.microsoftonline.com/common/oauth2/nativeclient?code=...&state=...`。用 `eval` 读 `location.href` 取 `code`（这个页面本身打不开，无所谓）。

```bash
pbpaste > /tmp/mm_code.txt          # 或把 code 写进文件，别放命令行参数（含 * ! $）
bun .agents/skills/outlook-token-import/scripts/exchange.ts \
  --code-file /tmp/mm_code.txt --key <第1步的key> \
  --email <目标邮箱> --import --verify
```

`exchange.ts` 会：换 token → 打印 scope/expires_in/RT 长度摘要 → `--import` 调 `/api/mailboxes/import`（format=json）→ `--verify` 拉 5 封邮件确认真的能取件。

## 5. 收尾（必做）

```bash
rm -f /tmp/mm_code.txt && rm -rf .mailman-oauth
```

- **绝不把 token 提交进仓库**：`data/` 已 gitignore，但 commit 前仍要扫一遍 `git diff | grep -i MsaArtifacts`。
- 微软会轮换 refresh token：导入后库里的是新值，浏览器 MSAL 缓存里那份可能随之失效（网页重新登录即可）。
- 目标账号会收到「New app(s) connected to your Microsoft account」安全通知，属正常。

## 6. 全自动兜底（Playwright）

内置浏览器 MCP 打不中元素时，用 Playwright 起一个有头 Chrome，脚本会拦截 `/token` 与 nativeclient 跳转，自动填邮箱/密码、自动从 Mail Man 读验证码、自动点 consent：

```bash
MAILMAN_PW='<密码>' python3 .agents/skills/outlook-token-import/scripts/flow.py \
  --email <目标邮箱> --code-mailbox <已导入的恢复邮箱> --import --verify
```

## 批量导入外部已有的 token（不登录，直接导）

如果用户直接给 `email----password----client_id----refresh_token` 行，先按下面判断，别直接往 `/api/mailboxes/import` 一丢就完事：

1. 失败报 `AADSTS70000` 且文案含 “scopes … unauthorized or expired” → **不是 token 无效，是刷新时传的 scope 不在已授予集合里**。本项目已改为刷新默认不传 scope。
2. 第 5 段形如 `rt.1.AAB…` 的不是 refresh token（拿去刷新报 `AADSTS9002313`），忽略是对的。
3. 导入后**必须逐个拉邮件验证**（`/api/messages?email=&top=2`），“导入 ok” 不等于 “能取件”。

```bash
bun -e 'const t=await Bun.file("<文件>").text(); /* POST /api/mailboxes/import */'
```

token/密码临时文件放 `/tmp` 并在导入后 `rm`，绝不进仓库。

## 不要做的事

- 不要为了「成功」而兜底：上游报错原样返回（这是本项目的既定原则，`hintFor()` 只在原始错误后追加一行建议）。
- 不要复用 SPA client 的 token；不要试图用 `Origin` 头绕过第一方 consent。
- 不要把密码写进命令行参数或提交到仓库（用 `MAILMAN_PW` 环境变量）。
