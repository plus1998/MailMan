# AADSTS 错误码对照表

Mail Man 不做兜底：`error` + `error_description` 原样返回。`src/outlook.ts` 的 `hintFor()` 只在后面追加一行建议。

| 错误码 | 含义 | 下一步 |
| --- | --- | --- |
| `AADSTS90023`（含 "Single-Page Application"） | client_id 是 SPA 类型，token 只能在浏览器里跨域兑换 | 换 native client（Thunderbird `9e5f94bc-e8a4-4e73-b8be-63364c29d753`）重新取 token |
| `AADSTS90023`（含 "first party application"） | 加了 `Origin` 头后触发：第一方应用不允许用户 consent | 同上，**不要**继续加 header 绕 |
| `AADSTS9002313` | 请求 malformed / RT 与 client_id 不匹配 / RT 已失效 | 确认导入的 client_id 与取 token 时是同一个；重新取 token。注意它包含 `AADSTS90023` 前缀，代码里必须先判它 |
| `AADSTS70043` | refresh token 已过期（MS 约 90 天，SPA/第一方会话更短） | 重新登录取 token |
| `AADSTS70000` | refresh token 解析失败 / 已被轮换 | 重新取 token。库里存的 RT 可能已被上一次刷新换掉 |
| `AADSTS7000218` | client_id 不存在或未在租户注册 | 检查 client_id 拼写 |
| `AADSTS50076` / `AADSTS50079` | 需要 MFA / 必须同意 | 走一次交互式登录（浏览器流程） |
| `AADSTS65001` | 用户未同意该应用 | 在 authorize 请求里去掉 `prompt=none`，手动点「接受」 |
| `invalid_grant`（无 AADSTS） | 通常是 redirect_uri 与取 code 时不一致 | 换 token 的 `redirect_uri` 必须与 authorize 完全一致（nativeclient） |

## 已实测的两条硬结论

1. `9199bf20-a13f-4107-85dc-02114787ef48`（outlook.live.com 自用，SPA + 第一方）：不带 `Origin` 报 SPA 限制；带 `Origin: https://outlook.office.com` 报第一方 consent 限制。**服务端不可用。**
2. `9e5f94bc-e8a4-4e73-b8be-63364c29d753`（Mozilla Thunderbird，native 公共客户端）：consent 页显示「Mozilla Thunderbird / 读取你的邮件」，返回的 `scope` 是全套 `Mail.Read / Mail.ReadWrite / EAS / EWS / POP / IMAP / SMTP.Send`（列表里没有 `offline_access`），但仍下发 refresh token 且能服务端刷新。
