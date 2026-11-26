#!/usr/bin/env python3
"""
全自动兜底：Playwright 起有头 Chrome 走完 OAuth 授权码流程，抓 code，再调 exchange.ts。
内置浏览器 MCP 打不中元素时才用这个。

  MAILMAN_PW='<密码>' python3 flow.py --email <目标邮箱> [--code-mailbox <已导入的恢复邮箱>] [--import] [--verify]

密码只从环境变量 MAILMAN_PW 读，不走命令行参数。
"""
import argparse
import base64
import hashlib
import json
import os
import re
import secrets
import subprocess
import sys
import time
import urllib.parse
import urllib.request

from playwright.sync_api import sync_playwright

CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"  # Mozilla Thunderbird（native）
REDIRECT = "https://login.microsoftonline.com/common/oauth2/nativeclient"
SCOPE = "https://graph.microsoft.com/Mail.Read offline_access"
AUTH_BASE = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE_DIR = os.path.abspath(os.path.join(SKILL_DIR, "..", "..", "..", ".mailman-oauth"))


def log(*a):
    print("[flow]", *a, flush=True)


def api_get(url):
    with urllib.request.urlopen(url, timeout=30) as r:
        return json.load(r)


def fetch_ms_code(api, mailbox, since_ts):
    """从已导入的恢复邮箱里取微软单用验证码（>= since_ts 的最新一封）。"""
    q = urllib.parse.urlencode({"email": mailbox, "top": 15})
    d = api_get(f"{api}/api/messages?{q}")
    if "messages" not in d:
        raise RuntimeError(f"读邮箱失败: {d}")
    cands = [m for m in d["messages"]
             if "accountprotection.microsoft.com" in (m.get("from") or {}).get("address", "")
             and m.get("receivedAt", "") >= since_ts]
    if not cands:
        return None
    m = sorted(cands, key=lambda x: x["receivedAt"])[-1]
    body = api_get(f"{api}/api/message?{urllib.parse.urlencode({'email': mailbox, 'id': m['id']})}")
    text = (body.get("bodyContent") or "") + " " + (body.get("bodyPreview") or "")
    codes = re.findall(r"\b(\d{4,8})\b", re.sub(r"<[^>]+>", " ", text))
    return codes[0] if codes else None


def visible(page, sel):
    try:
        el = page.query_selector(sel)
        return el if el and el.is_visible() else None
    except Exception:
        return None


def click_by_text(page, label):
    try:
        return page.evaluate(
            """(label) => {
              const bs = [...document.querySelectorAll('button,input[type=submit],a')]
                .filter(x => x.offsetParent !== null)
                .filter(x => (x.value || x.textContent || '').trim() === label);
              if (!bs.length) return false;
              bs[0].click();
              return true;
            }""",
            label,
        )
    except Exception as e:
        # Consent clicks navigate to nativeclient; Playwright can lose the old
        # execution context before the next loop reads the redirect URL.
        if "execution context was destroyed" in str(e).lower() or "navigation" in str(e).lower():
            return True
        raise


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--email", required=True, help="目标邮箱（要取 token 的账号）")
    p.add_argument("--code-mailbox", help="已导入 Mail Man 的恢复邮箱，用于自动读验证码")
    p.add_argument("--api", default="http://localhost:3000")
    p.add_argument("--client-id", default=CLIENT_ID)
    p.add_argument("--scope", default=SCOPE)
    p.add_argument("--import", action="store_true", dest="do_import")
    p.add_argument("--verify", action="store_true")
    p.add_argument("--timeout", type=int, default=240)
    a = p.parse_args()

    pw = os.environ.get("MAILMAN_PW")
    if not pw:
        log("未设置环境变量 MAILMAN_PW（目标账号密码），脚本只会把登录页留给你手动填")

    os.makedirs(STATE_DIR, exist_ok=True)
    key = re.sub(r"[^a-z0-9]", "", a.email.lower()) or "acct"
    verifier = base64.urlsafe_b64encode(secrets.token_bytes(48)).rstrip(b"=").decode()
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    with open(os.path.join(STATE_DIR, f"{key}.verifier.txt"), "w") as f:
        f.write(verifier)

    url = f"{AUTH_BASE}?" + urllib.parse.urlencode({
        "client_id": a.client_id, "response_type": "code", "redirect_uri": REDIRECT,
        "scope": a.scope, "state": f"mailman-{key}", "client_info": "1",
        "login_hint": a.email,
        "code_challenge": challenge, "code_challenge_method": "S256",
    })

    profile = os.path.join(STATE_DIR, "profile")
    code = None
    with sync_playwright() as sp:
        ctx = sp.chromium.launch_persistent_context(profile, channel="chrome", headless=False)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(url, wait_until="domcontentloaded")
        deadline = time.time() + a.timeout
        code_sent_at = None
        while time.time() < deadline:
            href = page.url
            if "nativeclient" in href and "code=" in href:
                code = urllib.parse.parse_qs(urllib.parse.urlparse(href).query).get("code", [None])[0]
                log("已收到授权码")
                break

            if visible(page, "#i0116") or visible(page, "input[name=loginfmt]"):
                sel = "#i0116" if visible(page, "#i0116") else "input[name=loginfmt]"
                page.fill(sel, a.email)
                if visible(page, "#idSIButton9"):
                    page.click("#idSIButton9")
                elif not click_by_text(page, "下一步") and not click_by_text(page, "Next"):
                    page.keyboard.press("Enter")
                log("已提交邮箱")
            elif visible(page, "#i0118") or visible(page, "input[name=passwd]") or visible(page, "input[type=password]"):
                if not pw:
                    log("到密码页了，但没设 MAILMAN_PW —— 请在窗口里手动输入")
                else:
                    sel = "#i0118" if visible(page, "#i0118") else (
                        "input[name=passwd]" if visible(page, "input[name=passwd]") else "input[type=password]"
                    )
                    page.fill(sel, pw)
                    if visible(page, "#idSIButton9"):
                        page.click("#idSIButton9")
                    elif not click_by_text(page, "下一步") and not click_by_text(page, "Next") and not click_by_text(page, "登录"):
                        page.keyboard.press("Enter")
                    log("已提交密码")
            elif visible(page, "#proofConfirmationText"):
                if not a.code_mailbox:
                    log("要邮箱验证码，但没给 --code-mailbox —— 请在窗口里手动完成")
                else:
                    page.fill("#proofConfirmationText", a.code_mailbox)
                    page.click("#primaryButton")
                    code_sent_at = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
                    log(f"已请求发送验证码到 {a.code_mailbox}")
                    got = None
                    for _ in range(12):
                        time.sleep(5)
                        try:
                            got = fetch_ms_code(a.api, a.code_mailbox, code_sent_at)
                        except Exception as e:
                            log("读验证码出错:", e)
                        if got:
                            break
                    if not got:
                        log("没从邮箱读到验证码，请在窗口里手动填")
                    else:
                        log("已填入验证码")
                        for sel in ("input[type=text]:not(#proofConfirmationText)", "input[type=number]", "input[type=tel]"):
                            el = visible(page, sel)
                            if el:
                                el.fill(got)
                                break
                        for bid in ("#primaryButton", "#idSIButton9"):
                            if visible(page, bid):
                                page.click(bid)
                                break
            elif visible(page, "#acceptButton"):
                page.click("#acceptButton")
                log("已点「保持登录」")
            elif click_by_text(page, "接受") or click_by_text(page, "Accept"):
                log("已点 consent「接受」")
            else:
                txt = page.evaluate("() => document.body ? document.body.innerText.slice(0,120) : ''")
                log("等待中:", txt.replace("\n", " / "))
            page.wait_for_timeout(1500)
        ctx.close()

    if not code:
        log("超时未拿到 code")
        sys.exit(1)

    code_file = os.path.join(STATE_DIR, "code.txt")
    with open(code_file, "w") as f:
        f.write(code)

    cmd = ["bun", os.path.join(SKILL_DIR, "scripts", "exchange.ts"),
           "--code-file", code_file, "--key", key, "--email", a.email, "--api", a.api,
           "--client-id", a.client_id, "--scope", a.scope]
    if a.do_import:
        cmd.append("--import")
    if a.verify:
        cmd.append("--verify")
    # 密码不上传命令行（会出现在进程列表）；exchange.ts 从继承的 MAILMAN_PW 环境变量读
    sys.exit(subprocess.call(cmd))


if __name__ == "__main__":
    main()
