#!/usr/bin/env python3
"""同步本目录到 GitHub: 593682963wd-wq/route-payload-web

第一次运行会自动创建仓库；之后每次运行做幂等增量上传（走 Contents API，
绕开本地 git push 偶尔的网络问题）。

依赖：Mac 钥匙串里已存有 github.com 的 PAT（与 obstacle-web、nofly-web 同账号）。
"""
import base64, json, os, subprocess, sys, urllib.error, urllib.parse, urllib.request

OWNER  = "593682963wd-wq"
REPO   = "route-payload-web"
BRANCH = "main"
DESC   = "✈ 航线载量分析系统 Web 版 · OFP TXT 批量解析输出 Word"
ROOT   = os.path.dirname(os.path.abspath(__file__))

SKIP_DIRS  = {".git", "__pycache__", ".venv", "venv", "node_modules", "输入", "输出"}
SKIP_FILES = {".DS_Store"}


def get_token() -> str:
    out = subprocess.run(
        ["git", "credential", "fill"],
        input="protocol=https\nhost=github.com\n\n",
        capture_output=True, text=True, check=True,
    ).stdout
    for line in out.splitlines():
        if line.startswith("password="):
            return line.split("=", 1)[1]
    sys.exit("找不到 github.com 的 PAT (Mac 钥匙串)")


TOKEN = get_token()


def api(method: str, path: str, body=None):
    req = urllib.request.Request(
        f"https://api.github.com/{path}",
        method=method,
        headers={
            "Authorization": f"token {TOKEN}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "route-payload-uploader",
        },
        data=json.dumps(body).encode() if body else None,
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status, json.loads(r.read() or b"{}")
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read() or b"{}")


def ensure_repo():
    code, _ = api("GET", f"repos/{OWNER}/{REPO}")
    if code == 200:
        print(f"  ==  仓库已存在: {OWNER}/{REPO}")
        return
    print(f"  ++  创建仓库: {OWNER}/{REPO}")
    code, data = api("POST", "user/repos", {
        "name": REPO,
        "description": DESC,
        "private": False,
        "auto_init": True,
        "default_branch": BRANCH,
    })
    if code >= 400:
        sys.exit(f"创建失败: {data}")


def upload(rel_path: str, abs_path: str) -> bool:
    with open(abs_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()
    enc = urllib.parse.quote(rel_path)
    code, data = api("GET", f"repos/{OWNER}/{REPO}/contents/{enc}?ref={BRANCH}")
    body = {
        "message": f"sync: {rel_path}",
        "content": content_b64,
        "branch": BRANCH,
    }
    if code == 200 and isinstance(data, dict) and "sha" in data:
        if data.get("content", "").replace("\n", "") == content_b64:
            print(f"  ==  {rel_path} (无变化)")
            return True
        body["sha"] = data["sha"]
    code, data = api("PUT", f"repos/{OWNER}/{REPO}/contents/{enc}", body)
    print(f"  {code}  {rel_path}")
    if code >= 400:
        print(f"      ERR: {data}")
        return False
    return True


def main():
    ensure_repo()
    ok = fail = 0
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fn in filenames:
            if fn in SKIP_FILES or fn.endswith(".pyc"):
                continue
            abs_p = os.path.join(dirpath, fn)
            rel_p = os.path.relpath(abs_p, ROOT)
            if upload(rel_p, abs_p):
                ok += 1
            else:
                fail += 1
    print(f"\n完成: 成功 {ok}, 失败 {fail}")
    print(f"\n→ 仓库:    https://github.com/{OWNER}/{REPO}")
    print( "→ 部署到:  https://share.streamlit.io  (New app → 选本仓库 → main → app.py)")


if __name__ == "__main__":
    main()
