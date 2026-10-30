# -*- coding: utf-8 -*-
"""文档分发助手 —— Release 上传工具。

用法：
    python release_tool.py ci                 # 查 CI run 状态与 Release 现状
    python release_tool.py upload             # 建/更新 Release，上传两个资产
    python release_tool.py upload 3.2.0       # 指定版本（默认取 lp_tool.py 的 APP_VERSION）

资产从项目根取：
    doc-mailer_v{版本}_portable.zip   （便携包，pack_portable.py 产出）
    doc-mailer_win64.zip              （单文件 exe 包，CI 传的同名未签名版会被自动替换）

⚠️ 三个必须记住的坑：
  1. **发布资产必须发到 uploads.github.com**，发 api.github.com 一律 404
     （建/改/删 Release、删资产发 api 域是正常的，所以症状很像"参数写错"）。
  2. **CI 会自己传一个未签名的 doc-mailer_win64.zip**。要用本地已签名的那份，
     得先 DELETE 同名资产再上传 —— 本脚本已自动处理。
  3. **github.com 本身可能被 IP 级屏蔽**（api.github.com 却能通）。
     git push / 上传走 github.com 时若连不上，先用本地代理（见 HANDOFF §17）。

token 从 Git Credential Manager 现场取用，不落盘、不打印明文。
"""
import json
import re
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

GCM = r"C:\Program Files\Git\mingw64\bin\git-credential-manager.exe"
REPO = "fundops-notes/doc-mailer"
API = "https://api.github.com"
UPLOADS = "https://uploads.github.com"
ROOT = Path(__file__).resolve().parent
if ROOT.name == 'tools' and ROOT.parent.name == 'lp-doc-mailer':
    ROOT = ROOT.parent.parent
BODY_FILE = ROOT / "release_body.md"


def app_version():
    """版本号从 lp_tool.py 里读，别写死 —— 写死过一次就漏发过一次。"""
    src = (ROOT / "lp_tool.py").read_text(encoding="utf-8")
    m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', src)
    if not m:
        raise SystemExit("[X] lp_tool.py 里找不到 APP_VERSION")
    return m.group(1)


def token():
    p = subprocess.run([GCM, "get"], input="protocol=https\nhost=github.com\n\n",
                       capture_output=True, text=True, timeout=90)
    if p.returncode != 0:
        raise SystemExit("[X] 取 GitHub 凭据失败 rc=%d" % p.returncode)
    for line in (p.stdout or "").splitlines():
        if line.startswith("password="):
            return line.split("=", 1)[1].strip()
    raise SystemExit("[X] 凭据里没有 password 字段")


TOK = None


def api(method, path, data=None, raw=None, headers=None):
    url = path if path.startswith("http") else API + path
    if method == "POST" and "/assets?" in path and url.startswith(API):
        url = UPLOADS + url[len(API):]          # ← 坑 1
    hdrs = {"User-Agent": "doc-mailer-release",
            "Accept": "application/vnd.github+json",
            "Authorization": "Bearer %s" % TOK}
    body = None
    if data is not None:
        body = json.dumps(data).encode("utf-8")
        hdrs["Content-Type"] = "application/json"
    if raw is not None:
        body = raw
        hdrs["Content-Type"] = headers or "application/zip"
    req = urllib.request.Request(url, data=body, headers=hdrs, method=method)
    try:
        with urllib.request.urlopen(req, timeout=600) as r:
            b = r.read()
            return r.status, (json.loads(b.decode("utf-8")) if b else None)
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "ignore")[:400]
    except Exception as e:
        return 0, str(e)


def show_ci(tag):
    st, d = api("GET", "/repos/%s/actions/runs?per_page=8" % REPO)
    if st != 200:
        print("[X] 查 CI 失败：%s" % d)
        return
    for r in d.get("workflow_runs", [])[:6]:
        print("  run: %s | %s | status=%s conclusion=%s | %s" % (
            r.get("id"), r.get("head_branch"), r.get("status"),
            r.get("conclusion"), r.get("html_url")))


def show_release(tag):
    st, d = api("GET", "/repos/%s/releases/tags/%s" % (REPO, tag))
    if st == 200:
        print("  Release %s：id=%s name=%s  url=%s" % (
            tag, d.get("id"), d.get("name"), d.get("html_url")))
        for a in d.get("assets", []):
            print("    资产: %s（%d 字节, 下载 %s 次）" % (
                a.get("name"), a.get("size"), a.get("download_count")))
        return d
    print("  Release %s 不存在（HTTP %s）" % (tag, st))
    return None


def upload(version):
    tag = "v%s" % version
    portable = ROOT / ("doc-mailer_v%s_portable.zip" % version)
    lightweight = ROOT / ("doc-mailer_v%s_update.zip" % version)
    win64 = ROOT / "doc-mailer_win64.zip"
    for p in (portable, lightweight, win64):
        if not p.exists():
            raise SystemExit("[X] 资产不存在：%s" % p)
        print("[OK] %s（%.1f MB）" % (p.name, p.stat().st_size / 1024 / 1024))
    body = BODY_FILE.read_text(encoding="utf-8") if BODY_FILE.exists() else ""

    st, rel = api("GET", "/repos/%s/releases/tags/%s" % (REPO, tag))
    if st == 200:
        rid = rel["id"]
        print("[OK] Release 已存在，id=%s（更新说明）" % rid)
        st, d = api("PATCH", "/repos/%s/releases/%s" % (REPO, rid),
                    data={"tag_name": tag, "name": "文档分发助手 v%s" % version,
                          "body": body, "draft": False, "prerelease": False})
        if st != 200:
            raise SystemExit("[X] 更新 Release 失败：%s" % d)
    else:
        st, d = api("POST", "/repos/%s/releases" % REPO,
                    data={"tag_name": tag, "name": "文档分发助手 v%s" % version,
                          "body": body, "draft": False, "prerelease": False})
        if st != 201:
            raise SystemExit("[X] 创建 Release 失败：%s" % d)
        rid = d["id"]
        print("[OK] Release 已创建，id=%s" % rid)

    # 同名旧资产先删（CI 传的 win64 是未签名的，要换成本地已签名那份）  ← 坑 2
    st, d = api("GET", "/repos/%s/releases/%s" % (REPO, rid))
    if st == 200:
        for a in d.get("assets", []):
            if a.get("name") in (portable.name, lightweight.name, win64.name):
                api("DELETE", "/repos/%s/releases/assets/%s" % (REPO, a["id"]))
                print("[OK] 删除同名旧资产 %s（%d 字节）" % (a.get("name"), a.get("size")))

    for p in (portable, lightweight, win64):
        raw = p.read_bytes()
        print("[..] 上传 %s（%.1f MB）…" % (p.name, len(raw) / 1024 / 1024))
        st, d = api("POST", "/repos/%s/releases/%s/assets?name=%s"
                    % (REPO, rid, p.name), raw=raw, headers="application/zip")
        if st != 201:
            raise SystemExit("[X] 上传 %s 失败 HTTP %s：%s" % (p.name, st, d))
        print("[OK] 资产已上传：%s（%d 字节，id=%s）"
              % (d.get("name"), d.get("size"), d.get("id")))


if __name__ == "__main__":
    TOK = token()
    cmd = sys.argv[1] if len(sys.argv) > 1 else "ci"
    version = sys.argv[2] if len(sys.argv) > 2 else app_version()
    tag = "v%s" % version
    if cmd == "ci":
        print("=== CI runs ===")
        show_ci(tag)
        print("=== Release ===")
        show_release(tag)
    elif cmd == "upload":
        upload(version)
        print("=== 上传后复核 ===")
        show_release(tag)
    else:
        print("用法：ci | upload [版本]（版本默认取 lp_tool.py 的 APP_VERSION）")
