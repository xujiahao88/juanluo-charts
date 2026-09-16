# -*- coding: utf-8 -*-
"""把「建材直供」看板的产物同步进钢材站（juanluo-charts）的 zhigong/ 子目录。

做三件事：
  1. 从直供流水线的 dashboard/ 目录复制 index.html + 两张 PNG 到本仓库 zhigong/
  2. 从 index.html 的最新数据日期刷新 assets/app.js 里 EMBED_DATASETS 的 asOf
     （该值显示在站点顶部信息栏「建材直供（至 MM-DD）」）
  3. 可选：调 deploy_repo.py 推送 zhigong/ 三个文件 + app.js 到 GitHub
     （Cloudflare Pages 会自动重新部署）

用法：
    python scripts/sync_zhigong.py                    # 仅本地同步
    python scripts/sync_zhigong.py --deploy           # 同步并推送（需环境变量 GITHUB_PAT）
    python scripts/sync_zhigong.py --src <dashboard>  # 指定直供 dashboard 目录
    python scripts/sync_zhigong.py --check            # 只比对，不写入

注意：直供看板页含中文文件名，脚本按「page1*.png / page2*.png」通配匹配，避免硬编码。
"""
import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent          # juanluo-charts/
ZHIGONG = ROOT / "zhigong"
APP_JS = ROOT / "assets" / "app.js"
DEPLOY = ROOT / "deploy_repo.py"
REPO = "juanluo-charts"

# 直供流水线的默认产物目录（换机时改这一处，或用 --src 覆盖）
DEFAULT_SRC = Path(r"C:/Users/Administrator/WorkBuddy/2026-08-11-23-07-07/dashboard")

DATE_RE = re.compile(r'<td class="date">(\d{4})/(\d{2})/(\d{2})</td>')
ASOF_RE = re.compile(r"(id:\s*'zhigong'[^}]*?asOf:\s*')([^']*)(')")


def latest_asof(html_text: str):
    """从直供页面的数字表里取最后一个日期，返回 'MM-DD'。"""
    hits = DATE_RE.findall(html_text)
    if not hits:
        return None
    _y, m, d = hits[-1]
    return f"{m}-{d}"


def collect_src_files(src: Path):
    """返回 {源文件: 目标文件名}，含 index.html 与 page1/page2 两张 PNG。"""
    files = {}
    idx = src / "index.html"
    if not idx.exists():
        raise SystemExit(f"[x] 找不到 {idx}（用 --src 指定直供 dashboard 目录）")
    files[idx] = "index.html"
    for pat in ("page1*.png", "page2*.png"):
        got = sorted(src.glob(pat))
        if not got:
            raise SystemExit(f"[x] {src} 下找不到 {pat}")
        if len(got) > 1:
            print(f"[!] {pat} 匹配到 {len(got)} 个文件，取最新的：{got[-1].name}")
        files[got[-1]] = got[-1].name
    return files


def main():
    ap = argparse.ArgumentParser(description="同步建材直供看板到钢材站")
    ap.add_argument("--src", default=str(DEFAULT_SRC), help="直供 dashboard 目录")
    ap.add_argument("--deploy", action="store_true", help="同步后推送到 GitHub")
    ap.add_argument("--check", action="store_true", help="只比对，不写入")
    args = ap.parse_args()

    src = Path(args.src)
    files = collect_src_files(src)

    html_text = (src / "index.html").read_text(encoding="utf-8")
    asof = latest_asof(html_text)
    print(f"📖 源目录: {src}")
    print(f"📅 直供最新数据日期: {asof or '<未识别>'}")

    # ---- 1) 复制文件 ----
    ZHIGONG.mkdir(exist_ok=True)
    for s, name in files.items():
        dst = ZHIGONG / name
        same = dst.exists() and dst.read_bytes() == s.read_bytes()
        if same:
            print(f"  = {name} 无变化")
            continue
        if args.check:
            print(f"  ~ {name} 有变化（--check 未写入）")
            continue
        shutil.copy2(s, dst)
        print(f"  ✅ {name} 已更新")

    # ---- 2) 刷新 asOf ----
    app_text = APP_JS.read_text(encoding="utf-8")
    m = ASOF_RE.search(app_text)
    if not m:
        print("[!] app.js 里没找到 EMBED_DATASETS 的 asOf 字段，跳过")
    elif not asof:
        print("[!] 未识别到日期，asOf 保持不变")
    elif m.group(2) == asof:
        print(f"  = app.js asOf 已是 {asof}")
    elif args.check:
        print(f"  ~ app.js asOf {m.group(2)} -> {asof}（--check 未写入）")
    else:
        APP_JS.write_text(ASOF_RE.sub(lambda x: x.group(1) + asof + x.group(3), app_text, count=1),
                          encoding="utf-8")
        print(f"  ✅ app.js asOf {m.group(2)} -> {asof}")

    if args.check:
        print("\n(--check 模式，未改动任何文件)")
        return

    # ---- 3) 推送 ----
    if not args.deploy:
        print("\n本地同步完成。要上线请加 --deploy（需环境变量 GITHUB_PAT）")
        return

    import os
    if not os.environ.get("GITHUB_PAT") and not os.environ.get("GH_TOKEN"):
        raise SystemExit("[x] 未设置 GITHUB_PAT / GH_TOKEN，无法推送")
    if os.environ.get("GH_TOKEN") and not os.environ.get("GITHUB_PAT"):
        os.environ["GITHUB_PAT"] = os.environ["GH_TOKEN"]

    targets = [f"zhigong/{n}" for n in files.values()] + ["assets/app.js"]
    cmd = [sys.executable, str(DEPLOY), REPO] + targets
    print("\n🚀 推送: " + " ".join(cmd))
    r = subprocess.run(cmd, cwd=str(ROOT))
    if r.returncode != 0:
        raise SystemExit(f"[x] 推送失败（exit {r.returncode}）")
    print("\n🎉 已推送，Cloudflare Pages 通常 1~3 分钟完成部署")


if __name__ == "__main__":
    main()
