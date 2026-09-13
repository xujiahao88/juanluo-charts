#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_chugang_pipeline.py — 出港「Excel 保存 → 更新公网 → 发微信」一条龙
=====================================================================
数据源：C:/Users/Administrator/Nutstore/1/小目标/出港.xlsx
站点：  https://xujiahao88.github.io/juanluo-charts/（juanluo-charts / GitHub Pages）

链路：
  1. 检查源 Excel 是否被占用、mtime 是否比上次处理新（防止空跑）
  2. build_daiguan_charts.py 重抽 → data/data.js + chugang.json + meta.json
  3. deploy_repo.py 推 GitHub（Pages 自动重建）
  4. shot_juanluo.py Edge headless 出「出港」长图
  5. send_to_wechat.py 发微信「文件传输助手」

用法：
  python scripts/run_chugang_pipeline.py                # 默认：有变化才跑，跑完发微信
  python scripts/run_chugang_pipeline.py --force        # 忽略 mtime，强制重跑
  python scripts/run_chugang_pipeline.py --dry-run      # 只预览
  python scripts/run_chugang_pipeline.py --no-push      # 不推公网
  python scripts/run_chugang_pipeline.py --no-shot      # 不出图
  python scripts/run_chugang_pipeline.py --no-wechat    # 不发微信
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import subprocess
import sys
import time

PY = r"C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe"
SITE = r"C:/Users/Administrator/juanluo-charts"
SRC = r"C:/Users/Administrator/Nutstore/1/小目标/出港.xlsx"
BUILD_SCRIPT = os.path.join(SITE, "scripts", "build_daiguan_charts.py")
DEPLOY_SCRIPT = os.path.join(SITE, "deploy_repo.py")
SHOT_SCRIPT = os.path.join(SITE, "scripts", "shot_juanluo.py")
SEND_SCRIPT = r"C:/Users/Administrator/iron-ore-charts/scripts/send_to_wechat.py"

SHOT_DIR = os.path.join(SITE, "shot")
IMG = os.path.join(SHOT_DIR, "出港_长图.png")
STATE = os.path.join(SHOT_DIR, ".chugang_state.json")

DEPLOY_FILES = [
    "index.html", "assets/app.js", "assets/style.css",
    "data/data.js", "data/meta.json", "data/chugang.json",
    "data/daiguan.json", "data/hanguan.json",
    "scripts/build_daiguan_charts.py", "scripts/run_chugang_pipeline.py",
]

_t0 = time.time()


def log(msg):
    print(f"[{time.time() - _t0:6.1f}s] {msg}", flush=True)


def run(cmd, cwd=None, env=None, check=True, timeout=300):
    r = subprocess.run(cmd, cwd=cwd, env=env, capture_output=True, text=True, timeout=timeout)
    out = (r.stdout or "") + (r.stderr or "")
    for line in out.splitlines():
        if line.strip():
            log("   " + line)
    if check and r.returncode != 0:
        raise SystemExit(f"命令失败 rc={r.returncode}: {' '.join(cmd[:4])}...")
    return r.returncode, out


def src_mtime():
    return os.path.getmtime(SRC) if os.path.exists(SRC) else 0


def load_state():
    try:
        with open(STATE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(mtime):
    os.makedirs(SHOT_DIR, exist_ok=True)
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump({"mtime": mtime, "processed_at": datetime.datetime.now().isoformat(timespec="seconds")},
                  f, ensure_ascii=False, indent=1)


def excel_locked():
    """源文件同目录下存在 ~$出港.xlsx 说明 Excel 正开着。"""
    d = os.path.dirname(SRC)
    return os.path.exists(os.path.join(d, "~$" + os.path.basename(SRC)))


def main():
    global SRC
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=SRC)
    ap.add_argument("--force", action="store_true", help="忽略 mtime，强制重跑")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--no-shot", action="store_true")
    ap.add_argument("--no-wechat", action="store_true")
    args = ap.parse_args()

    SRC = args.src
    mt = src_mtime()
    state = load_state()
    changed = args.force or mt > float(state.get("mtime", 0) or 0)

    if args.dry_run:
        log("--- [--dry-run] 预览 ---")
        log(f"源文件: {SRC} (mtime={mt})")
        log(f"上次处理 mtime={state.get('mtime')} → {'有变化，将执行' if changed else '无变化，将跳过'}")
        log(f"1) {PY} {BUILD_SCRIPT}")
        log(f"2) {PY} {DEPLOY_SCRIPT} juanluo-charts {' '.join(DEPLOY_FILES)}")
        log(f"3) {PY} {SHOT_SCRIPT} --ds chugang --out {IMG} …")
        log(f"4) {PY} {SEND_SCRIPT} --no-countdown {IMG}")
        return

    if not os.path.exists(SRC):
        raise SystemExit(f"找不到源文件: {SRC}")
    if not changed:
        log(f"⏭ 源文件未变化（mtime={mt}），跳过。需要强制请加 --force")
        return
    if excel_locked():
        raise SystemExit("⚠️ 检测到 ~$出港.xlsx，Excel 可能正打开着；请先保存并关闭后重试")

    for p in (BUILD_SCRIPT, DEPLOY_SCRIPT, SHOT_SCRIPT, SEND_SCRIPT, PY):
        if not os.path.exists(p):
            raise SystemExit(f"找不到必要文件: {p}")

    # 1. 抽数
    log("--- 步骤 1: 重抽数据 (build_daiguan_charts.py) ---")
    run([PY, BUILD_SCRIPT], cwd=SITE)

    # 读 asOf 用于出图标题
    as_of = ""
    try:
        with open(os.path.join(SITE, "data", "chugang.json"), encoding="utf-8") as f:
            as_of = json.load(f).get("asOf", "")
    except Exception:
        pass
    title = f"钢材出港&接单 · 数据截至 {as_of}" if as_of else "钢材出港&接单"
    subtitle = "数据来源：小目标/出港.xlsx"

    # 2. 推公网
    if args.no_push:
        log("--- 步骤 2: 跳过推送 (--no-push) ---")
    elif not os.environ.get("GITHUB_PAT"):
        log("  ⚠️ GITHUB_PAT 未设置，跳过推送")
    else:
        log("--- 步骤 2: 推 juanluo-charts ---")
        env = dict(os.environ, COMMIT_MSG=f"data: 出港更新 ({datetime.date.today().isoformat()})")
        run([PY, DEPLOY_SCRIPT, "juanluo-charts"] + DEPLOY_FILES, cwd=SITE, env=env)

    # 3. 出长图
    if args.no_shot:
        log("--- 步骤 3: 跳过出图 (--no-shot) ---")
    else:
        log("--- 步骤 3: 出「出港」长图 ---")
        os.makedirs(SHOT_DIR, exist_ok=True)
        run([PY, SHOT_SCRIPT, "--ds", "chugang", "--out", IMG,
             "--width", "1920", "--height", "2600",
             "--title", title, "--subtitle", subtitle], cwd=SITE)

    # 4. 发微信
    if args.no_shot:
        log("--- 步骤 4: 跳过（未出图）---")
    elif args.no_wechat:
        log(f"--- 步骤 4: 未发微信 (--no-wechat)，手动：{PY} {SEND_SCRIPT} --no-countdown \"{IMG}\" ---")
    else:
        log("--- 步骤 4: 发微信「文件传输助手」 ---")
        if not os.path.exists(IMG):
            log("  ⚠️ 找不到长图，跳过发送")
        else:
            rc, _ = run([PY, SEND_SCRIPT, "--no-countdown", IMG], cwd=SITE, check=False)
            if rc != 0:
                log("  ⚠️ 微信发送失败（可能未登录/未切到文件传输助手），图仍在: " + IMG)

    save_state(mt)
    log("✅ 完成")
    log(f"   总耗时 {time.time() - _t0:.1f}s")


if __name__ == "__main__":
    main()
