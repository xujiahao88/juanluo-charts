#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_biaowai_pipeline.py — 粗钢及表外「周更 → 公网 → 出图 → 发微信」一条龙
========================================================================
数据源：C:/Users/Administrator/Nutstore/1/小目标/粗钢及表外(非五大材)情况.xlsx（周频，用户每周五更新）
站点：  https://xujiahao88.github.io/juanluo-charts/（juanluo-charts / GitHub Pages）

链路（默认全跑，跑完即发微信——用户 2026-09-17 约定「run_* 默认发微信，--no-send 才跳过」）：
  1. 检查源 Excel mtime 是否比上次处理新（防止空跑）；Excel 占用检测
  2. build_biaowai_charts.py 重抽 → 3 个 dataset（biaowai / variety_demand / variety_inventory）+ data.js + meta.json
  3. bump index.html 缓存版本 ?v=N → N+1
  4. deploy_repo.py 推 GitHub（Pages 自动重建）
  5. shot_juanluo.py 出 3 张长图（表外 / 分品种表需 / 分品种库存，2K 密度）
  6. send_to_wechat.py 发 3 张图 + send_text_to_wechat.py 发点评文字（--text-file，可选）
     状态写入 shot/.biaowai_state.json（mtime 判重 + last_sent_asOf 防重复发送）

用法：
  python scripts/run_biaowai_pipeline.py                    # 有变化才跑，跑完发微信
  python scripts/run_biaowai_pipeline.py --force            # 忽略 mtime，强制重跑（可重发）
  python scripts/run_biaowai_pipeline.py --dry-run          # 只预览
  python scripts/run_biaowai_pipeline.py --no-push          # 不推公网
  python scripts/run_biaowai_pipeline.py --no-shot          # 不出图
  python scripts/run_biaowai_pipeline.py --no-send          # 不发微信（仍出图）
  python scripts/run_biaowai_pipeline.py --send-only --text-file <txt>
                                                            # 只发送：已有长图 + 文字（自动化流程第二段）
  python scripts/run_biaowai_pipeline.py --text-file <txt>  # 全链 + 图后补发点评文字
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import subprocess
import sys
import time

PY = r"C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe"
SITE = r"C:/Users/Administrator/juanluo-charts"
SRC = r"C:/Users/Administrator/Nutstore/1/小目标/粗钢及表外(非五大材)情况.xlsx"
BUILD_SCRIPT = os.path.join(SITE, "scripts", "build_biaowai_charts.py")
DEPLOY_SCRIPT = os.path.join(SITE, "deploy_repo.py")
SHOT_SCRIPT = os.path.join(SITE, "scripts", "shot_juanluo.py")
SEND_SCRIPT = r"C:/Users/Administrator/iron-ore-charts/scripts/send_to_wechat.py"
SEND_TEXT_SCRIPT = os.path.join(SITE, "scripts", "send_text_to_wechat.py")

SHOT_DIR = os.path.join(SITE, "shot")
STATE = os.path.join(SHOT_DIR, ".biaowai_state.json")

# (dataset id, 长图文件名, 图标题前缀, 窗口高度CSS px)——高度给足，截图脚本会自动裁白边
SHOTS = [
    ("biaowai",           "表外_长图.png",     "表外（非五大材）", 4600),
    ("variety_demand",    "分品种表需_长图.png", "分品种表需",      3000),
    ("variety_inventory", "分品种库存_长图.png", "分品种库存",      5600),
]

DEPLOY_FILES = [
    "index.html",
    "data/data.js", "data/meta.json",
    "data/biaowai.json", "data/variety_demand.json", "data/variety_inventory.json",
    "scripts/build_biaowai_charts.py",
    "scripts/run_biaowai_pipeline.py",
]

_t0 = time.time()


def log(msg):
    print(f"[{time.time() - _t0:6.1f}s] {msg}", flush=True)


def run(cmd, cwd=None, env=None, check=True, timeout=600):
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


def save_state(**kw):
    os.makedirs(SHOT_DIR, exist_ok=True)
    st = load_state()
    st.update(kw)
    st["processed_at"] = datetime.datetime.now().isoformat(timespec="seconds")
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1)


def excel_locked():
    """源文件同目录下存在 ~$粗钢及表外(非五大材)情况.xlsx 说明 Excel 正开着。"""
    d = os.path.dirname(SRC)
    return os.path.exists(os.path.join(d, "~$" + os.path.basename(SRC)))


def read_asof():
    """从 data/biaowai.json 读 asOf（构建后调用）。"""
    try:
        with open(os.path.join(SITE, "data", "biaowai.json"), encoding="utf-8") as f:
            return json.load(f).get("asOf", "")
    except Exception:
        return ""


def bump_cache():
    """index.html 里 ?v=N → N+1（全部替换），返回新版本号。"""
    p = os.path.join(SITE, "index.html")
    with open(p, encoding="utf-8") as f:
        html = f.read()
    nums = [int(x) for x in re.findall(r"\?v=(\d+)", html)]
    if not nums:
        log("  ⚠️ index.html 未找到 ?v=NN，跳过缓存 bump")
        return None
    new = max(nums) + 1
    html = re.sub(r"\?v=(\d+)", "?v=%d" % new, html)
    with open(p, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"  缓存版本 ?v={max(nums)} → ?v={new}（index.html）")
    return new


def shot_paths():
    return [os.path.join(SHOT_DIR, fn) for _, fn, _, _ in SHOTS]


def do_send(text_file=None, force=False):
    """发 3 张长图（+可选点评文字）到微信。返回 True/False。"""
    imgs = shot_paths()
    missing = [p for p in imgs if not os.path.exists(p)]
    if missing:
        log("  ❌ 缺长图，跳过发送：%s" % ", ".join(os.path.basename(m) for m in missing))
        return False

    # 防重复：同一 asOf 已经发过则跳过（除非 --force）
    asof = read_asof()
    st = load_state()
    if not force and asof and st.get("sent_asOf") == asof:
        log(f"  ⏭ 该版本（asOf={asof}）已发送过，跳过（--force 可强制重发）")
        return False

    log("--- 发送 3 张长图到微信「文件传输助手」 ---")
    rc, _ = run([PY, SEND_SCRIPT, "--no-countdown"] + imgs, cwd=SITE, check=False, timeout=420)
    ok = rc == 0
    if not ok:
        log("  ⚠️ 微信发图失败（可能未登录/未切到目标会话），图仍在 shot/ 目录")

    if text_file:
        if not os.path.exists(text_file):
            log("  ⚠️ 找不到点评文本: " + str(text_file))
        else:
            log("  📝 发送点评文字: " + os.path.basename(text_file))
            rc2, _ = run([PY, SEND_TEXT_SCRIPT, "--no-countdown", "--text-file", text_file],
                         cwd=SITE, check=False, timeout=120)
            if rc2 != 0:
                log("  ⚠️ 点评发送失败，文本仍在: " + str(text_file))

    if ok:
        save_state(sent_asOf=asof, sent_at=datetime.datetime.now().isoformat(timespec="seconds"))
    return ok


def main():
    global SRC
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=SRC)
    ap.add_argument("--force", action="store_true", help="忽略 mtime/已发送标记，强制重跑重发")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--no-shot", action="store_true")
    ap.add_argument("--no-send", "--no-wechat", dest="no_send", action="store_true",
                    help="不发微信（仍出图）")
    ap.add_argument("--send-only", action="store_true",
                    help="只发送：用 shot/ 里已有长图 + --text-file 文字（跳过 build/push/shot）")
    ap.add_argument("--text-file", default=None, help="点评文本文件（发完图后一发）")
    args = ap.parse_args()

    SRC = args.src

    # ---------- 只发送模式（自动化第二段：build 完成后写好评点再来） ----------
    if args.send_only:
        log("--- [--send-only] 只发送模式 ---")
        ok = do_send(text_file=args.text_file, force=args.force)
        log("✅ 完成" if ok else "⚠️ 未发送（见上方原因）")
        return

    mt = src_mtime()
    state = load_state()
    changed = args.force or mt > float(state.get("mtime", 0) or 0)

    if args.dry_run:
        log("--- [--dry-run] 预览 ---")
        log(f"源文件: {SRC} (mtime={mt})")
        log(f"上次处理 mtime={state.get('mtime')} → {'有变化，将执行' if changed else '无变化，将跳过'}")
        log(f"1) {PY} {BUILD_SCRIPT}（重建 3 个 tab）")
        log(f"2) bump index.html ?v=N → N+1")
        log(f"3) {PY} {DEPLOY_SCRIPT} juanluo-charts {' '.join(DEPLOY_FILES)}")
        for dsid, fn, title, h in SHOTS:
            log(f"4) {PY} {SHOT_SCRIPT} --ds {dsid} --out {os.path.join(SHOT_DIR, fn)} --height {h}")
        log(f"5) {PY} {SEND_SCRIPT} --no-countdown <3 png> [+ 文字：{args.text_file or '无'}]")
        log(f"   上次发送 asOf={state.get('sent_asOf')}")
        return

    if not os.path.exists(SRC):
        raise SystemExit(f"找不到源文件: {SRC}")
    if not changed:
        log(f"⏭ 源文件未变化（mtime={mt}），跳过。需要强制请加 --force")
        return
    if excel_locked():
        raise SystemExit("⚠️ 检测到 ~$粗钢及表外(非五大材)情况.xlsx，Excel 可能正打开着；"
                         "请先保存并关闭后重试")
    for p in (BUILD_SCRIPT, DEPLOY_SCRIPT, SHOT_SCRIPT, SEND_SCRIPT, SEND_TEXT_SCRIPT, PY):
        if not os.path.exists(p):
            raise SystemExit(f"找不到必要文件: {p}")

    # 1. 抽数（3 个 dataset 一起重建）
    log("--- 步骤 1: 重抽数据（build_biaowai_charts.py，3 个 tab） ---")
    run([PY, BUILD_SCRIPT], cwd=SITE)

    as_of = read_asof()
    log(f"   数据截至 asOf={as_of}")

    # 2. bump 缓存
    log("--- 步骤 2: bump index.html 缓存版本 ---")
    bump_cache()

    # 3. 推公网
    if args.no_push:
        log("--- 步骤 3: 跳过推送 (--no-push) ---")
    elif not os.environ.get("GITHUB_PAT"):
        log("  ⚠️ GITHUB_PAT 未设置，跳过推送")
    else:
        log("--- 步骤 3: 推 juanluo-charts ---")
        env = dict(os.environ, COMMIT_MSG=f"data(biaowai): 表外/分品种更新至 {as_of}")
        run([PY, DEPLOY_SCRIPT, "juanluo-charts"] + DEPLOY_FILES, cwd=SITE, env=env)

    # 4. 出长图
    if args.no_shot:
        log("--- 步骤 4: 跳过出图 (--no-shot) ---")
    else:
        log("--- 步骤 4: 出 3 张长图（2K 密度） ---")
        os.makedirs(SHOT_DIR, exist_ok=True)
        for dsid, fn, title, h in SHOTS:
            img = os.path.join(SHOT_DIR, fn)
            full_title = f"{title} · 数据截至 {as_of}" if as_of else title
            run([PY, SHOT_SCRIPT, "--ds", dsid, "--out", img,
                 "--width", "1920", "--height", str(h), "--scale", "1.333",
                 "--title", full_title], cwd=SITE, timeout=420)

    # 5. 发微信
    if args.no_shot:
        log("--- 步骤 5: 跳过（未出图） ---")
    elif args.no_send:
        log("--- 步骤 5: 未发微信 (--no-send)，手动发送：")
        log(f'    {PY} {SEND_SCRIPT} --no-countdown <shot 下 3 张长图>')
    else:
        log("--- 步骤 5: 发微信 ---")
        do_send(text_file=args.text_file, force=args.force)

    save_state(mtime=mt, asOf=as_of)
    log("✅ 完成")
    log(f"   总耗时 {time.time() - _t0:.1f}s")
    log(f"   公网: https://xujiahao88.github.io/juanluo-charts/")


if __name__ == "__main__":
    main()
