#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_mill_order_pipeline.py — 「钢厂日接单」一条龙调度器
=========================================================

输入 = 用户当天发来的「唐山钢厂当日接单量(勿转)」截图里读到的一天数据，
一次跑完：

  1. 清理 Excel 孤儿进程
  2. update_mill_order_source.py  写《钢厂日接单量统计.xlsx》「接单」
     （按表头名定位列、按日期定位行，命中覆盖 / 未命中追加；顺带维护「纵横+中铁」合并口径）
  3. build_mill_order_charts.py   抽数 → data/mill_order.json + data.js + meta.json
  4. 自动 bump index.html 缓存版本（?v=N → N+1，强制刷新）
  5. deploy_repo.py 推 juanluo-charts（GitHub Pages 自动重建） ‖ shot_juanluo.py 出长图（并行）
  6. send_to_wechat.py 发微信「文件传输助手」（默认；--no-wechat 跳过）

用法：
  python run_mill_order_pipeline.py --json '{"date":"2026-09-16","mills":{...},"profit":-40,"total":15.2}'
  python run_mill_order_pipeline.py --json-file row.json
  python run_mill_order_pipeline.py --skip-source              # 只重建+推站+出图（源已手改）
  python run_mill_order_pipeline.py --json '...' --no-wechat --no-push
  python run_mill_order_pipeline.py --json '...' --dry-run

说明：
  · --json / --json-file 至少给一个；给了才跑「写源」那步。
  · 长图存 shot/钢厂日接单_长图.png；新鲜度会校验（本轮生成的才发）。
  · 前提（发微信）：微信 PC 已登录 + 已切到「文件传输助手」+ 窗口在前台。
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import subprocess
import time

PY = r"C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe"
SITE = r"C:/Users/Administrator/juanluo-charts"
UPDATE_SCRIPT = os.path.join(SITE, "scripts", "update_mill_order_source.py")
BUILD_SCRIPT = os.path.join(SITE, "scripts", "build_mill_order_charts.py")
SHOT_SCRIPT = os.path.join(SITE, "scripts", "shot_juanluo.py")
DEPLOY_SCRIPT = os.path.join(SITE, "deploy_repo.py")
SEND_SCRIPT = os.path.join(SITE, "..", "iron-ore-charts", "scripts", "send_to_wechat.py")
SHOT = os.path.join(SITE, "shot")
LONG_IMG = os.path.join(SHOT, "钢厂日接单_长图.png")
INDEX = os.path.join(SITE, "index.html")

_t0 = time.time()


def log(msg):
    print(f"[{time.time() - _t0:6.1f}s] {msg}", flush=True)


def run(cmd, cwd=None, env=None, check=True, timeout=300):
    r = subprocess.run(cmd, cwd=cwd, env=env, capture_output=True, text=True, timeout=timeout)
    for line in ((r.stdout or "") + (r.stderr or "")).splitlines():
        if line.strip():
            log("   " + line)
    if check and r.returncode != 0:
        raise SystemExit(f"命令失败 rc={r.returncode}: {' '.join(str(c) for c in cmd[:4])}...")
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def launch(cmd, cwd=None, env=None):
    log("   ↳ (bg) " + " ".join(os.path.basename(x) if str(x).endswith('.py') else str(x) for x in cmd[:3]) + " ...")
    return subprocess.Popen(cmd, cwd=cwd, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)


def kill_orphan_excel():
    """清理没有可见窗口的孤儿 EXCEL.EXE（保留用户打开的）。"""
    try:
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                           capture_output=True, timeout=15)
    except Exception as e:
        log(f"  tasklist 失败：{e}")
        return
    pids = []
    for raw in (r.stdout or b"").splitlines():
        line = raw.decode("gbk", errors="ignore").strip().strip('"')
        if "EXCEL.EXE" not in line.upper():
            continue
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) >= 2:
            pids.append(parts[1])
    killed = 0
    for pid in pids:
        try:
            v = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH", "/V"],
                               capture_output=True, timeout=10)
            txt = (v.stdout or b"").decode("gbk", errors="ignore")
            if "可见" in txt or "Visible" in txt:
                continue
            subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True, timeout=10)
            killed += 1
        except Exception:
            pass
    log(f"  清理孤儿 Excel: {killed} 个" if killed else "  无孤儿（或用户已开 Excel）")


def bump_cache():
    """index.html 里 ?v=N → N+1（3 处一起），返回新版本号。"""
    with open(INDEX, encoding="utf-8") as f:
        html = f.read()
    nums = [int(x) for x in re.findall(r'\?v=(\d+)', html)]
    if not nums:
        return None
    new = max(nums) + 1
    html = re.sub(r'\?v=(\d+)', '?v=%d' % new, html)
    with open(INDEX, "w", encoding="utf-8") as f:
        f.write(html)
    return new


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--json", help='行数据 JSON（inline）')
    g.add_argument("--json-file", help='行数据 JSON 文件路径')
    ap.add_argument("--skip-source", action="store_true", help="跳过写源（源已手改，只重建+推站+出图）")
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--no-shot", action="store_true")
    ap.add_argument("--no-wechat", action="store_true", help="默认直接发微信；--no-wechat 才跳过")
    ap.add_argument("--no-countdown", action="store_true",
                    help="发微信时跳过 5s 倒计时（默认保留，供你切到「文件传输助手」）")
    ap.add_argument("--height", type=int, default=3000, help="截图窗口高（默认 3000，自动裁白边）")
    ap.add_argument("--scale", type=float, default=1.333, help="截图像素密度（默认 1.333 = 2K）")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not args.skip_source and not (args.json or args.json_file):
        raise SystemExit("需要 --json / --json-file（或用 --skip-source 跳过写源）")

    for p in (BUILD_SCRIPT, SHOT_SCRIPT, DEPLOY_SCRIPT, PY):
        if not os.path.exists(p):
            raise SystemExit(f"找不到必要文件: {p}")

    payload = None
    if args.json or args.json_file:
        payload = json.loads(args.json) if args.json else json.load(open(args.json_file, encoding="utf-8"))

    if args.dry_run:
        log("--- [--dry-run] 仅预览 ---")
        if payload:
            log(f"   写源: {PY} {UPDATE_SCRIPT} --json '{json.dumps(payload, ensure_ascii=False)}'")
        else:
            log("   写源: (跳过)")
        log(f"   抽数: {PY} {BUILD_SCRIPT}")
        log(f"   推:   {DEPLOY_SCRIPT} juanluo-charts index.html data/data.js data/meta.json "
            f"data/mill_order.json scripts/{os.path.basename(BUILD_SCRIPT)} ... （+bump 缓存）")
        log(f"   出图: {SHOT_SCRIPT} --ds mill_order --out {LONG_IMG} (scale {args.scale})")
        log(f"   发:   {SEND_SCRIPT} --no-countdown {LONG_IMG}")
        return

    # 1. 清理孤儿 Excel
    log("--- 步骤 1: 清理孤儿 Excel 进程 ---")
    kill_orphan_excel()

    # 2. 写源
    if args.skip_source:
        log("--- 步骤 2: 跳过写源（--skip-source）---")
    else:
        log("--- 步骤 2: 写《钢厂日接单量统计.xlsx》「接单」 ---")
        rc, _ = run([PY, UPDATE_SCRIPT, "--json", json.dumps(payload, ensure_ascii=False)], check=False)
        if rc != 0:
            raise SystemExit("  ❌ 写源失败，终止")

    # 3. 抽数
    log("--- 步骤 3: 抽数 build_mill_order_charts.py ---")
    run([PY, BUILD_SCRIPT], cwd=SITE)

    # 4. bump 缓存
    v = bump_cache()
    log(f"--- 步骤 4: index.html 缓存版本 → v={v} ---")

    # 5+6. 推送 与 出图【重叠】
    procs = []
    if args.no_push:
        log("--- 步骤 5: 跳过推送（--no-push）---")
    elif not os.environ.get("GITHUB_PAT"):
        log("  ⚠️ GITHUB_PAT 未设置，跳过推送")
    else:
        log("--- 步骤 5: 推 juanluo-charts（后台，与出图重叠）---")
        files = ["index.html", "assets/app.js", "assets/style.css",
                 "data/data.js", "data/meta.json", "data/mill_order.json",
                 "scripts/build_mill_order_charts.py", "scripts/update_mill_order_source.py",
                 "scripts/run_mill_order_pipeline.py"]
        asof = payload.get('date') if payload else datetime.date.today().isoformat()
        env = dict(os.environ, COMMIT_MSG=f"data(mill_order): 钢厂日接单更新 ({asof}) v{v}")
        procs.append(launch([PY, DEPLOY_SCRIPT, "juanluo-charts"] + files, cwd=SITE, env=env))

    if args.no_shot:
        log("--- 步骤 6: 跳过出图（--no-shot）---")
    else:
        log("--- 步骤 6: 出「钢厂日接单」长图（后台，与推送重叠）---")
        os.makedirs(SHOT, exist_ok=True)
        procs.append(launch([PY, SHOT_SCRIPT, "--ds", "mill_order", "--out", LONG_IMG,
                             "--width", "1920", "--height", str(args.height),
                             "--scale", str(args.scale)], cwd=SITE))

    if procs:
        log(f"⏳ 等待 {len(procs)} 个后台任务（推送/出图）完成...")
        for p in procs:
            out, _ = p.communicate()
            for line in (out or "").splitlines():
                if line.strip():
                    log("   " + line)
            if p.returncode != 0:
                log(f"  ⚠️ 后台任务 rc={p.returncode}")

    # 7. 发微信（默认）
    if args.no_shot:
        log("--- 步骤 7: 跳过发送（未出图）---")
    elif args.no_wechat:
        log("--- 步骤 7: 未发微信（--no-wechat）。手动发：---")
        log(f'   python "{SEND_SCRIPT}" --no-countdown "{LONG_IMG}"')
    else:
        log("--- 步骤 7: 发微信（默认）---")
        if not os.path.exists(LONG_IMG):
            log("  ⚠️ 找不到长图，跳过")
        elif os.path.getmtime(LONG_IMG) < _t0:
            log("  ⚠️ 长图不是本轮生成（新鲜度校验未过），跳过发送")
        else:
            log("  ⚠️ 将发 1 张图到微信「文件传输助手」，请确保已切到该会话")
            extra = ["--no-countdown"] if args.no_countdown else []
            run([PY, SEND_SCRIPT] + extra + [LONG_IMG], cwd=SITE, check=False)

    log("✅ 全部完成")
    log("   长图: %s" % LONG_IMG)
    log("   总耗时 %.1fs" % (time.time() - _t0))


if __name__ == "__main__":
    main()
