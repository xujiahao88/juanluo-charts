#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_gangyin_chain.py — 钢银库存「一条龙」统一调度器
====================================================

发数值 → 自动：
  1. 清理 Excel 孤儿进程
  2. update_gangyin.py 把库存 5 指标写进《钢银数据库.xlsx》F-J 列
     → K-O 公式自动算变化量 → 刷新 10 个透视表
  3. build_gangyin_charts.py 抽数 → 重生成 juanluo-charts 钢银 tab dataset
  4. deploy_repo.py 推 juanluo-charts（GitHub Pages 自动重建）
  5. shot_juanluo.py Edge headless 截图钢银 tab 长图
  6. send_to_wechat.py 发微信文件传输助手

用法：
  python run_gangyin_chain.py --date 2026-09-07 --values "541.27,251.62,78.05,126.20,997.14"
  python run_gangyin_chain.py --date 2026-09-07 --values "541.27,251.62,78.05,126.20,997.14" --no-wechat
  python run_gangyin_chain.py --date 2026-09-07 --values "541.27,251.62,78.05,126.20,997.14" --dry-run

⚠️ win32com 打开《钢银数据库.xlsx》约 5 秒（44 图表），整体 < 30s。
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time

PY = r"C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe"
SITE = r"C:/Users/Administrator/juanluo-charts"
SKILLS = r"C:/Users/Administrator/.workbuddy/skills"
UPDATE_SCRIPT = os.path.join(SKILLS, "gangyin-inventory-update", "update_gangyin.py")
DATA = os.path.join(SITE, "data")
SHOT = os.path.join(SITE, "shot")
SHOT_IMG = os.path.join(SHOT, "卷螺_钢银库存_长图.png")
DS_ID = "gangyin_stock"

_t0 = time.time()


def log(msg):
    print(f"[{time.time() - _t0:6.1f}s] {msg}", flush=True)


def run(cmd, cwd=None, env=None, check=True):
    r = subprocess.run(cmd, cwd=cwd, env=env, capture_output=True, text=True, timeout=300)
    out = (r.stdout or "") + (r.stderr or "")
    for line in out.splitlines():
        if line.strip():
            log("   " + line)
    if check and r.returncode != 0:
        raise SystemExit(f"命令失败 rc={r.returncode}: {' '.join(cmd[:4])}...")
    return r.returncode, out


def launch(cmd, cwd=None, env=None):
    log("   ↳ (bg) " + " ".join(os.path.basename(x) if x.endswith('.py') else x for x in cmd[:3]) + " ...")
    return subprocess.Popen(cmd, cwd=cwd, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)


def kill_orphan_excel():
    """清理没有可见窗口的孤儿 EXCEL.EXE 进程（保留用户打开的）

    注意：Windows 中文系统 tasklist 输出是 GBK；用 errors='ignore' 避免解码异常。
    """
    try:
        r = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=False, timeout=15)
        pids = []
        for raw in (r.stdout or b"").splitlines():
            try:
                line = raw.decode("gbk", errors="ignore").strip().strip('"')
            except Exception:
                continue
            if not line or "EXCEL.EXE" not in line.upper():
                continue
            parts = [p.strip().strip('"') for p in line.split(",")]
            if len(parts) < 2:
                continue
            pids.append(parts[1])
    except Exception as e:
        log(f"  tasklist 失败：{e}")
        return
    if not pids:
        log("  无 Excel 进程")
        return
    killed = 0
    for pid in pids:
        try:
            v = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH", "/V"],
                capture_output=True, timeout=10)
            txt = (v.stdout or b"").decode("gbk", errors="ignore")
            if "可见" in txt or "Visible" in txt:
                log(f"  ⚠️ PID {pid} 有可见窗口，跳过（可能是用户打开的 Excel）")
                continue
        except Exception:
            pass
        try:
            subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True, timeout=10)
            killed += 1
        except Exception:
            pass
    log(f"  清理孤儿 Excel: {killed} 个" if killed else "  无需清理（用户已开 Excel）")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", required=True, help="最新日期，如 2026-09-07")
    ap.add_argument("--values", required=True, help="5 个库存值，逗号分隔：建材,热卷,中厚板,冷轧涂镀,合计")
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--no-shot", action="store_true")
    ap.add_argument("--no-wechat", action="store_true",
                    help="默认直接发微信；--no-wechat 才跳过（仅出图并打印手动命令）")
    ap.add_argument("--all", action="store_true",
                    help="连卷螺源也一起重抽（卷螺 xlsm 更新时用；默认只重抽钢银，快 ~0.7s）")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(UPDATE_SCRIPT):
        raise SystemExit(f"找不到 update_gangyin.py: {UPDATE_SCRIPT}")
    if not os.path.exists(PY):
        raise SystemExit(f"找不到 Python: {PY}")

    if args.dry_run:
        log("--- [--dry-run] 仅预览，不写 Excel / 不出图 / 不推送 ---")
        log(f"   将执行: {PY} {UPDATE_SCRIPT} --date {args.date} --values {args.values} --dry-run")
        log(f"   将抽数:  python {SITE}/scripts/build_gangyin_charts.py")
        log(f"   将推:    {SITE}/deploy_repo.py juanluo-charts data/*")
        log(f"   将出图:  {SITE}/scripts/shot_juanluo.py --ds {DS_ID}")
        log(f"   将发:    {SITE}/../iron-ore-charts/scripts/send_to_wechat.py {SHOT_IMG}")
        return

    # 1. 清理孤儿 Excel
    log("--- 步骤 1: 清理孤儿 Excel 进程 ---")
    kill_orphan_excel()

    # 2. 更新 Excel（gangyin-inventory-update skill）
    log("--- 步骤 2: 更新《钢银数据库.xlsx》F-J + 刷 10 个透视表 ---")
    rc, _ = run([PY, UPDATE_SCRIPT, "--date", args.date, "--values", args.values], check=False)
    if rc != 0:
        log(f"  ⚠️ Excel 更新异常，仍继续抽数（用已有数据）")

    # 3. 抽数（默认只重抽钢银；--all 时连卷螺一起，慢 ~0.7s）
    log("--- 步骤 3: 抽数" + ("（全量：含卷螺）" if args.all else "（快速：仅钢银，卷螺复用磁盘）") + "---")
    build_cmd = [PY, os.path.join(SITE, "scripts", "build_gangyin_charts.py")]
    if args.all:
        build_cmd.append("--all")
    run(build_cmd, cwd=SITE)

    # 4+5. 推送 与 出图【重叠】
    procs = []
    if args.no_push:
        log("--- 步骤 4: 跳过推送（--no-push）---")
    elif not os.environ.get("GITHUB_PAT"):
        log("  ⚠️ GITHUB_PAT 未设置，跳过推送（需 classic PAT，repo+workflow）")
    else:
        log("--- 步骤 4: 推 juanluo-charts（后台，与出图重叠）---")
        files = ["data/data.js"]
        for f in sorted(os.listdir(DATA)):
            if f.endswith(".json"):
                files.append(os.path.join("data", f))
        env = dict(os.environ, COMMIT_MSG=f"data: 钢银库存周更 ({args.date})")
        procs.append(launch([PY, os.path.join(SITE, "deploy_repo.py"), "juanluo-charts"] + files,
                            cwd=SITE, env=env))

    if args.no_shot:
        log("--- 步骤 5: 跳过出图（--no-shot）---")
    else:
        log("--- 步骤 5: 出钢银长图（后台，与推送重叠）---")
        os.makedirs(SHOT, exist_ok=True)
        procs.append(launch([PY, os.path.join(SITE, "scripts", "shot_juanluo.py"),
                             "--ds", DS_ID, "--out", SHOT_IMG], cwd=SITE))

    if procs:
        log(f"⏳ 等待 {len(procs)} 个后台任务（推送/出图）完成...")
        for p in procs:
            out, _ = p.communicate()
            for line in (out or "").splitlines():
                if line.strip():
                    log("   " + line)
            if p.returncode != 0:
                log(f"  ⚠️ 后台任务 rc={p.returncode}")

    # 6. 发微信
    if args.no_shot:
        log("--- 步骤 6: 跳过（未出图）---")
    elif args.no_wechat:
        log("--- 步骤 6: 未发微信（--no-wechat），手动命令：---")
        if os.path.exists(SHOT_IMG):
            log(f"   python scripts/send_to_wechat.py \"{SHOT_IMG}\"")
    else:
        log("--- 步骤 6: 发微信（默认）---")
        if not os.path.exists(SHOT_IMG):
            log("  ⚠️ 找不到长图，跳过发送")
        else:
            log(f"  ⚠️ 将发 1 张图到微信「文件传输助手」，请确保已切到该会话")
            send_script = os.path.join(SITE, "..", "iron-ore-charts", "scripts", "send_to_wechat.py")
            run([PY, send_script, "--no-countdown", SHOT_IMG], cwd=SITE, check=False)

    log("✅ 全部完成")
    log(f"   总耗时 {time.time() - _t0:.1f}s")


if __name__ == "__main__":
    main()
