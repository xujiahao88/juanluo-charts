#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_juanluo_pipeline.py — 卷螺大样本「一条龙」统一调度器
=======================================================

解析 Mysteel《螺纹热卷全样本》xlsx → 自动：
  1. 清理 Excel 孤儿进程
  2. update_juanluo_source.py 把最新周写入《卷螺大样本.xlsm》【手抄】表
     （数值列手填 + 合计社库/总库/表需 拉公式，绝不写死数）
  3. build_juanluo_charts.py 抽数 → 写 data/*.json + data.js + meta.json
  4. deploy_repo.py 推 juanluo-charts（GitHub Pages 自动重建）
  5. shot_juanluo.py Edge headless 截图螺纹 / 热卷 两张长图（带标题）
  6. send_to_wechat.py 发微信文件传输助手（默认）

用法：
  # 老宽表（一个文件两 sheet）
  python run_juanluo_pipeline.py --src "<螺纹热卷全样本.xlsx>" --date 2026-09-09
  # 新长表（两个文件：螺纹 / 热卷 各一份）
  python run_juanluo_pipeline.py --src-luowen "<螺纹钢全样本生产数据.xlsx>" --src-hot "<热卷全样本生产数据.xlsx>"
  # 零参数：自动在微信落点找最新两份长表，日期取源里最新一周
  python run_juanluo_pipeline.py --discover
  # 常用开关
  python run_juanluo_pipeline.py ... --no-wechat / --no-push / --no-shot / --dry-run

⚠️ win32com 打开《卷螺大样本.xlsm》约 5 秒；整体 < 30s。
"""
from __future__ import annotations

import argparse
import datetime
import os
import re
import subprocess
import sys
import time

PY = r"C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe"
SITE = r"C:/Users/Administrator/juanluo-charts"
SKILLS = r"C:/Users/Administrator/.workbuddy/skills"
sys.path.insert(0, os.path.join(SITE, "scripts"))
try:
    import adapt_juanluo_source as ADAPT
except Exception:                                    # 适配器缺失时退化为「仅支持老宽表」
    ADAPT = None
UPDATE_SCRIPT = os.path.join(SITE, "scripts", "update_juanluo_source.py")
ADAPT_SCRIPT = os.path.join(SITE, "scripts", "adapt_juanluo_source.py")
BUILD_SCRIPT = os.path.join(SITE, "scripts", "build_juanluo_charts.py")
SHOT_SCRIPT = os.path.join(SITE, "scripts", "shot_juanluo.py")
DEPLOY_SCRIPT = os.path.join(SITE, "deploy_repo.py")
SEND_SCRIPT = os.path.join(SITE, "..", "iron-ore-charts", "scripts", "send_to_wechat.py")
SHOT = os.path.join(SITE, "shot")
LUOWEN_IMG = os.path.join(SHOT, "卷螺_螺纹_长图.png")
REJUAN_IMG = os.path.join(SHOT, "卷螺_热卷_长图.png")

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
    """清理没有可见窗口的孤儿 EXCEL.EXE 进程（保留用户打开的）。中文 Windows tasklist 输出 GBK。"""
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


def bump_cache_version():
    """把 index.html 里所有 ?v=NN 自增，保证公网强刷（Pages CDN 对静态资源有缓存）。"""
    path = os.path.join(SITE, "index.html")
    try:
        with open(path, encoding="utf-8") as f:
            txt = f.read()
    except Exception as e:
        log(f"  ⚠️ 读 index.html 失败，跳过缓存 bump：{e}")
        return None
    m = re.search(r"\?v=(\d+)", txt)
    if not m:
        log("  ⚠️ index.html 未找到 ?v=NN，跳过缓存 bump")
        return None
    old, new = int(m.group(1)), int(m.group(1)) + 1
    txt = re.sub(r"\?v=\d+", f"?v={new}", txt)
    with open(path, "w", encoding="utf-8") as f:
        f.write(txt)
    log(f"  缓存版本 ?v={old} → ?v={new}（index.html）")
    return new


def resolve_src(args):
    """把各种输入形态统一成一个「老宽表」临时文件路径。
    支持：
      --src <老宽表>                      → 原样使用
      --src-luowen <螺纹长表> --src-hot <热卷长表> → 适配成宽表
      （都不给 / --discover）             → 自动在微信落点找最新两份长表
    """
    src, lw, ht = args.src, args.src_luowen, args.src_hot
    if src and ADAPT and ADAPT.is_legacy(src):
        log(f"   源格式：老宽表（直接使用）→ {os.path.basename(src)}")
        return src
    if src and not lw and not ht and ADAPT:
        # 单个非宽表文件：按文件名判品种
        k = ADAPT.guess_kind(src)
        if k == "luowen":
            lw = src
        elif k == "hot":
            ht = src
        else:
            raise SystemExit(f"❌ 无法识别这个源文件格式/品种：{src}")
    if args.discover or (not lw and not ht and not src):
        if ADAPT is None:
            raise SystemExit("❌ 缺少 adapt_juanluo_source.py，无法自动发现源文件")
        dl, dh = ADAPT.discover_latest()
        lw = lw or dl
        ht = ht or dh
        log(f"   自动发现：螺纹={os.path.basename(lw) if lw else '无'} | 热卷={os.path.basename(ht) if ht else '无'}")
    if not lw and not ht:
        raise SystemExit("❌ 没找到任何源文件（试 --src / --src-luowen --src-hot / --discover）")
    if ADAPT is None:
        raise SystemExit("❌ 长表源需要 adapt_juanluo_source.py")
    out = os.path.join(SHOT, "_wide_tmp.xlsx")
    os.makedirs(SHOT, exist_ok=True)
    log("   源格式：长表（两份）→ 适配成老宽表临时文件")
    run([PY, ADAPT_SCRIPT, "--out", out,
         *( ["--luowen", lw] if lw else [] ),
         *( ["--hot", ht] if ht else [] )], check=True)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", help="Mysteel 老宽表《螺纹热卷全样本》xlsx（col3=日期、col4+ 区域块）")
    ap.add_argument("--src-luowen", help="长表：《螺纹钢全样本生产数据*.xlsx》")
    ap.add_argument("--src-hot", help="长表：《热卷全样本生产数据*.xlsx》")
    ap.add_argument("--discover", action="store_true", help="自动在微信落点找最新两份长表")
    ap.add_argument("--date", help="目标周 YYYY-MM-DD（缺省=从源里取最新一周）")
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--no-shot", action="store_true")
    ap.add_argument("--no-wechat", action="store_true",
                    help="默认直接发微信；--no-wechat 才跳过")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    for p in (UPDATE_SCRIPT, BUILD_SCRIPT, SHOT_SCRIPT, DEPLOY_SCRIPT, SEND_SCRIPT, PY):
        if not os.path.exists(p):
            raise SystemExit(f"找不到必要文件: {p}")

    # 0. 归一化源文件（老宽表 / 两份长表 / 自动发现）
    log("--- 步骤 0: 解析源文件（老宽表 or 长表适配）---")
    src_path = resolve_src(args)

    # 0b. --date 缺省 → 取源里最新一周
    date_str = args.date
    if not date_str:
        import openpyxl
        wb = openpyxl.load_workbook(src_path, read_only=True, data_only=True)
        latest = None
        for sn in wb.sheetnames:
            ws = wb[sn]
            for r in ws.iter_rows(min_row=3, values_only=True):
                d = r[2] if len(r) > 2 else None
                if isinstance(d, datetime.datetime):
                    d = d.date()
                if isinstance(d, datetime.date):
                    if latest is None or d > latest:
                        latest = d
        wb.close()
        if latest is None:
            raise SystemExit("❌ 无法从源文件推断最新周，请显式 --date")
        date_str = latest.isoformat()
        log(f"   --date 缺省 → 源里最新周 = {date_str}")

    target = datetime.date(*map(int, date_str.split("-")))
    week = target.isocalendar()[1]
    luowen_title = f"卷螺大样本 · 螺纹钢 · {date_str}（第{week}周）"
    rejuan_title = f"卷螺大样本 · 热轧卷板 · {date_str}（第{week}周）"
    # 长图顶部只要标题，不带「数据来源」副标题（用户 2026-09-18 要求去掉）

    if args.dry_run:
        log("--- [--dry-run] 仅预览，不写 Excel / 不出图 / 不推送 / 不发微信 ---")
        log(f"   源文件: {src_path}")
        log(f"   将写源: {PY} {UPDATE_SCRIPT} --src {src_path} --date {date_str}")
        log(f"   将抽数: {PY} {BUILD_SCRIPT}")
        log(f"   将推:   {DEPLOY_SCRIPT} juanluo-charts <data/*.json + data.js + index.html ...>")
        log(f"   将出图: {SHOT_SCRIPT} --ds juanluo_luowen --out {LUOWEN_IMG} --title \"{luowen_title}\"")
        log(f"   将出图: {SHOT_SCRIPT} --ds juanluo_rejuan --out {REJUAN_IMG} --title \"{rejuan_title}\"")
        log(f"   将发:   {SEND_SCRIPT} --no-countdown {LUOWEN_IMG} {REJUAN_IMG}")
        return

    # 1. 清理孤儿 Excel
    log("--- 步骤 1: 清理孤儿 Excel 进程 ---")
    kill_orphan_excel()

    # 2. 更新 Excel（写【手抄】表，数值列手填 + 合计公式下拉）
    log("--- 步骤 2: 更新《卷螺大样本.xlsm》【手抄】螺纹/热卷 ---")
    rc, _ = run([PY, UPDATE_SCRIPT, "--src", src_path, "--date", date_str], check=False)
    if rc != 0:
        raise SystemExit("  ❌ 源 xlsx 更新失败，终止（检查源路径与 --date 是否在源内）")

    # 3. 抽数
    log("--- 步骤 3: 抽数 build_juanluo_charts.py ---")
    run([PY, BUILD_SCRIPT], cwd=SITE)

    # 4+5. 推送 与 出图【重叠】
    procs = []
    if args.no_push:
        log("--- 步骤 4: 跳过推送（--no-push）---")
    elif not os.environ.get("GITHUB_PAT"):
        log("  ⚠️ GITHUB_PAT 未设置，跳过推送（需 classic PAT，repo+workflow）")
    else:
        log("--- 步骤 4: 推 juanluo-charts（后台，与出图重叠）---")
        bump_cache_version()          # 先 bump 缓存版本，再连 index.html 一起推
        files = ["index.html", "assets/app.js", "assets/style.css",
                 "data/data.js", "data/meta.json",
                 "data/juanluo_luowen.json", "data/juanluo_rejuan.json",
                 "scripts/build_juanluo_charts.py", "scripts/update_juanluo_source.py",
                 "scripts/adapt_juanluo_source.py", "scripts/run_juanluo_pipeline.py", "README.md"]
        env = dict(os.environ, COMMIT_MSG=f"data: 卷螺大样本周更 ({date_str})")
        procs.append(launch([PY, DEPLOY_SCRIPT, "juanluo-charts"] + files, cwd=SITE, env=env))

    if args.no_shot:
        log("--- 步骤 5: 跳过出图（--no-shot）---")
    else:
        log("--- 步骤 5: 出螺纹/热卷两张长图（后台，与推送重叠）---")
        os.makedirs(SHOT, exist_ok=True)
        procs.append(launch([PY, SHOT_SCRIPT, "--ds", "juanluo_luowen",
                             "--out", LUOWEN_IMG, "--width", "1920", "--height", "4200",
                             "--title", luowen_title], cwd=SITE))
        procs.append(launch([PY, SHOT_SCRIPT, "--ds", "juanluo_rejuan",
                             "--out", REJUAN_IMG, "--width", "1920", "--height", "4200",
                             "--title", rejuan_title], cwd=SITE))

    if procs:
        log(f"⏳ 等待 {len(procs)} 个后台任务（推送/出图）完成...")
        for p in procs:
            out, _ = p.communicate()
            for line in (out or "").splitlines():
                if line.strip():
                    log("   " + line)
            if p.returncode != 0:
                log(f"  ⚠️ 后台任务 rc={p.returncode}")

    # 6. 发微信（默认）
    if args.no_shot:
        log("--- 步骤 6: 跳过（未出图）---")
    elif args.no_wechat:
        log("--- 步骤 6: 未发微信（--no-wechat），手动命令：---")
        if os.path.exists(LUOWEN_IMG):
            log(f"   python scripts/send_to_wechat.py --no-countdown \"{LUOWEN_IMG}\" \"{REJUAN_IMG}\"")
    else:
        log("--- 步骤 6: 发微信（默认）---")
        if not (os.path.exists(LUOWEN_IMG) and os.path.exists(REJUAN_IMG)):
            log("  ⚠️ 找不到长图，跳过发送")
        else:
            log(f"  ⚠️ 将发 2 张图到微信「文件传输助手」，请确保已切到该会话")
            run([PY, SEND_SCRIPT, "--no-countdown", LUOWEN_IMG, REJUAN_IMG],
                cwd=SITE, check=False)

    log("✅ 全部完成")
    log(f"   总耗时 {time.time() - _t0:.1f}s")


if __name__ == "__main__":
    main()
