# -*- coding: utf-8 -*-
"""
update_juanluo_source.py — 把 Mysteel《螺纹热卷全样本》最新周写入《卷螺大样本.xlsm》

安全约定：
  · 目标 xlsm 含宏 + 多透视表 → 必须用 win32com 真实 Excel 写回（openpyxl 会丢宏/透视）。
  · 目标在坚果云同步目录 → 复制到本地临时副本改完 Save，再 copy2 回原路径（带重试）。
  · 只动【手抄】螺纹大样本 / 【手抄】热卷大样本。

列映射（已用 ws.cell 显式核对，1-indexed）：
  新 xlsx 区域块起始列 = 4 + 4*i（i=0..6 为 东北..西南；i=7 为 合计，合计块 7 列）
    每区：开工率,产能利用率,周产量,钢厂库存
    合计额外：社会库存=36, 库存(总库)=37, 表需=38
  源 sheet 区域块起始列 = 3 + 4*i
    每区：开工率,产能利用率,周产量,钢厂库存
    合计额外：社库=42, 总库=50, 表需=58
  ⚠️ 新 xlsx 不含分区域社库/总库/表需 → 这些列对新区留空（待社库周报补全）。
  ⚠️ 月日列(col2)强制文本，避免 Excel 把 "09-09" 自动转日期序列。

用法：
  python update_juanluo_source.py --src "<xlsx>" [--date 2026-09-09] [--dry-run]
"""
from __future__ import annotations

import argparse
import datetime
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl

SRC_XLSM = r"C:\Users\Administrator\Nutstore\1\小目标\卷螺大样本.xlsm"

SHEETS = [("螺纹全样本", "【手抄】螺纹大样本"),
          ("热卷全样本", "【手抄】热卷大样本")]

# 区域顺序（与新 xlsx / 源一致）：东北,华北,华东,华南,华中,西北,西南,合计
REGIONS = ["东北", "华北", "华东", "华南", "华中", "西北", "西南", "合计"]
# 新 xlsx 每区 4 列起始：开工率/产能利用率/周产量/钢厂库存
NEW_REGION_BASE = 4          # 东北开工率 = col4
SRC_REGION_BASE = 4          # 源 东北开工率 = col4（win32com 读表头：col3=日期,col4=东北开工率）
# ⚠️ 合计 社库/总库/表需（源 col43/51/59）= 公式列，绝不可当数值写入！
#    它们在主库是公式（合计社库=INDEX 螺纹热卷社库; 合计总库=钢厂+社库; 合计表需=周产+上周总库-总库），
#    写入死数会导致序列断裂/伪 spike。本脚本只填 cols4..35（周产/钢厂库存/开工率/产能利用率），
#    合计 社库/总库/表需 由下方 FORMULA 模板显式拉公式。

# 每个源 sheet 的「合计」公式模板（col43 合计社库 INDEX；col51/col59 结构化引用自动指本行）
SHEET_FORMULA = {
    "【手抄】螺纹大样本": {
        "P": "螺纹",
        "col43": "=INDEX(螺纹热卷社库!I:I,MATCH([@日期]+2,螺纹热卷社库!$A:$A,0),1)",
    },
    "【手抄】热卷大样本": {
        "P": "热卷",
        "col43": "=INDEX(螺纹热卷社库!W:W,MATCH([@日期]+2,螺纹热卷社库!$O:$O,0),1)",
    },
}

ERR_TOKENS = {"#N/A", "#N/A!", "#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", ""}


def kill_orphan_excel():
    try:
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                           capture_output=True, timeout=15)
        for raw in (r.stdout or b"").splitlines():
            line = raw.decode("gbk", errors="ignore").strip().strip('"')
            if "EXCEL.EXE" not in line.upper():
                continue
            pid = line.split(",")[1].strip().strip('"')
            try:
                v = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH", "/V"],
                                   capture_output=True, timeout=10)
                if "可见" in (v.stdout or b"").decode("gbk", errors="ignore"):
                    continue
            except Exception:
                pass
            try:
                subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True, timeout=10)
            except Exception:
                pass
    except Exception:
        pass


def clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ERR_TOKENS:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if v != v else float(v)
    return None


def find_target_row_new(ws, target):
    """在 新 xlsx 找目标日期行号（col3=日期）"""
    for r in range(3, ws.max_row + 1):
        d = ws.cell(row=r, column=3).value
        if isinstance(d, datetime.datetime) and d.date() == target:
            return r
        if isinstance(d, datetime.date) and d == target:
            return r
    return None


def find_target_row_src(ws, target):
    """找源里目标周的行号：优先命中日期；其次命中「同年但日期缺失」的不完整行（补写）。"""
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    hit = None
    for r in range(2, last + 1):
        d = ws.Cells(r, 3).Value
        if isinstance(d, datetime.datetime) and d.date() == target:
            return r
        if isinstance(d, datetime.date) and d == target:
            return r
        if ws.Cells(r, 1).Value == target.year and (d is None):
            hit = r
    return hit  # 可能 None → 调用方追加


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--date", default="2026-09-09", help="目标周 YYYY-MM-DD")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    target = datetime.date(*map(int, args.date.split("-")))
    if not os.path.exists(args.src) or not os.path.exists(SRC_XLSM):
        raise SystemExit("源文件或 xlsm 缺失")

    nwb = openpyxl.load_workbook(args.src, data_only=True)

    # 收集每个 sheet 要写入的 (src_col, value) 列表
    plan = {}
    for nsheet, _ in SHEETS:
        nws = nwb[nsheet]
        nr = find_target_row_new(nws, target)
        if nr is None:
            print(f"[跳过] {nsheet} 找不到 {target}")
            plan[nsheet] = None
            continue
        cells = {}
        for i, reg in enumerate(REGIONS):
            nb = NEW_REGION_BASE + 4 * i
            sb = SRC_REGION_BASE + 4 * i
            for a in range(4):
                cells[sb + a] = clean(nws.cell(row=nr, column=nb + a).value)
        # ⚠️ 注意：合计 社库/总库/表需(col43/51/59) 是公式列，不在此写入（见 SHEET_FORMULA）
        plan[nsheet] = (nr, cells)

    print(f"[计划] 目标周 {target}")
    for nsheet, _ in SHEETS:
        p = plan.get(nsheet)
        if not p:
            continue
        _, cells = p
        print(f"  {nsheet}: 合计周产={cells[34]} 合计厂库={cells[35]} "
              f"（合计社库/总库/表需=公式下拉，不写值）")

    if args.dry_run:
        print("[dry-run] 结束")
        return

    kill_orphan_excel()
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False).name
    shutil.copy2(SRC_XLSM, tmp)

    import win32com.client as w32
    xl = w32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(tmp)
        for nsheet, ssrc in SHEETS:
            p = plan.get(nsheet)
            if not p:
                continue
            _, cells = p
            ws = wb.Sheets(ssrc)
            r = find_target_row_src(ws, target)
            if r is None:
                r = ws.Cells(ws.Rows.Count, 3).End(-4162).Row + 1
            ws.Cells(r, 1).Value = target.year
            ws.Cells(r, 2).NumberFormat = "@"          # 强制文本，防 "09-09" 转日期
            ws.Cells(r, 2).Value = target.strftime("%m-%d")
            # ⚠️ 日期列必须写 Excel 序列号：win32com 写 datetime 会被时区挪 8 小时
            #    （曾把 2026-09-16 写成 2026-09-15 16:00 → 站点 asOf 少一天、社库 MATCH 失配）
            ws.Cells(r, 3).NumberFormat = "yyyy-mm-dd"
            ws.Cells(r, 3).Value2 = float((target - datetime.date(1899, 12, 30)).days)
            # 仅写数值列 cols4..35（周产/钢厂库存/开工率/产能利用率；含 合计 周产/厂库）
            for c, v in cells.items():
                ws.Cells(r, c).Value = v
            # 合计 社库/总库/表需 = 公式下拉（绝写死数，否则序列断裂/伪 spike）
            cfg = SHEET_FORMULA[ssrc]
            P = cfg["P"]
            ws.Cells(r, 43).Formula = cfg["col43"]
            ws.Cells(r, 51).Formula = f"=[@[{P}-合计-钢厂库存]]+[@[{P}-合计-社库]]"
            ws.Cells(r, 59).Formula = f"=[@[{P}-合计-周产量]]+AY{r-1}-[@[{P}-合计-总库]]"
            print(f"  ✅ 写入 {ssrc} 行{r} ({target})；合计社库/总库/表需已拉公式")
        wb.Save()
        wb.Close()
    finally:
        xl.Quit()

    for attempt in range(5):
        try:
            shutil.copy2(tmp, SRC_XLSM)
            print(f"[ok] 复制回: {SRC_XLSM}")
            break
        except Exception as e:
            if attempt == 4:
                raise SystemExit(f"复制回失败: {e}（临时副本: {tmp}）")
            import time
            time.sleep(1.5)
    try:
        os.remove(tmp)
    except Exception:
        pass
    print("✅ 源 xlsm 更新完成")


if __name__ == "__main__":
    main()
