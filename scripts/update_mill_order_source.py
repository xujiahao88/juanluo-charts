#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_mill_order_source.py — 把最新一天的钢厂接单数据写进《钢厂日接单量统计.xlsx》「接单」

数据来源：用户每天发来的「唐山钢厂当日接单量(勿转)」截图
（8 家钢厂接单量 + 平均利润率 + 合计成交量）。本脚本只负责「写源」：
按**表头名**定位列、按**日期**定位行（命中覆盖 / 未命中追加），顺带把
**合并口径**（纵横 + 中铁 → 「纵横+中铁」）维护好 —— 与用户 2026-09-16 的选择一致。

用法：
  python update_mill_order_source.py --json '{"date":"2026-09-16","mills":{"纵横+中铁":3.1,"安丰":1.5,"燕钢":1.4,"瑞丰":2.5,"新东海":2.5,"东华":0.1,"日照":4.1},"profit":-40,"total":15.2}'
  python update_mill_order_source.py --json-file "C:/path/row.json"
  python update_mill_order_source.py --json '...' --no-merge      # 不维护合并口径
  python update_mill_order_source.py --json '...' --no-backup     # 不备份（默认会备份）

JSON 字段：
  date    目标日期 YYYY-MM-DD（必填）
  mills   {钢厂名: 接单量}；名字可用「日照」= 「日钢」、「纵横」=「纵横+中铁」
  profit  平均利润率（可空）
  total   总接单 / 合计成交量（可空 → 各钢厂求和）
  rate    接单率（可空 → total / 日产合计）

⚠️ 关键坑（血泪）：
  1) 目标文件在**坚果云同步目录** → 直接 Open + Save 会 C 层崩溃：
     必须先复制到本地临时副本 → COM 改完保存 → shutil.copy2 复制回原路径（带重试）。
  2) win32com 写 datetime 会被**时区转换**（读回成前一天 16:00，日期错位一天）→
     必须写 **Excel 序列号** `.Value2 = (date - 1899-12-30).days`，并沿用上一行 NumberFormat。
  3) 「总接单(L)」「接单率(M)」是**公式列** → 追加行时写公式（=SUM(C:J)、=L/日产合计），不写死数。
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import shutil
import tempfile
import time

import win32com.client as win32

SRC = r'C:\Users\Administrator\Nutstore\1\小目标\钢厂日接单量统计.xlsx'
SHEET = '接单'
EPOCH = datetime.date(1899, 12, 30)     # Excel 序列号基准
XL_UP = -4162
XL_PASTE_FORMATS = -4122

# 图里/习惯叫法 → 表头名
ALIAS = {
    '日照': '日钢', '日钢': '日钢',
    '纵横': '纵横+中铁', '纵横+中铁': '纵横+中铁',
    '东海': '新东海', '新东海': '新东海',
}


def log(m):
    print(m, flush=True)


def norm_mill(name):
    n = str(name).strip()
    return ALIAS.get(n, n)


# ---------- COM 小工具 ----------
def find_header_row(ws):
    """返回表头行号（B 列含「日期」）。"""
    last = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row
    for r in range(1, min(last, 12) + 1):
        v = ws.Cells(r, 2).Value
        if v is not None and '日期' in str(v):
            return r
    raise SystemExit('未找到表头行（B 列含「日期」）')


def find_prod_row(ws):
    last = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row
    for r in range(1, min(last, 12) + 1):
        v = ws.Cells(r, 2).Value
        if v is not None and '日产' in str(v):
            return r
    raise SystemExit('未找到日产行（B 列含「日产」）')


def header_map(ws, hr, maxcol=17):
    """返回 {表头名: 列号}。"""
    out = {}
    for c in range(2, maxcol + 1):
        v = ws.Cells(hr, c).Value
        if v is not None and str(v).strip() and str(v).strip() not in out:
            out[str(v).strip()] = c
    return out


def rate_denominator(ws, hr, hmap):
    """从已有「接单率」公式里抠出分母（日产合计），默认 21.6。"""
    c = hmap.get('接单率')
    if c:
        for r in range(hr + 1, hr + 30):
            f = ws.Cells(r, c).Formula
            if isinstance(f, str) and f.startswith('='):
                m = re.search(r'/\s*([0-9.]+)', f)
                if m:
                    return float(m.group(1))
    return 21.6


def merge_zongheng_zhongtie(ws):
    """把 纵横、中铁 合并为「纵横+中铁」（幂等）。返回改动行数。"""
    hr = find_header_row(ws)
    pr = find_prod_row(ws)
    hmap = header_map(ws, hr)
    c_z = hmap.get('纵横+中铁') or hmap.get('纵横')
    c_t = hmap.get('中铁')
    if not c_z:
        log('  [merge] 未找到「纵横」列，跳过')
        return 0
    # 表头
    if str(ws.Cells(hr, c_z).Value).strip() != '纵横+中铁':
        ws.Cells(hr, c_z).Value = '纵横+中铁'
    if not c_t or c_t == c_z:
        log('  [merge] 无独立「中铁」列，表头已就位')
        return 0
    changed = 0
    # 日产行
    z0 = ws.Cells(pr, c_z).Value or 0
    t0 = ws.Cells(pr, c_t).Value or 0
    if t0:
        ws.Cells(pr, c_z).Value = round(z0 + t0, 4)
        ws.Cells(pr, c_t).Value = 0
        changed += 1
    last = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row
    rng = ws.Range(ws.Cells(hr + 1, c_z), ws.Cells(last, c_t)).Value
    # 逐行：C = C + T; T = 0（注意 c_z..c_t 之间可能还有别的列，只动首尾两列）
    for i in range(hr + 1, last + 1):
        z = ws.Cells(i, c_z).Value
        t = ws.Cells(i, c_t).Value
        if z is None and t is None:
            continue
        if t:
            ws.Cells(i, c_z).Value = round((z or 0) + (t or 0), 4)
            ws.Cells(i, c_t).Value = 0
            changed += 1
    log('  [merge] 纵横+中铁 合并完成，改动 %d 处' % changed)
    return changed


def find_or_append_row(ws, hmap, target_date):
    """按日期找行；找不到则在末尾追加（复制上一行格式）。返回行号。"""
    hr = find_header_row(ws)
    last = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row
    tgt = float((target_date - EPOCH).days)
    for r in range(hr + 1, last + 1):
        v = ws.Cells(r, 2).Value2
        if isinstance(v, (int, float)) and abs(float(v) - tgt) < 0.5:
            return r, False
    nr = last + 1
    ws.Rows(last).Copy()
    ws.Rows(nr).PasteSpecial(Paste=XL_PASTE_FORMATS)
    ws.Application.CutCopyMode = False
    ws.Cells(nr, 2).NumberFormat = ws.Cells(last, 2).NumberFormat
    return nr, True


def write_row(ws, hmap, row, payload, denom):
    """把 payload 写进 row。"""
    hr = find_header_row(ws)
    pr = find_prod_row(ws)
    # 日期（Excel 序列号，防时区错位）
    d = datetime.date(*map(int, payload['date'].split('-')))
    ws.Cells(row, 2).Value2 = float((d - EPOCH).days)
    if not ws.Cells(row, 2).NumberFormat or ws.Cells(row, 2).NumberFormat == 'General':
        ws.Cells(row, 2).NumberFormat = 'yyyy/mm/dd'

    # 钢厂
    mills = payload.get('mills') or {}
    written, missing = {}, []
    for k, v in mills.items():
        nm = norm_mill(k)
        c = hmap.get(nm)
        if c is None:
            missing.append(k)
            continue
        ws.Cells(row, c).Value = float(v)
        written[nm] = float(v)
    if missing:
        log('  ⚠️ 表头里没有这些钢厂，已跳过：%s' % '、'.join(missing))

    # 总接单（公式优先），接单率
    c_tot = hmap.get('总接单')
    c_rate = hmap.get('接单率')
    total = payload.get('total')
    if total is None:
        total = round(sum(written.values()), 4)
    c_profit = hmap.get('平均利润率')
    if payload.get('profit') is not None and c_profit:
        ws.Cells(row, c_profit).Value = float(payload['profit'])
    if c_tot:
        ws.Cells(row, c_tot).Formula = '=SUM(C%d:J%d)' % (row, row)
    if c_rate:
        ws.Cells(row, c_rate).Formula = '=L%d/%s' % (row, denom)
    log('  [row %d] 日期=%s 各厂=%s 合计=%s 率=合计/%s' %
        (row, payload['date'], written, total, denom))
    return total


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument('--json', help='行数据 JSON（inline）')
    g.add_argument('--json-file', help='行数据 JSON 文件路径')
    ap.add_argument('--src', default=SRC)
    ap.add_argument('--no-merge', action='store_true', help='不维护「纵横+中铁」合并口径')
    ap.add_argument('--no-backup', action='store_true')
    ap.add_argument('--keep-temp', action='store_true', help='保留临时副本（调试）')
    args = ap.parse_args()

    payload = json.loads(args.json) if args.json else json.load(open(args.json_file, encoding='utf-8'))
    if not payload.get('date'):
        raise SystemExit('payload 缺 date')

    if not os.path.exists(args.src):
        raise SystemExit('源文件不存在: %s' % args.src)

    # 备份
    if not args.no_backup:
        bak = os.path.join(os.path.dirname(args.src),
                           '%s.备份_%s.xlsx' % (os.path.splitext(os.path.basename(args.src))[0],
                                                datetime.datetime.now().strftime('%Y%m%d_%H%M%S')))
        shutil.copy2(args.src, bak)
        log('[bak] %s' % os.path.basename(bak))

    # 复制到本地临时副本（坚果云直存会崩）
    tmpdir = tempfile.mkdtemp(prefix='millorder_')
    tmp = os.path.join(tmpdir, os.path.basename(args.src))
    shutil.copy2(args.src, tmp)
    log('[tmp] %s' % tmp)

    app = win32.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(tmp)
        ws = wb.Worksheets(SHEET)
        if not args.no_merge:
            log('--- 维护合并口径（纵横+中铁）---')
            merge_zongheng_zhongtie(ws)
        hr = find_header_row(ws)
        hmap = header_map(ws, hr)
        denom = rate_denominator(ws, hr, hmap)
        log('[cols] %s | 接单率分母=%s' % (hmap, denom))
        row, appended = find_or_append_row(ws, hmap, datetime.date(*map(int, payload['date'].split('-'))))
        log('--- 目标行 %d（%s）---' % (row, '追加' if appended else '覆盖'))
        write_row(ws, hmap, row, payload, denom)
        app.CalculateFullRebuild()
        wb.Save()
        wb.Close(True)
    finally:
        app.Quit()

    # 复制回原路径（坚果云锁，带重试）
    last_err = None
    for i in range(6):
        try:
            shutil.copy2(tmp, args.src)
            last_err = None
            break
        except Exception as e:                     # noqa
            last_err = e
            time.sleep(1.2)
    if last_err:
        raise SystemExit('复制回原路径失败（坚果云占用？）：%s' % last_err)
    log('[ok] 已写回 %s' % args.src)

    if not args.keep_temp:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
    log('DONE')


if __name__ == '__main__':
    main()
