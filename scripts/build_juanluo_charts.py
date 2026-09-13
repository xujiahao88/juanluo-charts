# -*- coding: utf-8 -*-
"""
build_juanluo_charts.py — 卷螺大样本 → 交互站数据集（铁矿站同构）

生成与 iron-ore-charts 完全同构的 dataset（axes/charts/summary），前端直接复用
铁矿站的 index.html + assets/app.js（ECharts）。

结构：
  · 两个数据集（= 站点 tab）：螺纹 / 热卷
  · 每个数据集 40 张图 = 8 区域 × 5 指标（周产量/钢厂库存/社库/总库存/表需）
      图顺序按「区域优先」排列 —— 栅格一行 5 张，正好一个区域占一行
  · 每张图 = 季节图：x 轴 MM-DD（含 02-29），series = 年份（最近 5 年）
      配色沿用铁矿站 aubra：最老年虚线浅蓝 → 次老灰 → 上年蓝色平滑 → 当年红+圆点
  · 汇总表：列 = 区域分组(8) × 指标(5)，行 = 本期/上期/环比/同比/同比%

列偏移（基于 R1 表头，已逐列校验）：
  周产量 = 6 + idx*4 ；钢厂库存 = 7 + idx*4 ；社库 = 36 + idx
  总库存 = 44 + idx （源表自带「总库」列 = 厂库 + 社库，2026-09-08 加入）
  表需 = 52 + idx
  idx: 合计=7, 东北=0, 华北=1, 华东=2, 华南=3, 华中=4, 西北=5, 西南=6

用法：
  python build_juanluo_charts.py            # 写 data/*.json + data.js + meta.json
  python build_juanluo_charts.py --check    # 仅抽样校验
"""
import argparse
import datetime
import json
import os
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SCRIPTS = os.path.join(ROOT, 'scripts')
SRC = r'C:\Users\Administrator\Nutstore\1\小目标\卷螺大样本.xlsm'

if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)
from datasets_merge import merge_datasets  # noqa: E402  (共享 tab 合并器)

# 展示顺序：2026-09-08 用户要求「东北和总计换一下位置」→ 总计(合计)置首，东北置末。
# ⚠️ 元组第二项 idx 是「源表列块序号」，与展示顺序无关；换展示顺序时不要动 idx。
REGIONS = [('合计', 7), ('华北', 1), ('华东', 2), ('华南', 3),
           ('华中', 4), ('西北', 5), ('西南', 6), ('东北', 0)]
REGION_ORDER = [r for r, _ in REGIONS]

METRICS = ['周产量', '钢厂库存', '社库', '总库存', '表需']

SHEETS = [('螺纹', 'juanluo_luowen', '【手抄】螺纹大样本'),
          ('热卷', 'juanluo_rejuan', '【手抄】热卷大样本')]

UNIT = '万吨'
YOY_TOL_DAYS = 10          # 同比取「去年同周」的容忍天数
ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}
MAX_YEARS = 5               # 单数据集最多保留最近 N 年（用户 2026-09-13 要求：季节性图保留最新五年）


def col_for(metric, idx):
    if metric == '周产量':
        return 6 + idx * 4
    if metric == '钢厂库存':
        return 7 + idx * 4
    if metric == '社库':
        return 36 + idx
    if metric == '总库存':
        # 源表自带「总库」列（= 钢厂库存 + 社库），无需自行相加。
        # 已验算：东北 7.73 + 29.39 = 37.12(col44)；合计 291.03 + 658.03 = 949.06(col51)
        return 44 + idx
    if metric == '表需':
        return 52 + idx
    raise ValueError(metric)


def clean_num(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ERR_TOKENS:
            return None
        try:
            v = float(s)
        except ValueError:
            return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if v != v:
            return None
        return round(float(v), 2)
    return None


def doy_of_mmdd(mmdd):
    m, d = int(mmdd[:2]), int(mmdd[3:5])
    return (datetime.date(2025, m, d) - datetime.date(2025, 1, 1)).days


def series_style(i, n):
    """第 i 条(共 n 条)年份线的样式，对齐铁矿站 aubra 观感"""
    latest = (i == n - 1)
    prev = (i == n - 2)
    if latest:
        color = '#FF0000'
    elif prev:
        color = '#0070C0'
    elif i == n - 3:
        color = '#A5A5A5'
    else:
        color = '#9DC3E6'
    return {
        'color': color,
        'width': 1.5,
        'dash': 'dash' if i <= n - 4 else 'solid',
        'smooth': bool(prev),
        'marker': 'circle' if latest else 'none',
    }


def read_all():
    """返回 raw[page][region][metric] = [(date, year, mmdd, val), ...] 升序"""
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    raw = {}
    for page, _, sheet in SHEETS:
        ws = wb[sheet]
        per = {r: {m: [] for m in METRICS} for r in REGION_ORDER}
        for row in ws.iter_rows(values_only=True):
            if len(row) < 3:
                continue
            year = row[0]
            date = row[2]
            if isinstance(year, str) and year.isdigit():
                year = int(year)
            if not isinstance(year, (int, float)) or not isinstance(date, datetime.datetime):
                continue
            year = int(year)
            d = date.date()
            mmdd = '%02d-%02d' % (d.month, d.day)
            for reg, idx in REGIONS:
                for metric in METRICS:
                    c = col_for(metric, idx)
                    v = clean_num(row[c - 1])
                    if v is None:
                        continue
                    per[reg][metric].append((d, year, mmdd, v))
        for reg in REGION_ORDER:
            for metric in METRICS:
                per[reg][metric].sort(key=lambda x: x[0])
        raw[page] = per
    wb.close()
    return raw


def build_dataset(page, dsid, per):
    # 季节轴：该数据集内出现过的所有 MM-DD（含 02-29）
    mmdds = set()
    years = set()
    for reg in REGION_ORDER:
        for metric in METRICS:
            for d, y, mmdd, v in per[reg][metric]:
                mmdds.add(mmdd)
                years.add(y)
    # 季节轴：用「闰年 366 天」的 MM-DD 全日历轴（与铁矿站 aubra 一致），
    # 这样 x 间距才是真实日历比例、月份标签 /-01$/ 才能正常命中。
    # 周度数据只落在其中 ~52 个槽位，其余为 null，由前端 connectNulls 连线。
    axis = []
    _d = datetime.date(2024, 1, 1)
    while _d.year == 2024:
        axis.append('%02d-%02d' % (_d.month, _d.day))
        _d += datetime.timedelta(days=1)
    years_sorted = sorted(years)
    if len(years_sorted) > MAX_YEARS:
        years_sorted = years_sorted[-MAX_YEARS:]
    n = len(years_sorted)
    idx_of = {m: i for i, m in enumerate(axis)}

    charts = []
    for reg in REGION_ORDER:
        for metric in METRICS:
            # year -> {mmdd: val}
            byyear = {}
            for d, y, mmdd, v in per[reg][metric]:
                byyear.setdefault(y, {})[mmdd] = v
            series = []
            for i, y in enumerate(years_sorted):
                m = byyear.get(y, {})
                data = [m.get(mm) for mm in axis]
                st = series_style(i, n)
                series.append({
                    'name': str(y),
                    'color': st['color'],
                    'width': st['width'],
                    'dash': st['dash'],
                    'smooth': st['smooth'],
                    'marker': st['marker'],
                    'data': data,
                })
            charts.append({
                'key': '%s-%s' % (reg, metric),
                'title': '%s · %s' % (reg, metric),
                'type': 'line',
                'axis': 0,
                'group': reg,          # 前端按此字段分区块渲染（一个区域一个区块）
                'series': series,
            })

    # ---- 汇总表 ----
    columns = []
    for reg in REGION_ORDER:
        for metric in METRICS:
            columns.append({'key': '%s-%s' % (reg, metric),
                            'label': metric, 'group': reg})

    rows = {'本期': [], '上期': [], '环比': [], '同比': [], '同比%': []}
    cur_wk = prev_wk = ''
    for reg in REGION_ORDER:
        for metric in METRICS:
            seq = per[reg][metric]           # [(date, year, mmdd, val)]
            if len(seq) < 2:
                for k in rows:
                    rows[k].append(None)
                continue
            (d1, y1, m1, v1), (d2, y2, m2, v2) = seq[-1], seq[-2]
            if not cur_wk:
                cur_wk, prev_wk = m1, m2
            # 去年同期：去年内 doy 最接近 m1 的点（±10 天）
            yoy = None
            tgt = doy_of_mmdd(m1)
            best_gap = None
            for d, y, mm, v in seq:
                if y != y1 - 1:
                    continue
                gap = abs(doy_of_mmdd(mm) - tgt)
                if gap <= YOY_TOL_DAYS and (best_gap is None or gap < best_gap):
                    best_gap, yoy = gap, v
            rows['本期'].append(v1)
            rows['上期'].append(v2)
            rows['环比'].append(round(v1 - v2, 2))
            if yoy is None:
                rows['同比'].append(None)
                rows['同比%'].append(None)
            else:
                rows['同比'].append(round(v1 - yoy, 2))
                rows['同比%'].append(round((v1 - yoy) / yoy * 100, 2) if yoy else None)

    # asOf = 全数据集最新日期
    all_dates = [d for reg in REGION_ORDER for metric in METRICS
                 for (d, y, mm, v) in per[reg][metric]]
    last = max(all_dates) if all_dates else None
    as_of = '%02d-%02d' % (last.month, last.day) if last else ''

    return {
        'id': dsid,
        'name': page,
        'axes': [axis],
        'asOf': as_of,
        'charts': charts,
        'summary': {
            'columns': columns,
            'rows': rows,
            'rowOrder': ['本期', '上期', '环比', '同比', '同比%'],
            'currentWeek': cur_wk,
            'previousWeek': prev_wk,
            'unit': UNIT,
        },
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    raw = read_all()
    datasets = [build_dataset(page, dsid, raw[page]) for page, dsid, _ in SHEETS]

    if args.check:
        for ds in datasets:
            print('==', ds['id'], ds['name'], '| asOf', ds['asOf'],
                  '| charts', len(ds['charts']), '| axis', len(ds['axes'][0]))
            print('   axis head/tail:', ds['axes'][0][:3], ds['axes'][0][-3:])
            print('   years:', [s['name'] for s in ds['charts'][0]['series']])
            print('   chart titles(前8):', [c['title'] for c in ds['charts'][:8]])
            sm = ds['summary']
            print('   summary cols:', len(sm['columns']),
                  '| groups:', [c['group'] for c in sm['columns'][:5]], '...')
            print('   currentWeek/previousWeek:', sm['currentWeek'], sm['previousWeek'])
            for k in sm['rowOrder']:
                print('   ', k, sm['rows'][k][:6])
        return

    os.makedirs(DATA, exist_ok=True)
    stamp = datetime.datetime.now().isoformat(timespec='seconds')

    regen = datasets
    for ds in regen:
        p = os.path.join(DATA, ds['id'] + '.json')
        ds['updated'] = stamp
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(ds, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d axis=%d asOf=%s'
              % (ds['id'], os.path.getsize(p) / 1024, len(ds['charts']),
                 len(ds['axes'][0]), ds['asOf']))

    # ⚠️ 合并其他流水线的 tab（钢银/带钢/焊管/出港/出口/PSI…），
    #    否则本脚本写的 data.js + meta.json 只有 螺纹/热卷，会把别人的 tab 冲掉。
    #    详见 scripts/datasets_merge.py
    datasets = merge_datasets(regen)
    print('[ok] 合并后 tab：%s（共 %d）'
          % ('/'.join(d['id'] for d in datasets), len(datasets)))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_juanluo_charts.py @ ' + stamp + '\n')
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': datasets}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] 写 data.js')

    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']}
                         for d in datasets]}
    with open(os.path.join(DATA, 'meta.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] 写 meta.json datasets=%d' % len(meta['datasets']))


if __name__ == '__main__':
    main()
