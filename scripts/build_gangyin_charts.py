# -*- coding: utf-8 -*-
"""
build_gangyin_charts.py — 钢银库存 → 卷螺站(juanluo-charts) 数据集

数据源：《钢银数据库.xlsx》的「钢银库存」sheet（长表）
  列：A年(公式) B周(公式) C年文本(公式) D月日(公式) E日期
      F建材 G热卷 H中厚板 I冷轧涂镀 J合计库存
      K建材库存变化 L热卷库存变化 M中厚板库存变化 N冷轧涂镀库存变化 O合计库存变化

产出：与卷螺站同构的 dataset（366 天日历轴 + 年份叠加季节图）
  · 10 张图 = 5 个库存指标 + 5 个库存变化量（顺序与 Excel「阳历图」透视表块一致）
  · 分 2 个 group：「库存」(5 张) / 「库存变化」(5 张)，各占一行
  · 年份 2022–2026
  · 汇总表：列 = 10 指标，行 = 本期/上期/环比/同比/同比%

⚠️ 注意：sheet 里 2027–2032 行是透视表预留占位（全 #N/A），必须过滤掉。

用法：
  python build_gangyin_charts.py            # 生成钢银 dataset + 重生成卷螺两个 → 写 data.js/meta.json
  python build_gangyin_charts.py --check    # 仅抽样校验
  python build_gangyin_charts.py --only     # 只写钢银 dataset（不重生成卷螺，快）
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
SRC = r'C:\Users\Administrator\Nutstore\1\小目标\钢银数据库.xlsx'
SHEET = '钢银库存'

DS_ID = 'gangyin_stock'
DS_NAME = '钢银库存'
UNIT = '万吨'

# 年份范围（用户在 Excel「阳历图」透视表里看到的就是 2022–2026）
YEAR_FROM, YEAR_TO = 2022, 2026

# (指标名, 源表列号1-based, 所属 group) —— 顺序严格对齐 Excel 阳历图透视表块顺序
METRICS = [
    ('建材', 6, '库存'),
    ('热卷', 7, '库存'),
    ('中厚板', 8, '库存'),
    ('冷轧涂镀', 9, '库存'),
    ('合计库存', 10, '库存'),
    ('建材库存变化', 11, '库存变化'),
    ('热卷库存变化', 12, '库存变化'),
    ('中厚板库存变化', 13, '库存变化'),
    ('冷轧涂镀库存变化', 14, '库存变化'),
    ('合计库存变化', 15, '库存变化'),
]

YOY_TOL_DAYS = 10
ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}


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
    """返回 per[metric] = [(date, year, mmdd, val), ...] 升序（已过滤占位年）"""
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws = wb[SHEET]
    per = {m: [] for m, _, _ in METRICS}
    for row in ws.iter_rows(values_only=True):
        if len(row) < 15:
            continue
        d = row[4]                       # E 列才是真日期（A-D 是公式列）
        if not isinstance(d, datetime.datetime):
            continue
        y = d.year
        if not (YEAR_FROM <= y <= YEAR_TO):
            continue
        mmdd = '%02d-%02d' % (d.month, d.day)
        for m, c, _ in METRICS:
            v = clean_num(row[c - 1])
            if v is None:
                continue
            per[m].append((d.date(), y, mmdd, v))
    for m in per:
        per[m].sort(key=lambda x: x[0])
    wb.close()
    return per


def build_dataset(per):
    # 366 天日历轴（闰年），与卷螺站/铁矿站一致
    axis = []
    _d = datetime.date(2024, 1, 1)
    while _d.year == 2024:
        axis.append('%02d-%02d' % (_d.month, _d.day))
        _d += datetime.timedelta(days=1)

    years = set()
    for m in per:
        for _, y, _, _ in per[m]:
            years.add(y)
    years_sorted = sorted(y for y in years if YEAR_FROM <= y <= YEAR_TO)
    n = len(years_sorted)

    charts = []
    for m, _, grp in METRICS:
        byyear = {}
        for d, y, mmdd, v in per[m]:
            byyear.setdefault(y, {})[mmdd] = v
        series = []
        for i, y in enumerate(years_sorted):
            mm = byyear.get(y, {})
            st = series_style(i, n)
            series.append({
                'name': str(y),
                'color': st['color'],
                'width': st['width'],
                'dash': st['dash'],
                'smooth': st['smooth'],
                'marker': st['marker'],
                'data': [mm.get(x) for x in axis],
            })
        charts.append({
            'key': '%s-%s' % (DS_ID, m),
            'title': m,
            'type': 'line',
            'axis': 0,
            'group': grp,
            'series': series,
        })

    # ---- 汇总表 ----
    columns = [{'key': m, 'label': m, 'group': grp} for m, _, grp in METRICS]
    rows = {'本期': [], '上期': [], '环比': [], '同比': [], '同比%': []}
    cur_wk = prev_wk = ''
    for m, _, _ in METRICS:
        seq = per[m]
        if len(seq) < 2:
            for k in rows:
                rows[k].append(None)
            continue
        (d1, y1, m1, v1), (d2, y2, m2, v2) = seq[-1], seq[-2]
        if not cur_wk:
            cur_wk, prev_wk = m1, m2
        yoy = None
        tgt = doy_of_mmdd(m1)
        best = None
        for d, y, mm, v in seq:
            if y != y1 - 1:
                continue
            gap = abs(doy_of_mmdd(mm) - tgt)
            if gap <= YOY_TOL_DAYS and (best is None or gap < best):
                best, yoy = gap, v
        rows['本期'].append(v1)
        rows['上期'].append(v2)
        rows['环比'].append(round(v1 - v2, 2))
        if yoy is None:
            rows['同比'].append(None)
            rows['同比%'].append(None)
        else:
            rows['同比'].append(round(v1 - yoy, 2))
            # 「变化量」类指标可正可负、基数可能接近 0，算百分比会严重失真
            # （例：本期 -0.4 vs 去年同期 +3.18 → -112.58%，数学正确但解读误导）
            # → 变化量只保留同比绝对差（万吨），同比% 留空。
            if '变化' in m or not yoy:
                rows['同比%'].append(None)
            else:
                rows['同比%'].append(round((v1 - yoy) / yoy * 100, 2))

    all_dates = [d for m, _, _ in METRICS for (d, y, mm, v) in per[m]]
    last = max(all_dates) if all_dates else None
    as_of = '%02d-%02d' % (last.month, last.day) if last else ''

    return {
        'id': DS_ID,
        'name': DS_NAME,
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


def juanluo_datasets():
    """重生成卷螺（螺纹/热卷）两个 dataset，保证网站三个 tab 一致"""
    if SCRIPTS not in sys.path:
        sys.path.insert(0, SCRIPTS)
    import build_juanluo_charts as bjc
    raw = bjc.read_all()
    return [bjc.build_dataset(page, dsid, raw[page]) for page, dsid, _ in bjc.SHEETS]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--only', action='store_true', help='只写钢银 dataset（不重生成卷螺）')
    args = ap.parse_args()

    per = read_all()
    ds = build_dataset(per)

    if args.check:
        print('==', ds['id'], ds['name'], '| asOf', ds['asOf'],
              '| charts', len(ds['charts']), '| axis', len(ds['axes'][0]))
        print('   years:', [s['name'] for s in ds['charts'][0]['series']])
        print('   chart titles:', [c['title'] for c in ds['charts']])
        print('   groups:', [(c['group'], c['title']) for c in ds['charts']])
        sm = ds['summary']
        print('   currentWeek/previousWeek:', sm['currentWeek'], sm['previousWeek'])
        for k in sm['rowOrder']:
            print('   ', k, sm['rows'][k])
        for m, _, _ in METRICS:
            seq = per[m]
            print('   %-12s n=%d 最新 %s = %s' %
                  (m, len(seq), seq[-1][0] if seq else '-', seq[-1][3] if seq else '-'))
        return

    os.makedirs(DATA, exist_ok=True)
    stamp = datetime.datetime.now().isoformat(timespec='seconds')

    datasets = [ds]
    if not args.only:
        try:
            datasets = juanluo_datasets() + [ds]
            print('[ok] 已重生成卷螺 螺纹/热卷 两个 dataset')
        except Exception as e:
            print('[warn] 重生成卷螺 dataset 失败（%s），本次仅写钢银' % e)

    for d in datasets:
        p = os.path.join(DATA, d['id'] + '.json')
        d['updated'] = stamp
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
              % (d['id'], os.path.getsize(p) / 1024, len(d['charts']), d['asOf']))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_gangyin_charts.py @ ' + stamp + '\n')
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': datasets}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] 写 data.js (datasets=%d)' % len(datasets))

    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']}
                         for d in datasets]}
    with open(os.path.join(DATA, 'meta.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] 写 meta.json datasets=%d' % len(meta['datasets']))


if __name__ == '__main__':
    main()
