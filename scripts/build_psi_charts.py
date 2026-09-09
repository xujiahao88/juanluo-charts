# -*- coding: utf-8 -*-
"""
build_psi_charts.py — PSI 中联钢排产 → 卷螺站(juanluo-charts) 数据集

数据源：《PSI 中联钢排产.xlsx》的「计算-中钢联」sheet（月度长表）
  注：「汇总」sheet 只有 3 个标题（商品卷排产 / 分地区排产 / 分品种排产），
      数据都在嵌入图表里，openpyxl 读不到 —— 真正的数据源是「计算-中钢联」。

  列（1-based）：
    1日期 2合计排产 3华东 4东北 5华北 6中南 7西部
    8出口排产 9出口排产占比
    10华东出口排产 11山东出口排产 12华北出口排产 13内陆钢厂出口排产
    14普低花 15外销C料 16其他品种钢 17普低花占比 18外销C料占比 19其他品种钢占比

产出：与卷螺站同构的 dataset
  · **月度数据** → x 轴用 12 个月（'01'..'12'），不是 366 天日历轴
  · 10 张图 = 2 商品卷 + 5 分地区 + 3 分品种，按「汇总」sheet 的 3 个区块分组
  · 年份 2022–2026
  · 只取绝对量指标（万吨），占比类是小数(0.05=5%)量纲不同，不入图

用法：
  python build_psi_charts.py            # 生成 PSI dataset + 合并已有 tab → 写 data.js/meta.json
  python build_psi_charts.py --check    # 仅抽样校验
"""
import argparse
import datetime
import json
import os
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\小目标\PSI 中联钢排产.xlsx'
SHEET = '计算-中钢联'

DS_ID = 'psi_plan'
DS_NAME = '热卷排产（中联金）'
UNIT = '万吨'

YEAR_FROM, YEAR_TO = 2022, 2026

# (指标名, 源表列号1-based, group) —— group 对齐「汇总」sheet 的三个区块标题
METRICS = [
    ('合计排产', 2, '商品卷排产'),
    ('出口排产', 8, '商品卷排产'),
    ('华东', 3, '分地区排产'),
    ('东北', 4, '分地区排产'),
    ('华北', 5, '分地区排产'),
    ('中南', 6, '分地区排产'),
    ('西部', 7, '分地区排产'),
    ('普低花', 14, '分品种排产'),
    ('外销C料', 15, '分品种排产'),
    ('其他品种钢', 16, '分品种排产'),
]

# 网站 tab 顺序：螺纹 → 热卷 → 钢银库存 → PSI排产
OTHER_IDS = ['juanluo_luowen', 'juanluo_rejuan', 'gangyin_stock']

YOY_TOL_MONTHS = 1
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


def mon_of(s):
    """'1月' / '01' → 1（提取月份数字，用于同月比对）"""
    return int(''.join(ch for ch in str(s) if ch.isdigit()))


def read_all():
    """返回 per[metric] = [(date, year, month_str, val), ...] 升序"""
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws = wb[SHEET]
    per = {m: [] for m, _, _ in METRICS}
    for row in ws.iter_rows(values_only=True):
        if len(row) < 16:
            continue
        d = row[0]
        if not isinstance(d, datetime.datetime):
            continue
        y = d.year
        if not (YEAR_FROM <= y <= YEAR_TO):
            continue
        mm = '%d月' % d.month       # 与 axis 保持一致（'1月'..'12月'）
        for m, c, _ in METRICS:
            v = clean_num(row[c - 1])
            if v is None:
                continue
            per[m].append((d.date(), y, mm, v))
    for m in per:
        per[m].sort(key=lambda x: x[0])
    wb.close()
    return per


def build_dataset(per):
    # 月度轴：'1月' .. '12月'（月度数据用月份轴，比 366 天日历轴更贴合）
    axis = ['%d月' % m for m in range(1, 13)]

    years = set()
    for m, _, _ in METRICS:
        for _, y, _, _ in per[m]:
            years.add(y)
    years_sorted = sorted(y for y in years if YEAR_FROM <= y <= YEAR_TO)
    n = len(years_sorted)

    charts = []
    for m, _, grp in METRICS:
        byyear = {}
        for d, y, mm, v in per[m]:
            byyear.setdefault(y, {})[mm] = v
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
        # 同比：去年同月（±1 个月容差）
        yoy = None
        best = None
        for d, y, mm, v in seq:
            if y != y1 - 1:
                continue
            gap = abs(mon_of(mm) - mon_of(m1))
            if gap <= YOY_TOL_MONTHS and (best is None or gap < best):
                best, yoy = gap, v
        rows['本期'].append(v1)
        rows['上期'].append(v2)
        rows['环比'].append(round(v1 - v2, 2))
        if yoy is None:
            rows['同比'].append(None)
            rows['同比%'].append(None)
        else:
            rows['同比'].append(round(v1 - yoy, 2))
            rows['同比%'].append(round((v1 - yoy) / yoy * 100, 2) if yoy else None)

    all_dates = [d for m, _, _ in METRICS for (d, y, mm, v) in per[m]]
    last = max(all_dates) if all_dates else None
    as_of = '%d月' % last.month if last else ''

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


def load_cached_others():
    """从磁盘 data/*.json 复用其他 tab 的 dataset（不重解析源 Excel，
    也避免把已有 tab 从 data.js 里弄丢）"""
    out = []
    for dsid in OTHER_IDS:
        p = os.path.join(DATA, dsid + '.json')
        if not os.path.exists(p):
            print('[warn] 缺少 %s.json，该 tab 本次不会出现在 data.js' % dsid)
            continue
        try:
            with open(p, encoding='utf-8') as f:
                d = json.load(f)
            if isinstance(d, dict) and d.get('id') == dsid:
                out.append(d)
        except Exception as e:
            print('[warn] 读 %s.json 失败：%s' % (dsid, e))
    return out


def datasets_same(new_list):
    p = os.path.join(DATA, 'data.js')
    if not os.path.exists(p):
        return False
    try:
        with open(p, encoding='utf-8') as f:
            txt = f.read()
        old = json.loads(txt.split('window.CHART_DATA = ', 1)[1].rstrip().rstrip(';'))
    except Exception:
        return False
    old_list = old.get('datasets', [])
    if len(old_list) != len(new_list):
        return False
    for a, b in zip(new_list, old_list):
        x = {k: v for k, v in a.items() if k != 'updated'}
        y = {k: v for k, v in b.items() if k != 'updated'}
        if x != y:
            return False
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    per = read_all()
    ds = build_dataset(per)

    if args.check:
        print('==', ds['id'], ds['name'], '| asOf', ds['asOf'],
              '| charts', len(ds['charts']), '| axis', ds['axes'][0])
        print('   years:', [s['name'] for s in ds['charts'][0]['series']])
        print('   groups:', [(c['group'], c['title']) for c in ds['charts']])
        sm = ds['summary']
        print('   本期/上期:', sm['currentWeek'], sm['previousWeek'])
        for k in sm['rowOrder']:
            print('   ', k, sm['rows'][k])
        for m, _, _ in METRICS:
            seq = per[m]
            print('   %-8s n=%d 最新 %s = %s' %
                  (m, len(seq), seq[-1][0] if seq else '-', seq[-1][3] if seq else '-'))
        return

    os.makedirs(DATA, exist_ok=True)
    stamp = datetime.datetime.now().isoformat(timespec='seconds')

    datasets = load_cached_others() + [ds]
    if datasets_same(datasets):
        print('[skip] 数据未变化（忽略 updated 时间戳），不写盘')
        return

    for d in datasets:
        if d['id'] == DS_ID:
            p = os.path.join(DATA, d['id'] + '.json')
            d['updated'] = stamp
            with open(p, 'w', encoding='utf-8') as f:
                json.dump(d, f, ensure_ascii=False, separators=(',', ':'))
            print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
                  % (d['id'], os.path.getsize(p) / 1024, len(d['charts']), d['asOf']))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_psi_charts.py @ ' + stamp + '\n')
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
