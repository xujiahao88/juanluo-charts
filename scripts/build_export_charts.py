# -*- coding: utf-8 -*-
"""
build_export_charts.py — 「出口-分品种 / 出口-分国别」月度数据集生成器

源：《出港.xlsx》的 出口-分品种 / 出口-分国别 两个 sheet（中国海关月度出口量，万吨）。
产出 juanluo-charts 两个 dataset：
  · 出口-分品种 (export_variety)
  · 出口-分国别 (export_country)
每个 dataset 含：
  · 主要量级线图（按 2026 年 1-7 月累计量取 Top N）
  · 「当月环比增量排序」柱状图（正红负绿）
  · 「累计同比增量排序」柱状图（1-7 月，正红负绿）
  · 汇总表（本期/上期/环比/累计同比/累计增幅%）

并沿用 build_daiguan_charts.py 的合并策略：读 data/meta.json 现有顺序 →
本脚本 regenerated 的数据集就地替换、新数据集追加 → 重写 data.js + meta.json，
因此不会丢掉其他脚本生成的 tab。

用法：
  python scripts/build_export_charts.py            # 生成数据
  python scripts/build_export_charts.py --check    # 只看统计
"""
import argparse
import datetime
import json
import os
import re

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\小目标\出港.xlsx'
MAX_YEARS = 5

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
        return round(float(v), 4)
    return None


def short_name(full):
    """去掉指标前缀/后缀，留下品种或国别名。"""
    s = str(full).strip()
    for pre in ('中国海关: 钢材出口量: ', '中国海关: 钢材出口: '):
        if s.startswith(pre):
            s = s[len(pre):]
    for suf in (' : 月度', ': 月度', '：月度', ' 月度'):
        if s.endswith(suf):
            s = s[:-len(suf)]
    return s.strip()


def read_monthly(sheet):
    """返回 (series: {short: {(y,m): v}}, order: [short...], unit)"""
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    _ids = next(rows)          # 指标Id
    unit_row = next(rows)      # 单位
    _freq = next(rows)         # 频率
    unit = '万吨'
    if unit_row and len(unit_row) > 1 and unit_row[1] is not None:
        unit = str(unit_row[1]).strip()

    cols = {}
    for ci in range(1, len(header)):
        if header[ci] is not None:
            cols[ci] = short_name(header[ci])

    recs = []
    for row in rows:
        d = row[0]
        if not hasattr(d, 'year'):
            continue
        recs.append((d, row))
    recs.sort(key=lambda x: x[0])

    series = {name: {} for name in cols.values()}
    order = list(cols.values())
    for d, row in recs:
        for ci, name in cols.items():
            v = clean_num(row[ci] if ci < len(row) else None)
            if v is not None:
                series[name][(d.year, d.month)] = v
    wb.close()
    return series, order, unit


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


def cum(series, year, upto_month):
    return sum(v for (y, m), v in series.items() if y == year and m <= upto_month)


def build_dataset(sheet, dsid, name, dim_label, top_n=9, bar_half=10,
                  add_aggregates=False, bars_first=True, full_bars=True):
    series_map, order, unit = read_monthly(sheet)
    months = sorted({m for s in series_map.values() for m in s.keys()})
    if not months:
        return None
    last, prev = months[-1], months[-2]
    cur_year = last[0]

    # 分品种：额外加两条汇总线「总出口（钢材+钢坯）」「钢材（不含钢坯）」
    agg_names = set()
    if add_aggregates:
        billet_key = next((k for k in order if '钢坯' in k), None)
        total, steel = {}, {}
        for ym in months:
            vals = [series_map[k].get(ym) for k in order]
            ssum = sum(v for v in vals if v is not None)
            total[ym] = round(ssum, 2)
            if billet_key and series_map[billet_key].get(ym) is not None:
                steel[ym] = round(ssum - series_map[billet_key][ym], 2)
        g_total, g_steel = '总出口（钢材+钢坯）', '钢材（不含钢坯）'
        series_map[g_total] = total
        series_map[g_steel] = steel
        order = [g_total, g_steel] + order
        agg_names = {g_total, g_steel}

    stats = []
    for sname in order:
        s = series_map[sname]
        cur_v = s.get(last)
        prev_v = s.get(prev)
        c26 = cum(s, cur_year, last[1])
        c25 = cum(s, cur_year - 1, last[1])
        yoy = c26 - c25
        stats.append({
            'name': sname,
            'cur': cur_v,
            'prev': prev_v,
            'mom': (cur_v - prev_v) if (cur_v is not None and prev_v is not None) else None,
            'cum': c26,
            'yoy': yoy,
            'yoy_pct': (yoy / c25 * 100) if c25 else None,
            'agg': sname in agg_names,
        })

    # 主要量级：按累计量 TopN
    top = sorted(stats, key=lambda x: -(x['cum'] or 0))[:top_n]

    axis = [str(m) for m in range(1, 13)]
    lines = []
    for li, st in enumerate(top):
        byyear = {}
        for (y, m), v in series_map[st['name']].items():
            byyear.setdefault(y, {})[str(m)] = v
        years = sorted(byyear)[-MAX_YEARS:]
        n = len(years)
        ser = []
        for j, y in enumerate(years):
            so = series_style(j, n)
            ser.append({
                'name': str(y), 'color': so['color'], 'width': so['width'],
                'dash': so['dash'], 'smooth': so['smooth'], 'marker': so['marker'],
                'data': [byyear[y].get(mm) for mm in axis],
            })
        lines.append({
            'key': '%s-line-%d' % (dsid, li),
            'title': st['name'], 'type': 'line', 'axis': 0,
            'group': '主要量级（累计 Top%d）' % top_n,
            'series': ser,
        })

    def pick_bar(items, key):
        # 排序柱只排「品种/国别」本身，不带入汇总口径，避免与明细重复计算
        items = [x for x in items if x.get(key) is not None and not x.get('agg')]
        items.sort(key=lambda x: -x[key])
        if len(items) > bar_half * 2:
            items = items[:bar_half] + items[-bar_half:]
        return items

    def make_bar(title, items, key):
        cats, vals, pcts = [], [], []
        for it in items:
            cats.append(it['name'])
            vals.append(round(it[key], 2))
            pcts.append(round(it[key + '_pct'], 1) if it.get(key + '_pct') is not None else None)
        return {
            'key': '%s-bar-%s' % (dsid, key),
            'title': title, 'type': 'bar', 'axis': 0,
            'group': '环比 / 同比排序',
            'full': bool(full_bars),   # 占满整行
            'bar': {'categories': cats, 'values': vals, 'pcts': pcts, 'unit': unit},
        }

    bars = []
    bars.append(make_bar('%s · 当月环比增量排序（%d-%02d vs %d-%02d，万吨）'
                         % (dim_label, last[0], last[1], prev[0], prev[1]),
                         pick_bar([dict(x) for x in stats], 'mom'), 'mom'))
    bars.append(make_bar('%s · 累计同比增量排序（1-%d月，万吨）' % (dim_label, last[1]),
                         pick_bar([dict(x) for x in stats], 'yoy'), 'yoy'))

    charts = (bars + lines) if bars_first else (lines + bars)

    # 汇总表（列为主要量级 TopN）
    columns = [{'key': 'c%d' % i, 'label': st['name'], 'group': name}
               for i, st in enumerate(top)]
    rnd = lambda v: (None if v is None else round(float(v), 2))
    rows = {
        '本期': [rnd(st['cur']) for st in top],
        '上期': [rnd(st['prev']) for st in top],
        '环比': [rnd(st['mom']) for st in top],
        '累计同比': [rnd(st['yoy']) for st in top],
        '累计增幅%': [None if st['yoy_pct'] is None else round(st['yoy_pct'], 1) for st in top],
    }
    summary = {
        'columns': columns, 'rows': rows,
        'rowOrder': ['本期', '上期', '环比', '累计同比', '累计增幅%'],
        'currentWeek': '%d-%02d' % (last[0], last[1]),
        'previousWeek': '%d-%02d' % (prev[0], prev[1]),
        'unit': unit,
    }

    return {
        'id': dsid, 'name': name, 'axes': [axis],
        'asOf': '%d-%02d' % (last[0], last[1]),
        'charts': charts, 'summary': summary, 'unit': unit,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    ds_v = build_dataset('出口-分品种', 'export_variety', '出口-分品种', '分品种',
                         top_n=11, add_aggregates=True)
    ds_c = build_dataset('出口-分国别', 'export_country', '出口-分国别', '分国别', top_n=9)

    regenerated = {d['id']: d for d in (ds_v, ds_c) if d}

    if args.check:
        for ds in (ds_v, ds_c):
            if not ds:
                print('[warn] 无数据')
                continue
            print('== %s(%s) | asOf %s | charts %d' % (ds['name'], ds['id'], ds['asOf'], len(ds['charts'])))
            for c in ds['charts']:
                if c['type'] == 'bar':
                    print('   [bar]', c['title'], '| n=', len(c['bar']['categories']))
                else:
                    print('   [line]', c['title'])
        return

    # ---- 合并并重写 data.js + meta.json ----
    meta_path = os.path.join(DATA, 'meta.json')
    order = []
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as f:
                order = [d['id'] for d in json.load(f).get('datasets', [])]
        except Exception:
            pass
    if not order:
        order = ['juanluo_luowen', 'juanluo_rejuan', 'daiguan', 'hanguan', 'chugang']
    for did in regenerated:
        if did not in order:
            order.append(did)

    all_ds = []
    for did in order:
        if did in regenerated:
            all_ds.append(regenerated[did])
        else:
            p = os.path.join(DATA, did + '.json')
            if os.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    all_ds.append(json.load(f))
                print('[ok] 读取既有 %s.json' % did)
            else:
                print('[warn] 缺少 %s.json，跳过' % did)

    stamp = datetime.datetime.now().isoformat(timespec='seconds')
    for ds in regenerated.values():
        p = os.path.join(DATA, ds['id'] + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(ds, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
              % (ds['id'], os.path.getsize(p) / 1024, len(ds['charts']), ds['asOf']))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_export_charts.py @ ' + stamp + '\n')
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': all_ds}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] 写 data.js (datasets=%d)' % len(all_ds))

    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']}
                         for d in all_ds]}
    with open(meta_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] 写 meta.json datasets=%d' % len(meta['datasets']))


if __name__ == '__main__':
    main()
