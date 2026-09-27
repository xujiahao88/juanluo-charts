# -*- coding: utf-8 -*-
"""
build_biaowai_charts.py — 「表外（非五大材）」数据集生成器（季节性年对比版）

源：《粗钢及表外(非五大材)情况.xlsx》→ 「表外数据」sheet
  · A 列（第 1 列）= 日期，**周频（周五）**
  · 第 6/7/8/9 行 = 列名 / 指标名称 / 单位 / 指标说明（表头区）
  · 第 10 行起 = 数据

为什么用「表外数据」而不是「表外数据处理及图」：
  后者把同一批指标按「年内第 N 天」铺开成 5 个年份列（2022–2026），闰年会错位；
  前者有**真日期**，按 MM-DD 归年不会有任何错位，且口径一致。
  （已逐指标交叉验证：62 个指标块与源列 1:1 对齐，ratio=1）

年份策略（用户 2026-09-27 指定）：
  · YEAR_WINDOW = 5  → 图例/年份开关固定 5 格（2022、2023、2024、2025、2026），与全站其他 tab 一致；
  · PLOT_FROM  = 2023 → **只画 2023 起的线**，2022 那一格保留但不画线（表外口径最早 2023 年）。
    将来历史补齐后，把 PLOT_FROM 改成 2022 即可自动补上第五条线。

产出 juanluo-charts dataset「表外（非五大材）」(id=biaowai)：
  · 横轴 = 整年 MM-DD（01-01…12-31，366 点，含 02-29），与带钢/焊管/出港等 tab 同构；
  · 每张图按年份叠线，配色沿用全站：2022/2023 浅蓝虚线 → 2024 灰 → 2025 蓝 → 2026 红(带点)；
  · 全部 62 张图挂**同一组 5 个年份 series**，保证图例与顶部年份开关 5 格一致；
  · 汇总表：全部指标 本期/上期/环比/同比/同比%。

合并策略沿用 build_mill_order_charts.py：读 meta.json 现有顺序 → 本数据集就地替换/追加
→ 重写 data.js + meta.json，不丢其他 tab。

用法：
  python scripts/build_biaowai_charts.py            # 生成并合并
  python scripts/build_biaowai_charts.py --check    # 只看统计与结构，不落盘
"""
import argparse
import datetime
import json
import os

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\小目标\粗钢及表外(非五大材)情况.xlsx'
SHEET = '表外数据'

ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}

YEAR_WINDOW = 5        # 图例固定年份数（含预留的空年份）
PLOT_FROM = 2023       # 只画 >= 该年份的线（表外口径 2023 年起；将来补历史改这里）
DS_ID = 'biaowai'
DS_NAME = '表外（非五大材）'

# 单位换算：源表个别列为原始单位，需换算到「万吨」/「万吨·日」口径
#   col89 废钢日耗：源为吨/日，加工表(用户既有图)为万吨 → ×1e-4（已交叉验证 503983→50.40）
SCALE = {89: 1e-4}

# 指标定义：(列号(1-based), 显示名)  —— 顺序 / 命名对齐用户既有「表外数据处理及图」的 62 个块
GROUPS = [
    ('非五大材明细', [
        (32, '工角槽厂库'), (33, '工角槽社库'), (49, '工角槽库存'),
        (35, '彩涂厂库'), (36, '彩涂社库'), (34, '彩涂库存'),
        (38, '镀锌厂库'), (39, '镀锌社库'), (37, '镀锌库存'),
        (40, 'H型钢厂库'), (41, 'H型钢社库'), (50, 'H型钢库存'),
        (42, '带钢厂库'), (43, '带钢社库'), (51, '带钢库存'),
        (46, '钢坯仓库'), (47, '调坯库存'), (48, '钢坯库存'),
        (44, '焊管社库'), (45, '无缝管社库'),
        (52, '镀锌表需'), (53, '彩涂表需'), (54, '工角槽表需'),
    ]),
    ('总量', [
        (55, '总产量'), (56, '总厂库'), (57, '总社库'),
        (58, '总库存'), (59, '总表需'), (83, '总库销'),
    ]),
    ('五大材', [
        (60, '五大产量'), (61, '五大厂库'), (62, '五大社库'),
        (63, '五大库存'), (64, '五大表需'), (84, '五大库销'),
    ]),
    ('表外（非五大材）', [
        (65, '表外产量'), (66, '表外厂库'), (67, '表外社库'),
        (68, '表外库存'), (69, '表外表需'), (85, '表外库销'),
    ]),
    ('五大实（含样本外）', [
        (70, '五大实产量'), (71, '五大实厂库'), (72, '五大实社库'),
        (73, '五大实库存'), (74, '五大实表需'), (86, '五大实库销'),
    ]),
    ('非五大实（含样本外）', [
        (75, '非五大实产量'), (76, '非五大实厂库'), (77, '非五大实社库'),
        (78, '非五大实库存'), (79, '非五大实表需'), (87, '非五大实库销'),
        (80, '非五大厂库'), (81, '非五大社库'), (82, '非五大显性库存'),
    ]),
    ('供给端', [
        (90, '日均供给'), (91, '表内产量'), (92, '表外产量'), (89, '废钢日耗'),
    ]),
    ('需求结构（年度累计）', [
        (130, '五大材需求占比'), (131, '表外需求占比'),
    ]),
]

ALL_ITEMS = [(col, name) for _, items in GROUPS for col, name in items]

MONTH_DAYS = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]   # 2 月取 29 → 366 点轴


def year_axis():
    """整年 MM-DD 轴：01-01…12-31（含 02-29），共 366 点，与带钢/焊管等同构。"""
    return ['%02d-%02d' % (m, d)
            for m, nd in enumerate(MONTH_DAYS, 1)
            for d in range(1, nd + 1)]


def clean_num(v, col):
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
        if v != v:          # NaN
            return None
        f = SCALE.get(col, 1.0)
        return round(float(v) * f, 4)
    return None


def series_style(i, n):
    """与站点其他 tab 一致的年样式：最新红、前一年蓝、n-3 灰、更早浅蓝虚线。"""
    if i == n - 1:
        color, dash, marker = '#FF0000', 'solid', 'circle'
    elif i == n - 2:
        color, dash, marker = '#0070C0', 'solid', 'none'
    elif i == n - 3:
        color, dash, marker = '#A5A5A5', 'solid', 'none'
    else:
        color, dash, marker = '#9DC3E6', 'dash', 'none'
    return {'color': color, 'width': 1.5, 'dash': dash, 'smooth': False, 'marker': marker}


def parse(wb):
    """返回 (dates, series) ; series[列号] = {date: val}
    注意：按**列号**做键 —— 「表外产量」在 col65(表外口径) 与 col92(供给端口径) 重名，不能按名字存。"""
    ws = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))
    start = None
    for i, r in enumerate(rows):
        if r and isinstance(r[0], datetime.datetime):
            start = i
            break
    if start is None:
        raise SystemExit('「%s」未找到日期列（A 列应为日期）' % SHEET)

    series = {col: {} for col, _ in ALL_ITEMS}
    dates = []
    for r in rows[start:]:
        d = r[0] if r else None
        if not isinstance(d, datetime.datetime):
            continue
        dates.append(d)
        for col, _name in ALL_ITEMS:
            if col - 1 >= len(r):
                continue
            v = clean_num(r[col - 1], col)
            if v is not None:
                series[col][d] = v
    return sorted(set(dates)), series


def build_dataset():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    dates, series = parse(wb)
    wb.close()
    if not dates:
        return None

    plot_dates = [d for d in dates if d.year >= PLOT_FROM]
    if not plot_dates:
        raise SystemExit('PLOT_FROM=%d 之后无数据' % PLOT_FROM)
    latest = max(d.year for d in plot_dates)

    # 图例：固定 5 格，最新年在最后（app.js 用「最后一个 series 名」判最新年并置顶 z-index）
    legend_years = list(range(latest - YEAR_WINDOW + 1, latest + 1))
    n = len(legend_years)
    axis = year_axis()

    def seasonal(smap):
        byyear = {}
        for d, v in smap.items():
            if d.year in legend_years:
                byyear.setdefault(d.year, {})['%02d-%02d' % (d.month, d.day)] = v
        return byyear

    charts = []
    for group, items in GROUPS:
        for col, name in items:
            smap = series.get(col, {})
            if not smap:
                print('[warn] %s：无数据，跳过' % name)
                continue
            byyear = seasonal(smap)
            ser = []
            for i, y in enumerate(legend_years):
                st = series_style(i, n)
                row = byyear.get(y, {})
                # 早于 PLOT_FROM 的年份：挂空 series（图例保留、不画线）
                data = [row.get(md) for md in axis] if y >= PLOT_FROM else [None] * len(axis)
                ser.append({'name': str(y), 'color': st['color'], 'width': st['width'],
                            'dash': st['dash'], 'smooth': st['smooth'],
                            'marker': st['marker'], 'data': data})
            charts.append({'key': '%s-%d' % (DS_ID, col), 'title': name, 'type': 'line',
                           'axis': 0, 'group': group, 'series': ser})

    # —— 汇总表：本期 / 上期(周) / 环比 / 同比(去年同期周) / 同比% ——
    last = max(plot_dates)
    prev = last - datetime.timedelta(days=7)
    yoy = last - datetime.timedelta(days=364)

    def val_at(smap, target):
        best = None
        for d, v in smap.items():
            if d.year < PLOT_FROM or d > target:
                continue
            if best is None or d > best[0]:
                best = (d, v)
        return best[1] if best else None

    columns, cur = [], []
    # 重名指标（表外产量 col65 / col92）在汇总表里加组名前缀区分
    from collections import Counter as _C
    dup = _C(n for _, items in GROUPS for _, n in items)
    for group, items in GROUPS:
        for col, name in items:
            smap = series.get(col, {})
            if not smap:
                continue
            label = (group + '·' + name) if dup[name] > 1 else name
            columns.append({'key': 'c%d' % len(columns), 'label': label, 'group': group})
            c, p, y = (val_at(smap, last), val_at(smap, prev), val_at(smap, yoy))
            cur.append({'cur': c, 'prev': p,
                        'yoy': (c - y) if (c is not None and y is not None) else None,
                        'yoy_pct': (round((c - y) / abs(y) * 100, 1)
                                    if (c is not None and y not in (None, 0)) else None)})

    rnd = lambda v: (None if v is None else round(float(v), 2))
    today = lambda d: '%04d-%02d-%02d' % (d.year, d.month, d.day)
    summary = {
        'columns': columns,
        'rows': {
            '本期': [rnd(x['cur']) for x in cur],
            '上期': [rnd(x['prev']) for x in cur],
            '环比': [rnd(x['cur'] - x['prev']) if (x['cur'] is not None and x['prev'] is not None) else None
                     for x in cur],
            '同比': [rnd(x['yoy']) for x in cur],
            '同比%': [x['yoy_pct'] for x in cur],
        },
        'rowOrder': ['本期', '上期', '环比', '同比', '同比%'],
        'currentWeek': today(last), 'previousWeek': today(prev),
        'unit': '万吨 / 万吨·日 / 倍 / %',
    }

    note = ('口径：表外 = 总（钢联样本外推）− 五大材；非五大实含彩涂/镀锌/带钢/H型钢/工角槽/焊管/无缝管/钢坯等。'
            '周频（周五），源《粗钢及表外(非五大材)情况.xlsx》「表外数据」。'
            '图例固定 5 年（2022–2026），2022 年口径未覆盖故不画线，后续补齐历史后将自动补线。')
    return {
        'id': DS_ID, 'name': DS_NAME,
        'axes': [axis], 'asOf': today(last), 'charts': charts,
        'summary': summary, 'note': note, 'unit': 'mixed',
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    ds = build_dataset()
    if ds is None:
        print('[warn] 无数据')
        return

    if args.check:
        print('== %s(%s) | asOf %s | charts %d | axis %d'
              % (ds['name'], ds['id'], ds['asOf'], len(ds['charts']), len(ds['axes'][0])))
        g = None
        for c in ds['charts']:
            if c['group'] != g:
                g = c['group']
                print('  ── %s ──' % g)
            yrs = [s['name'] for s in c['series']]
            drew = [s['name'] for s in c['series'] if any(v is not None for v in s['data'])]
            npts = sum(1 for v in c['series'][-1]['data'] if v is not None)
            tail = [v for v in c['series'][-1]['data'] if v is not None][-1:]
            print('     %-16s 图例=%s 画线=%s 最新年点数=%d 末值=%s'
                  % (c['title'], ','.join(yrs), ','.join(drew) or '无', npts, tail))
        print('   summary cols(%d): %s' % (len(ds['summary']['columns']),
                                           [c['label'] for c in ds['summary']['columns']][:14]))
        print('   本期 %s / 上期 %s' % (ds['summary']['currentWeek'], ds['summary']['previousWeek']))
        return

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
    if ds['id'] not in order:
        order.append(ds['id'])

    all_ds = []
    for did in order:
        if did == ds['id']:
            all_ds.append(ds)
        else:
            p = os.path.join(DATA, did + '.json')
            if os.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    all_ds.append(json.load(f))
                print('[ok] 读取既有 %s.json' % did)
            else:
                print('[warn] 缺少 %s.json，跳过' % did)

    stamp = datetime.datetime.now().isoformat(timespec='seconds')
    p = os.path.join(DATA, ds['id'] + '.json')
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(ds, f, ensure_ascii=False, separators=(',', ':'))
    print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
          % (ds['id'], os.path.getsize(p) / 1024, len(ds['charts']), ds['asOf']))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_biaowai_charts.py @ ' + stamp + '\n')
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
