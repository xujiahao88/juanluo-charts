# -*- coding: utf-8 -*-
"""
build_daiguan_charts.py — 唐宋「管带数据库」→ 钢材站（原卷螺站）交互季节图数据集

从《唐宋管带数据库.xlsx》抽取带钢 / 焊管 / 管材 三条产品线的时间序列，
生成与卷螺站（build_juanluo_charts.py）完全同构的 dataset：
  · 每个产品线 = 站点一个 tab（dataset）
  · 每个 sheet = 一个区块（group），sheet 内每一列（除日期列）= 一张季节图
  · 每张图 = 季节图：x 轴 366 天日历轴(MM-DD)，series = 各年份
      配色沿用铁矿站 aubra：最老浅蓝虚线 → 灰 → 上年蓝平滑 → 当年红+圆点

降采样：所有序列统一做「周度降采样」（相邻点间隔≥6天才保留），
        日频价格(基准价)因此压成周频，周频基本面基本为空操作，
        既统一节奏又避免 JSON 膨胀。

年份裁剪：单数据集最多保留最近 5 个年份，避免长历史(如库存24年)糊成线团。

用法：
  python build_daiguan_charts.py            # 写 data/*.json + data.js + meta.json
  python build_daiguan_charts.py --check    # 仅抽样校验，不落盘
"""
import argparse
import datetime
import json
import os

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\我的坚果云\周度更新\唐宋管带数据库.xlsx'

# (dataset展示名, dataset id, [ (区块名 group, 源sheet名), ... ])
# 顺序即 tab / 区块顺序
# (dataset展示名, dataset id, [ (区块/图标题 group, 源sheet名, pick) ])
# pick=None：该 sheet 每列各出一张图（焊管保留原行为，约20张）
# pick=str ：该 sheet 只取「表头含 pick 的第一列」→ 每 sheet 仅 1 张（带钢按此精简为 9 张）
DATASETS = [
    ('带钢', 'daiguan', [
        ('带钢基准价',   '带钢基准价',   '全国'),
        ('带钢出厂价',   '带钢出厂价',   '瑞丰'),
        ('带钢市场价',   '带钢市场价',   '唐山市'),
        ('带钢开工率',   '带钢开工率',   '全国按条数调坯带钢开工率'),
        ('带钢产量',     '带钢产量',     '全国带钢日产量'),
        ('镀锌带订单',   '镀锌带订单',   '全国镀锌带企业订单量'),
        ('带钢库存',     '带钢库存',     '全国带钢库存'),
        ('带钢利润',     '带钢利润',     '唐山带钢利润'),
        ('带钢需求',     '带钢需求',     '带钢表需'),
    ]),
    ('焊管', 'hanguan', [
        ('焊管产量&开工率', '焊管产量&开工率', None),
        ('管厂库存',        '管厂库存',        None),
    ]),
]

ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}
MAX_YEARS = 5           # 单数据集最多保留最近 N 年


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


def read_sheet_series(wb, sheet_name, pick=None):
    """返回 [(header, [(datetime, val), ...]), ...]。
    pick=None：除日期列外每列一张图（焊管等保留原行为）。
    pick=str ：只取表头包含 pick 的第一列（带钢各 sheet 取全国/标杆系列，每 sheet 仅 1 张）。
    """
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    try:
        header = list(next(rows))
    except StopIteration:
        return []
    ncol = len(header)

    if pick is not None:
        target = None
        for ci in range(1, ncol):
            h = header[ci]
            if h is not None and pick in str(h):
                target = ci
                break
        if target is None:
            print('  [warn] %s 未找到含 %r 的列，跳过' % (sheet_name, pick))
            return []
        th = str(header[target]).strip()
        pts = []
        for row in rows:
            d = row[0]
            if not isinstance(d, datetime.datetime):
                continue
            v = clean_num(row[target] if target < len(row) else None)
            if v is None:
                continue
            pts.append((d, v))
        return [(th, pts)] if pts else []

    # ---- pick=None：原逻辑（每列一图）----
    out = []  # list of (header, list_of_points)
    for ci in range(1, ncol):
        h = header[ci]
        if h is None:
            continue
        hs = str(h).strip()
        if hs == '' or hs == '日期':   # 日期列（含分块后的第二个日期列）跳过
            continue
        out.append((hs, []))
    if not out:
        return []
    for row in rows:
        d = row[0]
        if not isinstance(d, datetime.datetime):
            # 非日期行（空行 / 标题行）跳过
            if d is None:
                continue
            continue
        for ci in range(1, ncol):
            if ci - 1 >= len(out):
                break
            v = clean_num(row[ci] if ci < len(row) else None)
            if v is None:
                continue
            out[ci - 1][1].append((d, v))
    # 去掉没有数据的列
    out = [(h, pts) for h, pts in out if pts]
    return out


def downsample_weekly(pts):
    """相邻保留点间隔≥6天 → 约周频；周频数据基本为空操作。"""
    if not pts:
        return []
    pts = sorted(pts, key=lambda x: x[0])
    out = [pts[0]]
    for p in pts[1:]:
        if (p[0] - out[-1][0]).days >= 6:
            out.append(p)
    return out


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


def build_dataset(name, dsid, groups):
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    charts = []
    all_years = set()
    all_dates = []
    idx = 0
    for item in groups:
        if len(item) == 3:
            group, sheet, pick = item
        else:
            group, sheet = item
            pick = None
        series = read_sheet_series(wb, sheet, pick)
        for header, pts in series:
            pts = downsample_weekly(pts)
            if not pts:
                continue
            byyear = {}
            for d, v in pts:
                byyear.setdefault(d.year, {})['%02d-%02d' % (d.month, d.day)] = v
                all_years.add(d.year)
                all_dates.append(d)
            # pick 模式：图标题用区块名(=sheet名)更简洁；原模式用列名
            title = group if pick is not None else header
            charts.append({
                'group': group,
                'header': title,
                'byyear': byyear,
            })
            idx += 1
    wb.close()

    if not charts:
        return None

    # 季节轴：闰年 366 天 MM-DD 全日历轴
    axis = []
    _d = datetime.date(2024, 1, 1)
    while _d.year == 2024:
        axis.append('%02d-%02d' % (_d.month, _d.day))
        _d += datetime.timedelta(days=1)

    years_sorted = sorted(all_years)
    if len(years_sorted) > MAX_YEARS:
        years_sorted = years_sorted[-MAX_YEARS:]
    n = len(years_sorted)

    out_charts = []
    for i, ch in enumerate(charts):
        series = []
        for j, y in enumerate(years_sorted):
            m = ch['byyear'].get(y, {})
            st = series_style(j, n)
            series.append({
                'name': str(y),
                'color': st['color'],
                'width': st['width'],
                'dash': st['dash'],
                'smooth': st['smooth'],
                'marker': st['marker'],
                'data': [m.get(mm) for mm in axis],
            })
        out_charts.append({
            'key': '%s-%d' % (dsid, i),
            'title': ch['header'],
            'type': 'line',
            'axis': 0,
            'group': ch['group'],
            'series': series,
        })

    last = max(all_dates) if all_dates else None
    as_of = '%02d-%02d' % (last.month, last.day) if last else ''

    return {
        'id': dsid,
        'name': name,
        'axes': [axis],
        'asOf': as_of,
        'charts': out_charts,
        'summary': None,        # 新数据集指标单位混杂(吨/%/元)，不生成汇总表
        'unit': 'mixed',
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    datasets = []
    for name, dsid, groups in DATASETS:
        ds = build_dataset(name, dsid, groups)
        if ds is None:
            print('[warn] %s(%s) 无数据，跳过' % (name, dsid))
            continue
        datasets.append(ds)
        if args.check:
            print('== %s(%s) | asOf %s | charts %d | years %d | axis %d'
                  % (ds['name'], ds['id'], ds['asOf'], len(ds['charts']),
                     len(ds['axes'][0]) and len([s for s in ds['charts'][0]['series']]),
                     len(ds['axes'][0])))
            # 区块分布
            from collections import Counter, OrderedDict
            grp = OrderedDict()
            for c in ds['charts']:
                grp.setdefault(c['group'], 0)
                grp[c['group']] += 1
            print('   区块:', dict(grp))
            print('   首图 title:', ds['charts'][0]['title'])
            print('   年份:', [s['name'] for s in ds['charts'][0]['series']])

    if args.check:
        return

    # ---- 合并现有 螺纹/热卷，重写 data.js + meta.json ----
    existing = []
    for fid in ('juanluo_luowen', 'juanluo_rejuan'):
        p = os.path.join(DATA, fid + '.json')
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                existing.append(json.load(f))
            print('[ok] 读取既有 %s.json' % fid)

    stamp = datetime.datetime.now().isoformat(timespec='seconds')
    all_ds = existing + datasets

    os.makedirs(DATA, exist_ok=True)
    for ds in datasets:
        p = os.path.join(DATA, ds['id'] + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(ds, f, ensure_ascii=False, separators=(',', ':'))
        print('[ok] 写 %s.json (%.0f KB) charts=%d asOf=%s'
              % (ds['id'], os.path.getsize(p) / 1024, len(ds['charts']), ds['asOf']))

    with open(os.path.join(DATA, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('// 自动生成，勿手改。build_daiguan_charts.py @ ' + stamp + '\n')
        f.write('window.CHART_DATA = ')
        json.dump({'updated': stamp, 'datasets': all_ds}, f,
                  ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print('[ok] 写 data.js (datasets=%d)' % len(all_ds))

    meta = {'updated': stamp,
            'datasets': [{'id': d['id'], 'name': d['name'],
                          'charts': len(d['charts']), 'asOf': d['asOf']}
                         for d in all_ds]}
    with open(os.path.join(DATA, 'meta.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print('[ok] 写 meta.json datasets=%d' % len(meta['datasets']))


if __name__ == '__main__':
    main()
