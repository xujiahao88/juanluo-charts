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
    """重生成卷螺（螺纹/热卷）两个 dataset（需解析《卷螺大样本.xlsm》，~0.3s）
    仅在 --all 或磁盘上没有卷螺 json 时使用。"""
    if SCRIPTS not in sys.path:
        sys.path.insert(0, SCRIPTS)
    import build_juanluo_charts as bjc
    raw = bjc.read_all()
    return [bjc.build_dataset(page, dsid, raw[page]) for page, dsid, _ in bjc.SHEETS]


def datasets_same(new_list):
    """与磁盘 data.js 的 datasets 对比（忽略 updated 时间戳）

    为什么要忽略 updated：每次跑脚本 stamp 都是当前时间，如果只因时间戳就重写
    data.js / gangyin json，deploy_repo.py 会判定"文件变化"而每次都真推送（~5s）。
    数据实质没变时就不写盘 → deploy 能直接 SKIP。
    """
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


# 网站 tab 顺序：螺纹 → 热卷 → 钢银
JUANLUO_IDS = ['juanluo_luowen', 'juanluo_rejuan']


def load_cached_juanluo():
    """从磁盘 data/*.json 直接读已有的卷螺 dataset（不解析源 xlsm，省 ~0.3s）

    ⚠️ 两个坑：
      1. 钢银周更时卷螺数据根本没变，没必要每次重读大 xlsm。
      2. 但也不能只写钢银 dataset —— 那会把网站的螺纹/热卷 tab 弄丢
         （2026-09-09 实测：--only 曾写出 datasets=1 的 data.js）。
    """
    out = []
    for dsid in JUANLUO_IDS:
        p = os.path.join(DATA, dsid + '.json')
        if not os.path.exists(p):
            return None                    # 缺任何一个 → 回退全量重生
        try:
            with open(p, encoding='utf-8') as f:
                d = json.load(f)
            if isinstance(d, dict) and d.get('id') == dsid:
                out.append(d)
            else:
                return None
        except Exception:
            return None
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--all', action='store_true',
                    help='全量：也重新解析《卷螺大样本.xlsm》重生螺纹/热卷（卷螺源更新时用）')
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

    # 默认【快速模式】：只重生钢银，卷螺两个 dataset 从磁盘 json 复用（省 ~0.3s，且不丢 tab）
    # --all：卷螺源也更新了，需重新解析《卷螺大样本.xlsm》全量重生
    head = []
    if args.all:
        head = juanluo_datasets()
        print('[ok] 全量重生成卷螺 螺纹/热卷（解析源 xlsm）')
    else:
        head = load_cached_juanluo()
        if head is None:
            try:
                head = juanluo_datasets()
                print('[ok] 磁盘无卷螺缓存 → 回退全量重生成')
            except Exception as e:
                print('[warn] 重生成卷螺失败（%s），本次仅写钢银' % e)
                head = []
        else:
            print('[ok] 复用磁盘卷螺 dataset（跳过解析 xlsm，省 ~0.3s）')
    datasets = head + [ds]

    # 幂等：数据实质未变（忽略 updated 时间戳）→ 完全不写盘，
    # 让后续 deploy_repo.py 能走 SKIP 分支（省 ~5s 推送）
    if datasets_same(datasets):
        print('[skip] 数据未变化（忽略 updated 时间戳），不写盘 → 推送可 SKIP')
        return

    # 写盘策略：
    #  · 全量模式(--all)：卷螺 json 也重写
    #  · 快速模式：卷螺 json 内容没变，【不重写】——否则 updated 时间戳一变，
    #    deploy_repo.py 会判定"文件变化"而重复上传 641KB，推送从 ~4s 拖到 ~6s。
    targets = datasets if args.all else [ds]
    for d in targets:
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
