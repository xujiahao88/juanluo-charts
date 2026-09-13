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

# (dataset展示名, dataset id, [ (区块名 group, 源sheet名, pick), ... ])
# 顺序即 tab / 区块顺序
# pick=None：该 sheet 每列各出一张图（焊管保留原行为，约20张）
# pick=str ：该 sheet 只取「表头含 pick 的第一列」→ 1 张
# pick=[str, ...]：取多列，逐行求和（用于“镀锌带+管厂带钢库存”这种组合指标）
DATASETS = [
    ('带钢', 'daiguan', [
        # 严格对齐用户截图「带钢周度高频跟踪」九宫格。
        # 每项 (group, sheet, pick, title_override)
        ('带钢周度高频跟踪', '带钢产量', '全国带钢日产量',       '带钢产量（唐宋口径）'),
        ('带钢周度高频跟踪', '带钢需求', '全国带钢库存（万吨）', '带钢社库（唐宋口径）'),
        ('带钢周度高频跟踪', '带钢需求', '带钢厂库（天津）',      '带钢厂库（唐宋口径）'),
        ('带钢周度高频跟踪', '带钢需求', '带钢下游库存', '带钢下游库存（唐宋口径）'),
        ('带钢周度高频跟踪', '带钢需求', '带钢总库存',           '全产业链带钢库存（唐宋口径）'),
        ('带钢周度高频跟踪', '带钢需求', '带钢表需',             '带钢表需（唐宋口径）'),
        ('带钢周度高频跟踪', '带钢需求', '带钢总消耗',           '带钢总消耗（唐宋口径）'),
        ('带钢周度高频跟踪', '镀锌带订单', '全国镀锌带企业订单量', '全国镀锌带企业订单（唐宋口径）'),
        ('带钢周度高频跟踪', '带钢利润', '唐山带钢利润',         '唐山带钢利润（唐宋口径）'),
    ]),
    ('焊管', 'hanguan', [
        ('焊管产量&开工率', '焊管产量&开工率', None, None),
        ('管厂库存',        '管厂库存',        None, None),
    ]),
]

# ---- 出港（Nutstore/1/小目标/出港.xlsx）----
CHUGANG_SRC = r'C:\Users\Administrator\Nutstore\1\小目标\出港.xlsx'
CHUGANG_DATASETS = [
    ('出港', 'chugang', [
        ('出港',     '出港',     '钢材国内主要港口出港汇总',          '国内主要港口出港汇总（万吨）'),
        ('出港',     '出港',     '全球钢材出港: 中国台湾',            '全球出港：中国台湾（万吨）'),
        ('出港',     '出港',     '全球钢材出港: 越南',                '全球出港：越南（万吨）'),
        ('出港',     '出港',     '全球钢材出港: 伊朗',                '全球出港：伊朗（万吨）'),
        ('出港',     '出港',     '钢材出港（除台湾）',                '钢材出港（除台湾，万吨）'),
        ('出口接单', '出口接单', 'SMM: 钢材出口接单: 31家出口商: 周度', '出口接单：总量（吨）'),
        ('出口接单', '出口接单', '板材接单',                          '出口接单：板材（吨）'),
        ('出口接单', '出口接单', '长材接单',                          '出口接单：长材（吨）'),
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
    pick=str ：只取表头包含 pick 的第一列（每 sheet 仅 1 张）。
    pick=list[str]：取多列，逐行求和（用于“镀锌带+管厂带钢库存”这种组合指标）。
    """
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    try:
        header = list(next(rows))
    except StopIteration:
        return []
    ncol = len(header)

    if pick is not None:
        picks = pick if isinstance(pick, (list, tuple)) else [pick]
        targets = []
        for p in picks:
            for ci in range(1, ncol):
                h = header[ci]
                if h is not None and p in str(h):
                    targets.append(ci)
                    break
        if len(targets) < len(picks):
            missed = [p for i, p in enumerate(picks) if i >= len(targets)]
            print('  [warn] %s 未找到含 %r 的列，跳过' % (sheet_name, missed))
            return []
        th = ' + '.join(str(header[t]).strip() for t in targets) if len(targets) > 1 else str(header[targets[0]]).strip()
        pts = []
        for row in rows:
            d = row[0]
            if not isinstance(d, datetime.datetime):
                continue
            vs = []
            for t in targets:
                v = clean_num(row[t] if t < len(row) else None)
                if v is not None:
                    vs.append(v)
            if not vs:
                continue
            pts.append((d, sum(vs)))
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


def build_summary(summary_series, group='带钢高频', unit='混合'):
    """根据各指标原始日度序列，生成「本期/上期/环比/同比/同比%」汇总表。
    summary_series: [(label, [(datetime, val), ...]), ...]"""
    if not summary_series:
        return None
    all_dates = sorted({d for _, pts in summary_series for d, _ in pts})
    if not all_dates:
        return None
    last = max(all_dates)
    prev = last - datetime.timedelta(days=7)
    yoy = last - datetime.timedelta(days=364)

    def val_at(pts, target):
        best = None
        for d, v in pts:
            if d <= target:
                if best is None or d > best[0]:
                    best = (d, v)
        return best[1] if best else None

    cur = [val_at(pts, last) for _, pts in summary_series]
    pv = [val_at(pts, prev) for _, pts in summary_series]
    yv = [val_at(pts, yoy) for _, pts in summary_series]

    def short_label(lab):
        for suf in ('（唐宋口径）', '（万吨）', '（吨）'):
            lab = lab.replace(suf, '')
        return lab

    columns = [{'key': 'c%d' % i, 'label': short_label(lab), 'group': group}
               for i, (lab, _) in enumerate(summary_series)]
    rnd = lambda v: (None if v is None else round(float(v), 2))
    rows = {
        '本期': [rnd(v) for v in cur],
        '上期': [rnd(v) for v in pv],
        '环比': [None if (c is None or p is None) else rnd(c - p) for c, p in zip(cur, pv)],
        '同比': [None if (c is None or y is None) else rnd(c - y) for c, y in zip(cur, yv)],
        '同比%': [None if (c is None or y is None or y == 0) else rnd((c - y) / abs(y) * 100)
                  for c, y in zip(cur, yv)],
    }
    return {
        'columns': columns,
        'rows': rows,
        'rowOrder': ['本期', '上期', '环比', '同比', '同比%'],
        'currentWeek': '%02d-%02d' % (last.month, last.day),
        'previousWeek': '%02d-%02d' % (prev.month, prev.day),
        'unit': unit,
    }


def build_dataset(name, dsid, groups, src=SRC):
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    charts = []
    summary_series = []
    all_years = set()
    all_dates = []
    idx = 0
    for item in groups:
        if len(item) >= 4:
            group, sheet, pick, title_override = item[:4]
        elif len(item) == 3:
            group, sheet, pick = item
            title_override = None
        else:
            group, sheet = item
            pick = None
            title_override = None
        series = read_sheet_series(wb, sheet, pick)
        for header, pts in series:
            raw = pts
            pts = downsample_weekly(pts)
            if not pts:
                continue
            byyear = {}
            for d, v in pts:
                byyear.setdefault(d.year, {})['%02d-%02d' % (d.month, d.day)] = v
                all_years.add(d.year)
                all_dates.append(d)
            # 标题优先级：自定义 > 列名（焊管原模式）
            title = title_override or header
            charts.append({
                'group': group,
                'header': title,
                'byyear': byyear,
            })
            summary_series.append((title, raw))
            idx += 1
    wb.close()

    if not charts:
        return None

    # 带钢：生成「本期/上期/环比/同比」汇总表（参考铁矿站样式）
    SUMMARY_CFG = {
        'daiguan': ('带钢高频', '混合（产量/库存:万吨, 利润:元/吨, 订单:吨）'),
        'chugang': ('出港高频', '混合（出港:万吨, 出口接单:吨）'),
    }
    if dsid in SUMMARY_CFG and summary_series:
        _g, _u = SUMMARY_CFG[dsid]
        summary = build_summary(summary_series, group=_g, unit=_u)
    else:
        summary = None

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
        # 只给「该图确实有数据的年份」建 series：这样出口接单(仅2025-2026)
        # 不会挂 2022-2024 三条空线，同时颜色仍按全局年份位次保持一致。
        present = [y for y in years_sorted if y in ch['byyear']]
        for y in present:
            m = ch['byyear'][y]
            j = years_sorted.index(y)
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
        'summary': summary,
        'unit': 'mixed',
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    args = ap.parse_args()

    datasets = []
    specs = [(name, dsid, groups, SRC) for name, dsid, groups in DATASETS]
    specs += [(name, dsid, groups, CHUGANG_SRC) for name, dsid, groups in CHUGANG_DATASETS]
    for name, dsid, groups, src in specs:
        ds = build_dataset(name, dsid, groups, src=src)
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

    # ---- 合并：按 meta.json 现有顺序保留其他 dataset，regenerated 就地替换 ----
    regen = {ds['id']: ds for ds in datasets}
    meta_path = os.path.join(DATA, 'meta.json')
    order = []
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as f:
                order = [d['id'] for d in json.load(f).get('datasets', [])]
        except Exception:
            pass
    if not order:
        order = ['juanluo_luowen', 'juanluo_rejuan']
    for did in regen:
        if did not in order:
            order.append(did)

    all_ds = []
    for did in order:
        if did in regen:
            all_ds.append(regen[did])
        else:
            p = os.path.join(DATA, did + '.json')
            if os.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    all_ds.append(json.load(f))
                print('[ok] 读取既有 %s.json' % did)
            else:
                print('[warn] 缺少 %s.json，跳过' % did)

    stamp = datetime.datetime.now().isoformat(timespec='seconds')

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
