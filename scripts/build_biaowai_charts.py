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

# 单位换算：源表个别列需换算到展示口径
#   col89 废钢日耗：源为吨/日，加工表(用户既有图)为万吨 → ×1e-4（已交叉验证 503983→50.40）
#   col119/123/127 需求累计同比 %：源为小数（0.0316），图上按百分数展示 → ×100（用户明确选「百分比 %」）
SCALE = {91: 1e-4, 119: 100, 123: 100, 127: 100}

# 指标定义：(列号(1-based), 显示名)  —— 顺序 / 命名对齐用户既有「表外数据处理及图」的 62 个块
# 2026-09-27 用户要求（第 2 次调整）：
#   ① 删「非五大实库销」(87)
#   ② 非五大材（系数折算）组内去掉「实」字；其中 col76/77 去「实」后与 col80/81 同名，
#      按用户「保留数字口径大的」→ 删 col80(270) / col81(705)，保留 col76(465) / col77(1259)
#   ③ 「总量 / 五大材 / 非五大材 / 供给端」四组提到最前
#   ④ 「需求结构（年度累计）」换成 3 张「累计同比 %」图（源表 col114-129 区，r4 标注为「累计同比」）
GROUPS = [
    ('总量', [
        (57, '总产量'), (58, '总厂库'), (59, '总社库'),
        (60, '总库存'), (61, '总表需'), (85, '总库销'),
    ]),
    ('五大材（系数折算）', [
        (72, '五大实产量'), (73, '五大实厂库'), (74, '五大实社库'),
        (75, '五大实库存'), (76, '五大实表需'), (88, '五大实库销'),
    ]),
    ('非五大材（系数折算）', [
        (77, '非五大产量'), (78, '非五大厂库'), (79, '非五大社库'),
        (80, '非五大库存'), (81, '非五大表需'), (84, '非五大显性库存'),
    ]),
    ('供给端', [
        (92, '日均供给'), (93, '表内产量'), (94, '表外产量'), (91, '废钢日耗'),
    ]),
    ('非五大材明细', [
        (32, '工角槽厂库'), (33, '工角槽社库'), (49, '工角槽库存'),
        (35, '彩涂厂库'), (36, '彩涂社库'), (34, '彩涂库存'),
        (38, '镀锌厂库'), (39, '镀锌社库'), (37, '镀锌库存'),
        (40, 'H型钢厂库'), (41, 'H型钢社库'), (50, 'H型钢库存'),
        (42, '带钢厂库'), (43, '带钢社库'), (51, '带钢库存'),
        (46, '钢坯仓库'), (47, '调坯库存'), (48, '钢坯库存'),
        (44, '焊管社库'), (45, '无缝管社库'),
        (54, '镀锌表需'), (55, '彩涂表需'), (56, '工角槽表需'),
    ]),
    # 分品种表需（2026-09-28 用户：替换原「需求累计同比」三张图）—— 含带钢，取自唐宋库
    ('分品种表需', [
        (52, '螺纹大样本表需'), (53, '热卷大样本表需'),
        (12, '线材表需'), (14, '冷轧表需'), (15, '中厚板表需'),
        ('DAIGUAN', '带钢表需'),
        (56, '型钢表需'), (54, '镀锌表需'), (55, '彩涂表需'),
    ]),
]

# 顶部数据表 = 需求三项 + 9 个品种表需（2026-09-28 用户要求：
#   值改成绝对值 + 行改成 本期/上期/环比/累计同比%；随后又把「粗钢/五大材/表外需求」放回最前面）
#   元素 = (来源, 显示名, 分组)；来源是「表外数据」列号，或 'DAIGUAN'（唐宋管带数据库）
#   需求三项的列已用「自算累计同比 vs 源表累计同比列」交叉验证：col61↔-0.85、col76↔-0.26、col71↔+0.55 完全一致
SUMMARY_ITEMS = [
    (61, '粗钢需求', '需求'), (76, '五大材需求', '需求'), (71, '表外需求', '需求'),
    (52, '螺纹大样本表需', '品种表需'), (53, '热卷大样本表需', '品种表需'),
    (12, '线材表需', '品种表需'), (14, '冷轧表需', '品种表需'), (15, '中厚板表需', '品种表需'),
    ('DAIGUAN', '带钢表需', '品种表需'),
    (56, '型钢表需', '品种表需'), (54, '镀锌表需', '品种表需'), (55, '彩涂表需', '品种表需'),
]
DAIGUAN_DB = r"C:/Users/Administrator/Nutstore/1/我的坚果云/周度更新/唐宋管带数据库.xlsx"
DAIGUAN_SHEET = '带钢需求'
DAIGUAN_COL = 9          # 「带钢需求」sheet 第 9 列 = 带钢表需（1-based）

# 某些分组单独设「一行几张」（给该组图表打 cols，app.js 用组内首张图的值）
# 2026-09-28：用户最终确认「整 tab 一行四张」→ 用 app.js 的 GRID_COLS[biaowai]=4 控制，此处留空备用
GROUP_COLS = {}

ALL_ITEMS = [(col, name) for _, items in GROUPS for col, name in items]

# 需要从「表外数据」解析的列 = 图表用的列 ∪ 顶部数据表用的列
# （汇总表里的 线材/冷轧/中厚/螺纹大样本/热卷大样本表需 不在 GROUPS 里，必须一并解析）
PARSE_COLS = sorted({c for c, _ in ALL_ITEMS if isinstance(c, int)} |
                    {c for c, _, _ in SUMMARY_ITEMS if isinstance(c, int)})

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


def load_daiguan_table():
    """读《唐宋管带数据库》「带钢需求」sheet 的「带钢表需」，返回 {(ISO年, ISO周): 值} → 用于按周对齐。

    2026-09-28 用户要求：顶部数据表新增「带钢表需」，数据取自唐宋库（站上直接读，不改用户 Excel）。
    唐宋库是**周三**、本表是**周五**，故按 ISO 周对齐而不是按日期精确匹配。
    """
    if not os.path.exists(DAIGUAN_DB):
        print('[warn] 唐宋库不存在: %s' % DAIGUAN_DB)
        return {}
    wb = openpyxl.load_workbook(DAIGUAN_DB, read_only=True, data_only=True)
    if DAIGUAN_SHEET not in wb.sheetnames:
        print('[warn] 唐宋库无「%s」sheet，现有: %s' % (DAIGUAN_SHEET, wb.sheetnames))
        return {}
    ws = wb[DAIGUAN_SHEET]
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r:
            continue
        d = r[0]
        v = r[DAIGUAN_COL - 1] if len(r) >= DAIGUAN_COL else None
        if isinstance(d, datetime.datetime) and isinstance(v, (int, float)) and not isinstance(v, bool):
            out[tuple(d.isocalendar()[:2])] = float(v)
    return out


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

    series = {col: {} for col in PARSE_COLS}
    dates = []
    for r in rows[start:]:
        d = r[0] if r else None
        if not isinstance(d, datetime.datetime):
            continue
        dates.append(d)
        for col in PARSE_COLS:
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

    # 带钢表需（唐宋管带数据库）：按 ISO 周对齐到「表外数据」的日期轴（季节图 + 汇总表共用）
    dg = load_daiguan_table()
    daiguan_smap = {}
    for d in plot_dates:
        k = tuple(d.isocalendar()[:2])
        if k in dg:
            daiguan_smap[d] = dg[k]
    if daiguan_smap:
        dlast = max(daiguan_smap)
        print('[ok] 带钢表需（唐宋库）对齐 %d 个周，最新 %s = %.2f'
              % (len(daiguan_smap), dlast.date(), daiguan_smap[dlast]))
    else:
        print('[warn] 唐宋库「带钢需求」未取到数据（检查路径/表名）')

    charts = []
    for group, items in GROUPS:
        for col, name in items:
            smap = daiguan_smap if col == 'DAIGUAN' else series.get(col, {})
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
            card = {'key': '%s-%s' % (DS_ID, col), 'title': name, 'type': 'line',
                    'axis': 0, 'group': group, 'series': ser}
            if group in GROUP_COLS:      # 该组一行几张（app.js 用组内首张图的 cols）
                card['cols'] = GROUP_COLS[group]
            charts.append(card)

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

    # —— 顶部数据表：需求三项 + 9 个品种表需（绝对值 + 累计同比%）——
    # （daiguan_smap 已在上面加载，季节图与汇总表共用）
    def cum_yoy(smap, target):
        """年内累计 vs 去年同期累计（%）；去年同期取 target-364 天。"""
        py = target - datetime.timedelta(days=364)
        a = sum(v for d, v in smap.items() if d.year == target.year and d <= target)
        b = sum(v for d, v in smap.items() if d.year == py.year and d <= py)
        if not b:
            return None
        return round((a / b - 1) * 100, 2)

    columns, cur = [], []
    for src, name, grp in SUMMARY_ITEMS:
        smap = daiguan_smap if src == 'DAIGUAN' else series.get(src, {})
        if not smap:
            print('[warn] 汇总列 %s 无数据，跳过' % name)
            continue
        columns.append({'key': 'c%d' % len(columns), 'label': name, 'group': grp})
        cur.append({'cur': val_at(smap, last), 'prev': val_at(smap, prev),
                    'cum': cum_yoy(smap, last)})

    rnd = lambda v: (None if v is None else round(float(v), 2))
    today = lambda d: '%04d-%02d-%02d' % (d.year, d.month, d.day)
    summary = {
        'columns': columns,
        'rows': {
            '本期': [rnd(x['cur']) for x in cur],
            '上期': [rnd(x['prev']) for x in cur],
            '环比': [rnd(x['cur'] - x['prev']) if (x['cur'] is not None and x['prev'] is not None) else None
                     for x in cur],
            '累计同比%': [x['cum'] for x in cur],
        },
        'rowOrder': ['本期', '上期', '环比', '累计同比%'],
        'currentWeek': today(last), 'previousWeek': today(prev),
        'unit': '万吨（累计同比行为 %）',
    }

    note = ('口径：五大材 / 非五大材（系数折算）= 钢联样本外推后按系数折算的全口径（原「含样本外」）；'
            '非五大材明细 = 彩涂/镀锌/带钢/H型钢/工角槽/焊管/无缝管/钢坯等分项。'
            '顶部数据表 = 9 个品种表需（绝对值，万吨），累计同比 = 年内累计 ÷ 去年同期累计 − 1；'
            '其中带钢表需取自《唐宋管带数据库》「带钢需求」（按其周频对齐）。'
            '周频（周五），源《粗钢及表外(非五大材)情况.xlsx》「表外数据」。'
            '图例固定 5 年（2022–2026），2022 年口径未覆盖故不画线，后续补齐历史后将自动补线；'
            '需求累计同比（%）源表仅 2025、2026 两年有值（2023/2024 为 #N/A）。')
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
