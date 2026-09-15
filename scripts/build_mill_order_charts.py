# -*- coding: utf-8 -*-
"""
build_mill_order_charts.py — 「钢厂日接单」数据集生成器（季节性年对比版）

源：《钢厂日接单量统计.xlsx》的「接单」sheet（唐宋/机构日度钢厂接单跟踪）。
  结构：
    第1行（日产）：B=日产标签，C..J = 8 家钢厂日产（万吨/日），K = 合计(21.6)
    第2行（表头）：B=日期/接单量，C..J = 8 家钢厂接单量，
                  K=平均利润率，L=总接单，M=接单率（=总接单/日产合计）
    第3行起：每日数据（2024-01 起 ~ 今日）

产出 juanluo-charts 一个 dataset「钢厂日接单」(id=mill_order)：
  · 横轴 = 日历轴（MM-DD 并集），**每个指标按年份叠线**（最新年红色、
    前一年蓝色、更早灰/浅蓝），与站点其他 tab（螺纹/热卷/带钢/出港）一致。
  · 图表：
      「钢厂接单（5日均值）」组：8 张 —— 纵横/中铁/安丰/燕钢/瑞丰/新东海/东华/日钢
        （日钢量级 10-35 万远大于其他 2-6 万，故单独一张，带日产参考线）
      「综合指标」组：3 张 —— 平均利润率（日频 + 盈亏平衡线）、总接单（5日均值 +
        日产合计参考线）、接单率（5日均值）
  · 5 日均值口径：**先在连续日期序列上做 MA5，再按年份拆分**，避免跨年错位。
  · 汇总表：8 钢厂 + 平均利润率 + 总接单 + 接单率 的 本期/上期/环比/同比/同比%

合并策略沿用 build_export_charts.py：读 meta.json 现有顺序 → 本数据集就地替换/追加 →
重写 data.js + meta.json，不丢其他 tab。

用法：
  python scripts/build_mill_order_charts.py            # 生成并合并
  python scripts/build_mill_order_charts.py --check    # 只看统计与结构
"""
import argparse
import datetime
import json
import os
import re

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
SRC = r'C:\Users\Administrator\Nutstore\1\小目标\钢厂日接单量统计.xlsx'

ERR_TOKENS = {'#N/A', '#N/A!', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', ''}

MAX_YEARS = 5           # 最多保留最近 N 年
MA_WINDOW = 5           # 接单量类指标的移动平均窗口（5 日）

# 源表头第 8 家（日产 4.0，接单量 10~35 万吨，量级远大于其他 2-6 万）→ 单独成图。
# 优先认固定名（源表头现为「日钢」），若源改名则自动退化为「日产最大的一家」，
# 避免表头一改就静默把大厂混进小量级图、还丢掉日产参考线。
BIG_MILL = '日钢'


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


def clean_num_loose(v):
    """宽松解析：从「1.5（含带）」「21.6（合计）」这类带注释的字符串中提取首个数字。"""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return round(float(v), 4)
    m = re.search(r'-?\d+(?:\.\d+)?', str(v))
    return float(m.group()) if m else None


def moving_avg(vals, window=5):
    """尾部窗口移动平均（窗口内非空值求均值；窗口全空则 None）。"""
    out = []
    for i in range(len(vals)):
        lo = max(0, i - window + 1)
        seg = [v for v in vals[lo:i + 1] if v is not None]
        out.append(round(sum(seg) / len(seg), 3) if seg else None)
    return out


def series_style(i, n):
    """与站点其他 tab 一致的年样式：最新红、前一年蓝、n-3 灰、更早浅蓝。"""
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


def parse(wb):
    """返回 (mills, daily_prod, daily_prod_total, dates, data)"""
    ws = wb['接单']
    rows = list(ws.iter_rows(values_only=True))

    # 找 日产行 与 表头行
    prod_row = None
    header_row = None
    for r in rows:
        if len(r) > 1 and r[1] is not None and '日产' in str(r[1]):
            prod_row = r
        elif len(r) > 1 and r[1] is not None and '日期' in str(r[1]):
            header_row = r
        if prod_row and header_row:
            break
    if prod_row is None or header_row is None:
        raise SystemExit('未找到 日产行 / 表头行')

    # 钢厂名（表头 C..J，索引 2..9）
    mills = []
    for ci in range(2, 10):
        nm = header_row[ci]
        if nm is None:
            continue
        mills.append(str(nm).strip())
    # 日产（日产行 C..J）—— 该行为「3.7」「1.5（含带）」等带注释字符串，用宽松解析
    daily_prod = {}
    for k, ci in enumerate(range(2, 10)):
        if k < len(mills):
            daily_prod[mills[k]] = clean_num_loose(prod_row[ci])
    daily_prod_total = clean_num_loose(prod_row[10])  # K 列「21.6（合计）」

    # 数据行：B(列索引1)=日期字符串，C..J=各钢厂，K=利润率, L=总接单, M=接单率
    data = {'mill': {m: {} for m in mills}, 'profit': {}, 'total': {}, 'rate': {}}
    dates = []
    for r in rows:
        ds_raw = r[1] if len(r) > 1 else None
        if not ds_raw or not str(ds_raw).strip():
            continue
        try:
            d = datetime.datetime.strptime(str(ds_raw)[:10], '%Y-%m-%d')
        except ValueError:
            continue
        dates.append(d)
        for k, m in enumerate(mills):
            v = clean_num(r[2 + k] if (2 + k) < len(r) else None)
            if v is not None:
                data['mill'][m][d] = v
        p = clean_num(r[10] if len(r) > 10 else None)
        if p is not None:
            data['profit'][d] = p
        t = clean_num(r[11] if len(r) > 11 else None)
        if t is not None:
            data['total'][d] = t
        rt = clean_num(r[12] if len(r) > 12 else None)
        if rt is not None:
            data['rate'][d] = rt

    dates = sorted(set(dates))
    return mills, daily_prod, daily_prod_total, dates, data


def build_dataset():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    mills, daily_prod, daily_prod_total, dates, data = parse(wb)
    wb.close()
    if not dates:
        return None

    # —— 季节性横轴：出现过的 MM-DD 并集，按 (月, 日) 升序 ——
    md_set = {(d.month, d.day) for d in dates}
    md_axis = ['%02d-%02d' % (mm, dd) for mm, dd in sorted(md_set)]

    # 每月在轴上的第一个位置（= 该月首个交易日）→ 供前端放月份标签/刻度线。
    # 日频数据 1/1、5/1、10/1 等假日无数据，若按「每月 1 号」定位会整月丢标签。
    month_first = []
    seen_mm = set()
    for mm, dd in sorted(md_set):
        if mm not in seen_mm:
            seen_mm.add(mm)
            month_first.append('%02d-%02d' % (mm, dd))

    years = sorted({d.year for d in dates})
    if len(years) > MAX_YEARS:
        years = years[-MAX_YEARS:]
    n = len(years)

    def seasonal(series_map, ma_window=0):
        """连续日频 {date: val} → 季节性 {year: {MM-DD: val}}
        ma_window>0 时先在连续序列上做移动平均，再按年拆分。"""
        raw = [series_map.get(d) for d in dates]
        vals = moving_avg(raw, ma_window) if ma_window else raw
        byyear = {}
        for d, v in zip(dates, vals):
            if v is None or d.year not in years:
                continue
            byyear.setdefault(d.year, {})['%02d-%02d' % (d.month, d.day)] = v
        return byyear

    charts = []

    def add_chart(key, title, group, byyear, ref=None, ref_label=None):
        present = [y for y in years if y in byyear]
        if not present:
            return
        ser = []
        for y in present:
            st = series_style(years.index(y), n)
            ser.append({
                'name': str(y), 'color': st['color'], 'width': st['width'],
                'dash': st['dash'], 'smooth': st['smooth'], 'marker': st['marker'],
                'data': [byyear[y].get(md) for md in md_axis],
            })
        ch = {
            'key': key, 'title': title, 'type': 'line', 'axis': 0,
            'group': group, 'series': ser,
        }
        if ref is not None:
            ch['refLine'] = ref
            ch['refLabel'] = ref_label or ('参考 %.1f' % ref)
        charts.append(ch)

    # —— 组1：钢厂接单（5日均值）——
    # 大厂（量级远大于其他）单独成图并带日产参考线，其余钢厂按源顺序排列
    if BIG_MILL in mills:
        big = BIG_MILL
    else:
        big = max(mills, key=lambda m: (daily_prod.get(m) or 0)) if mills else None
        if big:
            print('[warn] 源表头未找到「%s」，按日产最大自动选：%s' % (BIG_MILL, big))
    order = [m for m in mills if m != big] + ([big] if big else [])
    for m in order:
        if m not in data['mill']:
            continue
        title = '%s日接单（5日均值）' % m
        if m == big:
            add_chart('mill_order-%s' % m, title, '钢厂接单（5日均值）',
                      seasonal(data['mill'][m], MA_WINDOW),
                      ref=daily_prod.get(m),
                      ref_label='日产 %.1f' % (daily_prod.get(m) or 0))
        else:
            add_chart('mill_order-%s' % m, title, '钢厂接单（5日均值）',
                      seasonal(data['mill'][m], MA_WINDOW))

    # —— 组2：综合指标 ——
    add_chart('mill_order-profit', '平均利润率（日频）', '综合指标',
              seasonal(data['profit']),
              ref=0, ref_label='盈亏平衡')

    add_chart('mill_order-total', '总接单（5日均值）', '综合指标',
              seasonal(data['total'], MA_WINDOW),
              ref=daily_prod_total,
              ref_label='日产合计 %.1f' % (daily_prod_total or 0))

    add_chart('mill_order-rate', '接单率（5日均值·总接单/日产合计）', '综合指标',
              seasonal(data['rate'], MA_WINDOW))

    # —— 汇总表 ——
    last = dates[-1]
    prev = last - datetime.timedelta(days=7)
    yoy = last - datetime.timedelta(days=364)

    def val_at(series_map, target):
        best = None
        for d, v in series_map.items():
            if d <= target:
                if best is None or d > best[0]:
                    best = (d, v)
        return best[1] if best else None

    cols_def = [('mill', m, '钢厂接单') for m in mills] + \
               [('profit', None, '综合'), ('total', None, '综合'), ('rate', None, '综合')]
    columns = []
    cur_vals = []
    for kind, sub, grp in cols_def:
        if kind == 'mill':
            smap = data['mill'][sub]
            label = sub
        else:
            smap = data[kind]
            label = {'profit': '平均利润率', 'total': '总接单', 'rate': '接单率'}[kind]
        columns.append({'key': 'c%d' % len(columns), 'label': label, 'group': grp})
        c = val_at(smap, last)
        p = val_at(smap, prev)
        y = val_at(smap, yoy)
        cur_vals.append({
            'cur': c, 'prev': p, 'yoy': (c - y) if (c is not None and y is not None) else None,
            'yoy_pct': (round((c - y) / abs(y) * 100, 1) if (c is not None and y not in (None, 0)) else None),
        })

    rnd = lambda v: (None if v is None else round(float(v), 2))
    rows = {
        '本期': [rnd(x['cur']) for x in cur_vals],
        '上期': [rnd(x['prev']) for x in cur_vals],
        '环比': [rnd(x['cur'] - x['prev']) if (x['cur'] is not None and x['prev'] is not None) else None
                 for x in cur_vals],
        '同比': [rnd(x['yoy']) for x in cur_vals],
        '同比%': [x['yoy_pct'] for x in cur_vals],
    }
    summary = {
        'columns': columns, 'rows': rows,
        'rowOrder': ['本期', '上期', '环比', '同比', '同比%'],
        'currentWeek': '%04d-%02d-%02d' % (last.year, last.month, last.day),
        'previousWeek': '%04d-%02d-%02d' % (prev.year, prev.month, prev.day),
        'unit': '万吨·%·天',
    }

    prod_parts = ['%s %.1f' % (m, daily_prod.get(m) or 0) for m in mills]
    note = '日产参考（万吨/日）：' + '｜'.join(prod_parts) + '（合计 %.1f）' % (daily_prod_total or 0)

    as_of = '%04d-%02d-%02d' % (last.year, last.month, last.day)
    return {
        'id': 'mill_order', 'name': '钢厂日接单',
        'axes': [md_axis], 'monthFirst': month_first,
        'asOf': as_of, 'charts': charts, 'summary': summary,
        'dailyProd': daily_prod, 'note': note, 'unit': 'mixed',
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
              % (ds['name'], ds['id'], ds['asOf'], len(ds['charts']),
                 len(ds['axes'][0])))
        for c in ds['charts']:
            yrs = [s['name'] for s in c.get('series', [])]
            npt = len([v for v in c['series'][0]['data'] if v is not None]) if c.get('series') else 0
            print('   [%s] %s | 年份=%s | 非空点=%d refLine=%s'
                  % (c.get('group'), c['title'], ','.join(yrs), npt, c.get('refLine')))
        print('   summary cols:', [c['label'] for c in ds['summary']['columns']])
        print('   monthFirst:', ds['monthFirst'])
        print('   note:', ds['note'])
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
        f.write('// 自动生成，勿手改。build_mill_order_charts.py @ ' + stamp + '\n')
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
