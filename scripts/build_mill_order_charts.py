# -*- coding: utf-8 -*-
"""
build_mill_order_charts.py — 「钢厂日接单」日频时间序列数据集生成器

源：《钢厂日接单量统计.xlsx》的「接单」sheet（唐宋/机构日度钢厂接单跟踪）。
  结构：
    第1行（日产）：B=日产标签，C..J = 8 家钢厂日产（万吨/日），K = 合计(21.6)
    第2行（表头）：B=日期/接单量，C..J = 8 家钢厂接单量，
                  K=平均利润率，L=总接单，M=接单率（=总接单/日产合计）
    第3行起：每日数据（2024-01 起 ~ 今日）

产出 juanluo-charts 一个 dataset「钢厂日接单」(id=mill_order)：
  · 连续日期轴（YYYY-MM-DD），与现有「季节性年对比」模型不同 → ds.axisType='date'
  · 图表：
      「钢厂日接单（5日均值）」组：
        ① 各钢厂日接单（除日照）— 7 家钢厂 5 日均值多线（日照量级 35 万远大于其他 2-6 万，单放）
        ② 日照日接单 — 5 日均值 + 日产(4.0) 参考线
      「综合指标」组：
        ③ 平均利润率 — 原始日频 + 盈亏平衡(0) 参考线
        ④ 总接单（5日均值）— + 日产合计(21.6) 参考线
        ⑤ 接单率（5日均值）— 总接单/日产合计（覆盖天数）
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

# 7 家「小量级」钢厂配色（除日照外，量级 0-6 万，可同图比较）
MILL_COLORS = {
    '纵横': '#2563eb', '中铁': '#16a34a', '安丰': '#d97706', '燕钢': '#dc2626',
    '瑞丰': '#7c3aed', '新东海': '#0891b2', '东华': '#db2777',
}
RZ_COLOR = '#0f766e'          # 日照（单独图）


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
    """尾部窗口移动平均（窗口内非空值求均值；窗口空则 None）。"""
    out = []
    for i in range(len(vals)):
        lo = max(0, i - window + 1)
        seg = [v for v in vals[lo:i + 1] if v is not None]
        out.append(round(sum(seg) / len(seg), 3) if seg else None)
    return out


def parse(wb):
    """返回 (mills, daily_prod, daily_prod_total, dates, data)
    mills: [name,...] 顺序与表头一致
    daily_prod: {name: 日产}
    daily_prod_total: float
    dates: [datetime,...] 升序
    data: { 'mill':{date:val}, 'profit':{date:val}, 'total':{date:val}, 'rate':{date:val} }
    """
    ws = wb['接单']
    rows = list(ws.iter_rows(values_only=True))

    # 找 日产行 与 表头行
    prod_row = None
    header_row = None
    for r in rows:
        if r[1] is not None and '日产' in str(r[1]):
            prod_row = r
        elif r[1] is not None and '日期' in str(r[1]):
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

    # 数据行：B(列索引1)=日期字符串「YYYY-MM-DD 00:00:00」，C..J=各钢厂，
    #          K=利润率, L=总接单, M=接单率
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

    axis = ['%04d-%02d-%02d' % (d.year, d.month, d.day) for d in dates]

    def align(series_map):
        return [series_map.get(d) for d in dates]

    # 季度起始刻度（每季首个数据点）+ 第 0 个
    date_ticks = [0]
    last_q = None
    for i, d in enumerate(dates):
        q = d.year * 4 + (d.month - 1) // 3
        if q != last_q:
            if i != 0:
                date_ticks.append(i)
            last_q = q

    charts = []

    # —— 组1：钢厂日接单（5日均值）——
    small = [m for m in mills if m != '日照']
    small_series = []
    for m in small:
        raw = align(data['mill'][m])
        ma = moving_avg(raw, 5)
        small_series.append({
            'name': m, 'color': MILL_COLORS.get(m, '#64748b'),
            'width': 1.3, 'dash': 'solid', 'smooth': True, 'marker': 'none',
            'data': ma,
        })
    charts.append({
        'key': 'mill_order-combined',
        'title': '各钢厂日接单（除日照，5日均值）',
        'type': 'line', 'axis': 0,
        'group': '钢厂日接单（5日均值）',
        'tag': '5日均',
        'series': small_series,
    })

    # 日照单图（量级远大于其他）
    rz_raw = align(data['mill']['日照'])
    rz_ma = moving_avg(rz_raw, 5)
    charts.append({
        'key': 'mill_order-rz',
        'title': '日照日接单（5日均值）',
        'type': 'line', 'axis': 0,
        'group': '钢厂日接单（5日均值）',
        'tag': '5日均',
        'series': [{
            'name': '日照', 'color': RZ_COLOR,
            'width': 1.6, 'dash': 'solid', 'smooth': True, 'marker': 'none',
            'data': rz_ma,
        }],
        'refLine': daily_prod.get('日照'),
        'refLabel': '日产 %.1f' % (daily_prod.get('日照') or 0),
    })

    # —— 组2：综合指标 ——
    profit_raw = align(data['profit'])
    total_raw = align(data['total'])
    rate_raw = align(data['rate'])

    charts.append({
        'key': 'mill_order-profit',
        'title': '平均利润率（日频）',
        'type': 'line', 'axis': 0,
        'group': '综合指标',
        'tag': '日频',
        'series': [{
            'name': '平均利润率', 'color': '#1f2937',
            'width': 1.3, 'dash': 'solid', 'smooth': True, 'marker': 'none',
            'data': profit_raw,
        }],
        'refLine': 0,
        'refLabel': '盈亏平衡',
    })

    total_ma = moving_avg(total_raw, 5)
    charts.append({
        'key': 'mill_order-total',
        'title': '总接单（5日均值）',
        'type': 'line', 'axis': 0,
        'group': '综合指标',
        'tag': '5日均',
        'series': [{
            'name': '总接单', 'color': '#0b6bcb',
            'width': 1.6, 'dash': 'solid', 'smooth': True, 'marker': 'none',
            'data': total_ma,
        }],
        'refLine': daily_prod_total,
        'refLabel': '日产合计 %.1f' % (daily_prod_total or 0),
    })

    rate_ma = moving_avg(rate_raw, 5)
    charts.append({
        'key': 'mill_order-rate',
        'title': '接单率（5日均值，总接单/日产合计）',
        'type': 'line', 'axis': 0,
        'group': '综合指标',
        'tag': '5日均',
        'series': [{
            'name': '接单率', 'color': '#b45309',
            'width': 1.6, 'dash': 'solid', 'smooth': True, 'marker': 'none',
            'data': rate_ma,
        }],
    })

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

    # 汇总列：8 钢厂 + 平均利润率 + 总接单 + 接单率
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

    # 日产参考说明
    prod_parts = ['%s %.1f' % (m, daily_prod.get(m) or 0) for m in mills]
    note = '日产参考（万吨/日）：' + '｜'.join(prod_parts) + '（合计 %.1f）' % (daily_prod_total or 0)

    as_of = '%04d-%02d-%02d' % (last.year, last.month, last.day)
    return {
        'id': 'mill_order', 'name': '钢厂日接单',
        'axisType': 'date', 'axes': [axis], 'dateTicks': date_ticks,
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
        print('== %s(%s) | asOf %s | charts %d | axis %d | dateTicks %d'
              % (ds['name'], ds['id'], ds['asOf'], len(ds['charts']),
                 len(ds['axes'][0]), len(ds['dateTicks'])))
        for c in ds['charts']:
            print('   [%s] %s | series=%d tag=%s refLine=%s'
                  % (c.get('type'), c['title'], len(c.get('series', [])),
                     c.get('tag'), c.get('refLine')))
        print('   summary cols:', [c['label'] for c in ds['summary']['columns']])
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
