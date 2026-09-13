"""[共享] dataset 合并器 —— 防止某个流水线写 data.js 时把别家 tab 冲掉。

背景（2026-09-14 踩坑）：
  juanluo-charts 是「多流水线共用一个站点」：螺纹/热卷(juanluo)、钢银(gangyin)、
  带钢/焊管/出港/出口、PSI 排产 各自由不同 skill 维护，但都写同一份
  data/data.js + data/meta.json（前端 app.js 只认 window.CHART_DATA.datasets）。
  谁最后写谁定义 tab 全集 —— 早期 build_gangyin/build_juanluo/build_psi 各自
  只写自己的 2~3 个 dataset，于是每周把别人的 tab 全部冲掉
  （实测 2026-09-14：线上只剩 螺纹/热卷/钢银 3 个 tab，
   带钢/焊管/出港/出口品种/出口国别/PSI 6 个 tab 消失）。

用法（build_*.py 里）：
    from datasets_merge import merge_datasets
    regen = head + [ds]                       # 本次真正重生成的 dataset
    datasets = merge_datasets(regen, prefer=PREFER_ORDER)

返回顺序：① meta.json 现有顺序（保序）② prefer 里命中的补位
③ 其余磁盘上存在但 meta 没记的（字母序）④ 本次 regen 里 still 缺的。
不存在的 json 只 warn 不报错。
"""

import glob
import json
import os
import sys

DATA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')

# data/ 里不是 dataset 的文件
NON_DATASET = {'meta', 'data'}
# 已退役的 dataset（不要自动复活；前端 assets/app.js 无对应钩子）
RETIRED = {'guancai'}

# 站点 tab 的期望顺序（除本流水线自己负责的之外）；命中者按此顺序补位
PREFER_ORDER = ['daiguan', 'hanguan', 'chugang', 'export_variety', 'export_country',
                'psi_plan']


def merge_datasets(regen, prefer=None, data_dir=None, verbose=True):
    """把本次重生成的 dataset 与磁盘上其他 tab 的 dataset 合并。

    regen: list[dict] —— 本次重新生成的 dataset（含 id）
    prefer: list[str] —— 其他 tab 的期望顺序（默认 PREFER_ORDER）
    """
    dirp = data_dir or DATA
    regen_ids = [d['id'] for d in regen]

    def _log(msg):
        if verbose:
            print(msg)

    # ① meta.json 现有顺序
    order = []
    mp = os.path.join(dirp, 'meta.json')
    if os.path.exists(mp):
        try:
            with open(mp, encoding='utf-8') as f:
                order = [d['id'] for d in json.load(f).get('datasets', [])]
        except Exception as e:
            _log('[warn] 读 meta.json 失败：%s' % e)

    # ② 磁盘上存在、但 meta 没记录的 dataset（meta 被冲掉时用来恢复）
    on_disk = []
    for p in sorted(glob.glob(os.path.join(dirp, '*.json'))):
        did = os.path.splitext(os.path.basename(p))[0]
        if did in NON_DATASET or did in RETIRED:
            continue
        if did in regen_ids or did in order:
            continue
        on_disk.append(did)

    for did in (PREFER_ORDER if prefer is None else prefer):
        if did in on_disk:
            order.append(did)
            on_disk.remove(did)
    order += sorted(on_disk)

    # ③ 本次 regen 的补到末尾（理论上已在 meta 顺序里）
    for did in regen_ids:
        if did not in order:
            order.append(did)

    out = []
    by_id = {d['id']: d for d in regen}
    for did in order:
        if did in by_id:
            out.append(by_id[did])
            continue
        p = os.path.join(dirp, did + '.json')
        if not os.path.exists(p):
            _log('[warn] 缺少 %s.json，跳过' % did)
            continue
        try:
            with open(p, encoding='utf-8') as f:
                d = json.load(f)
            if isinstance(d, dict) and d.get('id') == did:
                out.append(d)
                _log('[ok] 保留既有 tab %s（%s）' % (did, d.get('name')))
            else:
                _log('[warn] %s.json 不是 dataset，跳过' % did)
        except Exception as e:
            _log('[warn] 读 %s.json 失败：%s' % (did, e))
    return out


if __name__ == '__main__':
    a = merge_datasets([])
    print('合并后 tab 顺序：', [d['id'] for d in a])
    print('共 %d 个 dataset' % len(a))
    sys.exit(0)
