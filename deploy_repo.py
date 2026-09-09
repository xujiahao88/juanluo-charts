#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
deploy_repo.py — 把本地静态站推到新建/已有的 GitHub 仓库（git-data API，绕过本地 git 客户端）。
用法：
  python deploy_repo.py <repo名> <相对路径1> <相对路径2> ...
例：
  python deploy_repo.py juanluo-seasonal index.html assets/app.js assets/style.css data/data.js scripts/build_seasonal.py README.md
说明：
  · 仓库不存在则自动建（auto_init 产生初始 README，main 分支），已存在则直接在其上追加 commit。
  · 仅重建并 PATCH 指定文件，其余文件(如初始 README)保留。
  · 2026-09-03 起代理已关，api.github.com 直连。
"""
import os, base64, json, sys, urllib.request, urllib.error, concurrent.futures

TOKEN = os.environ.get('GITHUB_PAT')
if not TOKEN:
    raise SystemExit('GITHUB_PAT 环境变量缺失')
OWNER = 'xujiahao88'
ROOT = os.path.dirname(os.path.abspath(__file__))
API = 'https://api.github.com'


def call(method, path, data=None, base=None):
    url = (base or API) + path
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header('Authorization', 'Bearer ' + TOKEN)
    req.add_header('Accept', 'application/vnd.github+json')
    req.add_header('User-Agent', 'deploy-repo')
    if body:
        req.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode()) if r.read else None
    except urllib.error.HTTPError as e:
        msg = e.read().decode(errors='replace')
        return {'_http': e.code, '_msg': msg}


def ensure_repo(name):
    res = call('POST', '/user/repos', {'name': name, 'auto_init': True, 'private': False})
    if isinstance(res, dict) and res.get('_http') == 422:
        print('[ok] 仓库已存在，跳过创建:', name)
        return
    if isinstance(res, dict) and 'full_name' in res:
        print('[ok] 已创建仓库:', res['full_name'])
        return
    if isinstance(res, dict) and res.get('_http'):
        raise SystemExit(f'建仓失败 HTTP {res["_http"]}: {res["_msg"]}')
    raise SystemExit('建仓返回异常: ' + str(res))


CACHE = os.path.join(ROOT, '.deploy_cache.json')


def _load_cache():
    try:
        with open(CACHE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_cache(c):
    try:
        with open(CACHE, 'w', encoding='utf-8') as f:
            json.dump(c, f)
    except Exception:
        pass


def blob_sha(data: bytes) -> str:
    """git blob sha1（与 GitHub 一致）：sha1('blob <len>\\0' + content)"""
    import hashlib
    h = hashlib.sha1()
    h.update(b'blob %d\0' % len(data))
    h.update(data)
    return h.hexdigest()


def make_blob(rel, content=None):
    abs_path = os.path.join(ROOT, rel)
    if content is None:
        with open(abs_path, 'rb') as f:
            content = f.read()
    b64 = base64.b64encode(content).decode()
    blob = call('POST', f'/repos/{OWNER}/{REPO}/git/blobs', {'content': b64, 'encoding': 'base64'})
    return {'path': rel, 'mode': '100644', 'type': 'blob', 'sha': blob['sha']}


if __name__ == '__main__':
    args = sys.argv[1:]
    flags = {a for a in args if a.startswith('--')}
    pos = [a for a in args if not a.startswith('--')]
    if len(pos) < 2:
        raise SystemExit('用法: python deploy_repo.py <repo名> <相对路径> ... [--force]')
    REPO = pos[0]
    files = pos[1:]
    force = '--force' in flags

    # ---- 第 0 层：本地缓存命中则零 API 调用直接跳过 ----
    # （单机推送场景可靠；远程若被别处改动，用 --force 强制走网络比对）
    local = {}
    for rel in files:
        with open(os.path.join(ROOT, rel), 'rb') as f:
            local[rel] = f.read()
    cache_all = _load_cache()
    cache = cache_all.get(REPO, {})
    if not force and all(cache.get(rel) == blob_sha(local[rel]) for rel in files):
        print(f'SKIP 本地缓存命中：{len(files)} 个文件均与上次推送一致（0 次 API 调用）')
        raise SystemExit(0)

    ensure_repo(REPO)

    ref = call('GET', f'/repos/{OWNER}/{REPO}/git/refs/heads/main')
    if isinstance(ref, dict) and ref.get('_http'):
        raise SystemExit(f'取 main 分支失败 HTTP {ref["_http"]}: {ref.get("_msg")}')
    base_sha = ref['object']['sha']

    # ---- 提速：先取远程 tree 对比 blob sha，未变化的文件/整体跳过 ----
    # 推送固定开销 = 6 次串行 API 往返（~4s）；钢银周更时卷螺两个大 json（640KB）根本没变，
    # 全未变时直接省掉 tree/commit/PATCH 三次往返（幂等重跑 4.3s → ~2s）。
    tree_remote = call('GET', f'/repos/{OWNER}/{REPO}/git/trees/{base_sha}?recursive=1')
    remote_sha = {}
    if isinstance(tree_remote, dict) and not tree_remote.get('_http'):
        remote_sha = {it['path']: it['sha'] for it in tree_remote.get('tree', [])
                      if it.get('type') == 'blob'}

    local = {}
    for rel in files:
        with open(os.path.join(ROOT, rel), 'rb') as f:
            local[rel] = f.read()
    changed = [rel for rel in files if blob_sha(local[rel]) != remote_sha.get(rel)]

    if not changed:
        print(f'SKIP 内容未变化，跳过推送（{len(files)} 个文件均与远程一致）')
        # 远程已一致，补上缓存 → 下次零 API 调用
        cache.update({rel: blob_sha(local[rel]) for rel in files})
        cache_all[REPO] = cache
        _save_cache(cache_all)
        raise SystemExit(0)

    if len(changed) < len(files):
        print(f'[fast] 仅 {len(changed)}/{len(files)} 个文件有变化：' + ', '.join(changed))

    tree_items = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(changed))) as ex:
        for item in ex.map(lambda r: make_blob(r, local[r]), changed):
            tree_items.append(item)

    new_tree = call('POST', f'/repos/{OWNER}/{REPO}/git/trees', {'base_tree': base_sha, 'tree': tree_items})
    new_commit = call('POST', f'/repos/{OWNER}/{REPO}/git/commits', {
        'message': os.environ.get('COMMIT_MSG', 'feat: 卷螺分区域季节性站点初始化'),
        'tree': new_tree['sha'],
        'parents': [base_sha],
    })
    call('PATCH', f'/repos/{OWNER}/{REPO}/git/refs/heads/main', {'sha': new_commit['sha']})
    print('PUSHED', new_commit['sha'], 'files:', ', '.join(changed))

    # 更新本地缓存（含本次未变的文件，保证下次命中）
    cache.update({rel: blob_sha(local[rel]) for rel in files})
    cache_all[REPO] = cache
    _save_cache(cache_all)
