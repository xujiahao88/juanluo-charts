# -*- coding: utf-8 -*-
"""
shot_juanluo.py — juanluo-charts 长图截图（Edge headless 单图工具）

用法：
  python shot_juanluo.py --ds gangyin_stock --out "C:/path/to/long.png"
  python shot_juanluo.py --ds juanluo_luowen  --out "C:/path/to/long.png" --width 1920 --height 3600

说明：
  · juanluo-charts 前端支持 ?ds=<id> 切 tab，?shot=1 截图模式（全量渲染、不懒加载）
  · 这里用 Edge headless 一次性出图（比铁矿站 html_to_image.py 简单，10 张图+汇总表够用）
  · Edge 路径自动探测（程序文件/Microsoft Edge + Program Files (x86)/Microsoft Edge）
"""
import argparse
import os
import shutil
import subprocess


def find_edge():
    for p in ("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
              "C:/Program Files/Microsoft/Edge/Application/msedge.exe"):
        if os.path.isfile(p):
            return p
    return shutil.which("msedge") or shutil.which("msedge.exe")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ds", required=True, help="数据集 id (如 gangyin_stock / juanluo_luowen)")
    ap.add_argument("--out", required=True, help="输出 PNG 绝对路径")
    ap.add_argument("--site", default=r"C:\Users\Administrator\juanluo-charts",
                    help="本地站点根目录（含 index.html）")
    ap.add_argument("--width", type=int, default=1920)
    ap.add_argument("--height", type=int, default=2200, help="钢银 2200、卷螺 3600+")
    ap.add_argument("--url", default=None, help="自定义 URL（覆盖默认 site 构造）")
    args = ap.parse_args()

    edge = find_edge()
    if not edge:
        raise SystemExit("未找到 Edge，无法截图")

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    # 注：某些沙箱环境禁用 os.remove（回收站不可用 → SAFE_DELETE_FAIL_CLOSED 抛错）。
    # 删除失败无所谓 —— Edge --screenshot 会直接覆盖同名文件。
    if os.path.exists(args.out):
        try:
            os.remove(args.out)
        except Exception:
            pass

    if args.url:
        url = args.url
    else:
        index = os.path.join(args.site, "index.html").replace("\\", "/")
        url = "file:///" + index + f"?ds={args.ds}&shot=1"

    cmd = [
        edge, "--headless=new", "--disable-gpu", "--hide-scrollbars",
        "--virtual-time-budget=9000",
        f"--screenshot={args.out}",
        f"--window-size={args.width},{args.height}",
        url,
    ]
    print("$ " + " ".join(cmd[:5]) + f" … --window-size={args.width},{args.height} {url}")
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        raise SystemExit(f"Edge 截图失败 rc={r.returncode}: {(r.stderr or r.stdout)[:300]}")
    if not os.path.exists(args.out) or os.path.getsize(args.out) < 5000:
        raise SystemExit(f"截图未生成或过小：{args.out}")
    sz = os.path.getsize(args.out)
    print(f"✅ 截图完成: {args.out} ({sz:,} bytes, {args.width}x{args.height})")


if __name__ == "__main__":
    main()
