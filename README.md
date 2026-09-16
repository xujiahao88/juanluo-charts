# 钢材 · 分品种季节性图谱

交互式站点，与 [iron-ore-charts](https://github.com/xujiahao88/iron-ore-charts)（铁矿石发运·到港）**同构**：同一套 ECharts 前端、同一套数据集契约。

## 看板结构
站内含 **11 个 tab**：

| tab | 内容 | 类型 |
|-----|------|------|
| 螺纹 | 卷螺大样本·螺纹（8 区域 × 5 指标） | ECharts 季节图 |
| 热卷 | 卷螺大样本·热卷（8 区域 × 5 指标） | ECharts 季节图 |
| 钢银库存 | 钢银全国城市库存 | ECharts 季节图 |
| 带钢 | 唐宋管带数据库·带钢（周度基本面 + 基准价） | ECharts 季节图 |
| 焊管 | 唐宋管带数据库·焊管 | ECharts 季节图 |
| 钢材出港&接单 | 出港量 / 钢厂接单 | ECharts 季节图 |
| 出口-分品种 / 出口-分国别 | 海关出口 | ECharts 季节图 |
| 热卷排产（中联钢） | 中联钢热卷排产 | ECharts 季节图 |
| 钢厂日接单 | 钢厂日度接单 | ECharts 季节图 |
| **建材直供** | 直供&出库日度跟踪（数字表 + 近 5 年多年度叠加图） | **iframe 嵌入静态看板** |

- 每个数据集图按「区块（group）优先」排列，栅格一行 5 张。
- 每张图是**季节图**：x 轴为 366 天日历轴，年份叠加成线
  - 当年（2026）红色 + 圆点标记；上年蓝色平滑线；更早年份灰/虚线浅蓝
  - 日频价格序列统一做「周度降采样」，单数据集最多保留近 12 年
- 顶部工具：年份开关、连断点、Y 轴含 0、同行同 Y、全部展开

## 建材直供 tab（iframe 嵌入）
「建材直供」不是 ECharts 数据集，而是一个**独立的静态看板页**，用 iframe 原样嵌入，
因此它自身的版式（深蓝页头 + 数字表 + matplotlib 多年度叠加图）与独立站完全一致，且样式互不干扰。

- 页面文件：`zhigong/index.html` + `zhigong/page1_直供出库.png` + `zhigong/page2_直供占比.png`
- 数据集声明位于 `assets/app.js` 顶部的 `EMBED_DATASETS` 常量（**刻意不写进 `data/data.js` 与
  `data/meta.json`**，这样任何 `build_*.py` 重建数据都不会把它冲掉）。其中的 `asOf` 显示在顶部信息栏，
  更新直供数据时同步改这一处即可。
- iframe 高度由 `app.js` 的 `fitEmbedHeight()` 读取内部页面 `scrollHeight` 自适应（需要 http/https
  同源环境；直接以 `file://` 打开时浏览器禁止跨文档读取，会退回 `min-height`）。
- 更新直供数据：由建材直供流水线（`WorkBuddy/…/run_pipeline.py`）生成 dashboard 后，
  把 `index.html` 与两张 PNG 覆盖到本目录 `zhigong/`，再按下方部署命令推送。

## 数据源
- **螺纹 / 热卷**：`卷螺大样本.xlsm` 的【手抄】螺纹大样本 / 【手抄】热卷大样本 主表（周度）
- **带钢 / 焊管 / 管材**：`唐宋管带数据库.xlsx`（唐宋数据，日/周频，含产量、开工率、库存、表需、利润、基准价、焊管、管厂库存等）

## 重新生成数据
更新源 Excel 后：
```
# 螺纹 + 热卷
python scripts/build_juanluo_charts.py
# 带钢 + 焊管 + 管材（从唐宋管带数据库.xlsx）
python scripts/build_daiguan_charts.py
```
`build_daiguan_charts.py` 会读取既有 螺纹/热卷 json 并合并重写 `data/data.js` 与 `data/meta.json`。

## 本地预览
```
python -m http.server 8002
# 打开 http://127.0.0.1:8002
```

## 部署
```
python deploy_repo.py juanluo-charts index.html assets/app.js assets/style.css \
  assets/vendor/echarts.min.js data/data.js data/meta.json \
  data/juanluo_luowen.json data/juanluo_rejuan.json \
  data/daiguan.json data/hanguan.json data/guancai.json \
  zhigong/index.html zhigong/page1_直供出库.png zhigong/page2_直供占比.png \
  scripts/build_juanluo_charts.py scripts/build_daiguan_charts.py README.md
```
推送走 GitHub git-data API（需 classic PAT，环境变量 `GITHUB_PAT`）。
接 Cloudflare Pages：连本仓库 → 生产分支 `main` → 构建命令留空、输出目录留空（根即 index.html）。
站点地址：https://juanluo-charts.pages.dev

---
内容由 AI 基于公开与用户数据整理，仅供参考，不构成投资建议。
