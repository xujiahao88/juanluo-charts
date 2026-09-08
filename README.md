# 卷螺大样本 · 分区域季节性图谱

交互式站点，与 [iron-ore-charts](https://github.com/xujiahao88/iron-ore-charts)（铁矿石发运·到港）**同构**：同一套 ECharts 前端、同一套数据集契约。

## 看板结构
- **两个数据集 tab**：`螺纹` / `热卷`
- 每个数据集 **32 张图** = 8 个区域（东北/华北/华东/华南/华中/西北/西南/合计）× 4 个指标（周产量 / 钢厂库存 / 社库 / 表需）
- 图按「区域优先」排列，栅格一行 4 张 → **一个区域正好占一行**
- 每张图是**季节图**：x 轴为 366 天日历轴，年份叠加成线
  - 当年（2026）红色 + 圆点标记；上年蓝色平滑线；更早年份灰/虚线浅蓝
- **汇总表**：列 = 区域分组 × 指标，行 = 本期 / 上期 / 环比 / 同比 / 同比%，单位 **万吨**
- 顶部工具：年份开关、连断点、Y 轴含 0、同行同 Y、全部展开

## 数据源
`卷螺大样本.xlsm` 的【手抄】螺纹大样本 / 【手抄】热卷大样本 主表（周度）。
列偏移：周产量 = 6+idx*4，钢厂库存 = 7+idx*4，社库 = 36+idx，表需 = 52+idx
（idx：合计=7，东北=0，华北=1，华东=2，华南=3，华中=4，西北=5，西南=6）

## 重新生成数据
更新源 Excel 后：
```
python scripts/build_juanluo_charts.py
```
产出 `data/juanluo_luowen.json`、`data/juanluo_rejuan.json`、`data/data.js`、`data/meta.json`。

## 本地预览
```
python -m http.server 8002
# 打开 http://127.0.0.1:8002
```

## 部署
```
python deploy_repo.py juanluo-charts index.html assets/app.js assets/style.css \
  assets/vendor/echarts.min.js data/data.js data/meta.json \
  data/juanluo_luowen.json data/juanluo_rejuan.json scripts/build_juanluo_charts.py README.md
```
推送走 GitHub git-data API（需 classic PAT，环境变量 `GITHUB_PAT`）。
接 Cloudflare Pages：连本仓库 → 生产分支 `main` → 构建命令留空、输出目录留空（根即 index.html）。

---
内容由 AI 基于公开与用户数据整理，仅供参考，不构成投资建议。
