# 钢材 · 分品种季节性图谱

交互式站点，与 [iron-ore-charts](https://github.com/xujiahao88/iron-ore-charts)（铁矿石发运·到港）**同构**：同一套 ECharts 前端、同一套数据集契约。

## 看板结构
站内含 **5 个数据集 tab**，覆盖主要钢材品种：

| tab | 内容 | 区块（group） |
|-----|------|---------------|
| 螺纹 | 卷螺大样本·螺纹（8 区域 × 5 指标） | 合计/华北/华东/华南/华中/西北/西南/东北 |
| 热卷 | 卷螺大样本·热卷（8 区域 × 5 指标） | 同上 |
| 带钢 | 唐宋管带数据库·带钢（周度基本面 + 基准价） | 带钢产量 / 开工率 / 库存 / 表需 / 利润 / 基准价 |
| 焊管 | 唐宋管带数据库·焊管 | 焊管产量&开工率 / 管厂库存 |
| 管材 | 唐宋管带数据库·管材 | 管材基准价 |

- 每个数据集图按「区块（group）优先」排列，栅格一行 5 张。
- 每张图是**季节图**：x 轴为 366 天日历轴，年份叠加成线
  - 当年（2026）红色 + 圆点标记；上年蓝色平滑线；更早年份灰/虚线浅蓝
  - 日频价格序列统一做「周度降采样」，单数据集最多保留近 12 年
- 顶部工具：年份开关、连断点、Y 轴含 0、同行同 Y、全部展开

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
  scripts/build_juanluo_charts.py scripts/build_daiguan_charts.py README.md
```
推送走 GitHub git-data API（需 classic PAT，环境变量 `GITHUB_PAT`）。
接 Cloudflare Pages：连本仓库 → 生产分支 `main` → 构建命令留空、输出目录留空（根即 index.html）。
站点地址：https://juanluo-charts.pages.dev

---
内容由 AI 基于公开与用户数据整理，仅供参考，不构成投资建议。
