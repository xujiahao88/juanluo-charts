/* 铁矿石发运·到港 季节性图谱
 * 数据来自 data/data.js（内联，file:// 可直接打开），缺失时回退 fetch data/*.json
 * URL 参数：?ds=global|aubra|arrival 指定数据集；?shot=1 截图模式（全量渲染，不懒加载）
 */
(function () {
  'use strict';

  var S = {
    data: null,
    dsId: null,
    items: [],
    io: null,
    shot: false,
    hidden: {},        // 年份 -> true 表示隐藏
    years: [],
    opt: { connect: true, zero: false, sync: false },
    expanded: false
  };

  // 各数据集「区块内一行几张图」（默认 5 张，见 style.css .grid.rgrid）
  var GRID_COLS = { psi_plan: 3, daiguan: 3, chugang: 3 };

  // 横坐标按「1月…12月」显示（每月 1 号一个刻度）的数据集
  var MONTH_AXIS = { daiguan: 1, chugang: 1, juanluo_luowen: 1, juanluo_rejuan: 1 };

  var DASH = {
    solid: 'solid', dash: 'dashed', sysDash: 'dashed',
    dot: 'dotted', sysDot: 'dotted',
    dashDot: [6, 3, 2, 3], sysDashDot: [6, 3, 2, 3]
  };

  function $(id) { return document.getElementById(id); }

  function dashOf(v) { return DASH[v] || 'solid'; }

  /* ---------------- 数据加载 ---------------- */
  function loadData(cb) {
    if (window.CHART_DATA && window.CHART_DATA.datasets) {
      S.data = window.CHART_DATA;
      cb();
      return;
    }
    fetch('data/meta.json').then(function (r) { return r.json(); }).then(function (meta) {
      var jobs = meta.datasets.map(function (d) {
        return fetch('data/' + d.id + '.json').then(function (r) { return r.json(); });
      });
      return Promise.all(jobs).then(function (arr) {
        S.data = { updated: meta.updated, datasets: arr };
        cb();
      });
    }).catch(function (e) {
      $('loading').textContent = '数据加载失败：' + e.message + '（请先运行 scripts/extract.py）';
    });
  }

  /* ---------------- 图表配置 ---------------- */
  function buildOption(chart, ds, yRange) {
    var axis = ds.axes[chart.axis];
    var series = chart.series.filter(function (s) { return !S.hidden[s.name]; });
    var maxYear = series.length ? series[series.length - 1].name : '';

    return {
      animation: false,
      title: {
        text: chart.title, left: 'center', top: 6,
        textStyle: { fontSize: 13, fontWeight: 600, color: '#1f2937' }
      },
      tooltip: {
        trigger: 'axis', confine: true,
        textStyle: { fontSize: 11 },
        axisPointer: { type: 'line', lineStyle: { color: '#cbd5e1', width: 1 } }
      },
      legend: {
        top: 27, itemGap: 6, itemWidth: 16, itemHeight: 8,
        textStyle: { fontSize: 9, color: '#6b7280' },
        // legend 图标默认不继承 series 的线型，必须逐项显式带 lineStyle，
        // 否则 2022/2023 的虚线系列在图例里仍显示成实线。
        data: series.map(function (s) {
          return {
            name: s.name,
            itemStyle: { color: s.color },
            lineStyle: {
              color: s.color,
              type: dashOf(s.dash),
              width: s.width || 1.5
            }
          };
        })
      },
      grid: { left: 54, right: 20, top: 54, bottom: 32 },
      xAxis: {
        type: 'category', data: axis, boundaryGap: false,
        axisLine: { lineStyle: { color: '#d5dbe6' } },
        axisTick: { show: false },
        axisLabel: {
          // 月度轴（带钢/出港/螺纹/热卷）标签多，字号调小到 8
          fontSize: MONTH_AXIS[S.dsId] ? 8 : 9, color: '#94a3b8',
          // 月度轴：横坐标显示「1月…12月」，每月 1 号一个刻度
          formatter: function (v) {
            if (MONTH_AXIS[S.dsId]) {
              var mm = /^(\d{2})-/.exec(v);
              if (mm) return parseInt(mm[1], 10) + '月';
              return v;
            }
            return v;
          },
          // 同时兼容两种轴，避免标签交叠且只留季度刻度：
          //  · 月份轴（'1月'..'12月' 或 '01'..'12'，月度数据如中联钢排产）→ 季度首月
          //  · 日历轴（MM-DD，周度数据如钢材各品种/钢银）→ 季度首月 1 号 + 首点
          interval: function (i, v) {
            // 中联钢排产（psi_plan）：纯数字 1–12 月份轴，12 个刻度全显示
            if (S.dsId === 'psi_plan') return true;
            // 月度轴（带钢/出港/螺纹/热卷）：每月 1 号显示
            if (MONTH_AXIS[S.dsId]) return /-01$/.test(v);
            var m = /^\s*(\d{1,2})\s*月?\s*$/.exec(v);
            if (m) { var n = parseInt(m[1], 10); return n === 1 || n === 4 || n === 7 || n === 10; }
            return i === 0 || /^(01|04|07|10)-01$/.test(v);
          }
        }
      },
      yAxis: {
        type: 'value',
        scale: !S.opt.zero,
        min: yRange ? yRange.min : null,
        max: yRange ? yRange.max : null,
        splitLine: { lineStyle: { color: '#eef1f6' } },
        axisLine: { show: false },
        axisTick: { show: false },
        axisLabel: { fontSize: 9, color: '#94a3b8' }
      },
      series: series.map(function (s) {
        return {
          name: s.name,
          type: 'line',
          data: s.data,
          connectNulls: S.opt.connect,
          smooth: !!s.smooth,
          showSymbol: !!s.marker && s.marker !== 'none',
          symbol: 'circle',
          symbolSize: 3.5,
          lineStyle: { width: s.width || 1.5, type: dashOf(s.dash), color: s.color },
          itemStyle: { color: s.color },
          emphasis: { focus: 'series' },
          z: s.name === maxYear ? 6 : 2
        };
      })
    };
  }

  /* ---------------- 汇总表 ---------------- */
  function fmtNum(v, suffix) {
    // 绝对值 >=100 整数，否则保留 1 位小数；自动千分位
    var abs = Math.abs(v);
    var fixed = abs >= 100 ? 0 : (abs >= 10 ? 1 : 2);
    var s = v.toFixed(fixed);
    var parts = s.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.') + (suffix || '');
  }

  function renderSummaryTable(ds) {
    var s = ds.summary;
    if (!s || !s.columns || !s.columns.length) return null;
    var asPct = !!s.asPercent;   // 盈利率：比例值→前端 ×100 显示百分比

    // 合并分组（连续相同 group 名算一段）
    var groups = [];
    var cur = null;
    s.columns.forEach(function (c) {
      if (!cur || cur.name !== c.group) {
        cur = { name: c.group, start: groups.length === 0 ? 0 : cur.end + 1, end: 0 };
        cur.start = cur.end = s.columns.indexOf(c);
        groups.push(cur);
      } else {
        cur.end = s.columns.indexOf(c);
      }
    });

    var wrap = document.createElement('div');
    wrap.className = 'summary';

    var head = document.createElement('div');
    head.className = 'summary-head';
    head.innerHTML =
      '<strong>' + ds.name + ' · 本期汇总</strong>' +
      '<span class="summary-dates">本期 ' + s.currentWeek +
      ' · 上期 ' + s.previousWeek + ' · 单位：' + (s.unit || '万吨') + '</span>';
    wrap.appendChild(head);

    var scroll = document.createElement('div');
    scroll.className = 'summary-scroll';

    var table = document.createElement('table');
    table.className = 'summary-table';

    // 列宽策略：
    //   全球发运 / 分省份盈利率 / 分省份铁水 → 固定像素宽 + 横向滚动（列多不可压扁）
    //     全球用长图截图量出的像素；省份用 标签列190 + 数据列58（够装数值与省名）
    //   其余（分区域盈利率/铁水、澳巴、到港、15港）→ 等百分比宽（#35 原始行为）
    var GLOBAL_FIT_PX = [120, 109, 84, 85, 90, 91, 91, 92, 106, 90, 90, 127, 90, 91, 109, 109, 109, 91, 84, 84];
    var PX_FIT_IDS = { global: true, profit_province: true, ironwater_province: true, juanluo_luowen: true, juanluo_rejuan: true };
    var colgroup = document.createElement('colgroup');
    var totalCols = s.columns.length + 1;
    if (PX_FIT_IDS[ds.id]) {
      table.classList.add('global-fit');
      if (ds.id === 'global') {
        for (var gi = 0; gi < totalCols; gi++) {
          var gc = document.createElement('col');
          gc.style.width = (GLOBAL_FIT_PX[gi] || 90) + 'px';
          colgroup.appendChild(gc);
        }
      } else {
        var LABEL_W = 190, DATA_W = 58;
        var lc = document.createElement('col');
        lc.style.width = LABEL_W + 'px';
        colgroup.appendChild(lc);
        for (var di = 0; di < s.columns.length; di++) {
          var dc = document.createElement('col');
          dc.style.width = DATA_W + 'px';
          colgroup.appendChild(dc);
        }
      }
    } else {
      var each = (100 / totalCols).toFixed(3);
      for (var ci = 0; ci < totalCols; ci++) {
        var cc = document.createElement('col');
        cc.style.width = each + '%';
        colgroup.appendChild(cc);
      }
    }
    table.appendChild(colgroup);

    // thead: 分组行 + 列名行
    var thead = document.createElement('thead');
    var grpRow = document.createElement('tr');
    grpRow.className = 'grp-row';
    var th0 = document.createElement('th');
    th0.className = 'row-label';
    th0.rowSpan = 2;
    th0.textContent = '指标';
    grpRow.appendChild(th0);
    groups.forEach(function (g) {
      var th = document.createElement('th');
      th.colSpan = g.end - g.start + 1;
      th.className = 'grp grp-' + g.name;
      th.textContent = g.name;
      grpRow.appendChild(th);
    });
    thead.appendChild(grpRow);

    var labRow = document.createElement('tr');
    labRow.className = 'lab-row';
    s.columns.forEach(function (c) {
      var th = document.createElement('th');
      th.textContent = c.label;
      th.title = c.label;
      labRow.appendChild(th);
    });
    thead.appendChild(labRow);
    table.appendChild(thead);

    // tbody: 5 行（本期/上期/环比/累计同比/累计增幅%）
    var ROW_META = {
      '本期':      {},
      '上期':      {},
      '环比':      {signed: true},
      '累计同比':  {signed: true},
      '累计增幅%': {signed: true, suffix: '%'},
      '较年初变化': {signed: true},
      '较年初':    {signed: true},
      '同比':      {signed: true},
      '同比%':     {signed: true, suffix: '%'}
    };
    var rowOrder = s.rowOrder || ['本期', '上期', '环比', '累计同比', '累计增幅%'];
    var rowDefs = rowOrder.map(function (k) {
      // 复合键 '指标·视图'（profit/ironwater 用）：完整键作 label，让4 行字面就不同
      // （如「盈利率·本期（08-28）」「盈利率·较年初」），不再剥掉视图段。
      // 纯键（aubra/global/arrival/portinv 的 本期/上期/环比 等）label = k 自身。
      var seg = k.split('·');
      var view = seg[seg.length - 1];
      var label = k;
      var meta = ROW_META[view] || {};
      var d = { key: k, label: label };
      if (view === '本期') d.date = s.currentWeek;
      if (view === '上期') d.date = s.previousWeek;
      if (meta.signed) d.signed = true;
      if (meta.suffix) d.suffix = meta.suffix;
      if (asPct) d.suffix = (d.suffix || '') + '%';   // 盈利率：比例值→百分比显示
      // 差值行：环比 / 较年初 / 同比 / 累计同比 / 累计增幅% — 加 r-delta class，CSS 给浅灰背景
      d.isDelta = meta.signed === true;
      return d;
    });
    var tbody = document.createElement('tbody');
    rowDefs.forEach(function (rd) {
      var tr = document.createElement('tr');
      tr.className = 'r-' + rd.key + (rd.isDelta ? ' r-delta' : '');
      var lbl = document.createElement('td');
      lbl.className = 'row-label';
      lbl.textContent = rd.label + (rd.date ? ' (' + rd.date + ')' : '');
      tr.appendChild(lbl);

      var vals = s.rows[rd.key] || [];
      s.columns.forEach(function (c, i) {
        var td = document.createElement('td');
        var raw = vals[i];
        if (raw === null || raw === undefined) {
          td.textContent = '—';
          td.className = 'na';
        } else {
          var v = asPct ? raw * 100 : raw;   // 盈利率：比例值 ×100 显示百分比
          var sign = (rd.signed && v > 0) ? '+' : '';
          td.textContent = sign + fmtNum(v, rd.suffix || '');
          if (rd.signed) {
            if (v > 0) td.className = 'up';
            else if (v < 0) td.className = 'down';
          }
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    scroll.appendChild(table);
    wrap.appendChild(scroll);
    return wrap;
  }

  /* ---------------- 渲染 ---------------- */
  function datasetById(id) {
    for (var i = 0; i < S.data.datasets.length; i++) {
      if (S.data.datasets[i].id === id) return S.data.datasets[i];
    }
    return S.data.datasets[0];
  }

  function initChart(item) {
    if (item.inst) return;
    var inst = echarts.init(item.el, null, { renderer: 'canvas' });
    item.inst = inst;
    inst.setOption(buildOption(item.chart, item.ds, null));
  }

  function refreshAll() {
    S.items.forEach(function (it) {
      if (!it.inst) return;
      it.inst.setOption(buildOption(it.chart, it.ds, null), true);
    });
    if (S.opt.sync) syncY();
  }

  /* 同行同 Y：按实际 offsetTop 分组，组内共享 Y 范围 */
  function syncY() {
    var rows = {};
    S.items.forEach(function (it) {
      if (!it.inst || !it.el) return;
      var top = it.card.offsetTop;
      (rows[top] = rows[top] || []).push(it);
    });
    Object.keys(rows).forEach(function (k) {
      var list = rows[k], mn = Infinity, mx = -Infinity;
      list.forEach(function (it) {
        it.chart.series.forEach(function (s) {
          if (S.hidden[s.name]) return;
          s.data.forEach(function (v) {
            if (v === null || v === undefined) return;
            if (v < mn) mn = v;
            if (v > mx) mx = v;
          });
        });
      });
      if (mn === Infinity) return;
      var pad = (mx - mn) * 0.06 || 1;
      var r = { min: S.opt.zero ? 0 : Math.max(0, mn - pad), max: mx + pad };
      list.forEach(function (it) {
        it.inst.setOption({ yAxis: { min: r.min, max: r.max } });
      });
    });
  }

  function renderDataset(dsId) {
    S.dsId = dsId;
    var ds = datasetById(dsId);
    var main = $('main');

    if (S.io) { S.io.disconnect(); S.io = null; }
    S.items.forEach(function (it) { if (it.inst) it.inst.dispose(); });
    S.items = [];

    // 汇总表（在图表网格之前）
    var sumEl = renderSummaryTable(ds);
    main.innerHTML = '';
    main.setAttribute('data-ds', dsId);   // 数据集级样式钩子（如带钢加宽）
    if (sumEl) main.appendChild(sumEl);

    // 按 ch.group 分区块渲染（钢材站：一个区块一组指标，区块内 5 图一行）
    var curGroup = null, grid = null;
    ds.charts.forEach(function (ch) {
      var g = ch.group || '';
      if (g !== curGroup) {
        curGroup = g;
        var host = main;
        if (g) {
          var sec = document.createElement('section');
          sec.className = 'rgroup';
          var h = document.createElement('h2');
          h.className = 'rgroup-title';
          h.textContent = g;
          sec.appendChild(h);
          main.appendChild(sec);
          host = sec;
        }
        grid = document.createElement('div');
        // 区块列数：默认 5 列；按数据集例外（PSI 热卷排产用户要求一排 3 张）
        grid.className = 'grid rgrid' + (GRID_COLS[ds.id] ? (' cols' + GRID_COLS[ds.id]) : '');
        host.appendChild(grid);
      }
      var card = document.createElement('div');
      card.className = 'card' + (S.expanded ? ' wide' : '');

      var head = document.createElement('div');
      head.className = 'card-head';
      head.innerHTML = '<span>' + ch.title + '</span>';
      var tag = document.createElement('span');
      tag.className = 'tag';
      tag.textContent = ch.series.length + ' 年对比';
      head.appendChild(tag);

      var box = document.createElement('div');
      box.className = 'chart';

      card.appendChild(head);
      card.appendChild(box);
      grid.appendChild(card);

      S.items.push({ key: ch.key, chart: ch, ds: ds, el: box, card: card, inst: null });
    });

    // 年份开关（按当前数据集的系列名）
    var years = [];
    if (ds.charts[0]) {
      ds.charts[0].series.forEach(function (s) { years.push({ name: s.name, color: s.color }); });
    }
    S.years = years;
    renderYearToggles();

    if (S.shot) {
      document.body.classList.add('shot');
      S.items.forEach(function (it) { initChart(it); });
      S.items.forEach(function (it) { if (it.inst) it.inst.resize(); });
    } else {
      document.body.classList.remove('shot');
      if ('IntersectionObserver' in window) {
        S.io = new IntersectionObserver(function (entries) {
          entries.forEach(function (e) {
            if (e.isIntersecting) {
              var it = findByEl(e.target);
              if (it) initChart(it);
              S.io.unobserve(e.target);
            }
          });
        }, { rootMargin: '300px' });
        S.items.forEach(function (it) { S.io.observe(it.el); });
      } else {
        S.items.forEach(function (it) { initChart(it); });
      }
    }

    if (S.opt.sync) setTimeout(syncY, 60);
    window.scrollTo(0, 0);
  }

  function findByEl(el) {
    for (var i = 0; i < S.items.length; i++) if (S.items[i].el === el) return S.items[i];
    return null;
  }

  function renderYearToggles() {
    var box = $('yearToggles');
    box.innerHTML = '';
    S.years.forEach(function (y) {
      var on = !S.hidden[y.name];
      var b = document.createElement('button');
      b.className = 'ytoggle' + (on ? ' on' : '');
      b.innerHTML = '<span class="dot" style="background:' + y.color + '"></span>' + y.name;
      b.onclick = function () {
        S.hidden[y.name] = !S.hidden[y.name];
        b.className = 'ytoggle' + (!S.hidden[y.name] ? ' on' : '');
        refreshAll();
      };
      box.appendChild(b);
    });
  }

  function renderTabs() {
    var box = $('tabs');
    box.innerHTML = '';
    S.data.datasets.forEach(function (d) {
      var b = document.createElement('button');
      b.className = 'tab' + (d.id === S.dsId ? ' active' : '');
      b.innerHTML = d.name + '<span class="cnt">' + d.charts.length + '</span>';
      b.onclick = function () {
        if (d.id === S.dsId) return;
        renderDataset(d.id);
        renderTabs();
      };
      box.appendChild(b);
    });
  }

  function bindToolbar() {
    $('optConnect').onchange = function () { S.opt.connect = this.checked; refreshAll(); };
    $('optZero').onchange = function () { S.opt.zero = this.checked; refreshAll(); };
    $('optSync').onchange = function () {
      S.opt.sync = this.checked;
      refreshAll();
      if (!this.checked) {
        S.items.forEach(function (it) {
          if (it.inst) it.inst.setOption({ yAxis: { min: null, max: null } });
        });
      }
    };
    $('btnExpand').onclick = function () {
      S.expanded = !S.expanded;
      this.textContent = S.expanded ? '收起全部' : '全部展开';
      S.items.forEach(function (it) {
        it.card.classList.toggle('wide', S.expanded);
      });
      setTimeout(function () {
        S.items.forEach(function (it) { if (it.inst) it.inst.resize(); });
        if (S.opt.sync) syncY();
      }, 120);
    };
    $('btnTop').onclick = function () { window.scrollTo({ top: 0, behavior: 'smooth' }); };
  }

  function boot() {
    var params = new URLSearchParams(location.search);
    S.shot = params.get('shot') === '1';
    var want = params.get('ds');

    loadData(function () {
      var ds = (want && datasetById(want).id === want) ? datasetById(want) : S.data.datasets[0];
      S.dsId = ds.id;

      var d = new Date(S.data.updated.replace('T', ' ').replace(/-/g, '/'));
      var stamp = isNaN(d) ? S.data.updated : d.toLocaleString('zh-CN', { hour12: false });
      var parts = S.data.datasets.map(function (x) {
        return x.name + ' ' + x.charts.length + ' 图（至 ' + x.asOf + '）';
      });
      $('sub').textContent = '数据更新 ' + stamp + ' · ' + parts.join(' ｜ ');
      $('footNote').textContent = '共 ' +
        S.data.datasets.reduce(function (a, b) { return a + b.charts.length; }, 0) +
        ' 张图 · 数据提取自本地 Excel 数据库';

      renderTabs();
      renderDataset(S.dsId);
      bindToolbar();

      window.addEventListener('resize', function () {
        S.items.forEach(function (it) { if (it.inst) it.inst.resize(); });
        if (S.opt.sync) syncY();
      });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
