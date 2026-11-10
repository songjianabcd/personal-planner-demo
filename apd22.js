/* apd22.js — APD22_VERSION=20261023x */
/* 2026-10-23：KPI 预留位3 → 「小型化每日下达」（大数字+环比；悬浮面板=范围筛选 +
   多日期柱状图[红色虚线 60/150] + 图号构成[具体数字+占当天小型化占比]） */
/* 2026-10-23：预留位5 取消（5 卡变 4 卡）→ 末位改「超大安全钳关注」：卡面=标题（居中）
   + 分类汇总计数（按关注图号的箱数、降序；用户二次明确不要总数字/汇总行、不要副标题）；
   悬浮面板=时间筛选器（默认/最近1年/最近3年/最近5年/全部，均按箱子完工日期筛选）
   + 下达日期/完工日期/图号 三维分类汇总表 + 明细行表（拖选复制 + 一键复制 TSV，
   表格统一底色不逐行变色——参考明细表）。 */
/* ============================================================
 * apd22.js —— 3号箱APD及安全钳看板（深色指挥塔大屏）
 * 数据源：MySQL sap_table.03_apd（接口 /api/apd22/*）
 * 视觉：Operations Control Tower 深海军蓝 + 青绿辉光；参考 Tableau 高级图形
 *   （旭日四级下钻 / 玫瑰图 / 帕累托 / 完工日期分割堆叠柱+折线）。
 * 布局规则（用户明令）：从上往下颗粒度越来越细，从左到右越来越细。
 * 交互：自定义时间范围 + 梯种/框架属性/系列/型号 chips + 图内点击多级下钻
 *   （角标/梯种/旭日环层/BG 柱 → 全局过滤，KPI 与全部图联动刷新）。
 * 依赖：app/echarts.min.js（本地）、TableUX（index.html 内置）。
 * ============================================================ */
(function () {
  'use strict';

  var ACC = '#2E8FE8';            // 主蓝
  var CY  = '#7CD8FF';            // 亮青
  var GRN = '#2EE6A8';            // 青绿
  var GOLD= '#F5C86A';            // 金
  var RED = '#FF6B87';            // 警示
  var VIO = '#9D8CFF';
  var PALETTE = ['#2E8FE8', '#2EE6A8', '#7CD8FF', '#F5C86A', '#9D8CFF', '#FF6B87', '#37D0B2', '#5E7CE2', '#F2994A', '#63C9F5', '#B07CF5', '#6FCF97'];
  var FIN_PALETTE = ['#2E8FE8', '#2EE6A8', '#F5C86A', '#9D8CFF', '#FF6B87', '#37D0B2', '#7CD8FF', '#F2994A'];
  var TIP_BG = 'rgba(8,20,42,.94)';
  /* 用户自定义配色（布局编辑器「🎨」写入 window.A22_CHART_PREFS）：per 图表键 or '*' 全局。
     pal(key) 返回生效色板，无覆盖时返回 null → 调用方用原 PALETTE/FIN_PALETTE。 */
  function pal(key) {
    var prefs = window.A22_CHART_PREFS || {};
    var p = prefs[key] || prefs['*'];
    if (!p || !p.base) return null;
    var PRESETS = {
      '默认（指挥塔蓝绿）': PALETTE,
      '科技蓝紫': ['#5B8DEF', '#8F7CFF', '#63C9F5', '#B07CF5', '#4ECDC4', '#7C8EF5', '#5E7CE2', '#A0E7E5', '#9D8CFF', '#6FCF97'],
      '活力橙金': ['#F2994A', '#F5C86A', '#FF6B87', '#F5A65B', '#FFD166', '#E8734A', '#F2C94C', '#FF9A62', '#FBE38E', '#F2994A'],
      '清爽青绿': ['#2EE6A8', '#37D0B2', '#63C9F5', '#6FCF97', '#4ECDC4', '#7CD8FF', '#A0E7E5', '#2E8FE8', '#5FB2F5', '#9BE8C5'],
      '极光幻彩': ['#FF6B87', '#9D8CFF', '#63C9F5', '#2EE6A8', '#F5C86A', '#B07CF5', '#FF9A62', '#7CD8FF', '#F2994A', '#6FCF97'],
    };
    var list = (p.preset && PRESETS[p.preset]) ? PRESETS[p.preset].slice() : PALETTE.slice();
    var idx = list.indexOf(p.base);
    if (idx > 0) { list.splice(idx, 1); list.unshift(p.base); }
    return list;
  }

  function fmt(n) { return (n == null ? '-' : Number(n).toLocaleString('en-US')); }
  function pct(a, b) { if (!b) return '—'; var p = (a - b) / b * 100; return (p >= 0 ? '▲ ' : '▼ ') + Math.abs(p).toFixed(1) + '%'; }
  function el(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  var state = {
    data: null, options: null, charts: {},
    f: { start: '', end: '', fstart: '', fend: '', tizi: [], attr: [], s_series: [], s_typ: [], s_fig: [], f_series: [], f_fig: [], finish: [], corner: [], bg: [] },
    detail: { rows: [], headers: [], date: '', page: 1, size: 300, sortCol: null, sortDir: 1, filterText: '', savedQueries: [] },
    tizi: { range: 'day', from: '', to: '' },   // 梯种玫瑰范围筛选（真源：控件随 #apd22Kpi innerHTML 重建，值只信 state）
    mini: { range: '7', from: '', to: '' },   // 小型化每日下达 KPI 悬浮面板范围筛选（默认近7天；同上，值只信 state）
    mega: { range: 'default', from: '', to: '', seq: 0, data: null, loading: false },   // 超大安全钳关注 KPI：范围（默认=完工今天-7~今天+180；custom 用 from/to）+ 请求序号防竞态 + 响应缓存
    drill: null   // 下钻路径描述（状态栏显示）
  };

  /* ---------- 图表公共样式 ---------- */
  var AXIS = { axisLine: { lineStyle: { color: 'rgba(84,130,204,.4)' } }, axisTick: { show: false }, axisLabel: { color: '#7E96C2', fontSize: 10 } };
  function baseGrid(pos) { return Object.assign({ left: 8, right: 16, top: 30, bottom: 4, containLabel: true }, pos || {}); }
  function mk(id) {
    var dom = el(id);
    if (!dom) return null;
    if (state.charts[id] && !state.charts[id].isDisposed() && state.charts[id].getDom() === dom) return state.charts[id];
    /* 容器被 innerHTML 重建后旧实例已随 DOM 消亡——disposed 或 dom 变了都要重新 init */
    if (state.charts[id] && !state.charts[id].isDisposed()) { try { state.charts[id].dispose(); } catch (e) { } }
    var c = echarts.init(dom, null, { renderer: 'canvas' });
    state.charts[id] = c;
    return c;
  }
  window.addEventListener('resize', function () {
    Object.keys(state.charts).forEach(function (k) { state.charts[k].resize(); });
  });

  /* ---------- 数据加载 ---------- */
  /* v5（2026-10-18 首屏 3s 优化）：加载条+骨架微光提示；renderAll 改渐进渲染（图表逐个出，
     每张之间让出主线程，首张图尽早可见）；detail 与图表并行拉取 */
  function showLoading(on) {
    var bar = el('a22LoadBar'), pg = el('apd22Page');
    if (bar) bar.classList.toggle('show', !!on);
    if (pg) pg.classList.toggle('a22-loading', !!on);
    if (on) { var s0 = el('apd22Status'); if (s0) s0.textContent = '3号箱APD及安全钳看板 · 数据加载中…'; }
  }
  function load(retry) {
    var btn = el('apd22Refresh');
    if (btn) btn.disabled = true;
    showLoading(true);
    var payload = JSON.parse(JSON.stringify(state.f));
    var _cR = document.getElementById('a22CornerRange');
    if (_cR && _cR.value === 'custom') {
      payload.cstart = (document.getElementById('a22CornerFrom') || {}).value || '';
      payload.cend = (document.getElementById('a22CornerTo') || {}).value || '';
    }
    var _tR = document.getElementById('a22TiziRange');
    if (_tR) { state.tizi.range = _tR.value; state.tizi.from = (document.getElementById('a22TiziFrom') || {}).value || ''; state.tizi.to = (document.getElementById('a22TiziTo') || {}).value || ''; }
    if (state.tizi.range === 'custom') {
      payload.tstart = state.tizi.from;
      payload.tend = state.tizi.to;
    }
    fetch('/api/apd22/data', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) })
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(function (d) {
        state.data = d;
        loadDetail(d.kpi ? d.kpi.latest : '');   // 明细与图表并行拉（原来排在 renderAll 之后）
        renderAll();
      })
      .catch(function (e) {
        console.error('apd22 load fail', e);
        showLoading(false);
        if (!retry && /HTTP 5\d\d|Failed to fetch/i.test(String(e.message))) {
          setTimeout(function () { load(true); }, 2000); return;
        }
        var s = el('apd22Status'); if (s) s.textContent = '数据加载失败：' + e.message + '（可点「刷新」重试）';
      })
      .finally(function () { if (btn) btn.disabled = false; });
  }
  function loadOptions() {
    return fetch('/api/apd22/options').then(function (r) { return r.json(); }).then(function (o) {
      state.options = o;
      renderFilterBar();
    }).catch(function () { });
  }

  /* ---------- 筛选条（时间范围 + chips 多选） ---------- */
  function chip(group, val, label, active) {
    return '<button class="apd22-chip' + (active ? ' on' : '') + '" data-g="' + group + '" data-v="' + esc(val) + '">' + esc(label) + '</button>';
  }
  function renderFilterBar() {
    var o = state.options; if (!o) return;
    var wrap = el('apd22Filters'); if (!wrap) return;
    var html = '<div class="apd22-fgroup"><span class="apd22-fgroup-label">下达时间</span>' +
      '<input type="date" class="apd22-dateinput" id="a22FStart" value="' + esc(state.f.start) + '"> <span style="color:#5A7398;font-size:11px">至</span> ' +
      '<input type="date" class="apd22-dateinput" id="a22FEnd" value="' + esc(state.f.end) + '">' +
      '<span style="font-size:10.5px;color:#64809F">（日期格式 2026.09.30 可直接点 chips 快选）</span>' +
      o.dates.slice(0, 5).map(function (d) { return chip('day', d, d, false); }).join('') + '</div>';
    html += '<div class="apd22-fgroup"><span class="apd22-fgroup-label">完工时间</span>' +
      '<input type="date" class="apd22-dateinput" id="a22FFStart" value="' + esc(state.f.fstart) + '"> <span style="color:#5A7398;font-size:11px">至</span> ' +
      '<input type="date" class="apd22-dateinput" id="a22FFEnd" value="' + esc(state.f.fend) + '">' +
      (o.finishes || []).slice(0, 5).map(function (d) { return chip('fday', d, d, false); }).join('') + '</div>';
    html += '<div class="apd22-fgroup"><span class="apd22-fgroup-label">梯种</span>' +
      o.tizi.slice(0, 10).map(function (t) { return chip('tizi', t, t, state.f.tizi.indexOf(t) >= 0); }).join('') + '</div>';
    html += '<div class="apd22-fgroup"><span class="apd22-fgroup-label">安全钳筛选</span>' +
      '<span class="apd22-fsub">系列</span>' +
      (o.s_series || []).slice(0, 8).map(function (t) { return chip('s_series', t, t, state.f.s_series.indexOf(t) >= 0); }).join('') +
      '<span class="apd22-fsub">TYP</span>' +
      (o.s_typs || []).slice(0, 8).map(function (t) { return chip('s_typ', t, t, state.f.s_typ.indexOf(t) >= 0); }).join('') +
      '<span class="apd22-fsub">图号</span>' +
      (o.s_figs || []).slice(0, 8).map(function (t) { return chip('s_fig', t, t, state.f.s_fig.indexOf(t) >= 0); }).join('') + '</div>';
    html += '<div class="apd22-fgroup"><span class="apd22-fgroup-label">框架筛选</span>' +
      '<span class="apd22-fsub">分类</span>' +
      (o.f_series || []).slice(0, 8).map(function (t) { return chip('f_series', t, t, state.f.f_series.indexOf(t) >= 0); }).join('') +
      '<span class="apd22-fsub">图号</span>' +
      (o.f_figs || []).slice(0, 8).map(function (t) { return chip('f_fig', t, t, state.f.f_fig.indexOf(t) >= 0); }).join('') + '</div>';
    wrap.innerHTML = html;
    wrap.querySelectorAll('.apd22-chip').forEach(function (b) {
      b.addEventListener('click', function () {
        var g = b.dataset.g, v = b.dataset.v;
        if (g === 'day') {
          state.f.start = v; state.f.end = v;
          var si = el('a22FStart'), ei = el('a22FEnd');
          if (si) si.value = v.replace(/\./g, '-');
          if (ei) ei.value = v.replace(/\./g, '-');
          load(); return;
        }
        if (g === 'fday') {
          state.f.fstart = v; state.f.fend = v;
          var fi = el('a22FFStart'), fe = el('a22FFEnd');
          if (fi) fi.value = v.replace(/\./g, '-');
          if (fe) fe.value = v.replace(/\./g, '-');
          load(); return;
        }
        var arr = state.f[g];
        var i = arr.indexOf(v);
        if (i >= 0) arr.splice(i, 1); else arr.push(v);
        b.classList.toggle('on');
        scheduleApply();
      });
    });
    var si = el('a22FStart'), ei = el('a22FEnd'), fi = el('a22FFStart'), fe = el('a22FFEnd');
    function onDate() {
      state.f.start = (si && si.value) ? si.value.replace(/-/g, '.') : '';
      state.f.end = (ei && ei.value) ? ei.value.replace(/-/g, '.') : '';
      scheduleApply();
    }
    function onFDate() {
      state.f.fstart = (fi && fi.value) ? fi.value.replace(/-/g, '.') : '';
      state.f.fend = (fe && fe.value) ? fe.value.replace(/-/g, '.') : '';
      scheduleApply();
    }
    if (si) si.addEventListener('change', onDate);
    if (ei) ei.addEventListener('change', onDate);
    if (fi) fi.addEventListener('change', onFDate);
    if (fe) fe.addEventListener('change', onFDate);
  }
  var applyTimer = null;
  function scheduleApply() {
    clearTimeout(applyTimer);
    applyTimer = setTimeout(load, 350);
  }
  function resetFilters() {
    state.f = { start: '', end: '', fstart: '', fend: '', tizi: [], attr: [], s_series: [], s_typ: [], s_fig: [], f_series: [], f_fig: [], finish: [], corner: [], bg: [] };
    state.drill = null;
    renderFilterBar();
    load();
  }
  function drillLabel() {
    var f = state.f, parts = [];
    if (f.start || f.end) parts.push('时间 ' + (f.start || '…') + '~' + (f.end || '…'));
    if (f.tizi.length) parts.push('梯种 ' + f.tizi.join('/'));
    if (f.attr.length) parts.push('属性 ' + f.attr.join('/'));
    if (f.s_series.length) parts.push('钳系列 ' + f.s_series.join('/'));
    if (f.s_typ.length) parts.push('钳TYP ' + f.s_typ.join('/'));
    if (f.s_fig.length) parts.push('钳图号 ' + f.s_fig.join('/'));
    if (f.f_series.length) parts.push('框分类 ' + f.f_series.join('/'));
    if (f.f_fig.length) parts.push('框图号 ' + f.f_fig.join('/'));
    if (f.finish.length) parts.push('完工 ' + f.finish.join('/'));
    if (f.corner.length) parts.push('角标 ' + f.corner.join('/'));
    if (f.bg.length) parts.push('BG ' + f.bg.join('/'));
    return parts.length ? '下钻：' + parts.join(' · ') : '';
  }

  /* ---------- KPI 卡标题下方「下达日期」小字行（2026-10-25 用户要求） ----------
     当日下达箱数 / 小型化每日下达 两张数字卡，在**标题下面**跟一行该数字对应的下达日期，
     写成带小括号的 yyyy.mm.dd（如（2026.10.24）），字号小于标题。
     ⚠️ 位置是标题下方、数字上方——用户明确纠正过「跑数字下面」不对，别再挪。 */
  function _kpiDate(s) {
    var t = String(s == null ? '' : s).trim().slice(0, 10).replace(/[\/\-\.]/g, '.');
    var p = t.split('.');
    if (p.length !== 3 || !p[0]) return '';
    return p[0] + '.' + ('0' + p[1]).slice(-2) + '.' + ('0' + p[2]).slice(-2);
  }
  function _kpiDateRow(s) {
    var t = _kpiDate(s);
    return t ? '<div class="a22-kpi-date">（' + esc(t) + '）</div>' : '';
  }

  /* ---------- KPI 区：仅第 1 个数字卡（悬浮展示每日数量）+ 2~5 空位（待用户添加） ---------- */
  function renderKpi(d) {
    var k = d.kpi;
    var daily = (d.finishDays && d.finishDays.length ? d.finishDays : (d.finishDaily || d.daily15).slice(0, 8)).slice(0, 8);
    var _vals = daily.map(function (x) { return x.count !== undefined ? x.count : x.total; });
    var _max = Math.max.apply(null, _vals.concat([1]));
    var tipRows = daily.map(function (x, _i) {
      var v = _vals[_i];
      var barH = Math.max(3, Math.round(v / _max * 46));
      var dLabel = String(x.date).slice(5).replace('-', '.');
      var srcs = x.srcs || [];
      var srcRows = srcs.length ? srcs.map(function (s) {
        return '<div style="display:flex;justify-content:space-between;gap:14px;line-height:17px">' +
          '<span style="color:#EAF3FF">' + esc(s.date) + '</span><b style="color:#F5C86A">' + fmt(s.count) + ' 箱</b></div>';
      }).join('') : '<div style="color:#54688C;line-height:17px">无 03#箱记录</div>';
      return '<div class="a22bar" data-i="' + _i + '" style="display:flex;flex-direction:column;align-items:center;gap:2px;flex:1;position:relative;cursor:default">' +
        '<span style="color:#F5C86A;font-weight:700;font-size:10.5px">' + fmt(v) + '</span>' +
        '<div style="width:16px;height:' + barH + 'px;background:linear-gradient(180deg,#5FB2F5,#2E8FE8);border-radius:2px 2px 0 0"></div>' +
        '<span style="color:#8FA6CC;font-size:9.5px">' + dLabel + '</span>' +
        '<div class="a22sub" style="display:none;position:absolute;bottom:calc(100% + 2px);left:50%;transform:translateX(-50%);width:190px;background:#0D2547;border:1px solid #2A4A78;border-radius:6px;padding:7px 9px;box-shadow:0 8px 22px rgba(0,0,0,.45);z-index:60;white-space:nowrap">' +
        '<div style="color:#7CD8FF;font-weight:700;margin-bottom:3px">完工 ' + esc(x.date) + ' · 共 ' + fmt(v) + ' 箱</div>' + srcRows +
        '</div>' +
        '</div>';
    }).join('');
    // 迷你折线（终版v2）：SVG 拉伸铺满（线/面积不变形 non-scaling-stroke），HTML 圆点百分比定位（正圆），CSS hover 气泡（日期+数量）
    var _last7 = daily.slice(0, 7).slice().reverse();
    var _v7 = _last7.map(function (x) { return x.count !== undefined ? x.count : x.total; });
    var _d7 = _last7.map(function (x) { return String(x.date).slice(5).replace('-', '.'); });
    var _m7 = Math.max.apply(null, _v7.concat([1]));
    var _min7 = Math.min.apply(null, _v7);
    var _SW = 300, _SH = 100, _PL = 12, _PR = 12, _PT = 16, _PB = 26;
    var _xy = _v7.map(function (v, i2) {
      var x = _PL + i2 * ((_SW - _PL - _PR) / Math.max(1, _v7.length - 1));
      var y = _PT + (1 - (v - _min7) / Math.max(1, _m7 - _min7)) * (_SH - _PT - _PB);
      return [x, y];
    });
    var _path = '';
    for (var _pi = 0; _pi < _xy.length; _pi++) {
      if (_pi === 0) { _path = 'M' + _xy[0][0].toFixed(1) + ' ' + _xy[0][1].toFixed(1); continue; }
      var _p0 = _xy[_pi - 1], _p1 = _xy[_pi];
      var _cx = (_p0[0] + _p1[0]) / 2;
      _path += ' C' + _cx.toFixed(1) + ' ' + _p0[1].toFixed(1) + ' ' + _cx.toFixed(1) + ' ' + _p1[1].toFixed(1) + ' ' + _p1[0].toFixed(1) + ' ' + _p1[1].toFixed(1);
    }
    var _area = _path + ' L' + _xy[_xy.length - 1][0].toFixed(1) + ' ' + (_SH - 4) + ' L' + _xy[0][0].toFixed(1) + ' ' + (_SH - 4) + ' Z';
    var _dots = '';
    for (var _ni = 0; _ni < _xy.length; _ni++) {
      var _lf = (_xy[_ni][0] / _SW * 100).toFixed(2);
      var _tp = (_xy[_ni][1] / _SH * 100).toFixed(2);
      var _isLast = _ni === _xy.length - 1;
      var _below = _xy[_ni][1] < 34 ? ' tip-below' : '';
      _dots += '<div class="a22dot' + (_isLast ? ' a22dot-last' : '') + _below + '" style="left:' + _lf + '%;top:' + _tp + '%">' +
        '<div class="a22dot-tip"><b>完工 ' + _d7[_ni] + '</b><span>' + fmt(_v7[_ni]) + ' 箱</span>' +(_ni > 0 && _v7[_ni - 1] ? '<em class="a22dot-hb ' + (_v7[_ni] >= _v7[_ni - 1] ? 'hb-up' : 'hb-down') + '">环比 ' + (_v7[_ni] >= _v7[_ni - 1] ? '+' : '') + ((_v7[_ni] - _v7[_ni - 1]) / _v7[_ni - 1] * 100).toFixed(1) + '%</em>' : '<em class="a22dot-hb">环比 —</em>') +'</div>' +
        '</div>';
    }
    var _spark = '<svg width="100%" height="100%" viewBox="0 0 ' + _SW + ' ' + _SH + '" preserveAspectRatio="none" style="display:block;position:absolute;left:0;top:0">' +
      '<defs>'+
      '<linearGradient id="a22sg" x1="0" y1="0" x2="0" y2="1">'+
      '<stop offset="0%" stop-color="rgba(124,216,255,.50)"/><stop offset="100%" stop-color="rgba(124,216,255,.04)"/></linearGradient>'+
      '<filter id="a22glow" x="-40%" y="-40%" width="180%" height="180%">'+
      '<feGaussianBlur stdDeviation="3" result="b1"/><feMerge><feMergeNode in="b1"/><feMergeNode in="SourceGraphic"/></feMerge></filter>'+
      '</defs>'+
      '<path d="' + _area + '" fill="url(#a22sg)" stroke="none"/>'+
      '<path d="' + _path + '" fill="none" stroke="#2E8FE8" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" filter="url(#a22glow)" vector-effect="non-scaling-stroke"/>'+
      '<line x1="' + _PL + '" y1="' + (_SH - 4) + '" x2="' + (_SW - _PR) + '" y2="' + (_SH - 4) + '" stroke="rgba(84,130,204,.35)" stroke-width="1" vector-effect="non-scaling-stroke"/>'+
      '</svg>' + _dots +
      '<span class="a22spark-x1">' + _d7[0] + '</span>'+
      '<span class="a22spark-x2">' + _d7[_d7.length - 1] + '</span>';
    /* ---------- KPI 预留位3：小型化每日下达（2026-10-23 用户要求） ----------
       卡面=标题+大数字+环比（参考「当日下达箱数」样式，副标题按用户要求去掉）；
       悬浮面板=范围筛选（当天/近7/15/30/全部/自定义）+ 多日期柱状图（红色虚线水平线 60 / 150），
       图号构成只在悬浮柱的 tooltip 内展示（2026-10-23 用户二次要求：面板只留柱状图）。 */
    function miniCardHtml(d) {
      var md = d.miniDaily || {};
      var last = md.latestCount || 0, prev = md.prevCount || 0;
      var dCls = last >= prev ? 'up' : 'down';
      var dTxt = prev > 0 ? pct(last, prev) : '—';
      return '<div class="apd22-kpi" id="a22Kpi3" style="cursor:default">' +
        '<div class="apd22-kpi-title">小型化每日下达</div>' +
        _kpiDateRow(md.latest) +
        '<div class="a22-kpi-mid"><span class="apd22-kpi-value">' + fmt(last) + '</span>' +
        '<span class="apd22-kpi-delta ' + dCls + '">' + dTxt + '</span></div>' +
        '<div class="apd22-hover-panel a22-mini-panel" id="a22MiniPanel">' +
        '<div class="a22-mini-head">' +
        '<span class="a22-mini-head-t">小型化 · 按下达日期柱状图</span>' +
        '<span style="display:inline-flex;align-items:center;gap:6px;flex-wrap:wrap">' +
        '<span id="a22MiniCustom" style="display:none;white-space:nowrap"><input type="date" id="a22MiniFrom" class="apd22-mini-date"><span style="color:#8FA6CC;margin:0 2px">→</span><input type="date" id="a22MiniTo" class="apd22-mini-date"></span>' +
        '<select id="a22MiniRange" class="apd22-mini-select" title="选择柱状图统计的时间范围"><option value="day">当天</option><option value="7">近 7 天</option><option value="15">近 15 天</option><option value="30">近 30 天</option><option value="all">全部</option><option value="custom">自定义</option></select>' +
        '</span></div>' +
        '<div class="a22-mini-chart" id="a22MiniChart"></div>' +
        '</div></div>';
    }
    function _miniSlice(byDate) {
      var r = state.mini.range;
      var arr = (byDate || []).slice();
      if (r === 'custom') {
        var f = state.mini.from ? _dkey(state.mini.from) : '';
        var t = state.mini.to ? _dkey(state.mini.to) : '';
        if (!f && !t) return [];   // 自定义未选日期 → 不出图（面板提示选日期）
        if (f) arr = arr.filter(function (x) { return _dkey(x.date) >= f; });
        if (t) arr = arr.filter(function (x) { return _dkey(x.date) <= t; });
        return arr;
      }
      if (r === 'all') return arr;
      if (r === 'day') return arr.slice(0, 1);
      return arr.slice(0, parseInt(r, 10) || 15);
    }
    /* 空范围时的图内提示文案（自定义未选日期 / 起止填反 / 确实无记录） */
    function _miniEmptyHint() {
      var r = state.mini.range;
      if (r === 'custom' && !state.mini.from && !state.mini.to) return '请选择开始 / 结束日期';
      if (r === 'custom' && state.mini.from && state.mini.to && _dkey(state.mini.from) > _dkey(state.mini.to)) return '开始日期晚于结束日期，请调整';
      return '该范围无小型化下达记录';
    }
    /* 渲染悬浮面板：范围筛选（默认近7天）+ 多日期堆叠柱状图（markLine 60/150 红色虚线）。
       柱内按图号堆叠体现构成（同色蓝 + 段间底色实线间隙，2026-10-23 用户要求：像左侧堆叠图那样
       看出组成，但不要多种颜色、不要虚线分隔）；图号构成的具体数字与占比同时保留在 tooltip 内。
       面板用 visibility 隐藏（保留布局），故 ECharts 可在 innerHTML 重建后立即 init。 */
    function renderMiniDaily(d) {
      var md = (d && d.miniDaily) || { byDate: [] };
      var sel = el('a22MiniRange');
      if (sel) {
        sel.value = state.mini.range;
        sel.onchange = function () {
          state.mini.range = this.value;
          renderMiniDaily(state.data);
        };
      }
      var cf = el('a22MiniCustom');
      if (cf) cf.style.display = state.mini.range === 'custom' ? 'inline' : 'none';
      ['a22MiniFrom', 'a22MiniTo'].forEach(function (id2) {
        var e2 = el(id2);
        if (!e2) return;
        e2.value = (id2 === 'a22MiniFrom' ? state.mini.from : state.mini.to) || '';
        e2.onchange = function () {
          state.mini[id2 === 'a22MiniFrom' ? 'from' : 'to'] = this.value;
          if (state.mini.range === 'custom') renderMiniDaily(state.data);
        };
      });
      var rows = _miniSlice(md.byDate);
      var _emptyHint = _miniEmptyHint();
      var c = mk('a22MiniChart');
      if (c) {
        if (!rows.length) {
          /* 空范围：图内居中文字提示（而非空白画布） */
          c.setOption({
            graphic: [{ type: 'text', left: 'center', top: 'middle', style: { text: _emptyHint, fill: '#7E96C2', fontSize: 12 } }],
            xAxis: { show: false }, yAxis: { show: false },
            series: [{ type: 'bar', data: [] }]
          }, { notMerge: true });
        }
        else {
          var dates = rows.map(function (x) { return String(x.date).slice(5).replace('-', '.'); });
          var totals = rows.map(function (x) { return x.total; });
          var _bar = ACC;
          var _xa = Object.assign({ type: 'category', data: dates }, AXIS);
          _xa.axisLabel = { color: '#7E96C2', fontSize: 10, rotate: 35, interval: 'auto' };
          /* 柱内按图号堆叠体现构成（用户 2026-10-23 要求：像左侧「近15个下达日箱数」那样堆叠看出组成，
             但不要多种颜色、也不要虚线分隔）：全部段同一蓝色系渐变；段与段之间描「面板底色 2px 实线」
             当分隔（边框色与背景同色 → 左右侧边天然隐形；实线在视觉上就是一道干净缺口，比虚线耐看）；
             当天缺该图号的段值=0 且单独关掉边框，防止 0 高度段画出漂浮线。 */
          var _figKeys = [], _figSum = {}, _figName = {};
          rows.forEach(function (x) {
            (x.figs || []).forEach(function (g) {
              var k = g.typ + '||' + g.fig;
              if (_figKeys.indexOf(k) < 0) { _figKeys.push(k); _figSum[k] = 0; _figName[k] = g.typ + ' · ' + g.fig; }
              _figSum[k] += g.value;
            });
          });
          _figKeys.sort(function (a, b) { return _figSum[b] - _figSum[a]; });   // 区间量大的图号放底部
          var _series = _figKeys.map(function (k, si) {
            var isTop = si === _figKeys.length - 1;
            var _sty = {
              color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: '#5FB2F5' }, { offset: 1, color: _bar }] },
              borderRadius: isTop ? [3, 3, 0, 0] : 0
            };
            if (!isTop) { _sty.borderColor = '#123059'; _sty.borderWidth = 2; _sty.borderType = 'solid'; }
            return {
              name: _figName[k], type: 'bar', stack: 'mini', barMaxWidth: 26, z: 10,
              data: rows.map(function (x) {
                var v = 0;
                (x.figs || []).forEach(function (g) { if (g.typ + '||' + g.fig === k) v = g.value; });
                return v > 0 ? v : { value: 0, itemStyle: { borderWidth: 0 } };
              }),
              itemStyle: _sty,
              label: isTop ? {
                show: true, position: 'top', color: '#EAF3FF', fontSize: 9.5, fontWeight: 600,
                textBorderColor: '#0B1F3A', textBorderWidth: 2,
                formatter: function (p) { return totals[p.dataIndex]; }   // 顶端段顶上的数字=当天合计
              } : { show: false },
              markLine: isTop ? {
                symbol: 'none', silent: true,
                data: [
                  { yAxis: 60, lineStyle: { color: '#FF4D5E', type: 'dashed', width: 1.6 }, label: { formatter: '60', color: '#FF6B87', fontSize: 10, fontWeight: 700, position: 'insideEndTop' } },
                  { yAxis: 150, lineStyle: { color: '#FF4D5E', type: 'dashed', width: 1.6 }, label: { formatter: '150', color: '#FF6B87', fontSize: 10, fontWeight: 700, position: 'insideEndTop' } }
                ]
              } : undefined
            };
          });
          if (!_series.length) {   // 兜底：无图号构成数据 → 退回单系列总量柱
            _series = [{
              type: 'bar', data: totals, barMaxWidth: 26, z: 10,
              itemStyle: { color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: '#5FB2F5' }, { offset: 1, color: _bar }] }, borderRadius: [3, 3, 0, 0] },
              label: { show: true, position: 'top', color: '#EAF3FF', fontSize: 9.5, fontWeight: 600, textBorderColor: '#0B1F3A', textBorderWidth: 2 },
              markLine: {
                symbol: 'none', silent: true,
                data: [
                  { yAxis: 60, lineStyle: { color: '#FF4D5E', type: 'dashed', width: 1.6 }, label: { formatter: '60', color: '#FF6B87', fontSize: 10, fontWeight: 700, position: 'insideEndTop' } },
                  { yAxis: 150, lineStyle: { color: '#FF4D5E', type: 'dashed', width: 1.6 }, label: { formatter: '150', color: '#FF6B87', fontSize: 10, fontWeight: 700, position: 'insideEndTop' } }
                ]
              }
            }];
          }
          c.setOption({
            tooltip: {
              trigger: 'axis', axisPointer: { type: 'shadow' },
              backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF', fontSize: 11 },
              formatter: function (ps) {
                var x = rows[ps[0].dataIndex];
                if (!x) return '';
                var html = '<b>' + esc(x.date) + '</b> · 小型化 <b style="color:#F5C86A">' + fmt(x.total) + '</b> 箱';
                if (x.figs && x.figs.length) {
                  html += '<br><span style="color:#9CCDFF;font-size:10px">图号构成（占当天小型化占比）：</span>';
                  x.figs.forEach(function (g) {
                    var p = x.total ? (g.value / x.total * 100).toFixed(1) : '0.0';
                    html += '<br>' + esc(g.typ) + ' · ' + esc(g.fig) + '：<b style="color:#F5C86A">' + fmt(g.value) + '</b> 箱（' + p + '%）';
                  });
                }
                return html;
              }
            },
            grid: baseGrid({ top: 24, left: 4, right: 10 }),
            xAxis: _xa,
            yAxis: Object.assign({ type: 'value', name: '箱数', minInterval: 1, splitLine: { lineStyle: { color: 'rgba(84,130,204,.15)' } } }, AXIS),
            series: _series
          }, { notMerge: true });
        }
      }
    }

    el('apd22Kpi').innerHTML =
      '<div class="apd22-kpi" id="a22Kpi1" style="cursor:default">' +
      '<div class="apd22-kpi-title">当日下达箱数</div>' +
      _kpiDateRow(k.latest) +
      '<div class="a22-kpi-mid"><span class="apd22-kpi-value" >' + fmt(k.latestCount) + '</span>' +
      '<span class="apd22-kpi-delta ' + (k.latestCount >= k.prevCount ? 'up' : 'down') + '">' + pct(k.latestCount, k.prevCount) + '</span></div>' +
      '<div class="apd22-hover-panel" style="left:auto;right:8px;min-width:150px">' +
      '<div style="color:#7CD8FF;font-weight:700;margin-bottom:4px">昨日下达数量</div>' +
      '<div style="font-size:20px;font-weight:800;color:#EAF3FF">' + fmt(k.prevCount) + '</div>' +
      '<div style="color:#8FA6CC;margin-top:3px">下达日 ' + esc(k.prev) + ' · 今日对比 ' + pct(k.latestCount, k.prevCount) + '</div>' +
      '</div>' +
      '</div>' +
      '<div class="apd22-kpi" id="a22Kpi2" style="cursor:default">' +
      '<div class="apd22-kpi-title">近7天完工日数量波动</div>' +
      '<div class="a22spark-wrap">' + _spark + '</div>' +
      '<div class="apd22-hover-panel" id="a22KpiHover">' +
      '<div style="color:#7CD8FF;font-weight:700;margin-bottom:4px">最近 8 个完工日期 · 03#箱数量</div>' +
      '<div style="color:#FFFFFF;font-size:10.5px;margin-bottom:5px">悬浮柱可看各下达日期贡献</div>' +
      '<div style="display:flex;align-items:flex-end;gap:7px;padding:2px 2px 0">' + tipRows + '</div>' +
      '<div style="border-top:1px solid rgba(84,130,204,.35);margin-top:5px;padding-top:4px;display:flex;justify-content:space-between;gap:18px">' +
      '<span style="color:#8FA6CC">合计</span><b style="color:#2EE6A8">' + fmt(daily.reduce(function (a, x) { return a + (x.count !== undefined ? x.count : x.total); }, 0)) + ' 箱</b></div>' +
      '</div></div>' +
      (d.miniDaily ? miniCardHtml(d) :
        '<div class="apd22-kpi apd22-kpi-empty" data-slot="3"><div class="apd22-kpi-title">预留位 3</div><div class="apd22-kpi-empty-hint">待添加指标</div></div>') +
      megaCardHtml();
    if (d.miniDaily) renderMiniDaily(d);
    megaLoad();   // 卡面大数字异步绘制（面板数据同源，悬浮时已就绪）
  }

  /* ---------- 图1：15 个下达日 · 完工日期分割堆叠柱 + 下达量折线 ---------- */
  /* ⚠️ 日期比较必须归一化：daily15 的 date 是点分 '2026.10.14'，input[type=date] 是 ISO
   * '2026-10-05'——字符串直比时 '.'(46)>'-'(45) 导致上界恒 false、筛选恒空（2026-10-17 用户
   * 报"筛选的数据对吗"实测复现）。统一转成 yyyymmdd 数字比。 */
  function _dkey(s) { return String(s || '').replace(/[-.]/g, ''); }

  /* ---------- KPI 预留位5 取消 → 「超大安全钳关注」（2026-10-23 用户要求） ----------
     卡面=标题（居中）+ 汇总数字（当前范围命中总箱数，居中）+ 分类汇总四列
     （下达时间 / 完工时间 / 图号 / 汇总计数，按完工日期升序=最紧急在前）；
     悬浮面板=时间筛选器（默认=完工今天-7天~今天+180天 / 最近1年 / 最近3年 / 最近5年 / 全部 /
     自定义完工日期）+ 明细行表（user-select:text 拖选即可复制；另有一键复制按钮，TSV 含表头，
     可直接粘进 Excel；表格统一底色、不逐行变色——用户 2026-10-23 多次强调）。
     用户 2026-10-23 明确：悬浮里**不要**分类汇总表，只要明细行。 */
  function megaCardHtml() {
    return '<div class="apd22-kpi" id="a22Kpi4" style="cursor:default">' +
      /* 标题行（2026-10-23）：只放「超大安全钳关注 + 范围下拉」并整组居中；
         命中箱数（6箱）移到下面表格「汇总计数」列的表头里，与计数上下同列对齐。 */
      '<div class="apd22-kpi-title a22-mega-title">' +
      '<span class="a22-mega-title-t">超大安全钳关注</span>' +
      '<select id="a22MegaCardRange" class="a22-mega-card-sel" title="按箱子完工日期筛选">' +
      '<option value="default">默认</option><option value="1y">最近1年</option>' +
      '<option value="3y">最近3年</option><option value="5y">最近5年</option>' +
      '<option value="all">全部</option><option value="custom">自定义</option>' +
      '</select>' +
      '</div>' +
      '<div class="a22-mega-face-wrap">' +
      '<div class="a22-mega-face" id="a22MegaFace"></div>' +
      '<div class="a22-mega-rail" id="a22MegaRail" title="拖动查看全部"><i class="a22-mega-thumb" id="a22MegaThumb"></i></div>' +
      '</div>' +
      '<div class="apd22-hover-panel a22-mega-panel" id="a22MegaPanel">' +
      '<div class="a22-mega-head">' +
      '<span class="a22-mega-head-t">超大安全钳关注 · 明细</span>' +
      '<span style="display:inline-flex;align-items:center;gap:6px;flex-wrap:wrap">' +
      '<span id="a22MegaCustom" style="display:none;white-space:nowrap">' +
      '<input type="date" id="a22MegaFrom" class="apd22-mini-date">' +
      '<span style="color:#8FA6CC;margin:0 2px">→</span>' +
      '<input type="date" id="a22MegaTo" class="apd22-mini-date"></span>' +
      '<select id="a22MegaRange" class="apd22-mini-select" title="按箱子完工日期筛选关注范围">' +
      '<option value="default">默认</option><option value="1y">最近1年</option>' +
      '<option value="3y">最近3年</option><option value="5y">最近5年</option>' +
      '<option value="all">全部</option><option value="custom">自定义完工日期</option>' +
      '</select></span></div>' +
      '<div class="a22-mega-stats" id="a22MegaStats"></div>' +
      '<div class="a22-mega-sec"><span class="a22-mega-sec-t">明细行（可选中复制）</span>' +
      '<button class="a22-mega-copy" id="a22MegaCopy" type="button">复制明细</button></div>' +
      '<div class="a22-mega-det-wrap" id="a22MegaDet"></div>' +
      '</div></div>';
  }
  /* 卡面：分类汇总计数四列（下达时间 / 完工时间 / 图号 / 汇总计数），按完工日期升序（最紧急在前）。
     用户 2026-10-23 多次明确：这几个条件都要展示在大看板卡面上；不要汇总行、不要副标题。
     用户 2026-10-23 最新：「6箱 要和 汇总计数 在一列」→ 命中总箱数（#a22MegaTotal）放进
     「汇总计数」表头里，跟下方各行的计数数字上下同一列对齐，不再飘在标题行右侧。 */
  function _megaFaceHtml(d) {
    var gs = d.groups || [];
    var tot = (d.total == null || d.total === '') ? '—' : fmt(d.total);
    if (!gs.length) {
      return '<div class="a22-mega-face-empty">该范围无关注记录</div>';
    }
    var h = '<table class="a22-mega-face-tab"><thead><tr>' +
      '<th>下达时间</th><th>完工时间</th><th>图号</th>' +
      '<th class="num a22-mega-th-total">' +
      '<span class="a22-mega-th-name">汇总计数</span>' +
      '<span class="a22-mega-sum"><span id="a22MegaTotal">' + esc(tot) + '</span><em>箱</em></span>' +
      '</th>' +
      '</tr></thead><tbody>';
    gs.forEach(function (g) {
      h += '<tr><td>' + esc(g.order) + '</td><td>' + esc(g.finish) + '</td>' +
        '<td>' + esc(g.fig) + '</td><td class="num">' + fmt(g.count) + '</td></tr>';
    });
    return h + '</tbody></table>';
  }
  /* 面板宽度自适应：祖先 .apd22-page / .content 是 overflow:hidden/auto，面板若超出其左缘会被裁切
     （实测裁掉 41px，左侧「下达日期」整列看不见）。这里按「卡片右缘 - 容器左缘」动态算 max-width，
     保证面板左缘始终落在容器内；15 列明细超宽时由 .a22-mega-det-wrap 内部横向滚动。 */
  function megaFitPanel() {
    var panel = el('a22MegaPanel'), card = el('a22Kpi4');
    if (!panel || !card) return;
    var clipL = 0, node = panel.parentElement;
    while (node && node !== document.body) {
      var cs = getComputedStyle(node);
      if (cs.overflowX !== 'visible' || cs.overflow !== 'visible') {
        clipL = Math.max(clipL, node.getBoundingClientRect().left);
      }
      node = node.parentElement;
    }
    var right = card.getBoundingClientRect().right - 8;
    var avail = right - clipL - 12;      /* 12px 安全余量 */
    if (avail > 240) panel.style.maxWidth = Math.round(avail) + 'px';
  }
  /* 卡面自绘滚动条：Win11 Edge 用 overlay 滚动条（不占宽、悬停才显，原生态几乎看不见），
     用户 2026-10-23 明确「放不下就加垂直滚动条」→ 自绘常驻细条，滑块可按比例拖动，
     滚轮照旧由容器原生滚动处理（只在 scroll 时同步滑块位置）。 */
  function megaRailSync() {
    var f = el('a22MegaFace'), rail = el('a22MegaRail'), th = el('a22MegaThumb');
    if (!f || !rail || !th) return;
    var sh = f.scrollHeight, ch = f.clientHeight;
    if (sh <= ch + 1) { rail.style.display = 'none'; return; }
    rail.style.display = 'block';
    var hR = rail.clientHeight || ch;
    var hT = Math.max(14, Math.round(hR * ch / sh));
    th.style.height = hT + 'px';
    var max = sh - ch;
    th.style.top = Math.round((hR - hT) * (f.scrollTop / max)) + 'px';
  }
  /* 绑定滚动同步 + 滑块拖动（一次性绑定；面板每次重绘只更新内容，不重建 rail 元素） */
  function megaRailBind() {
    var f = el('a22MegaFace'), rail = el('a22MegaRail'), th = el('a22MegaThumb');
    if (!f || !rail || !th || rail.dataset.bound === '1') return;
    rail.dataset.bound = '1';
    f.addEventListener('scroll', megaRailSync, { passive: true });
    th.addEventListener('mousedown', function (ev) {
      ev.preventDefault(); ev.stopPropagation();
      var startY = ev.clientY, startTop = f.scrollTop;
      var hR = rail.clientHeight || f.clientHeight;
      var hT = th.offsetHeight;
      var max = f.scrollHeight - f.clientHeight;
      var ratio = max / Math.max(1, hR - hT);
      var move = function (e2) {
        f.scrollTop = Math.max(0, Math.min(max, startTop + (e2.clientY - startY) * ratio));
        megaRailSync();
      };
      var up = function () {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    });
    /* 点轨道空白处翻页 */
    rail.addEventListener('mousedown', function (ev) {
      if (ev.target === th) return;
      ev.preventDefault(); ev.stopPropagation();
      var r = rail.getBoundingClientRect();
      f.scrollTop += (ev.clientY < r.top + th.offsetTop + th.offsetHeight / 2) ? -f.clientHeight : f.clientHeight;
      megaRailSync();
    });
    window.addEventListener('resize', function () { megaRailSync(); megaFitPanel(); });
  }
  /* 明细一键复制：TSV（含表头，\r\n 换行，Excel 直接粘贴成格）。clipboard API 失败时
     execCommand 兜底（file:// 等非安全上下文）。 */
  function megaCopyDetails() {
    var d = state.mega.data;
    if (!d || !d.details || !d.details.length) return;
    var hd = d.headers || [];
    var lines = [hd.join('\t')];
    d.details.forEach(function (r) { lines.push(r.join('\t')); });
    var txt = lines.join('\r\n');
    var btn = el('a22MegaCopy');
    var done = function (ok) {
      if (!btn) return;
      var t = btn.textContent;
      btn.textContent = ok ? '已复制 ✓' : '复制失败';
      setTimeout(function () { btn.textContent = t; }, 1500);
    };
    var fallback = function () {
      try {
        var ta = document.createElement('textarea');
        ta.value = txt;
        ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.select();
        var ok = document.execCommand('copy');
        document.body.removeChild(ta);
        return ok;
      } catch (e) { return false; }
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(txt).then(function () { done(true); }, function () { done(fallback()); });
    } else {
      done(fallback());
    }
  }
  /* 拉取当前范围的关注数据并绘制（卡面 + 面板）。state.mega.loading 防并发重复请求。
   range=custom 时把 from/to 一起带上（ISO，单边可空；两端都空=完工日期不限）。 */
  function megaLoad() {
    var qs = 'range=' + encodeURIComponent(state.mega.range);
    if (state.mega.range === 'custom') {
      if (state.mega.from) qs += '&start=' + encodeURIComponent(state.mega.from);
      if (state.mega.to) qs += '&end=' + encodeURIComponent(state.mega.to);
    }
    var seq = (state.mega.seq || 0) + 1;
    state.mega.seq = seq;
    fetch('/api/apd22/mega-clamp?' + qs)
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(function (d) {
        if (state.mega.seq !== seq) return;   // 已有更新的请求发出，丢弃过期响应
        state.mega.data = d;
        megaPaint(d);
      })
      .catch(function (e) {
        if (state.mega.seq !== seq) return;
        console.error('apd22 mega load fail', e);
        var fc = el('a22MegaFace');
        if (fc) fc.innerHTML = '<div class="a22-mega-face-empty">加载失败：' + esc(e.message) + '</div>';
        var tv = el('a22MegaTotal'); if (tv) tv.textContent = '—';
      });
  }
  /* 用响应绘制卡面（分类汇总四列，命中箱数在「汇总计数」表头内）与面板（筛选器回填 + 统计条 + 明细表）。 */
  function megaPaint(d) {
    /* #a22MegaTotal 现在位于「汇总计数」表头里，由 _megaFaceHtml 一并渲染，这里不再单独赋值 */
    var fc = el('a22MegaFace');
    if (fc) { fc.innerHTML = _megaFaceHtml(d); fc.scrollTop = 0; }
    megaRailBind();
    /* 下一帧再算：innerHTML 刚写入时 scrollHeight 还没稳定 */
    requestAnimationFrame(function () { megaRailSync(); megaFitPanel(); });
    /* 两处范围下拉（卡面 + 悬浮面板）共用同一 state，任一改动都要同步另一处与自定义日期框 */
    ['a22MegaRange', 'a22MegaCardRange'].forEach(function (sid) {
      var sel = el(sid);
      if (!sel) return;
      sel.value = state.mega.range;
      sel.onchange = function () {
        state.mega.range = this.value;
        var cf = el('a22MegaCustom');
        if (cf) cf.style.display = state.mega.range === 'custom' ? 'inline' : 'none';
        megaLoad();
      };
    });
    var cf2 = el('a22MegaCustom');
    if (cf2) cf2.style.display = state.mega.range === 'custom' ? 'inline' : 'none';
    [['a22MegaFrom', 'from'], ['a22MegaTo', 'to']].forEach(function (p) {
      var e2 = el(p[0]);
      if (!e2) return;
      e2.value = state.mega[p[1]] || '';
      e2.onchange = function () {
        state.mega[p[1]] = this.value;
        if (state.mega.range === 'custom') megaLoad();
      };
    });
    var st = el('a22MegaStats');
    if (st) {
      st.innerHTML = '命中 <b>' + fmt(d.total) + '</b> 箱 · 图号 <b>' + fmt(d.figCount) + '</b> 种 · ' +
        '汇总 <b>' + fmt(d.groupCount) + '</b> 组 · ' + esc(d.window || '');
    }
    var det = el('a22MegaDet');
    if (det) {
      if (!d.details || !d.details.length) {
        det.innerHTML = '<div class="a22-mega-empty">该范围无关注记录</div>';
      } else {
        /* 明细展示与主明细表同口径的全部列（headers 由后端给出，行是等长数组） */
        var hd = d.headers || [];
        var h2 = '<table class="a22-mega-tab"><thead><tr>';
        hd.forEach(function (c, ci) { h2 += '<th' + (ci === 0 ? '' : '') + '>' + esc(c) + '</th>'; });
        h2 += '</tr></thead><tbody>';
        d.details.forEach(function (r) {
          h2 += '<tr>';
          r.forEach(function (v) { h2 += '<td>' + esc(v) + '</td>'; });
          h2 += '</tr>';
        });
        det.innerHTML = h2 + '</tbody></table>';
      }
    }
    var cb = el('a22MegaCopy');
    if (cb) cb.onclick = function () { megaCopyDetails(); };
  }
  function chartDaily15(d) {
    var c = mk('a22Daily15');
    if (!c) return;
    var _sel = document.getElementById('a22D15Range');
    var _rngVal = _sel ? _sel.value : '15';
    var _cust = document.getElementById('a22D15Custom');
    if (_cust) _cust.style.display = _rngVal === 'custom' ? 'inline' : 'none';
    var _d15 = (d.daily15 || []);
    if (_rngVal === 'custom') {
      var _f = (document.getElementById('a22D15From') || {}).value || '';
      var _t = (document.getElementById('a22D15To') || {}).value || '';
      if (_f) _d15 = _d15.filter(function (x) { return _dkey(x.date) >= _dkey(_f); });
      if (_t) _d15 = _d15.filter(function (x) { return _dkey(x.date) <= _dkey(_t); });
    } else if (_rngVal !== 'all') {
      _d15 = _d15.slice(0, parseInt(_rngVal, 10) || 15);
    }
    /* 标题三段式：前缀/数字/后缀各自独立 span，custom 时前缀藏「近」、后缀藏「个下达日箱数」，
     * 区间文案走独立副标题 a22D15CustomLabel——避免整段塞进 label 把「近 N 个」拆得难看（2026-10-17 用户打回）。 */
    var _lb = document.getElementById('a22D15RangeLabel');
    var _pre = document.getElementById('a22D15Prefix');
    var _suf = document.getElementById('a22D15Suffix');
    var _clb = document.getElementById('a22D15CustomLabel');
    var _hint = document.getElementById('a22D15Hint');
    if (_lb) {
      if (_rngVal === 'custom') {
        if (_pre) _pre.style.display = 'none';
        if (_suf) _suf.style.display = 'none';
        if (_hint) _hint.style.display = 'none';
        _lb.textContent = '自定义';
        var _f2 = (document.getElementById('a22D15From') || {}).value || '';
        var _t2 = (document.getElementById('a22D15To') || {}).value || '';
        if (_clb) _clb.textContent = (_f2 || _t2 ? (_f2 ? _f2.replace(/-/g, '.') : '…') + '~' + (_t2 ? _t2.replace(/-/g, '.') : '…') + ' · ' : '') + '共 ' + _d15.length + ' 个下达日';
      } else {
        if (_hint) _hint.style.display = '';
        if (_rngVal === 'all') {
          if (_pre) _pre.style.display = 'none';
          if (_suf) { _suf.style.display = ''; _suf.textContent = ' 全部下达日箱数'; }
          _lb.textContent = '';
        } else {
          if (_pre) _pre.style.display = '';
          if (_suf) { _suf.style.display = ''; _suf.textContent = ' 个下达日箱数'; }
          _lb.textContent = _rngVal === '7' ? '7' : (_rngVal === '30' ? '30' : '15');
        }
        if (_clb) _clb.textContent = '';
      }
    }
    var dates = _d15.map(function (x) { return x.date; });

    var finNames = [];
    _d15.forEach(function (x) { x.stacks.forEach(function (s) { if (finNames.indexOf(s.date) < 0) finNames.push(s.date); }); });
    var _pal15 = pal('a22Daily15');
    var finColor = {};
    finNames.forEach(function (n, i) { finColor[n] = _pal15 ? _pal15[i % _pal15.length] : FIN_PALETTE[i % FIN_PALETTE.length]; });
    var series = finNames.map(function (fn) {
      return {
        name: fn, type: 'bar', stack: 'total',
        itemStyle: { color: finColor[fn], borderRadius: fn === finNames[finNames.length - 1] ? [3, 3, 0, 0] : 0 },
        emphasis: { focus: 'series' },
        data: _d15.map(function (x) {
          var hit = x.stacks.filter(function (s) { return s.date === fn; })[0];
          return hit ? hit.value : 0;
        })
      };
    });
    var _gold15 = _pal15 ? _pal15[0] : GOLD;
    series.push({
      name: '下达量', type: 'line', yAxisIndex: 1,
      data: _d15.map(function (x) { return x.total; }),
      smooth: true, symbol: 'circle', symbolSize: 6,
      label: { show: true, position: 'top', color: _gold15, fontSize: 10, fontWeight: 600,
        textBorderColor: '#0B1F3A', textBorderWidth: 2,
        formatter: function (p) { return p.value > 0 ? p.value : ''; } },
      lineStyle: { color: _gold15, width: 2.4, shadowColor: 'rgba(245,200,106,.5)', shadowBlur: 8 },
      itemStyle: { color: _gold15, borderColor: '#0B1F3A', borderWidth: 1 },
      z: 10
    });
    c.setOption({
      tooltip: {
        trigger: 'axis', axisPointer: { type: 'shadow' },
        backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF', fontSize: 12 },
        formatter: function (ps) {
          var idx = ps[0].dataIndex, day = _d15[idx] || { total: 0 };
          var finCount = 0, rows = [];
          ps.forEach(function (p) {
            if (p.seriesName === '下达量') return;
            if (p.value > 0) { finCount++; rows.push('<span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:' + p.color + ';margin-right:6px"></span>' + p.seriesName + '：' + fmt(p.value)); }
          });
          return '<b>' + day.date + '</b>（下达 <b style="color:#F5C86A">' + fmt(day.total) + '</b> 箱）<br>完工日期分割：' + finCount + ' 种<br>' + rows.join('<br>');
        }
      },
      legend: { top: 2, left: 2, type: 'scroll', textStyle: { color: '#7E96C2', fontSize: 10 }, itemWidth: 10, itemHeight: 8, pageIconColor: ACC, inactiveColor: '#3D5578' },
      grid: baseGrid({ top: 34, left: 6, right: 8 }),
      xAxis: Object.assign({ type: 'category', data: dates, axisLabel: { color: '#7E96C2', fontSize: 10, rotate: 35 } }, AXIS),
      yAxis: [
        Object.assign({ type: 'value', name: '箱数', splitLine: { lineStyle: { color: 'rgba(84,130,204,.15)' } } }, AXIS),
        Object.assign({ type: 'value', name: '下达量', splitLine: { show: false } }, AXIS)
      ],
      series: series
    }, { notMerge: true });   // 替换模式：筛选区间变化后旧完工日 series 必须清掉，否则图例残留旧日期
  }

  /* ---------- 图2：角标 甜甜圈（点击下钻） ---------- */
  function chartCorner(d) {
    var c = mk('a22Corner');
    if (!c) return;
    var _rngEl = document.getElementById('a22CornerRange');
    var _rng = _rngEl ? _rngEl.value : 'day';
    var _data = (d.cornerByRange && d.cornerByRange[_rng]) ? d.cornerByRange[_rng] : d.corner.slice(0, 8);
    var _scope2 = ({ day: '当天', "7": '近 7 天', "15": '近 15 天', "30": '近 30 天', all: '全部', custom: '自定义' })[_rng];
    var items = _data.slice(0, 6);
    var _scope = document.getElementById('a22CornerScope');
    if (_scope) _scope.textContent = _scope2;
    c.setOption({
      tooltip: { trigger: 'item', backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF' }, formatter: '{b} 角<br/>{c} 箱（{d}%）<br><span style="color:#F5C86A;font-size:10px">点击下钻</span>' },
      legend: { bottom: 0, left: 'center', textStyle: { color: '#7E96C2', fontSize: 10 }, itemWidth: 10, itemHeight: 8 },
      series: [{
        type: 'pie', radius: ['50%', '74%'], center: ['50%', '42%'],
        itemStyle: { borderColor: '#0E2347', borderWidth: 2, borderRadius: 5 },
        label: { show: true, formatter: '{b}\n{d}%', color: '#A9BFE4', fontSize: 10, lineHeight: 13 },
        labelLine: { length: 8, length2: 8, lineStyle: { color: 'rgba(84,130,204,.5)' } },
        data: items.map(function (x, i) { var _pc = pal('a22Corner'); return { name: x.name, value: x.value, itemStyle: { color: _pc ? _pc[i % _pc.length] : PALETTE[i % PALETTE.length] } }; })
      }]
    });
    c.off('click');
    c.on('click', function (p) {
      if (p.componentType !== 'series') return;
      var name = p.name.replace(' 角', '');
      toggleFilter('corner', name);
    });
  }

  /* ---------- 图3：梯种 玫瑰图（迷你版，住在 KPI 预留位5；点击下钻） ---------- */
  function chartTizi(d) {
    var c = mk('a22Tizi');
    if (!c) return;
    var _rng = state.tizi.range;
    var _src = (d.tiziByRange && d.tiziByRange[_rng]) ? d.tiziByRange[_rng] : (d.tizi || []);
    var _scope2 = ({ day: '当天', "7": '近 7 天', "15": '近 15 天', "30": '近 30 天', all: '全部', custom: '自定义' })[_rng];
    var items = _src.slice(0, 12);
    var _scope = document.getElementById('a22TiziScope');
    if (_scope) _scope.textContent = '箱号第 3~5 位 · ' + _scope2;
    c.setOption({
      tooltip: { trigger: 'item', backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF', fontSize: 11 }, formatter: '梯种 {b}<br/>{c} 箱（{d}%）<br><span style="color:#F5C86A;font-size:10px">点击下钻</span>' },
      series: [{
        type: 'pie', roseType: 'area', radius: ['14%', '72%'], center: ['50%', '52%'],
        itemStyle: { borderColor: '#0E2347', borderWidth: 1.5, borderRadius: 4 },
        label: { color: '#A9BFE4', fontSize: 10, formatter: '{b}' },
        labelLine: { length: 6, length2: 6, lineStyle: { color: 'rgba(84,130,204,.5)' } },
        data: items.map(function (x, i) { var _pt = pal('a22Tizi'); return { name: x.name, value: x.value, itemStyle: { color: _pt ? _pt[(i + 1) % _pt.length] : PALETTE[(i + 1) % PALETTE.length] } }; })
      }]
    }, { notMerge: true });
    c.off('click');
    c.on('click', function (p) {
      if (p.componentType !== 'series') return;
      toggleFilter('tizi', p.name);
    });
  }

  /* ---------- 透视折叠树（安全钳/框架共用） ---------- */
  var EXPAND = {};
  function pivotTreeHtml(nodes, level, keyPrefix, colorClass, parentTotal, rootTotal) {
    var html = '';
    var sibTotal = 0; (nodes || []).forEach(function (x) { sibTotal += x.value; });
    (nodes || []).forEach(function (n, i) {
      var key = keyPrefix + i;
      var hasKids = n.children && n.children.length;
      // 占上级 = 本值 / 上一层级合计（parentTotal 由递归传入：顶层=总合计，子层=父节点值）
      var pParent = parentTotal ? n.value / parentTotal * 100 : 0;
      var pRoot = rootTotal ? n.value / rootTotal * 100 : 0;
      html += '<div class="apd22-pv-row lv' + level + '" data-pk="' + key + '">' +
        (hasKids ? '<span class="apd22-pv-toggle">' + (EXPAND[key] ? '−' : '+') + '</span>' : '<span class="apd22-pv-toggle none"></span>') +
        '<span class="apd22-pv-name c' + colorClass + '" title="' + esc(n.name) + '">' + esc(n.name) + '</span>' +
        '<span class="apd22-pv-bar"><i style="width:' + pvWidth(nodes, n) + '%"></i></span>' +
        '<span class="apd22-pv-val">' + fmt(n.value) + '</span>' +
        '<span class="apd22-pv-pct" title="占上级 ' + pParent.toFixed(1) + '% · 占总计 ' + pRoot.toFixed(1) + '%">' + pParent.toFixed(1) + '%<em>' + pRoot.toFixed(1) + '%</em></span></div>';
      if (hasKids && EXPAND[key]) html += pivotTreeHtml(n.children, level + 1, key + '-', colorClass, n.value, rootTotal);
    });
    return html;
  }
  function pvWidth(nodes, n) {
    var max = 0; nodes.forEach(function (x) { if (x.value > max) max = x.value; });
    return max ? Math.max(2, Math.round(n.value / max * 100)) : 0;
  }
  function bindPivot(hostId, rerender) {
    var host = el(hostId); if (!host) return;
    host.querySelectorAll('.apd22-pv-row').forEach(function (r) {
      r.addEventListener('click', function () {
        if (!r.querySelector('.apd22-pv-toggle')) return;
        var k = r.dataset.pk;
        EXPAND[k] = !EXPAND[k];
        rerender();
      });
    });
  }
  function buildPivotSel(selId, arr) {
    /* 「全区间」条目排到最前并默认选中（用户 2026-10-16：区间段要能选） */
    var ordered = arr.slice().sort(function (a, b) { return (a.date === '全区间' ? -1 : 0) - (b.date === '全区间' ? -1 : 0); });
    var sel = el(selId);
    if (sel && !sel.dataset.built) {
      sel.innerHTML = ordered.map(function (x, i) { return '<option' + (i === 0 ? ' selected' : '') + '>' + x.date + '</option>'; }).join('');
      sel.dataset.built = '1';
    } else if (sel) {
      var cur = sel.value;
      if (cur !== '全区间') cur = '全区间';
      sel.innerHTML = ordered.map(function (x) { return '<option' + (x.date === cur ? ' selected' : '') + '>' + x.date + '</option>'; }).join('');
    }
    return (sel && sel.value) ? (ordered.filter(function (x) { return x.date === sel.value; })[0] || ordered[0]) : ordered[0];
  }
  /* ---------- 图4：安全钳按下达日期透视 ---------- */
  function chartSafetyPivot(d) {
    var dd = buildPivotSel('a22SafetyPivotDate', d.safetyPivot);
    var host = el('a22SafetyPivot'); if (!host) return;
    host.innerHTML = '<div class="apd22-pv-head"><span>系列 / TYP / 图号</span><span>数量</span></div>' +
      '<div class="apd22-pv-total">' + (dd ? dd.date : '') + ' 合计 <b>' + fmt(dd ? dd.total : 0) + '</b> 箱 · 点击 + 展开</div>' +
      pivotTreeHtml(dd ? dd.children : [], 0, 'S', 'blue', dd ? dd.total : 0, dd ? dd.total : 0);
    bindPivot('a22SafetyPivot', function () { chartSafetyPivot(state.data); });
  }

  /* ---------- 矩形树图（treemap）公共：大方块=上级，内嵌小方块=下级，点击=下钻过滤 ----------
     排列规则（用户 2026-10-16 明令）：每层按数值降序——最大的块在左上，从上往下、从左往右越来越小 */
  /* treemap 专用浅色板（用户 2026-10-16：矩形图要浅色一点）：原色向白混合 35% */
  function lighten(hex, ratio) {
    var n = parseInt(hex.slice(1), 16), r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
    r = Math.round(r + (255 - r) * ratio); g = Math.round(g + (255 - g) * ratio); b = Math.round(b + (255 - b) * ratio);
    return 'rgb(' + r + ',' + g + ',' + b + ')';
  }
  var TM_PALETTES = {};
  function paleOf(palette) {
    var key = palette.join(',');
    if (!TM_PALETTES[key]) TM_PALETTES[key] = palette.map(function (c) { return lighten(c, 0.35); });
    return TM_PALETTES[key];
  }
  function treemapData(nodes, palette) {
    var sorted = (nodes || []).slice().sort(function (a, b) { return (b.value || 0) - (a.value || 0); });
    return sorted.map(function (n, i) {
      var o = { name: n.name, value: n.value, itemStyle: { color: palette[i % palette.length] } };
      if (n.children && n.children.length) {
        var kids = n.children.slice().sort(function (a, b) { return (b.value || 0) - (a.value || 0); });
        o.children = kids.map(function (c, j) {
          var co = { name: c.name, value: c.value, itemStyle: { color: palette[(i + 1 + j) % palette.length] } };
          if (c.children && c.children.length) {
            var gs = c.children.slice().sort(function (a, b) { return (b.value || 0) - (a.value || 0); });
            co.children = gs.map(function (g, k) {
              return { name: g.name, value: g.value, itemStyle: { color: palette[(i + j + 2 + k) % palette.length] } };
            });
          }
          return co;
        });
      }
      return o;
    });
  }
  function treemapOption(treeData, drillMap) {
    return {
      tooltip: {
        trigger: 'item', backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF', fontSize: 12 },
        formatter: function (p) { return '<b>' + esc(p.name) + '</b><br>' + fmt(p.value) + ' 箱<br><span style="color:#F5C86A;font-size:10px">点击下钻该层</span>'; }
      },
      series: [{
        type: 'treemap', roam: false, nodeClick: false, breadcrumb: { show: false },
        sort: 'desc', width: '100%', height: '100%', top: 4, left: 4, right: 4, bottom: 4,
        label: { color: '#12325C', fontSize: 11, fontWeight: 600 },
        upperLabel: { show: true, height: 20, color: '#2A4A78', fontSize: 10, fontWeight: 600, bgColor: 'rgba(227,238,252,.95)', borderWidth: 0 },
        levels: [
          { itemStyle: { borderWidth: 0, gapWidth: 3 }, upperLabel: { show: false } },
          { itemStyle: { borderColor: '#FFFFFF', borderWidth: 2, gapWidth: 2 }, upperLabel: { show: true, bgColor: 'rgba(227,238,252,.95)' } },
          { itemStyle: { borderWidth: 1.5, gapWidth: 1.5 } },
          { itemStyle: { borderWidth: 1, gapWidth: 1 } },
        ],
        data: treeData
      }]
    };
  }
  function bindTreemapDrill(chartId, treeData, layerOf) {
    var c = mk(chartId);
    if (!c) return;
    c.off('click');
    c.on('click', function (p) {
      // 顶层块 → 第 1 层过滤；子块 → 按其层级对应字段过滤
      var depth = 0, d = p.data;
      // treemap click 事件没有直接层级，用名称在树里查找
      function find(nodes, lvl) {
        for (var i = 0; i < (nodes || []).length; i++) {
          if (nodes[i].name === p.name) return lvl;
          var r = find(nodes[i].children, lvl + 1);
          if (r) return r;
        }
        return 0;
      }
      depth = find(treeData, 1);
      if (!depth) return;
      var field = layerOf(depth);
      if (field) toggleFilter(field, p.name);
    });
  }

  /* ---------- 图5：安全钳系列环形图（内圈=系列饼 · 外圈=TYP径向柱状图·柱面积∝用量 · 悬浮看数量/比例/图号明细） ----------
     数据修正（2026-10-17）：safetyFinishTree 顶层是完工日期，先聚合成 系列→TYP→图号 再画。
     视觉 v5（2026-10-17，用户指定色板）：
     - 内圈系列饼用海蓝渐变板 #0b2a3f…#ffffff（跳过最深与纯白，等距取样保证相邻瓣对比度），
       瓣间白色描边分界；浅色瓣自动配深字。
     - 外圈 polar 径向柱状图用洋红渐变板 #da4194…#410b29 按排名 亮→暗 插值；
       柱宽=槽位59%（原80%）；柱端不贴标签，TYP名/数量/比例/图号明细全部收进悬浮 tooltip。
     - 铺满布局：内饼0~38%、外环41~90%（中间空白带仅3%）；卡片改为近方形列
       （grid 2.6fr:1fr + 图表容器 aspect-ratio:1/1），圆环按短边撑满、左右不留白。
     交互 v6（2026-10-17 晚，用户要求）：
     - 内圈标签缩小到 10.5px，按瓣内弦宽估算，放不下就不显示——文字不得超出图形；
     - 内圈悬浮（任一位置）= 全量分层明细：所有系列→TYP→图号，数量+占比，全部倒序；
     - 外圈悬浮 = TYP 摘要 + 图号方块图（SVG treemap：上→下左→右递减、无间隙、
       洋红色板、字号自动收缩适配块大小、放不下只留色块并用 <title> 二次悬浮看明细）。 */
  function chartSafetyFinishTree(d) {
    var c = mk('a22SafetyFinishTree');
    if (!c) return;
    // 聚合：完工日期→系列→TYP→图号  ⇒  系列→TYP→图号
    var seriesMap = {};
    var total = 0;
    (d.safetyFinishTree || []).forEach(function (fd) {
      (fd.children || []).forEach(function (s) {
        var m = seriesMap[s.name] || (seriesMap[s.name] = { name: s.name, value: 0, typs: {} });
        m.value += (s.value || 0);
        total += (s.value || 0);
        (s.children || []).forEach(function (t) {
          var tm = m.typs[t.name] || (m.typs[t.name] = { name: t.name, value: 0, figs: {} });
          tm.value += (t.value || 0);
          (t.children || []).forEach(function (g) {
            tm.figs[g.name] = (tm.figs[g.name] || 0) + (g.value || 0);
          });
        });
      });
    });
    var seriesArr = Object.keys(seriesMap).map(function (k) { return seriesMap[k]; });
    seriesArr.sort(function (a, b) { return b.value - a.value; });

    /* ===== 视觉风格 v5（用户指定色板）=====
     1) 内圈系列饼：用户给的海蓝渐变板 #0b2a3f…#ffffff，按占比 从大到小、从深到浅 取色；
        跳过最深的 #0b2a3f 与纯白 #ffffff（用户允许不从最深开始），等距取样保证
        相邻瓣对比度足够（两块都深会糊成一片）；瓣与瓣之间用白色描边明确分界。
     2) 外圈 TYP 径向柱状图：用户给的洋红渐变板 #da4194…#410b29，按排名从亮到暗插值；
        柱宽=槽位 59%（原 74% 的 80%）；柱端不再贴标签（GSB-282 等全部收进悬浮明细），
        数量/比例/下钻图号明细都在 tooltip 里展示。
     3) 铺满布局：内饼 0~38%、外环 41~90%（中间空白带仅 3%）；图表容器为正方（CSS aspect-ratio），
        圆环按短边 90% 撑满，四周只剩一圈细边。 */
    function hex2rgb(hex) {
      var n = parseInt(hex.slice(1), 16);
      return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
    }
    function lum(hex) {  // 相对亮度 0~1，用于决定瓣内文字用白还是深色
      var c = hex2rgb(hex);
      return (0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]) / 255;
    }
    // 多停止点渐变插值（t: 0~1）
    function ramp(stops, t) {
      var x = Math.max(0, Math.min(1, t)) * (stops.length - 1);
      var i = Math.min(Math.floor(x), stops.length - 2), f = x - i;
      var a = hex2rgb(stops[i]), b = hex2rgb(stops[i + 1]);
      return 'rgb(' + Math.round(a[0] + (b[0] - a[0]) * f) + ',' +
        Math.round(a[1] + (b[1] - a[1]) * f) + ',' +
        Math.round(a[2] + (b[2] - a[2]) * f) + ')';
    }
    // 内圈：等距取样（跳过 #0b2a3f 与 #ffffff），保证相邻瓣有色差
    var CAND = ['#217ebc', '#3b98d6', '#51a3da', '#7cbae3', '#92c6e8', '#beddf1', '#e9f4fa'];
    var SN = seriesArr.length;
    var innerPicks = [];
    if (SN <= 1) innerPicks = [CAND[0]];
    else if (SN <= CAND.length) {
      for (var pi = 0; pi < SN; pi++) innerPicks.push(CAND[Math.round(pi * (CAND.length - 1) / (SN - 1))]);
    } else {
      var FULL = ['#0b2a3f'].concat(CAND, ['#ffffff']);
      for (var pj = 0; pj < SN; pj++) innerPicks.push(FULL[Math.round(pj * (FULL.length - 1) / (SN - 1)) % FULL.length]);
    }
    var innerData = seriesArr.map(function (s, i) {
      var col = innerPicks[i] || CAND[0];
      return {
        name: s.name, value: s.value,
        itemStyle: { color: col, borderColor: '#FFFFFF', borderWidth: 2.5 },
        // 浅色瓣配深字，深色瓣配白字
        label: { color: lum(col) > 0.62 ? '#0B2A3F' : '#FFFFFF' }
      };
    });
    // 外圈 TYP 全局降序（多的在 12 点起顺时针排前面）
    var typAll = [];
    seriesArr.forEach(function (s) {
      Object.keys(s.typs).forEach(function (k) { typAll.push({ name: k, value: s.typs[k].value, series: s.name }); });
    });
    typAll.sort(function (a, b) { return b.value - a.value; });
    var MAXBARS = 20;
    var bars = typAll.slice(0, MAXBARS);
    var otherFigs = null;
    if (typAll.length > MAXBARS) {
      var rest = typAll.slice(MAXBARS);
      otherFigs = {};
      var restSum = 0;
      rest.forEach(function (t) {
        restSum += t.value;
        var ty = seriesMap[t.series] && seriesMap[t.series].typs[t.name];
        if (ty) Object.keys(ty.figs).forEach(function (f) { otherFigs[f] = (otherFigs[f] || 0) + ty.figs[f]; });
      });
      bars.push({ name: '其他(' + rest.length + '个TYP)', value: restSum, series: '__other__' });
    }
    var N = bars.length, stepAng = 360 / Math.max(1, N);
    var maxV = bars.length ? bars[0].value : 1;
    var sqrtMax = Math.sqrt(maxV) * 1.02;          // 半径∝√用量（柱面积∝用量，玫瑰图惯例）
    var MAGENTA = ['#da4194', '#da2c88', '#a4278a', '#ae235d', '#991f5f', '#831a52', '#6d1644', '#571236', '#410b29'];
    var typNames = bars.map(function (t) { return t.name; });
    var barData = bars.map(function (t, rank) {
      var lt = N > 1 ? rank / (N - 1) : 0;         // 排名0最亮，越往后越暗（洋红渐变板）
      return {
        name: t.name, value: Math.sqrt(t.value), raw: t.value, series: t.series,
        itemStyle: { color: ramp(MAGENTA, lt), borderColor: 'rgba(255,255,255,.5)', borderWidth: 0.8 }
      };
    });

    /* ---- tooltip 明细排版工具 ----
       estW：文本宽度估算（CJK=1em、ASCII=0.58em），用于判断文字是否放得进图形内；
       treemapRows：行式 treemap——每行从左到右、整块从上到下按值递减铺满，无间隙；
       figTreemapSvg：把图号列表画成 SVG 方块图，字号自动收缩适配块大小，
       放不下文字的块只留色块，靠 <title> 原生悬浮显示二次明细。 */
    function estW(text, fs) {
      var w = 0;
      for (var i = 0; i < text.length; i++) w += (text.charCodeAt(i) > 255 ? 1 : 0.58) * fs;
      return w;
    }
    function lumRgb(c) {
      var m = (c.match(/\d+/g) || [0, 0, 0]);
      return (0.299 * +m[0] + 0.587 * +m[1] + 0.114 * +m[2]) / 255;
    }
    function treemapRows(items, W, H) {
      var out = [], i = 0, x = 0, y = 0, w = W, h = H;
      var remain = items.reduce(function (a, b) { return a + b.value; }, 0) || 1;
      var guard = 0;
      while (i < items.length && w > 0.5 && h > 0.5 && guard++ < 500) {
        // 贪心组行：逐项加入，取"最差长宽比"最小的行
        var sum = 0, bestWorst = Infinity, bestCount = 1;
        for (var k = 0; k < items.length - i; k++) {
          sum += items[i + k].value;
          var rh = h * sum / remain, worst = 0;
          for (var m = 0; m <= k; m++) {
            var iw = w * items[i + m].value / sum;
            var asp = Math.max(iw / rh, rh / iw);
            if (asp > worst) worst = asp;
          }
          if (worst <= bestWorst) { bestWorst = worst; bestCount = k + 1; } else break;
        }
        var rowSum = 0;
        for (var n2 = 0; n2 < bestCount; n2++) rowSum += items[i + n2].value;
        var rh2 = h * rowSum / remain, cx = x;
        for (var q = 0; q < bestCount; q++) {
          var it = items[i + q], iw2 = w * it.value / rowSum;
          if (it.value > 0) out.push({ name: it.name, value: it.value, x: cx, y: y, w: iw2, h: rh2 });
          cx += iw2;
        }
        y += rh2; h -= rh2; remain -= rowSum; i += bestCount;
      }
      return out;
    }
    function figTreemapSvg(figs) {
      var W = 480, H = 240, n = figs.length;
      var sum = figs.reduce(function (a, b) { return a + b.value; }, 0) || 1;
      var blocks = treemapRows(figs, W, H);
      var sv = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%" height="' + H + '" style="display:block;border-radius:4px" xmlns="http://www.w3.org/2000/svg">';
      blocks.forEach(function (b, rank) {
        var pct = b.value / sum * 100;
        var col = ramp(MAGENTA, n > 1 ? rank / (n - 1) : 0);
        var txtCol = lumRgb(col) > 0.62 ? '#3A0A26' : '#FFFFFF';
        var l1 = b.name, l2 = fmt(b.value) + ' · ' + pct.toFixed(1) + '%';
        // 字号适配：11→7 递减，优先双行（图号+数量占比），退化为单行，再退化不显示
        var fs = 0, two = false;
        for (var f = 11; f >= 7; f -= 0.5) {
          if (b.h >= f * 2.8 && b.w >= Math.max(estW(l1, f), estW(l2, f)) + 8) { fs = f; two = true; break; }
          if (b.h >= f * 1.7 && b.w >= estW(l1, f) + 6) { fs = f; two = false; break; }
        }
        sv += '<g><title>' + esc(l1) + '：' + fmt(b.value) + ' 箱 · 占该 TYP ' + pct.toFixed(1) + '%</title>' +
          '<rect x="' + b.x.toFixed(1) + '" y="' + b.y.toFixed(1) + '" width="' + b.w.toFixed(1) + '" height="' + b.h.toFixed(1) +
          '" fill="' + col + '" stroke="#FFFFFF" stroke-width="1" stroke-opacity=".85"/>';
        var cxp = (b.x + b.w / 2).toFixed(1);
        if (fs > 0 && two) {
          sv += '<text x="' + cxp + '" y="' + (b.y + b.h / 2 - 1).toFixed(1) + '" text-anchor="middle" dominant-baseline="middle" font-size="' + fs + '" font-weight="700" fill="' + txtCol + '" font-family="Inter,sans-serif">' + esc(l1) + '</text>' +
            '<text x="' + cxp + '" y="' + (b.y + b.h / 2 + fs * 1.05).toFixed(1) + '" text-anchor="middle" dominant-baseline="middle" font-size="' + (fs - 1.5) + '" fill="' + txtCol + '" opacity=".85" font-family="Inter,sans-serif">' + esc(l2) + '</text>';
        } else if (fs > 0) {
          sv += '<text x="' + cxp + '" y="' + (b.y + b.h / 2).toFixed(1) + '" text-anchor="middle" dominant-baseline="middle" font-size="' + fs + '" font-weight="700" fill="' + txtCol + '" font-family="Inter,sans-serif">' + esc(l1) + '</text>';
        }
        sv += '</g>';
      });
      return sv + '</svg>';
    }

    c.setOption({
      tooltip: {
        trigger: 'item', backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF', fontSize: 12 },
        confine: true, extraCssText: 'max-height:72vh; overflow:auto;',
        formatter: function (p) {
          /* ---- 内圈系列饼：悬浮任一位置 → 全量分层明细（系列→TYP→图号，全部倒序） ---- */
          if (p.seriesIndex === 0) {
            var h = '<div style="width:310px">';
            seriesArr.forEach(function (s) {
              h += '<div style="display:flex;justify-content:space-between;gap:14px;font-weight:700;color:#EAF3FF;margin-top:7px;line-height:17px">' +
                '<span>' + esc(s.name) + '</span><span>' + fmt(s.value) + ' 箱 · ' + (total ? (s.value / total * 100).toFixed(1) : 0) + '%</span></div>';
              var typs = Object.keys(s.typs).map(function (k) { return { name: k, value: s.typs[k].value, figs: s.typs[k].figs }; });
              typs.sort(function (a, b) { return b.value - a.value; });
              typs.forEach(function (t) {
                h += '<div style="display:flex;justify-content:space-between;gap:14px;color:#7CD8FF;padding-left:12px;line-height:16px">' +
                  '<span>└ ' + esc(t.name) + '</span><span>' + fmt(t.value) + ' · ' + (total ? (t.value / total * 100).toFixed(1) : 0) + '%</span></div>';
                var figs = Object.keys(t.figs).map(function (k) { return { name: k, value: t.figs[k] }; });
                figs.sort(function (a, b) { return b.value - a.value; });
                figs.forEach(function (g) {
                  h += '<div style="display:flex;justify-content:space-between;gap:14px;color:#C9D8EE;padding-left:26px;font-size:11px;line-height:15px">' +
                    '<span>' + esc(g.name) + '</span><span>' + fmt(g.value) + ' · ' + (total ? (g.value / total * 100).toFixed(1) : 0) + '%</span></div>';
                });
              });
            });
            h += '<div style="color:#F5C86A;font-size:10px;margin-top:5px">悬浮饼图任一位置显示全部 · 数量与占比均为全区间口径 · 点击下钻该系列</div></div>';
            return h;
          }
          /* ---- 外圈 TYP 柱：头部摘要 + 图号方块图（treemap） ---- */
          var figSrc = null;
          if (p.data && p.data.series === '__other__') {
            figSrc = Object.keys(otherFigs || {}).map(function (k) { return { name: k, value: otherFigs[k] }; });
          } else {
            var s2 = seriesMap[p.data.series];
            var hit = s2 ? s2.typs[p.name] : null;
            if (hit) figSrc = Object.keys(hit.figs).map(function (k) { return { name: k, value: hit.figs[k] }; });
          }
          var isOther = !!(p.data && p.data.series === '__other__');
          var sName = isOther ? '尾部聚合' : p.data.series;
          var head = '<b>' + esc(p.name) + '</b>（' + esc(sName) + ' · ' + fmt(p.data.raw) + ' 箱 · 全区间占 ' + (total ? (p.data.raw / total * 100).toFixed(1) : 0) + '%' +
            (!isOther && seriesMap[p.data.series] ? ' · 占所属系列 ' + (p.data.raw / seriesMap[p.data.series].value * 100).toFixed(1) + '%' : '') + '）' +
            '<div style="color:#8FA6CC;font-size:10.5px;margin:2px 0 5px">图号用量方块图（上→下 · 左→右 递减 · 悬浮方块看明细）：</div>';
          if (!figSrc || !figSrc.length) return head + '<span style="color:#F5C86A;font-size:10px">点击下钻该 TYP</span>';
          figSrc.sort(function (a, b) { return b.value - a.value; });
          return '<div style="width:504px">' + head + figTreemapSvg(figSrc) +
            '<div style="color:#F5C86A;font-size:10px;margin-top:4px">点击下钻该 TYP</div></div>';
        }
      },
      polar: { center: ['50%', '50%'], radius: ['41%', '90%'] },
      angleAxis: {
        type: 'category', data: typNames, startAngle: 90, clockwise: true,
        axisLine: { show: false }, axisTick: { show: false }, axisLabel: { show: false },
        splitLine: { show: false }
      },
      radiusAxis: {
        min: 0, max: sqrtMax,
        axisLine: { show: false }, axisTick: { show: false }, axisLabel: { show: false },
        splitLine: { show: true, lineStyle: { color: 'rgba(255,255,255,.10)', width: 0.8 } },
        splitArea: { show: false }
      },
      series: [
        {
          name: '系列', type: 'pie', radius: ['0%', '38%'], center: ['50%', '50%'],
          startAngle: 90, clockwise: true, z: 3,
          label: {
            position: 'inside', fontSize: 10.5, fontWeight: 700,
            textBorderColor: 'rgba(10,26,52,.55)', textBorderWidth: 2,
            formatter: function (p) {
              // 文字不得超出瓣面：按瓣内弦宽估算——单行放得下就单行；
              // 单行放不下但分行放得下就两行（名称/百分比各一行）；都放不下就不显示
              if (p.percent < 2) return '';
              var label = p.name + ' ' + p.percent + '%';
              var chord = 2 * (0.19 * Math.min(c.getWidth(), c.getHeight())) * Math.sin(p.percent / 100 * Math.PI);
              if (estW(label, 10.5) <= chord * 0.92) return label;
              if (Math.max(estW(p.name, 10.5), estW(p.percent + '%', 10.5)) <= chord * 0.92) return p.name + '\n' + p.percent + '%';
              return '';
            }
          },
          data: innerData
        },
        {
          name: 'TYP', type: 'bar', coordinateSystem: 'polar', z: 2,
          barCategoryGap: '41%',
          label: { show: false },
          data: barData
        }
      ]
    });
    c.off('click');
    c.on('click', function (p) {
      if (p.componentType !== 'series') return;
      if (p.seriesIndex === 0) { toggleFilter('s_series', p.name); return; }  // 内圈：系列下钻
      if (p.data && p.data.series === '__other__') return;                     // 聚合柱不下钻
      toggleFilter('s_typ', p.name);                                           // 外圈：TYP 下钻
    });
  }

  /* ---------- 图6：框架按下达日期透视 ---------- */
  function chartFramePivot(d) {
    var dd = buildPivotSel('a22FramePivotDate', d.framePivot);
    var host = el('a22FramePivot'); if (!host) return;
    host.innerHTML = '<div class="apd22-pv-head"><span>系列 / TYP / 图号</span><span>数量</span></div>' +
      '<div class="apd22-pv-total">' + (dd ? dd.date : '') + ' 合计 <b>' + fmt(dd ? dd.total : 0) + '</b> 箱 · 点击 + 展开</div>' +
      pivotTreeHtml(dd ? dd.children : [], 0, 'F', 'green', dd ? dd.total : 0, dd ? dd.total : 0);
    bindPivot('a22FramePivot', function () { chartFramePivot(state.data); });
  }

  /* ---------- 图7：框架 系列→图号 矩形树图（Tableau 参考样式 · 2026-10-18 重做 v3） ----------
     参考用户给的 Tableau 树形图（每天学点Tableau 03）：
     ① 每个系列一个色系家族：父块=中色调，子块=同色系由深入浅渐变（取代原来的彩虹轮盘取色）；
     ② 块内多行标注：名称 / 占比% · 数值——块越大行数越多；小块只留名称；放不下就不显示
        （文字不得超出块面，明细靠悬浮）；
     ③ 所有块白色描边 + 间隙，边界清晰；
     ④ 按数值降序：大块在左上，从上往下、从左往右越来越小；
     ⑤ 悬浮显示 系列/图号 路径 + 箱数 + 占比；点击下钻（系列→f_series，图号→f_fig）。
     数据修正（v2）：frameSun 的 TYP 层与系列同名且每系列仅 1 项（完全冗余），
     直接拍平成 系列→图号 两层，图号明细才能展示（旧版把 TYP 当子块，图号全丢了）。
     标签贴合（v3）：文字放不放得下不能靠面积估算（窄高块宽度远小于 √面积），改为两遍渲染——
     第一遍不出标签，读 ECharts 真实布局矩形后按块宽/块高精确回填，超宽则隐藏（不留截断残句）；
     父块眉标走 levels[1].upperLabel.formatter（数据项级 upperLabel 会被 ECharts 忽略）。 */
  var FRAME_HUES = {
    'E系列': '#2E6FD8', 'M系列': '#7C3AED', '无机房': '#0E9F9F', '汽车梯': '#DE8A00',
    '分体式700S': '#D84A6B', '分体式500S': '#3FA7F5', '导靴机架': '#C2570F',
    '限速器机架': '#5E9E2F', '小型化': '#B8448E',
    '小型化无机房': '#3FA7F5', '小型化E系列': '#C2570F', '小型化有机房': '#5E9E2F'
  };
  var FRAME_HUE_LIST = ['#2E6FD8', '#7C3AED', '#0E9F9F', '#DE8A00', '#D84A6B', '#3FA7F5', '#C2570F', '#5E9E2F', '#B8448E', '#8A94A6', '#C9A227'];
  function tmLum(c) { var m = (c.match(/\d+/g) || [0, 0, 0]); return (0.299 * +m[0] + 0.587 * +m[1] + 0.114 * +m[2]) / 255; }
  function tmEstW(t, fs) { var w = 0; for (var i = 0; i < t.length; i++) w += (t.charCodeAt(i) > 255 ? 1 : 0.62) * fs; return w; }
  function tmFind(nodes, name) { for (var i = 0; i < (nodes || []).length; i++) if (nodes[i].name === name) return nodes[i]; return null; }
  function chartFrameSun(d) {
    var c = mk('a22FrameSun');
    if (!c) return;
    var raw = (d.frameSun || []).slice().sort(function (a, b) { return (b.value || 0) - (a.value || 0); });
    var total = raw.reduce(function (a, x) { return a + (x.value || 0); }, 0) || 1;
    var area = Math.max(1, c.getWidth() * c.getHeight());
    var UPPER_TXT = {};                                   // 系列名 → 父块眉标文本（'' = 不显示）
    var treeData = raw.map(function (s, i) {
      var base = FRAME_HUES[s.name] || FRAME_HUE_LIST[i % FRAME_HUE_LIST.length];
      // 拍平：TYP 层冗余（每系列仅 1 个同名 TYP），把各 TYP 下的图号提为子块（同名图号合并）
      var figMap = {}, figs = [];
      (s.children || []).forEach(function (t) {
        var gs = t.children || [];
        if (!gs.length) gs = [{ name: t.name, value: t.value }];   // 无图号时退回 TYP 本身占位
        gs.forEach(function (g) {
          if (!figMap[g.name]) { figMap[g.name] = { name: g.name, value: 0 }; figs.push(figMap[g.name]); }
          figMap[g.name].value += (g.value || 0);
        });
      });
      figs.sort(function (a, b) { return b.value - a.value; });
      var sPct = s.value / total * 100;
      var sideP = Math.sqrt(sPct / 100 * area);            // 父块近似边长（仅无子块的叶子兜底用）
      var pCol = lighten(base, 0.10);
      return {
        name: s.name, value: s.value,
        itemStyle: { color: pCol, borderColor: '#FFFFFF', borderWidth: 1 },
        label: figs.length ? { show: false } : {           // 无子块时父块自己当叶子展示
          show: sideP >= 24 && tmEstW(s.name, 10) <= sideP * 1.7, position: 'inside',
          formatter: s.name + '\n' + sPct.toFixed(1) + '% · ' + fmt(s.value),
          color: tmLum(pCol) > 0.55 ? '#1B2A45' : '#FFFFFF',
          fontSize: 10, fontWeight: 600, lineHeight: 14
        },
        children: figs.map(function (k, j) {
          var lt = figs.length > 1 ? j / (figs.length - 1) : 0;    // 子块同色系由深入浅
          var col = lighten(base, 0.30 + 0.40 * lt);
          return {
            name: k.name, value: k.value, __series: s.name,
            itemStyle: { color: col, borderColor: '#FFFFFF', borderWidth: 1 },
            label: {
              show: false, position: 'inside',                 // 文本由 patchLabels 按真实矩形回填
              color: tmLum(col) > 0.55 ? '#1B2A45' : '#FFFFFF',
              fontSize: 10, fontWeight: 600, lineHeight: 14
            }
          };
        })
      };
    });
    function buildOption() {
      return {
        animation: false,
        tooltip: {
          trigger: 'item', backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF', fontSize: 12 },
          formatter: function (p) {
            var isParent = !!(p.data && p.data.children && p.data.children.length);
            var label = p.data && p.data.__series ? (p.data.__series + ' / ' + p.name) : p.name;
            var pct = (p.value / total * 100).toFixed(1);
            return '<b>' + esc(label) + '</b><br>' + fmt(p.value) + ' 箱 · 占全区间 ' + pct + '%' +
              '<br><span style="color:#F5C86A;font-size:10px">点击下钻' + (isParent ? '该系列' : '该图号') + '</span>';
          }
        },
        series: [{
          type: 'treemap', roam: false, nodeClick: false, breadcrumb: { show: false },
          sort: 'desc', width: '100%', height: '100%', top: 3, left: 3, right: 3, bottom: 3,
          levels: [
            { itemStyle: { borderWidth: 0, gapWidth: 0 }, upperLabel: { show: false } },
            { itemStyle: { borderColor: '#FFFFFF', borderWidth: 1, gapWidth: 1.5 },
              // 注意：眉标带底色=父块描边色（白），故文字必须用深色，白字会隐形
              upperLabel: { show: true, color: '#1B2A45', fontSize: 12, fontWeight: 700, height: 18,
                formatter: function (p) { return UPPER_TXT[p.name] || ''; } } },
            { itemStyle: { borderColor: '#FFFFFF', borderWidth: 1, gapWidth: 1 }, upperLabel: { show: false } }
          ],
          data: treeData
        }]
      };
    }
    /* 按真实布局矩形回填标签：眉标（父）与块内文字（子）都按块宽/块高精确判断，放不下就不显示 */
    function patchLabels() {
      try {
        var s = c.getModel().getSeriesByIndex(0);
        var tree = s.getData().tree;
        if (!tree || !tree.root || !tree.root.children) return false;
        var changed = false;
        tree.root.children.forEach(function (inNode) {
          var myS = tmFind(treeData, inNode.name);
          if (!myS) return;
          var hasKids = !!(myS.children && myS.children.length);
          var r = typeof inNode.getLayout === 'function' ? inNode.getLayout() : null;
          if (r && !r.invisible) {
            var txt = myS.name + '  ' + (myS.value / total * 100).toFixed(1) + '%';
            var show = hasKids && r.height >= 30 && tmEstW(txt, 12) * 1.12 + 10 <= r.width;  // 18px 眉标 + 至少 12px 子块
            var ut = show ? txt : '';
            if (myS.__putxt !== ut) { myS.__putxt = ut; UPPER_TXT[myS.name] = ut; changed = true; }
          }
          (inNode.children || []).forEach(function (inKid) {
            var myK = hasKids ? tmFind(myS.children, inKid.name) : null;
            if (!myK || !myK.label) return;
            var kr = typeof inKid.getLayout === 'function' ? inKid.getLayout() : null;
            if (!kr || kr.invisible) return;
            var availW = kr.width - 16, availH = kr.height - 12;
            var l2 = (myK.value / total * 100).toFixed(1) + '% · ' + fmt(myK.value);
            var t = '';
            if (availW >= 30 && availH >= 30 && tmEstW(l2, 9.5) * 1.12 <= availW && tmEstW(myK.name, 10) * 1.12 <= availW) t = myK.name + '\n' + l2;
            else if (availW >= 24 && availH >= 16 && tmEstW(myK.name, 10) * 1.12 <= availW) t = myK.name;
            if (myK.__ptxt !== t) {
              myK.__ptxt = t; myK.label.show = !!t; myK.label.formatter = t; changed = true;
            }
          });
        });
        return changed;
      } catch (e) { return false; }
    }
    /* 回填后必须 clear() 再全量重渲：ECharts 增量更新不会重绘已用空文本创建过的 upperLabel
       （实测：直接 setOption 第二遍，眉标 formatter 有返回也不渲染） */
    var drillLayer = function (depth) { return depth === 1 ? 'f_series' : (depth === 2 ? 'f_fig' : ''); };
    function relabel() {
      if (!patchLabels()) return;
      c.clear();
      c.setOption(buildOption(), true);
      bindTreemapDrill('a22FrameSun', treeData, drillLayer);
    }
    c.setOption(buildOption());
    [40, 150, 400, 900].forEach(function (ms) { setTimeout(relabel, ms); });   // 渲染完成后回填（多次兜底）
    bindTreemapDrill('a22FrameSun', treeData, drillLayer);
    // 视口变化后 ECharts 会自动重排版，标签需按新矩形重新回填
    var dom = el('a22FrameSun');
    if (dom && !dom.__tmRz) {
      dom.__tmRz = 1;
      var tmr = null;
      window.addEventListener('resize', function () {
        clearTimeout(tmr);
        tmr = setTimeout(relabel, 220);
      });
    }
    window.__frameSun = { upper: UPPER_TXT, tree: treeData };   // 调试/自检句柄
  }

/* ---------- 图8：BG 帕累托（悬浮看构成 · 点击下钻） ---------- */
  function chartBgPareto(d) {
    var c = mk('a22BgPareto');
    if (!c) return;
    var items = d.bgPareto;
    var cums = [], acc = 0, total = d.bgTotal || 1;
    items.forEach(function (x) { acc += x.value; cums.push(+(acc / total * 100).toFixed(1)); });
    c.setOption({
      tooltip: {
        trigger: 'item', backgroundColor: TIP_BG, borderWidth: 0, textStyle: { color: '#E5F0FF' },
        formatter: function (p) {
          var it = items[p.dataIndex];
          if (!it || !it.breakdown || !it.breakdown.length) return 'BG' + it.bg + '：' + fmt(it.value) + ' 箱';
          var rows = it.breakdown.slice(0, 6).map(function (b) {
            return b.series + ' / ' + b.typ + ' / ' + b.fig + '：<b>' + fmt(b.value) + '</b>';
          }).join('<br>');
          var more = it.extra > 0 ? '<br>… 另有 ' + it.extra + ' 组' : '';
          return '<b>BG' + it.bg + '</b>（' + fmt(it.value) + ' 箱 · 占 ' + (it.value / total * 100).toFixed(1) + '% · 累计 ' + it.cumPct + '%）<br>' + rows + more +
            '<br><span style="color:#F5C86A;font-size:10px">点击柱体下钻该 BG</span>';
        }
      },
      grid: baseGrid({ top: 30, left: 6, right: 54 }),
      xAxis: Object.assign({ type: 'category', data: items.map(function (x) { return x.bg; }), axisLabel: { color: '#7E96C2', fontSize: 10, rotate: 35 } }, AXIS),
      yAxis: [
        Object.assign({ type: 'value', name: '箱数', splitLine: { lineStyle: { color: 'rgba(84,130,204,.15)' } } }, AXIS),
        Object.assign({ type: 'value', name: '累计 %', max: 100, splitLine: { show: false }, axisLabel: { formatter: '{value}%', color: '#7E96C2', fontSize: 10 } }, AXIS)
      ],
      series: [
        {
          type: 'bar', data: items.map(function (x, i) {
            // 同色系渐变递淡：按排名从主蓝逐渐变淡（不透明度 1 → 0.25），「其他」用灰
            var isOther = items[i] && items[i].bg === '其他';
            var alpha = isOther ? 0.25 : Math.max(0.25, 1 - i * (0.75 / Math.max(1, items.length - 1)));
            var _pb = pal('a22BgPareto');
            var col = isOther ? '#5A7398' : (_pb ? _pb[0] : ACC);
            return { value: x.value, itemStyle: { color: { type: 'linear', x: 0, y: 0, x2: 0, y2: 1, colorStops: [{ offset: 0, color: col, opacity: alpha }, { offset: 1, color: col, opacity: alpha * 0.45 }] }, borderRadius: [3, 3, 0, 0] } };
          }),
          barMaxWidth: 30
        },
        { type: 'line', yAxisIndex: 1, data: cums, symbol: 'circle', symbolSize: 4.5, lineStyle: { color: GRN, width: 1.8, shadowColor: 'rgba(46,230,168,.4)', shadowBlur: 6 }, itemStyle: { color: GRN } }
      ]
    });
    c.off('click');
    c.on('click', function (p) {
      if (p.componentType === 'series' && p.seriesIndex === 0) {
        var it = items[p.dataIndex];
        if (it && it.bg !== '其他') toggleFilter('bg', it.bg);
      }
    });
  }

  function openBgModal(it, total) {
    var m = el('a22BgModal');
    if (!m || !it) return;
    el('a22BgTitle').textContent = 'BG' + it.bg + ' · 系列/型号/图号 分布';
    var sum = it.breakdown.reduce(function (a, b) { return a + b.value; }, 0) || 1;
    el('a22BgBody').innerHTML = '<table class="apd22-bgtable"><thead><tr><th>系列</th><th>TYP 型号</th><th>图号</th><th>数量</th><th>占比</th><th style="width:200px">分布</th></tr></thead><tbody>' +
      it.breakdown.map(function (b) {
        var p = (b.value / sum * 100).toFixed(1);
        return '<tr><td>' + esc(b.series) + '</td><td>' + esc(b.typ) + '</td><td>' + esc(b.fig) + '</td><td class="num">' + fmt(b.value) + '</td><td class="num">' + p + '%</td><td class="td-bar"><i style="width:' + p + '%"></i></td></tr>';
      }).join('') +
      (it.extra > 0 ? '<tr class="muted"><td colspan="3">其他组合</td><td class="num">' + fmt(it.value - sum) + '</td><td class="num">' + ((it.value - sum) / sum * 100).toFixed(1) + '%</td><td></td></tr>' : '') +
      '</tbody><tfoot><tr><td colspan="3">合计</td><td class="num">' + fmt(it.value) + '</td><td class="num">100%</td><td></td></tfoot></table>';
    m.classList.add('show');
  }

  /* ---------- 图9/10：框架属性 × 图号 透视（下达/完工 双视角） ---------- */
  var ATTR_EXPAND = {};
  function renderAttrPivot(d, dataKey, hostId, selId, expandKey) {
    var arr = d[dataKey];
    if (!arr || !arr.length) { var h0 = el(hostId); if (h0) h0.innerHTML = '<div class="apd22-pv-total">无数据</div>'; return; }
    var sel = el(selId);
    if (sel && !sel.dataset.built) {
      /* 「全区间」置顶默认（用户 2026-10-16） */
      var ordered = arr.slice().sort(function (a, b) { return (a.date === '全区间' ? -1 : 0) - (b.date === '全区间' ? -1 : 0); });
      sel.innerHTML = ordered.map(function (x, i) { return '<option' + (i === 0 ? ' selected' : '') + '>' + x.date + '</option>'; }).join('');
      sel.dataset.built = '1';
      sel.__ordered = ordered;
    } else if (sel) {
      var ordered = sel.__ordered || arr;
      var cur = sel.value || '全区间';
      sel.innerHTML = ordered.map(function (x) { return '<option' + (x.date === cur ? ' selected' : '') + '>' + x.date + '</option>'; }).join('');
    }
    var orderedAll = (sel && sel.__ordered) || arr;
    var dd = (sel && sel.value) ? (orderedAll.filter(function (x) { return x.date === sel.value; })[0] || orderedAll[0]) : orderedAll[0];
    var dd = (sel && sel.value) ? (arr.filter(function (x) { return x.date === sel.value; })[0] || arr[0]) : arr[best !== undefined ? best : 0];
    var host = el(hostId);
    if (!host) return;
    host.innerHTML = '<div class="apd22-pv-head"><span>框架属性 / 图号</span><span>数量 · 占比</span></div>' +
      '<div class="apd22-pv-total">合计 <b>' + fmt(dd ? dd.total : 0) + '</b> 箱 · 点击属性行展开图号</div>' +
      (dd ? dd.children : []).map(function (a, i) {
        var key = expandKey + i;
        var open = ATTR_EXPAND[key];
        var cls = a.name === 'X' ? 'attr-x' : (a.name === 'F' ? 'attr-f' : 'attr-n');
        var html = '<div class="apd22-pv-row lv0" data-ak="' + key + '">' +
          '<span class="apd22-pv-toggle">' + (open ? '−' : '+') + '</span>' +
          '<span class="apd22-pv-name"><b class="apd22-badge ' + cls + '">' + esc(a.name) + '</b></span>' +
          '<span class="apd22-pv-bar"><i style="width:' + (a.pct || 0) + '%"></i></span>' +
          '<span class="apd22-pv-val">' + fmt(a.value) + ' · ' + a.pct + '%</span></div>';
        if (open && a.children && a.children.length) {
          html += a.children.map(function (fg) {
            return '<div class="apd22-pv-row lv1 static"><span class="apd22-pv-toggle none"></span>' +
              '<span class="apd22-pv-name dim">' + esc(fg.name) + '</span>' +
              '<span class="apd22-pv-bar"><i style="width:' + Math.max(2, Math.round(fg.value / a.value * (a.pct || 1) * 2)) + '%"></i></span>' +
              '<span class="apd22-pv-val">' + fmt(fg.value) + '</span></div>';
          }).join('');
        }
        return html;
      }).join('');
    host.querySelectorAll('.apd22-pv-row[data-ak]').forEach(function (r) {
      r.addEventListener('click', function () {
        var k = r.dataset.ak;
        ATTR_EXPAND[k] = !ATTR_EXPAND[k];
        renderAttrPivot(state.data, dataKey, hostId, selId, expandKey);
      });
    });
  }
  function chartAttrPivots(d) {
    renderAttrPivot(d, 'attrByIssue', 'a22AttrIssue', 'a22AttrIssueDate', 'AI');
    renderAttrPivot(d, 'attrByFinish', 'a22AttrFinish', 'a22AttrFinishDate', 'AF');
  }

  /* ---------- 下钻联动 ---------- */
  function toggleFilter(group, val) {
    var arr = state.f[group];
    var i = arr.indexOf(val);
    if (i >= 0) arr.splice(i, 1); else arr.push(val);
    renderFilterBar();
    load();
  }

  /* ---------- 明细表 ---------- */
  var DETAIL_KEY = 'apd22_detail';
  function loadDetail(date) {
    console.log('[dbg] loadDetail 开始 date=' + date);
    return fetch('/api/apd22/detail?date=' + encodeURIComponent(date))
      .then(function (r) { console.log('[dbg] detail 响应 ' + r.status); return r.json(); })
      .then(function (d) {
        console.log('[dbg] detail 数据 rows=' + (d.rows || []).length);
        state.detail.rows = d.rows; state.detail.headers = d.headers; state.detail.date = d.date;
        state.detail.filtered = null; state.detail.page = 1;
        renderDetail();
        console.log('[dbg] renderDetail 完成');
      }).catch(function (e) { console.log('[dbg] loadDetail 异常: ' + e); });
  }
  function detailRowsView() {
    var rows = state.detail.rows;
    var f = state.detail.filterText;
    if (f) {
      var fl = f.toLowerCase();
      rows = rows.filter(function (r) { return r.join('¬').toLowerCase().indexOf(fl) >= 0; });
    }
    if (state.detail.sortCol != null) {
      var ci = state.detail.sortCol, dir = state.detail.sortDir;
      rows = rows.slice().sort(function (a, b) {
        var x = a[ci], y = b[ci];
        var nx = parseFloat(x), ny = parseFloat(y);
        if (!isNaN(nx) && !isNaN(ny) && x !== '' && y !== '') return (nx - ny) * dir;
        return String(x).localeCompare(String(y), 'zh-CN') * dir;
      });
    }
    return rows;
  }
  function renderDetail() {
    var host = el('a22DetailWrap');
    if (!host) return;
    var headers = state.detail.headers;
    var view = detailRowsView();
    var st = TableUX.pageSlice(view, DETAIL_KEY);
    var sortHtml = function (i) { return state.detail.sortCol === i ? (state.detail.sortDir > 0 ? ' ↑' : ' ↓') : ''; };
    host.innerHTML = '<table class="apd22-dtable"><thead><tr>' +
      headers.map(function (h, i) { return '<th data-si="' + i + '" title="点击排序">' + esc(h) + sortHtml(i) + '</th>'; }).join('') +
      '</tr></thead><tbody>' +
      st.rows.map(function (r) {
        return '<tr style="background:#0D1E3C">' + r.map(function (v, i) {
          var cls = /^\d+$/.test(v) ? ' class="num"' : '';
          return '<td' + cls + '>' + esc(v) + '</td>';
        }).join('') + '</tr>';
      }).join('') + '</tbody></table>';
    /* 列宽默认自适应各列字数长度（用户 2026-10-16 明令） */
    try {
      (function () {
        var tbl = host.querySelector('.apd22-dtable');
        if (!tbl) return;
        var min = 64, max = 320, pad = 24;
        var widths = headers.map(function (h, ci) {
          var w = TableUX.textW(h, 12) + 18;
          var rows = st.rows;
          var step = Math.max(1, Math.ceil(rows.length / 300));
          for (var i = 0; i < rows.length; i += step) {
            var v = rows[i][ci] == null ? '' : String(rows[i][ci]);
            w = Math.max(w, TableUX.textW(v, 12));
          }
          return Math.min(max, Math.max(min, w + pad));
        });
        var cg = '<colgroup>' + widths.map(function (w) { return '<col style="width:' + w + 'px">'; }).join('') + '</colgroup>';
        tbl.insertAdjacentHTML('afterbegin', cg);
        tbl.style.tableLayout = 'fixed';
        tbl.style.width = 'max-content';
        tbl.style.minWidth = '100%';
      })();
    } catch (e) { /* 自适应失败不阻塞渲染 */ }
    var metaEl = el('a22DetailMeta');
    if (metaEl) metaEl.textContent = '下达日期 ' + (state.detail.date || '-') + ' · 共 ' + st.total + ' 条 · 第 ' + st.page + '/' + st.pages + ' 页 · 每页 ' + st.size + ' 条 · 点击列头排序';
    el('a22DetailPager').innerHTML = TableUX.pagerHtml(DETAIL_KEY, st);
    TableUX.bindPager(el('a22DetailPager'), DETAIL_KEY, st.pages, renderDetail);
    host.querySelectorAll('th[data-si]').forEach(function (th) {
      th.addEventListener('click', function () {
        var i = +th.dataset.si;
        if (state.detail.sortCol === i) { state.detail.sortDir = -state.detail.sortDir; if (state.detail.sortDir === 1 && th.dataset.c === '2') { state.detail.sortCol = null; } }
        else { state.detail.sortCol = i; state.detail.sortDir = 1; }
        th.dataset.c = ((+th.dataset.c || 0) + 1);
        renderDetail();
      });
    });
    var meta = el('a22DetailMeta');
    if (meta) meta.textContent = '下达日期 ' + state.detail.date + ' · 共 ' + fmt(view.length) + ' 行' + (state.detail.filterText ? '（筛选后）' : '');
  }

  /* ---------- 自定义查询（条件组保存） ---------- */
  function loadSavedQueries() {
    try { state.detail.savedQueries = JSON.parse(localStorage.getItem('apd22_savedQueries') || '[]'); } catch (e) { state.detail.savedQueries = []; }
  }
  function saveSavedQueries() {
    try { localStorage.setItem('apd22_savedQueries', JSON.stringify(state.detail.savedQueries)); } catch (e) { }
  }
  function renderSavedQueries() {
    var host = el('a22SavedQueries');
    if (!host) return;
    host.innerHTML = state.detail.savedQueries.map(function (q, i) {
      return '<span class="apd22-sq" data-i="' + i + '" title="' + esc(q.desc) + '">' + esc(q.name) + '<i data-del="' + i + '">✕</i></span>';
    }).join('');
    host.querySelectorAll('.apd22-sq').forEach(function (s) {
      s.addEventListener('click', function (e) {
        if (e.target.dataset.del != null) {
          state.detail.savedQueries.splice(+e.target.dataset.del, 1);
          saveSavedQueries(); renderSavedQueries(); return;
        }
        var q = state.detail.savedQueries[+s.dataset.i];
        el('a22DetailSearch').value = q.expr || '';
        state.detail.filterText = q.expr || '';
        renderDetail();
      });
    });
  }

  /* ---------- 总渲染 ---------- */
  function renderAll() {
    var d = state.data;
    if (!d || d.empty) {
      var s = el('apd22Status'); if (s) s.textContent = '筛选条件下无数据，请放宽筛选。';
      showLoading(false);
      return;
    }
    var dl = drillLabel();
    el('apd22Status').textContent = 'MySQL sap_table.03_apd · ' + fmt(d.kpi.total) + ' 行 · ' + d.kpi.dates + ' 个下达日' + (dl ? ' · ' + dl : '') + ' · 更新于 ' + new Date().toLocaleTimeString('zh-CN', { hour12: false });
    renderKpi(d);
    if (window.Apd22LayoutRearm) window.Apd22LayoutRearm();
    /* v5 渐进渲染：KPI 与首图立即出，其余图表逐个渲染、每张之间让出主线程（rAF+setTimeout），
       避免 10 张 ECharts 同步 setOption 把首屏堵到几千毫秒（2026-10-18 首屏 3s 优化） */
    var steps = [chartDaily15, chartCorner, chartTizi, chartSafetyPivot, chartSafetyFinishTree,
                 chartFramePivot, chartFrameSun, chartBgPareto, chartAttrPivots];
    var i = 0;
    window.__a22Steps = [];   // 调试用：各图渲染耗时（控制台/自动化可读）
    var _t0 = Date.now();
    function next() {
      if (i >= steps.length) { showLoading(false); return; }
      var fn = steps[i++];
      var _ts = Date.now();
      try { fn(d); } catch (e) { console.error('apd22 chart fail', e); }
      window.__a22Steps.push(fn.name + ':' + (Date.now() - _ts) + 'ms@' + (Date.now() - _t0));
      if (i < steps.length) requestAnimationFrame(function () { setTimeout(next, 0); });
      else showLoading(false);
    }
    next();
  }

  /* ---------- 全量刷新：重建 MySQL 03#箱databoard 数据处理表 + 重拉看板 ---------- */
  function fullRefresh() {
    var btn = el('apd22FullRefresh');
    var s = el('apd22Status');
    if (btn) { btn.disabled = true; btn.textContent = '刷新中…'; }
    if (s) s.textContent = '全量刷新中：正在重建 03#箱databoard 全部数据处理表（约 16 万行，请稍候）…';
    fetch('/api/apd22/refresh', { method: 'POST' })
      .then(function (r) { if (!r.ok) return r.json().then(function (j) { throw new Error(j.detail || ('HTTP ' + r.status)); }); return r.json(); })
      .then(function (d) {
        var parts = Object.keys(d.tables || {}).map(function (k) { return k + ' ' + fmt(d.tables[k]) + ' 行'; });
        if (s) s.textContent = '全量刷新完成 ' + (d.ts || '') + ' · 03#箱databoard：' + parts.join(' · ');
        return loadOptions();
      })
      .then(function () { load(); })
      .catch(function (e) {
        if (s) s.textContent = '全量刷新失败：' + e.message;
      })
      .finally(function () { if (btn) { btn.disabled = false; btn.textContent = '全量刷新'; } });
  }

  /* ---------- 自动刷新：3s 轮询后端数据版本号，MySQL 03_apd 数据变化 → 看板自动重载 ----------
   * 后端 /api/apd22/auto-sync/status 每次响应前顺带做库变化检测（毫秒级 COUNT），
   * 有变化当次即带新 version，故端到端延迟≈3s 轮询间隔（用户要求 5 秒内重载）。
   * 后端守护线程绝不自动入库；数据入库唯一入口是「立即同步」按钮（manualSync）。
   * 重载走 load()：state.f 是筛选真源，load() 以 state.f 为 payload，
   * 因此自动刷新不会清掉用户已选的筛选/下钻条件。 */
  var autoSync = { version: -1, timer: null, busy: false, lastSync: '', lastInserted: 0 };

  function autoSyncBadge(text, tone) {
    var b = el('apd22AutoSync');
    if (!b) return;
    b.textContent = text;
    b.className = 'a22-autosync' + (tone ? ' ' + tone : '');
  }

  function pollAutoSync() {
    fetch('/api/apd22/auto-sync/status', { cache: 'no-store' })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d || d.ok === false) return;
        autoSync.lastSync = d.last_sync || '';
        autoSync.lastInserted = d.last_inserted || 0;
        if (autoSync.version < 0) {
          autoSync.version = d.version || 0;
          autoSyncBadge('自动刷新已启用 · 检查于 ' + (d.last_check || '').slice(-8) +
                        (autoSync.lastSync ? ' · 最近同步 ' + autoSync.lastSync.slice(5) : ''), '');
          return;
        }
        if ((d.version || 0) !== autoSync.version) {
          autoSync.version = d.version || 0;
          autoSyncBadge('检测到新数据，正在自动刷新…', 'busy');
          loadOptions().then(function () { return load(); }).then(function () {
            autoSyncBadge('已自动刷新 · ' + (autoSync.lastSync || ''), 'ok');
          }).catch(function (e) {
            autoSyncBadge('自动刷新失败：' + e.message, 'err');
          });
          return;
        }
        if (d.running) { autoSyncBadge('正在同步源数据…', 'busy'); return; }
        if (d.error) { autoSyncBadge('同步异常：' + d.error, 'err'); return; }
        autoSyncBadge('自动刷新已启用 · 检查于 ' + (d.last_check || '').slice(-8) +
                      (autoSync.lastSync ? ' · 最近同步 ' + autoSync.lastSync.slice(5) : ''), '');
      })
      .catch(function () { /* 后端未起或网络抖动：静默，不打扰用户 */ });
  }

  function startAutoSyncPoll() {
    if (autoSync.timer) return;
    pollAutoSync();
    autoSync.timer = setInterval(pollAutoSync, 3000);   // 3s 一次（5 秒内完成自动重载）
  }

  /* 手工同步数据并更新看板（唯一数据入库入口：源目录订单下达记录 → sap_table.03_apd） */
  function manualSync() {
    var btn = el('apd22AutoSyncBtn');
    if (btn) { btn.disabled = true; btn.textContent = '同步中…'; }
    autoSyncBadge('正在同步源数据…', 'busy');
    fetch('/api/apd22/auto-sync/run', { method: 'POST' })
      .then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error(j.detail || ('HTTP ' + r.status)); return j; }); })
      .then(function (j) {
        autoSyncBadge('同步完成 · 新增 ' + (j.inserted || 0) + ' 行', 'ok');
        if (j.version !== autoSync.version) { autoSync.version = j.version; return loadOptions().then(function () { return load(); }); }
      })
      .catch(function (e) { autoSyncBadge('同步失败：' + e.message, 'err'); })
      .finally(function () { if (btn) { btn.disabled = false; btn.textContent = '立即同步'; } });
  }

  /* ---------- 初始化 ---------- */
  function init() {
    if (!el('apd22Page')) return;
    if (typeof echarts === 'undefined') { var s = el('apd22Status'); if (s) s.textContent = 'echarts.min.js 未加载'; return; }
    loadSavedQueries(); renderSavedQueries();
    el('apd22Refresh').addEventListener('click', function () { load(); });
    el('apd22FullRefresh').addEventListener('click', fullRefresh);
    el('apd22Reset').addEventListener('click', resetFilters);
    /* 自动刷新：进入看板即开启 3s 轮询；页面切到后台时暂停，回到前台立即查一次 */
    startAutoSyncPoll();
    (function () {
      var btnSync = el('apd22AutoSyncBtn');
      if (btnSync) btnSync.addEventListener('click', manualSync);
      document.addEventListener('visibilitychange', function () {
        if (document.hidden) { if (autoSync.timer) { clearInterval(autoSync.timer); autoSync.timer = null; } }
        else { startAutoSyncPoll(); pollAutoSync(); }
      });
    })();
    // 冻结筛选区折叠/展开（默认展开；每次进入页面都回到展开态）
        // KPI 悬浮柱：二次悬浮显示该完工日的下达日期贡献（事件委托）


    el('apd22Fold').addEventListener('click', function () {
      var wrap = document.querySelector('.apd22-sticky-wrap');
      if (!wrap) return;
      var folded = wrap.classList.toggle('folded');
      this.textContent = folded ? '展开筛选 ▼' : '收起筛选 ▲';
      this.title = folded ? '展开筛选区' : '折叠筛选区（冻结区收起后只保留工具栏）';
    });
    // 默认收起筛选（用户 2026-10-17 要求）：进页面即折叠，按钮显示「展开筛选 ▼」
    (function () {
      var wrap0 = document.querySelector('.apd22-sticky-wrap');
      if (!wrap0) return;
      wrap0.classList.add('folded');
      var btn0 = el('apd22Fold');
      if (btn0) { btn0.textContent = '展开筛选 ▼'; btn0.title = '展开筛选区'; }
    })();
    el('a22SafetyPivotDate').addEventListener('change', function () { chartSafetyPivot(state.data); });
    el('a22D15Range').addEventListener('change', function () { if (state.data) chartDaily15(state.data); });
    el('a22CornerRange').addEventListener('change', function () {
      var _c = document.getElementById('a22CornerCustom');
      if (_c) _c.style.display = this.value === 'custom' ? 'inline' : 'none';
      if (state.data) chartCorner(state.data);
    });
    ['a22D15From', 'a22D15To'].forEach(function (id2) {
      var e2 = el(id2);
      if (e2) e2.addEventListener('change', function () { if (state.data) chartDaily15(state.data); });
    });
    ['a22CornerFrom', 'a22CornerTo'].forEach(function (id2) {
      var e2 = el(id2);
      if (e2) e2.addEventListener('change', function () {
        var s2 = el('a22CornerRange');
        if (s2 && s2.value === 'custom') { load(); }
      });
    });
    /* 梯种玫瑰（每日下达区第三卡，控件在 index.html 静态存在）：与角标图同套路直接绑定。
     * 交互顺序：先写 state.tizi（真源）→ 普通范围/无日期 custom 本地重绘；custom+日期 → load() 重请求（约 4s，scope 先行提示） */
    el('a22TiziRange').addEventListener('change', function () {
      state.tizi.range = this.value;
      state.tizi.from = (el('a22TiziFrom') || {}).value || '';
      state.tizi.to = (el('a22TiziTo') || {}).value || '';
      var _c = document.getElementById('a22TiziCustom');
      if (_c) _c.style.display = this.value === 'custom' ? 'inline' : 'none';
      var _hasDates = state.tizi.from || state.tizi.to;
      if (this.value === 'custom' && _hasDates) {
        var _sc0 = document.getElementById('a22TiziScope');
        if (_sc0) _sc0.textContent = '箱号第 3~5 位 · 自定义（加载中…）';
        if (state.data) load();
        return;
      }
      if (state.data) chartTizi(state.data);
    });
    ['a22TiziFrom', 'a22TiziTo'].forEach(function (id2) {
      var e2 = el(id2);
      if (e2) e2.addEventListener('change', function () {
        state.tizi.from = (el('a22TiziFrom') || {}).value || '';
        state.tizi.to = (el('a22TiziTo') || {}).value || '';
        var s2 = el('a22TiziRange');
        if (s2 && s2.value === 'custom') {
          var _sc = document.getElementById('a22TiziScope');
          if (_sc) _sc.textContent = '箱号第 3~5 位 · 自定义（加载中…）';
          load();
        }
      });
    });
    el('a22FramePivotDate').addEventListener('change', function () { chartFramePivot(state.data); });
    el('a22AttrIssueDate').addEventListener('change', function () { renderAttrPivot(state.data, 'attrByIssue', 'a22AttrIssue', 'a22AttrIssueDate', 'AI'); });
    el('a22AttrFinishDate').addEventListener('change', function () { renderAttrPivot(state.data, 'attrByFinish', 'a22AttrFinish', 'a22AttrFinishDate', 'AF'); });
    el('a22DetailSearch').addEventListener('input', function () {
      state.detail.filterText = this.value.trim(); state.detail.page = 1; renderDetail();
    });
    el('a22DetailSaveQ').addEventListener('click', function () {
      var name = prompt('条件组名称：', '查询' + (state.detail.savedQueries.length + 1));
      if (!name) return;
      state.detail.savedQueries.push({ name: name, expr: el('a22DetailSearch').value.trim(), desc: '全文筛选：' + el('a22DetailSearch').value.trim() });
      saveSavedQueries(); renderSavedQueries();
    });
    
  /* ---------- 导出：Excel（后端 openpyxl 多工作表）/ HTML（前端原样快照） ---------- */
  function bindApd22Export() {
    var btn = el('apd22Export'), menu = el('apd22ExportMenu');
    if (!btn || !menu || btn.dataset.bound) return;
    btn.dataset.bound = '1';
    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      menu.style.display = (menu.style.display === 'none') ? 'flex' : 'none';
    });
    document.addEventListener('click', function () { menu.style.display = 'none'; });
    menu.addEventListener('click', function (e) { e.stopPropagation(); });
    el('apd22ExportXlsx').addEventListener('click', function () {
      menu.style.display = 'none';
      var body = JSON.stringify(state.filter || {});
      btn.textContent = '导出中…';
      fetch('/api/apd22/export_xlsx', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: body })
        .then(function (r) {
          if (!r.ok || (r.headers.get('content-type') || '').indexOf('sheet') < 0) return r.json().then(function (j) { throw new Error(j.error || '导出失败'); });
          return r.blob();
        })
        .then(function (blob) {
          var a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = '3号箱APD及安全钳看板_' + new Date().toISOString().slice(0, 10) + '.xlsx';
          a.click();
        })
        .catch(function (err) { alert('Excel 导出失败：' + err.message); })
        .finally(function () { btn.textContent = '导出 ▼'; });
    });
    el('apd22ExportHtml').addEventListener('click', function () {
      menu.style.display = 'none';
      btn.textContent = '快照中…';
      try {
        var page = el('apd22Page');
        var clone = page.cloneNode(true);
        /* 图表 canvas → 内嵌图片 */
        var srcCanvases = page.querySelectorAll('canvas');
        var dstCanvases = clone.querySelectorAll('canvas');
        srcCanvases.forEach(function (cv, i) {
          try {
            var img = document.createElement('img');
            img.src = cv.toDataURL('image/png');
            img.style.width = cv.style.width || cv.offsetWidth + 'px';
            img.style.height = cv.style.height || cv.offsetHeight + 'px';
            if (dstCanvases[i] && dstCanvases[i].parentNode) dstCanvases[i].parentNode.replaceChild(img, dstCanvases[i]);
          } catch (e) { /* 跨域 canvas 忽略 */ }
        });
        /* 收集全部样式表 */
        var css = '';
        document.querySelectorAll('style').forEach(function (s) { css += s.textContent + '\n'; });
        var doc = '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8">\n<title>3号箱APD及安全钳看板 · 快照</title>\n<style>' + css + '</style>\n';
        doc += '<style>body{margin:0;background:#071830;} .apd22-page{display:block !important;}</style>\n</head>\n<body>\n';
        doc += clone.outerHTML;
        doc += '\n</body>\n</html>';
        var blob = new Blob([doc], { type: 'text/html;charset=utf-8' });
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = '3号箱APD及安全钳看板_快照_' + new Date().toISOString().slice(0, 10) + '.html';
        a.click();
      } catch (e) { alert('HTML 快照失败：' + e.message); }
      btn.textContent = '导出 ▼';
    });
  }
  bindApd22Export();

  el('a22DetailExport').addEventListener('click', function () {
      var headers = state.detail.headers, rows = detailRowsView();
      var csv = '\ufeff' + headers.join(',') + '\n' + rows.map(function (r) {
        return r.map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(',');
      }).join('\n');
      var a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = '3号箱APD明细222_' + state.detail.date + '.csv';
      a.click();
    });
    el('a22BgClose').addEventListener('click', function () { el('a22BgModal').classList.remove('show'); });
    el('a22BgModal').addEventListener('click', function (e) { if (e.target === this) this.classList.remove('show'); });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') el('a22BgModal').classList.remove('show'); });
    A22Layout.init();
    loadOptions();
    load();
  }

  /* ---------- 布局编辑器（2026-10-17 用户要求，作用于本3号箱APD及安全钳看板）----------
     v2（2026-10-17 追加）：文字可编辑/删除/拖动移位；图表可改配色（预设色板+主色）。
     v3（2026-10-18 用户明令「保存才生效 + Ctrl+Z 撤销 + Ctrl+Y 恢复 + 一键还原」）：
       · 草稿态：拖拽/改文字/改配色只标记未保存（保存钮亮红点），绝不自动落盘；
       · 保存（Ctrl+S 或点「保存」）：快照压历史栈并 POST，此刻才「生效」；
       · 撤销（Ctrl+Z）：有未保存草稿先丢弃草稿；否则回退一步并持久化；
       · 恢复（Ctrl+Y 或 Ctrl+Shift+Z）：重做被撤销的一步；
       · 还原默认：confirm 后清空全部偏好，该步同样可被 Ctrl+Z 撤销；
       · applyPrefs(快照) = resetDom（清内联样式回 CSS 默认）+ applyAll（重放快照）。
     默认关闭：关闭态不加任何把手/画笔，看板渲染与原版完全一致；「编辑布局」按钮开启。 */
  var A22Layout = {
    NS: 'apd22',
    on: false,
    prefs: { cards: {}, texts: {}, charts: {} },   // 当前草稿（内存态；未保存时也可能含改动）
    unsaved: false,                                // 有未保存修改
    history: [],                                   // 已保存快照栈
    histIdx: 0,                                    // 当前指针：history[histIdx] = 屏幕与服务端的当前态
    clone: function (o) { return JSON.parse(JSON.stringify(o || {})); },
    /* 归一化 + 深拷贝：norm 的产物会成为 this.prefs，若浅拷贝则 collect() 的原地写
       会污染 history 里的已保存快照（2026-10-18 E2E 抓出：草稿值漏进旧快照） */
    norm: function (p) {
      p = p || {};
      return { cards: this.clone(p.cards), texts: this.clone(p.texts), charts: this.clone(p.charts) };
    },
    init: function () {
      var btn = el('a22LayoutEdit');
      if (!btn) return;
      var self = this;
      fetch('/api/board/layout?ns=' + this.NS).then(function (r) { return r.json(); }).then(function (p) {
        self.prefs = self.norm(p);
        self.history = [self.clone(self.prefs)];
        self.histIdx = 0;
        self.prefsLoaded = true;
        window.A22_CHART_PREFS = self.prefs.charts;   // 图表取色 pal() 在首轮 renderAll 前就要能读到
        self.applyAll();
        self.updateButtons();
      }).catch(function () {
        self.prefsLoaded = true;
        self.history = [self.norm(self.prefs)];
        self.histIdx = 0;
      });
      btn.addEventListener('click', function () { self.toggle(); });
      var bind = function (id, fn) { var b = el(id); if (b) b.addEventListener('click', function () { fn.call(self); }); };
      var fb0 = el('a22LayoutFold');
      if (fb0) fb0.addEventListener('click', function (e) { e.stopPropagation(); self.togglePop(); });
      bind('a22LayoutSave', this.save);
      bind('a22LayoutUndo', this.undo);
      bind('a22LayoutRedo', this.redo);
      bind('a22LayoutReset', this.restoreDefault);
      // 快捷键（本页可见且焦点不在输入框时）：Ctrl+S 保存 / Ctrl+Z 撤销 / Ctrl+Y(或 Ctrl+Shift+Z) 恢复
      document.addEventListener('keydown', function (e) {
        var pg = el('apd22Page');
        if (!pg || pg.offsetParent === null) return;          // 3号箱APD及安全钳看板 不在当前页
        var t = e.target || {};
        var tag = (t.tagName || '').toUpperCase();
        if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || t.isContentEditable) return;
        if (!(e.ctrlKey || e.metaKey)) return;
        var k = (e.key || '').toLowerCase();
        if (k === 's') { e.preventDefault(); self.save(); }
        else if (k === 'z' && !e.shiftKey) { e.preventDefault(); self.undo(); }
        else if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); self.redo(); }
      });
    },
    popOpen: true,                                 // 布局操作面板展开态（默认展开）
    /* 折叠/展开「保存·撤销·恢复·还原默认」面板（▲ 收起 / ▼ 下拉） */
    togglePop: function (force) {
      // 手动点击（无 force）只在编辑态响应；toggle() 退出编辑时会用 force=false 强制收起
      if (force === undefined && !this.on) return;
      this.popOpen = (force === undefined) ? !this.popOpen : !!force;
      var pop = el('a22LayoutPop');
      var fb = el('a22LayoutFold');
      if (pop) pop.classList.toggle('show', this.popOpen);
      if (fb) {
        fb.textContent = this.popOpen ? '▲' : '▼';
        fb.title = (this.popOpen ? '收起 ▲' : '展开 ▼') + ' 布局操作面板（保存 / 撤销 / 恢复 / 还原默认）';
        fb.classList.toggle('a22-pop-folded', !this.popOpen);
      }
      this.updateButtons();
    },
    toggle: function () {
      if (this.on && this.unsaved) {
        if (!confirm('有未保存的布局修改，退出编辑将丢弃。\n确定退出？（建议先 Ctrl+S 保存）')) return;
        this.unsaved = false;
        this.applyPrefs(this.history[this.histIdx]);   // 丢弃草稿，回到已保存态
      }
      this.on = !this.on;
      el('apd22Page').classList.toggle('a22-editing', this.on);
      var btn = el('a22LayoutEdit');
      btn.classList.toggle('a22-edit-on', this.on);
      btn.textContent = this.on ? '退出编辑' : '编辑布局';
      var fb = el('a22LayoutFold');
      if (fb) fb.classList.toggle('show', this.on);   // 折叠钮随编辑态出现/隐藏
      this.togglePop(this.on);                        // 进入编辑默认展开，退出编辑自动收起
      if (this.on) this.arm(); else this.disarm();
      this.updateButtons();
    },
    // 稳定键：卡自身 id → 卡内第一个带 id 的元素 id（a22Daily15 / a22Kpi1 / a22SafetyPivot…）
    keyOf: function (card) {
      if (!card) return null;
      if (card.dataset.lzKey) return card.dataset.lzKey;
      if (card.id) { card.dataset.lzKey = card.id; return card.id; }
      var inner = card.querySelector('[id]');
      if (inner && inner.id) { card.dataset.lzKey = inner.id; return inner.id; }
      return null;
    },
    /* 卡内 canvas → ECharts 图表键（mk() 注册用的容器 id，如 a22Daily15）。
       卡的 keyOf 是标题 id（a22D15Title），与图表键不同，必须映射。 */
    chartKeyOf: function (card) {
      if (!card) return null;
      if (card.dataset.lzChartKey) return card.dataset.lzChartKey;
      var CHART_IDS = ['a22Daily15', 'a22Corner', 'a22Tizi', 'a22SafetyFinishTree', 'a22FrameSun', 'a22BgPareto'];
      for (var i = 0; i < CHART_IDS.length; i++) {
        var dom = document.getElementById(CHART_IDS[i]);
        if (dom && card.contains(dom) && dom.querySelector('canvas')) { card.dataset.lzChartKey = CHART_IDS[i]; return CHART_IDS[i]; }
      }
      return null;
    },
    arm: function () {
      var self = this;
      document.querySelectorAll('#apd22Page .apd22-card, #apd22Page .apd22-kpi').forEach(function (card, ci) {
        if (card.dataset.lzArmed) return;
        card.dataset.lzArmed = '1';
        if (!card.dataset.lzKey) card.dataset.lzKey = self.keyOf(card) || ('idx#' + ci);
        var rz = document.createElement('i'); rz.className = 'a22lz-r'; rz.title = '拖动调宽度（本排所有卡同步变）';
        var bz = document.createElement('i'); bz.className = 'a22lz-b'; bz.title = '拖动调高度（仅本卡）';
        card.appendChild(rz); card.appendChild(bz);
        self.wireW(rz, card); self.wireH(bz, card);
        // 图表配色按钮（仅 ECharts 卡）
        var ck = self.chartKeyOf(card);
        if (ck) {
          var cb = document.createElement('button');
          cb.className = 'a22-chart-btn'; cb.textContent = '🎨'; cb.title = '改图表配色（色板+主色，也可只调主色）';
          card.appendChild(cb);
          cb.addEventListener('click', function (e) { e.stopPropagation(); self.openChartPanel(ck, cb); });
        }
      });
      // 文字：画笔（改）+ 移动手柄（拖动移位）+ 删除钮
      var armText = function (t, card, kind) {
        if (!t || t.dataset.penArmed) return;
        t.dataset.penArmed = '1';
        var tKey = self.keyOf(card) + '.' + kind;
        t.style.position = t.style.position || 'relative';
        var pen = document.createElement('button');
        pen.className = 'a22-pen'; pen.textContent = '✎'; pen.title = '改文字颜色/字号/内容';
        var mv = document.createElement('button');
        mv.className = 'a22-pen a22-mv'; mv.textContent = '✥'; mv.title = '拖动移位（按住拖）';
        var del = document.createElement('button');
        del.className = 'a22-pen a22-del'; del.textContent = '✕'; del.title = '删除此文字';
        t.appendChild(pen); t.appendChild(mv); t.appendChild(del);
        pen.addEventListener('click', function (e) { e.stopPropagation(); self.openPen(tKey, t, pen); });
        del.addEventListener('click', function (e) {
          e.stopPropagation();
          if (!confirm('确定删除「' + (t.textContent.replace(/[✎✥✕]/g, '') || '此文字') + '」？（还原=编辑布局面板里清空偏好）')) return;
          t.style.display = 'none';
          self.prefs.texts[tKey] = Object.assign({}, self.prefs.texts[tKey] || {}, { hidden: true });
          self.dirty();
        });
        self.wireMove(mv, t, tKey);
        var p = self.prefs.texts[tKey];
        if (p) self.styleText(t, p);
      };
      document.querySelectorAll('#apd22Page .apd22-card-head').forEach(function (head) {
        var card = head.closest('.apd22-card');
        if (!card) return;
        var b = head.querySelector('b'); if (b) armText(b, card, 'title');
        var sub = head.querySelector('.apd22-card-sub'); if (sub) armText(sub, card, 'sub');
      });
      document.querySelectorAll('#apd22Page .apd22-kpi').forEach(function (card) {
        var t = card.querySelector('.apd22-kpi-title'); if (t) armText(t, card, 'title');
        var s = card.querySelector('.apd22-kpi-sub'); if (s) armText(s, card, 'sub');
      });
      this.applyChartPrefs();
    },
    // 文字拖动移位：absolute 偏移 dx/dy 持久化（transform 不与原布局冲突）
    wireMove: function (mv, t, tKey) {
      var self = this;
      mv.addEventListener('pointerdown', function (e) {
        e.preventDefault(); e.stopPropagation();
        var p = self.prefs.texts[tKey] || {};
        var dx = p.dx || 0, dy = p.dy || 0;
        var sx = e.clientX, sy = e.clientY;
        try { mv.setPointerCapture(e.pointerId); } catch (err) { }
        document.body.classList.add('a22-lz-resizing');
        var move = function (ev) {
          dx = (p.dx || 0) + (ev.clientX - sx);
          dy = (p.dy || 0) + (ev.clientY - sy);
          t.style.transform = 'translate(' + dx + 'px,' + dy + 'px)';
        };
        var up = function () {
          document.body.classList.remove('a22-lz-resizing');
          mv.removeEventListener('pointermove', move); mv.removeEventListener('pointerup', up); mv.removeEventListener('pointercancel', up);
          if (dx || dy) {
            self.prefs.texts[tKey] = Object.assign({}, self.prefs.texts[tKey] || {}, { dx: dx, dy: dy });
          } else {
            if (self.prefs.texts[tKey]) { delete self.prefs.texts[tKey].dx; delete self.prefs.texts[tKey].dy; }
          }
          self.dirty();
        };
        mv.addEventListener('pointermove', move);
        mv.addEventListener('pointerup', up);
        mv.addEventListener('pointercancel', up);
      });
    },
    disarm: function () {
      // 把手/画笔/配色钮一起清：旧版只清把手，Rearm 会重复挂出一堆 ✎✥✕（2026-10-18 修复）
      document.querySelectorAll('#apd22Page .a22lz-r, #apd22Page .a22lz-b, #apd22Page .a22-pen, #apd22Page .a22-chart-btn').forEach(function (e) { e.remove(); });
    },
    /* 编辑态重挂：清把手/画笔 → 清 armed 标记 → 重新 arm */
    rearm: function () {
      this.disarm();
      document.querySelectorAll('#apd22Page [data-lz-armed], #apd22Page [data-pen-armed]').forEach(function (e) {
        delete e.dataset.lzArmed; delete e.dataset.penArmed;
      });
      this.arm();
    },
    styleText: function (t, p) {
      if (p.color) t.style.color = p.color;
      if (p.size) t.style.fontSize = p.size + 'px';
      if (p.dx || p.dy) t.style.transform = 'translate(' + (p.dx || 0) + 'px,' + (p.dy || 0) + 'px)';
      if (p.hidden) { t.style.display = 'none'; return; }
      if (p.text != null && p.text !== '') {
        if (t.dataset.origText == null) t.dataset.origText = t.textContent.replace(/[✎✥✕]/g, '');
        t.textContent = p.text;
      }
    },
    wireW: function (rz, card) {
      var self = this;
      rz.addEventListener('pointerdown', function (e) {
        e.preventDefault();
        var grid = card.parentElement;
        if (!grid || !grid.classList.contains('apd22-grid')) return;   // KPI 条等非栅格卡不拖宽
        var cols = (getComputedStyle(grid).gridTemplateColumns || '').split(' ').filter(Boolean);
        var idx = Array.prototype.indexOf.call(grid.children, card);
        if (idx < 0 || cols.length <= 1) return;
        var startX = e.clientX;
        var gridW = grid.getBoundingClientRect().width - (cols.length - 1) * 10;   // gap 10
        var startFr = cols.map(function (c) { return parseFloat(c) / gridW; });
        try { rz.setPointerCapture(e.pointerId); } catch (err) { /* 合成事件无真实 pointer，忽略 */ }
        rz.classList.add('active');
        document.body.classList.add('a22-lz-resizing');
        var move = function (ev) {
          var dx = ev.clientX - startX;
          var nf = startFr.slice();
          nf[idx] = Math.max(0.15, startFr[idx] + dx / gridW);
          var rest = 1 - nf[idx];
          var restSum = startFr.reduce(function (s, f, i) { return i === idx ? s : s + f; }, 0);
          nf = nf.map(function (f, i) { return i === idx ? f : (restSum ? f / restSum * rest : f); });
          var val = nf.map(function (f) { return f.toFixed(4) + 'fr'; }).join(' ');
          grid.style.gridTemplateColumns = val;
          self.dirty();
        };
        var up = function () {
          rz.classList.remove('active');
          document.body.classList.remove('a22-lz-resizing');
          rz.removeEventListener('pointermove', move); rz.removeEventListener('pointerup', up); rz.removeEventListener('pointercancel', up);
          self.resizeCharts();
        };
        rz.addEventListener('pointermove', move);
        rz.addEventListener('pointerup', up);
        rz.addEventListener('pointercancel', up);
      });
    },
    wireH: function (bz, card) {
      var self = this;
      bz.addEventListener('pointerdown', function (e) {
        e.preventDefault();
        var startY = e.clientY;
        var startH = card.getBoundingClientRect().height;
        try { bz.setPointerCapture(e.pointerId); } catch (err) { /* 合成事件无真实 pointer，忽略 */ }
        bz.classList.add('active');
        document.body.classList.add('a22-lz-resizing');
        var move = function (ev) {
          var h = Math.max(90, Math.round(startH + (ev.clientY - startY)));
          card.style.height = h + 'px';
          card.dataset.lzH = String(h); self.dirty();
        };
        var up = function () {
          bz.classList.remove('active');
          document.body.classList.remove('a22-lz-resizing');
          bz.removeEventListener('pointermove', move); bz.removeEventListener('pointerup', up); bz.removeEventListener('pointercancel', up);
          self.resizeCharts();
        };
        bz.addEventListener('pointermove', move);
        bz.addEventListener('pointerup', up);
        bz.addEventListener('pointercancel', up);
      });
    },
    openPen: function (tKey, target, pen) {
      var self = this;
      document.querySelectorAll('.a22-pen-panel').forEach(function (p) { p.remove(); });
      var cur = self.prefs.texts[tKey] || {};
      var panel = document.createElement('div');
      panel.className = 'a22-pen-panel';
      var rect = pen.getBoundingClientRect();
      panel.style.left = Math.min(window.innerWidth - 240, Math.max(6, rect.left - 100)) + 'px';
      panel.style.top = (rect.bottom + 6) + 'px';
      panel.innerHTML =
        '<div class="pp-row"><span class="pp-lab">文字内容</span><input type="text" class="pp-text" value="' + esc(target.textContent.replace(/[✎✥✕]/g, '')) + '"></div>' +
        '<div class="pp-row"><span class="pp-lab">颜色</span><input type="color" class="pp-color" value="' + (cur.color || '#EAF3FF') + '"></div>' +
        '<div class="pp-row"><span class="pp-lab">字号 px</span><input type="number" class="pp-size" min="9" max="60" value="' + (cur.size || parseInt(getComputedStyle(target).fontSize, 10) || 13) + '"></div>' +
        '<div class="pp-btns"><button class="pp-restore">还原文字</button><button class="pp-apply">应用</button></div>';
      document.body.appendChild(panel);
      var close = function () { panel.remove(); document.removeEventListener('click', onDoc); };
      var onDoc = function (e) { if (!panel.contains(e.target) && e.target !== pen) close(); };
      setTimeout(function () { document.addEventListener('click', onDoc); }, 0);
      panel.querySelector('.pp-apply').addEventListener('click', function () {
        var text = panel.querySelector('.pp-text').value.trim();
        var color = panel.querySelector('.pp-color').value;
        var size = parseInt(panel.querySelector('.pp-size').value, 10) || 0;
        if (target.dataset.origText == null) target.dataset.origText = target.textContent.replace(/[✎✥✕]/g, '');
        target.textContent = text || target.dataset.origText;
        target.appendChild(pen);
        target.style.color = color;
        if (size) target.style.fontSize = size + 'px';
        target.style.display = '';
        self.prefs.texts[tKey] = Object.assign({}, self.prefs.texts[tKey] || {}, { color: color, size: size, text: text, hidden: false });
        self.dirty(); close();
      });
      panel.querySelector('.pp-restore').addEventListener('click', function () {
        // 还原：清掉该文字的全部自定义（内容/颜色/字号/位移/隐藏）
        if (target.dataset.origText != null) target.textContent = target.dataset.origText;
        target.appendChild(pen);
        target.style.color = ''; target.style.fontSize = '';
        target.style.transform = ''; target.style.display = '';
        delete self.prefs.texts[tKey];
        self.dirty(); close();
      });
    },
    /* ---------- 图表配色：色板在全局 pal()（chartKey 或 '*' 生效）；这里只弹面板+落盘+重绘 ---------- */
    openChartPanel: function (chartKey, anchor) {
      var self = this;
      document.querySelectorAll('.a22-pen-panel').forEach(function (p) { p.remove(); });
      var cur = self.prefs.charts[chartKey] || {};
      var panel = document.createElement('div');
      panel.className = 'a22-pen-panel';
      var rect = anchor.getBoundingClientRect();
      panel.style.left = Math.min(window.innerWidth - 250, Math.max(6, rect.left - 110)) + 'px';
      panel.style.top = (rect.bottom + 6) + 'px';
      var opts = Object.keys(self.CHART_PALETTES).map(function (k) {
        return '<option value="' + k + '"' + (cur.preset === k ? ' selected' : '') + '>' + k + '</option>';
      }).join('');
      panel.innerHTML =
        '<div class="pp-row"><span class="pp-lab">色板</span><select class="pp-preset" style="flex:1;background:#0D2547;color:#EAF3FF;border:1px solid #2A4A78;border-radius:6px;padding:3px 6px;font-size:12px">' +
        '<option value="">（跟随看板默认）</option>' + opts + '</select></div>' +
        '<div class="pp-row"><span class="pp-lab">主色</span><input type="color" class="pp-base" value="' + (cur.base || '#2E8FE8') + '"></div>' +
        '<div class="pp-btns"><button class="pp-restore">还原</button><button class="pp-apply">应用</button></div>';
      document.body.appendChild(panel);
      var close = function () { panel.remove(); document.removeEventListener('click', onDoc); };
      var onDoc = function (e) { if (!panel.contains(e.target) && e.target !== anchor) close(); };
      setTimeout(function () { document.addEventListener('click', onDoc); }, 0);
      panel.querySelector('.pp-apply').addEventListener('click', function () {
        var preset = panel.querySelector('.pp-preset').value;
        var base = panel.querySelector('.pp-base').value;
        self.prefs.charts[chartKey] = { preset: preset, base: base };
        self.dirty(); self.applyChartPrefs(); close();
      });
      panel.querySelector('.pp-restore').addEventListener('click', function () {
        delete self.prefs.charts[chartKey];
        self.dirty(); self.applyChartPrefs(); close();
      });
    },
    /* 把用户配色发布到全局（pal() 读 window.A22_CHART_PREFS）并重绘全部图表。
       注意：只重绘图表本体，绝不走 renderAll —— renderAll→rearm→arm→applyChartPrefs→renderAll 会无限递归（2026-10-17 卡死根因） */
    applyChartPrefs: function () {
      window.A22_CHART_PREFS = this.prefs.charts || {};
      var d = state.data;
      if (!d || d.empty) return;
      try { chartDaily15(d); chartCorner(d); chartTizi(d); chartSafetyFinishTree(d); chartFrameSun(d); chartBgPareto(d); } catch (e) { }
    },
    applyAll: function (snap) {
      var P = this.norm(snap || this.prefs);
      var self = this;
      // 卡高
      document.querySelectorAll('#apd22Page .apd22-card, #apd22Page .apd22-kpi').forEach(function (card) {
        var key = self.keyOf(card);
        var pref = key && P.cards[key];
        if (pref && pref.h) card.style.height = pref.h + 'px';
      });
      // 栅格轨道
      document.querySelectorAll('#apd22Page .apd22-grid').forEach(function (grid) {
        var first = grid.children[0];
        if (!first) return;
        var key = self.keyOf(first);
        var pref = key && P.cards[key];
        if (pref && pref.w) grid.style.gridTemplateColumns = pref.w;
      });
      // 文字样式
      Object.keys(P.texts).forEach(function (k) {
        var p = P.texts[k];
        var dot = k.lastIndexOf('.');
        if (dot < 0) return;
        var cardKey = k.slice(0, dot), kind = k.slice(dot + 1);
        var host = document.getElementById(cardKey);
        if (!host) return;
        var card = host.classList.contains('apd22-kpi') || host.classList.contains('apd22-card') ? host : host.closest('.apd22-card, .apd22-kpi');
        if (!card) return;
        var sel = kind === 'title'
          ? (card.classList.contains('apd22-kpi') ? '.apd22-kpi-title' : '.apd22-card-head b')
          : (card.classList.contains('apd22-kpi') ? '.apd22-kpi-sub' : '.apd22-card-sub');
        var t = card.querySelector(sel);
        if (!t) return;
        if (p.color) t.style.color = p.color;
        if (p.size) t.style.fontSize = p.size + 'px';
        if (p.dx || p.dy) t.style.transform = 'translate(' + (p.dx || 0) + 'px,' + (p.dy || 0) + 'px)';
        if (p.hidden) { t.style.display = 'none'; return; }
        if (p.text != null && p.text !== '') {
          if (t.dataset.origText == null) t.dataset.origText = t.textContent.replace(/[✎✥✕]/g, '');
          t.textContent = p.text;
        }
      });
      this.resizeCharts();
    },
    resizeCharts: function () {
      setTimeout(function () { Object.keys(state.charts).forEach(function (k) { try { state.charts[k].resize(); } catch (e) { } }); }, 60);
    },
    /* ---------- v3（2026-10-18）：草稿态 + 保存 / 撤销 / 恢复 / 一键还原 ---------- */
    /* 清零：内联样式全部清回 CSS 默认（卡高/栅格/文字还原原文），快照回放前必须先清零 */
    resetDom: function () {
      document.querySelectorAll('#apd22Page .apd22-card, #apd22Page .apd22-kpi').forEach(function (card) {
        card.style.height = '';
        delete card.dataset.lzH;
      });
      document.querySelectorAll('#apd22Page .apd22-grid').forEach(function (grid) {
        grid.style.gridTemplateColumns = '';
      });
      document.querySelectorAll('#apd22Page .apd22-card-head b, #apd22Page .apd22-card-head .apd22-card-sub, #apd22Page .apd22-kpi-title, #apd22Page .apd22-kpi-sub').forEach(function (t) {
        t.querySelectorAll('.a22-pen').forEach(function (b) { b.remove(); });
        if (t.dataset.origText != null) { t.textContent = t.dataset.origText; }
        else {
          var tx = t.textContent.replace(/[✎✥✕]/g, '');
          if (tx !== t.textContent) t.textContent = tx;
        }
        t.style.color = ''; t.style.fontSize = ''; t.style.transform = ''; t.style.display = '';
      });
    },
    /* 快照回放：设为当前草稿 → 清零 → 重放 → 重挂把手（编辑态）→ 按新配色重绘图表 */
    applyPrefs: function (snap) {
      var P = this.norm(snap);
      this.prefs = P;
      window.A22_CHART_PREFS = P.charts;
      this.resetDom();
      this.applyAll(P);
      this.applyChartPrefs();
      if (this.on) this.rearm();
    },
    /* 收集 DOM 当前态（栅格宽 + 卡高）进 prefs.cards（文字/配色是即时写入 prefs 的，无需收集） */
    collect: function () {
      var self = this;
      var gridRecs = {};
      document.querySelectorAll('#apd22Page .apd22-grid').forEach(function (grid) {
        var first = grid.children[0];
        if (!first) return;
        if (grid.style.gridTemplateColumns) {
          var gk = self.keyOf(first);
          if (gk) gridRecs[gk] = { w: grid.style.gridTemplateColumns };
        }
      });
      Object.keys(gridRecs).forEach(function (k) {
        self.prefs.cards[k] = Object.assign({}, self.prefs.cards[k] || {}, gridRecs[k]);
      });
      document.querySelectorAll('#apd22Page .apd22-card, #apd22Page .apd22-kpi').forEach(function (card) {
        var key = self.keyOf(card);
        if (key && card.dataset.lzH) {
          self.prefs.cards[key] = Object.assign({}, self.prefs.cards[key] || {}, { h: parseInt(card.dataset.lzH, 10) });
        }
      });
    },
    /* 标记未保存：只亮红点 + 启用保存钮，绝不自动落盘（v3 起只有 save() 才 POST） */
    dirty: function () {
      this.collect();
      this.unsaved = true;
      this.updateButtons();
    },
    persist: function (snap) {
      var P = this.norm(snap);
      fetch('/api/board/layout', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ns: this.NS, cards: P.cards, texts: P.texts, charts: P.charts }) }).catch(function () { });
    },
    toast: function (msg) { if (typeof showToast === 'function') { showToast(msg); } else { console.log('[布局] ' + msg); } },
    /* 保存：快照压栈（作废后面的恢复步）+ POST，此刻修改才「生效」 */
    save: function () {
      if (!this.unsaved) { this.toast('没有未保存的修改'); return; }
      this.collect();
      var snap = this.clone(this.prefs);
      this.history = this.history.slice(0, this.histIdx + 1);
      this.history.push(snap);
      this.histIdx = this.history.length - 1;
      this.unsaved = false;
      this.persist(snap);
      this.updateButtons();
      this.toast('布局已保存并生效（Ctrl+Z 可撤销这一步）');
    },
    /* 撤销：有草稿先丢草稿；否则回退一步并持久化 */
    undo: function () {
      if (this.unsaved) {
        this.applyPrefs(this.history[this.histIdx]);
        this.unsaved = false;
        this.updateButtons();
        this.toast('已放弃未保存的修改');
        return;
      }
      if (this.histIdx <= 0) { this.toast('没有可撤销的步骤了'); return; }
      this.histIdx--;
      var snap = this.history[this.histIdx];
      this.applyPrefs(snap);
      this.persist(snap);
      this.updateButtons();
      this.toast('已撤销一步（Ctrl+Y 可恢复）');
    },
    /* 恢复（重做）：把被撤销的一步找回来并持久化 */
    redo: function () {
      if (this.unsaved) { this.toast('有未保存的修改，请先保存（Ctrl+S）或撤销'); return; }
      if (this.histIdx >= this.history.length - 1) { this.toast('没有可恢复的步骤了'); return; }
      this.histIdx++;
      var snap = this.history[this.histIdx];
      this.applyPrefs(snap);
      this.persist(snap);
      this.updateButtons();
      this.toast('已恢复一步');
    },
    /* 一键还原默认：清空全部自定义；还原前屏幕态也入栈，Ctrl+Z 可撤销这次还原 */
    restoreDefault: function () {
      if (!confirm('一键还原默认布局？\n将清空全部自定义（卡宽/卡高/文字/配色），此步之后仍可用 Ctrl+Z 撤销。')) return;
      this.collect();
      var before = this.clone(this.prefs);
      var empty = { cards: {}, texts: {}, charts: {} };
      this.history = this.history.slice(0, this.histIdx + 1);
      this.history.push(before);
      this.history.push(empty);
      this.histIdx = this.history.length - 1;
      this.applyPrefs(empty);
      this.persist(empty);
      this.unsaved = false;
      this.updateButtons();
      this.toast('已还原默认布局（Ctrl+Z 可撤销）');
    },
    updateButtons: function () {
      var sv = el('a22LayoutSave'), un = el('a22LayoutUndo'), rd = el('a22LayoutRedo');
      if (sv) { sv.disabled = !this.unsaved; sv.classList.toggle('a22-unsaved', !!this.unsaved); }
      if (un) un.disabled = !this.unsaved && this.histIdx <= 0;
      if (rd) rd.disabled = this.unsaved || this.histIdx >= this.history.length - 1;
      // 面板收起时保存钮不可见 → 把「未保存」红点挪到折叠钮上，避免改动被遗忘
      var fb = el('a22LayoutFold');
      if (fb) fb.classList.toggle('a22-unsaved', !!this.unsaved && !this.popOpen);
    },
  };
  // 数据刷新会重画 KPI/透视卡 → 重挂把手/画笔（编辑态）+ 重放文字/卡高样式（恒定，编辑态与否都要）
  window.Apd22LayoutRearm = function () {
    A22Layout.resetDom();      // 先清零（去旧把手/画笔、文字还原原文）再重放，杜绝样式/按钮累积
    A22Layout.applyAll();
    if (!A22Layout.on) return;
    A22Layout.rearm();
  };

  /* ---------- 打开页面时惰性初始化 ---------- */
  var booted = false;
  function boot() {
    if (booted) return;
    booted = true;
    showLoading(true);   // v5：一进页面就给加载反馈，不等数据回来
    init();
    /* 偏好恢复需等：①网络偏好加载完成 ②首轮 renderAll 产出卡片 —— 两者齐了才 applyAll */
    (function waitPrefs() {
      if (A22Layout.prefsLoaded && document.querySelector('#apd22Page .apd22-card, #apd22Page .apd22-kpi')) A22Layout.applyAll();
      else setTimeout(waitPrefs, 250);
    })();
    setTimeout(function () { Object.keys(state.charts).forEach(function (k) { state.charts[k].resize(); }); }, 350);
  }
  window.Apd22Board = { boot: boot, reload: load };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', maybeAutoBoot);
  else setTimeout(maybeAutoBoot, 400);
  function maybeAutoBoot() {
    var pg = el('apd22Page');
    if (pg && pg.offsetParent !== null) boot();
  }
})();

/* 版本自检：旧 HTML 缓存强制刷新一次（会话内仅一次）。
   自适应版：与本脚本自身的 src 版本比对，不再硬编码版本号（2026-10-18 v5 起） */
(function () {
  var me = (document.currentScript && document.currentScript.getAttribute('src')) || '';
  var myVer = (me.match(/[?&]v=([\w.]+)/) || [])[1];
  if (!myVer) return;
  var fresh = false;
  document.querySelectorAll('script[src*=apd22]').forEach(function (t) {
    if ((t.getAttribute('src') || '').indexOf('v=' + myVer) >= 0) fresh = true;
  });
  if (!fresh && !sessionStorage.getItem('apd22_v_reloaded_' + myVer)) {
    sessionStorage.setItem('apd22_v_reloaded_' + myVer, '1');
    location.reload();
  }
})();
