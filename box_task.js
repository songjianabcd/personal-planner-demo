/* 维护物料模块自动化（BOX TASK）· 页面脚本
   2026-10-24 集成 box_level_task 脚本（第二轮：抬头一行化 + 通知对象多选框 + 历史账号下拉 +
   步骤灯 + 结果精简 + 表格右键复制 + 日志可全选复制）。
   只与本机 /api/box/* 打交道；不引用其它模块的内部状态。 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var API = {
    config: '/api/box/config',
    options: '/api/box/options',
    run: '/api/box/run',
    status: '/api/box/status',
    stop: '/api/box/stop',
    state: '/api/box/state',
    log: '/api/box/log',
    latest: '/api/box/latest',
    notify: '/api/box/notify'
  };
  var 通道 = ['钉钉消息', '群消息', '邮箱'];
  // 步骤灯：名 + 子步骤数（与脚本内部动作一一对应，用于徽标显示）
  var 阶段 = [
    { 名: '连接浏览器并登录', 步: 3, 子步: ['按 CDP(9222) 连接 / 拉起 Chrome', '打开 ECS「物料维护流程」任务页', '未直达任务页则走登录流程（填域账号密码）'] },
    { 名: '任务进程（读取 / 回写）', 步: 5, 子步: ['展开「任务进程」面板', '读当前任务状态', '点「读取任务」并等状态回到完成', '点「回写任务」并等状态回到完成', '记录「任务流程结果」'] },
    { 名: '清空操作人员筛选', 步: 2, 子步: ['定位「操作人员」输入框', '清空 Vue 绑定值并回读校验为空'] },
    { 名: '查询 + 切换 600 条/页', 步: 2, 子步: ['点「查询」', '切换每页 600 条并等表格刷新'] },
    { 名: '提取并筛选待维护物料', 步: 3, 子步: ['逐页提取（校分页总数=实际提取）', '按「上级物料号」X-X-X / X-X-X-X 正则筛选', '输出「页面提取 N 条 / 筛选后 M 条」两条计数'] },
    { 名: '生成 Excel', 步: 1, 子步: ['openpyxl 写 xlsx（4 行表头 + 数据行）到 output 目录'] }
  ];
  // 每阶段判据（关键字全部取自脚本真实输出，可在运行日志里一一对上）
  var 判据 = [
    { 成: ['已连接浏览器标签页', '登录表单填充: filled'], 败: ['[FATAL] 登录失败'] },
    { 成: ['任务流程结果:'], 败: ['读取出错', '回写出错'] },
    { 成: ['操作人员已为空', 'Vue值清空成功'], 败: [] },
    { 成: ['已切换为600条/页'], 败: [] },
    { 成: ['筛选后 '], 败: ['提取失败'] },
    { 成: ['Excel 已生成'], 败: [] }
  ];
  var 开始标记 = '开始执行：箱子及装箱级物料未维护';

  var state = {
    booted: false, cfg: null, opts: null, 对象: [], runId: '', cursor: 0,
    日志: [], 行数: 5, 展开更多: false, 轮询句柄: null, 存盘句柄: null, 运行中: false,
    步骤折叠: true, 表格折叠: false, 结果数据: null,    // 用户定规：步骤默认折叠、表格默认展开
    时间线: null, 片: []
  };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function toast(msg) { if (typeof window.showToast === 'function') window.showToast(msg); }
  function tip(文本, ok) {
    var el = $('boxSaveTip'); if (!el) return;
    el.textContent = 文本; el.className = ok ? 'box-save-tip ok' : 'box-save-tip';
    clearTimeout(tip.句柄);
    tip.句柄 = setTimeout(function () { el.textContent = '配置自动保存'; el.className = 'box-save-tip'; }, 2200);
  }
  function 取值(v) { return (v === null || v === undefined || v === '') ? '—' : v; }

  // 判「事件是否发生在某容器内」：菜单项点击会用 innerHTML 重建列表，被点节点中途脱离 DOM，
  // 此时 e.target.closest() 会失效（自验抓到：菜单自己把菜单关了）→ 用 composedPath（派发时快照）
  function 在里(e, sel) {
    var path = (e.composedPath && e.composedPath()) || [];
    for (var i = 0; i < path.length; i++) {
      var n = path[i];
      if (n && n.nodeType === 1 && n.matches && n.matches(sel)) return true;
    }
    return !!(e.target && e.target.closest && e.target.closest(sel));
  }

  function 复制文本(文本, 提示) {
    var done = function () { toast(提示 || '已复制到剪贴板'); };
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(文本).then(done, function () { 回退(); });
        return;
      }
    } catch (e) { /* 落到回退 */ }
    回退();
    function 回退() {
      try {
        var ta = document.createElement('textarea');
        ta.value = 文本; ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.select();
        document.execCommand('copy'); document.body.removeChild(ta); done();
      } catch (e) { toast('复制失败，请手动 Ctrl+C'); }
    }
  }

  // ═══════════ 通知对象（一个多选框：勾选 + 自定义） ═══════════
  function 从配置读对象(cfg) {
    var out = [], 通知 = (cfg || {})['通知'] || {};
    通道.forEach(function (t) {
      var arr = (通知[t] || {})['对象'] || [];
      arr.forEach(function (v) {
        v = String(v || '').trim(); if (!v) return;
        out.push({ 值: v, 键: v, 类型: t });
      });
    });
    return out;
  }
  // 类型判定（用户 2026-10-24 定规）：**不猜名字是群还是人**。优先级：
  //   ① 自带后缀 （群消息）/（钉钉消息）/（邮箱）  ② 候选表里已有的对象  ③ 裸邮箱地址（含 @ 且没有后缀）
  //   ④ 都认不出 → null，前端提示用户补后缀
  var 后缀正则 = /[（(]\s*(群消息|钉钉消息|邮箱)\s*[)）]\s*$/;
  function 解析类型(v) {
    var s = String(v || '').trim();
    var m = s.match(后缀正则);
    if (m) return m[1];
    var 候 = 找候选(s);
    if (候 && 候.类型) return 候.类型;
    if (s.indexOf('@') >= 0) return '邮箱';
    return null;
  }
  // 统一成「名字（通道）」全角后缀写法；裸邮箱地址原样保留
  function 规范化(v) {
    var s = String(v || '').trim();
    var t = 解析类型(s);
    if (!t) return s;
    var 去 = s.replace(后缀正则, '').trim();
    if (!去) return s;
    if (t === '邮箱' && !后缀正则.test(s)) return 去;
    return 去 + '（' + t + '）';
  }
  function 找候选(v) {
    var all = (state.opts && state.opts.全部) || [];
    for (var i = 0; i < all.length; i++) if (all[i].值 === v) return all[i];
    return null;
  }
  function renderTags() {
    var host = $('boxNselTags'); if (!host) return;
    host.innerHTML = state.对象.map(function (o, i) {
      return '<span class="box-tag t-' + esc(o.类型) + '" title="' + esc(o.类型 + '：' + o.值) + '">' +
        esc(o.值) + '<span class="x" data-i="' + i + '">×</span></span>';
    }).join('');
  }
  function 加对象(v) {
    var s = String(v || '').trim();
    if (!s) return false;
    var 类型 = 解析类型(s);
    if (!类型) {
      toast('认不出是群还是人：请自己带后缀——计划科（群消息）／王锦志（钉钉消息）／a@b.com（邮箱）');
      开菜单(true);
      return false;
    }
    var 值 = 规范化(s);
    for (var i = 0; i < state.对象.length; i++) if (state.对象[i].值 === 值) return false;
    var 候 = 找候选(值);
    state.对象.push(候 ? { 值: 值, 键: 候.值, 类型: 类型, 群名: 候.群名, 会话ID: 候.会话ID, 钉钉ID: 候.钉钉ID, 姓名: 候.姓名 }
                     : { 值: 值, 键: 值, 类型: 类型 });
    renderTags(); 存盘(false);
    return true;
  }
  function 删对象(v) { state.对象 = state.对象.filter(function (o) { return o.值 !== v; }); renderTags(); 存盘(false); }
  function 有对象(v) { return state.对象.some(function (o) { return o.值 === v; }); }

  function renderNselMenu(过滤) {
    var host = $('boxNselMenu'); if (!host) return;
    var opts = state.opts || {}, 词 = String(过滤 || '').trim().toLowerCase();
    var html = '';
    if (词) {
      var 词类型 = 解析类型(词);
      html += '<div class="box-nsel-grp">自定义</div>' +
        '<div class="box-nsel-opt" data-custom="1"><span class="cb">' + (词类型 ? '+' : '!') + '</span>' +
        '<span class="nm">' + (词类型 ? '加入自定义：' + esc(规范化(词)) : '要加后缀才能加：' + esc(词) + '（群消息／钉钉消息／邮箱）') +
        '</span><span class="use">' + esc(词类型 || '待定') + '</span></div>';
    }
    通道.forEach(function (t) {
      var arr = (opts[t] || []).filter(function (it) {
        return !词 || String(it.值).toLowerCase().indexOf(词) >= 0 || String(it.姓名 || '').toLowerCase().indexOf(词) >= 0;
      }).slice(0, 词 ? 40 : 18);
      if (!arr.length) return;
      html += '<div class="box-nsel-grp">' + esc(t) + '（按常用度）</div>';
      arr.forEach(function (it) {
        var on = 有对象(it.值);
        html += '<div class="box-nsel-opt' + (on ? ' on' : '') + '" data-val="' + esc(it.值) + '">' +
          '<span class="cb">' + (on ? '✓' : '') + '</span><span class="nm">' + esc(it.值) + '</span>' +
          '<span class="use">' + (it.用量 >= 100000000 ? '常用' : (it.用量 ? esc(it.用量) : '')) + '</span></div>';
      });
    });
    if (!html) html = '<div class="box-nsel-empty">没有匹配的候选；直接输入后回车可加为自定义对象。</div>';
    host.innerHTML = html;
  }
  // 下拉菜单都是 position:fixed（避开 .box-bar 的 overflow 裁剪），打开时手动贴到锚点下方
  function 定位菜单(m, 锚) {
    if (!m || !锚) return;
    var r = 锚.getBoundingClientRect();
    m.style.display = 'block';
    var w = m.offsetWidth, h = m.offsetHeight;
    m.style.left = Math.max(4, Math.min(r.left, window.innerWidth - w - 8)) + 'px';
    m.style.top = Math.max(4, Math.min(r.bottom + 4, window.innerHeight - h - 8)) + 'px';
  }
  function 开菜单(开) {
    var m = $('boxNselMenu'); if (!m) return;
    if (开) { renderNselMenu($('boxNselInp').value); 定位菜单(m, $('boxNsel')); }
    else m.style.display = 'none';
  }
  function 关菜单() { 开菜单(false); $('boxUserMenu').style.display = 'none'; $('boxPassMenu').style.display = 'none'; }

  // ═══════════ 历史账号下拉 ═══════════
  function 历史账号() { return ((state.cfg || {})['历史账号'] || []).filter(function (x) { return x && x.账号; }); }
  function renderHistMenu(which) {
    var m = $('boxUserMenu'), mp = $('boxPassMenu');
    var host = which === 'pass' ? mp : m;
    var arr = 历史账号();
    if (!arr.length) { host.innerHTML = '<div class="box-nsel-empty">还没有历史账号（运行一次后自动记住）</div>'; return; }
    host.innerHTML = arr.map(function (it, i) {
      return '<div class="box-combo-item" data-i="' + i + '">' + esc(it.账号) +
        '<span class="cm-sub">' + esc(it.域 || '') + '｜' + esc(it.最近 || '') + '</span></div>';
    }).join('');
  }
  function 选历史(i) {
    var it = 历史账号()[i]; if (!it) return;
    $('boxUser').value = it.账号 || '';
    $('boxPass').value = it.密码 || '';
    if (it.域) { state.cfg['登录'] = state.cfg['登录'] || {}; state.cfg['登录']['域'] = it.域; }
    $('boxUserMenu').style.display = 'none'; $('boxPassMenu').style.display = 'none';
    renderNselMenu($('boxNselInp').value); 存盘(false);
    toast('已套用历史账号：' + (it.账号 || ''));
  }

  // ═══════════ 步骤灯 ═══════════
  function 阶段态(行, 运行中) {
    var at = -1, i;
    for (i = 行.length - 1; i >= 0; i--) if (行[i].indexOf(开始标记) >= 0) { at = i; break; }
    var scope = at >= 0 ? 行.slice(at) : 行;
    var 态 = 阶段.map(function () { return 'pending'; });
    scope.forEach(function (t) {
      判据.forEach(function (判, k) {
        if (判.败.some(function (s) { return t.indexOf(s) >= 0; })) { 态[k] = 'failed'; return; }
        if (态[k] !== 'failed' && 判.成.some(function (s) { return t.indexOf(s) >= 0; })) 态[k] = 'success';
      });
    });
    if (运行中) {
      for (i = 0; i < 态.length; i++) if (态[i] === 'pending') { 态[i] = 'running'; break; }
    }
    return 态;
  }
  // 步数徽标悬浮气泡：列出该阶段的具体子步骤（用户要求「5步要悬浮展示具体步骤」）
  function 关步提示() { var t = $('boxStepsTip'); if (t && t.parentNode) t.parentNode.removeChild(t); }
  function 显步提示(徽标) {
    var st = 阶段[+徽标.getAttribute('data-bs')]; if (!st || !st.子步) return;
    关步提示();
    var t = document.createElement('div');
    t.className = 'box-steps-tip'; t.id = 'boxStepsTip';
    t.innerHTML = '<span class="t">' + esc(st.名) + '（' + st.步 + ' 步）</span>' +
      st.子步.map(function (s, k) { return '<span class="s"><i>' + (k + 1) + '</i>' + esc(s) + '</span>'; }).join('');
    document.body.appendChild(t);
    var r = 徽标.getBoundingClientRect();
    t.style.left = Math.max(8, Math.min(r.left, window.innerWidth - t.offsetWidth - 8)) + 'px';
    t.style.top = Math.max(8, Math.min(r.bottom + 8, window.innerHeight - t.offsetHeight - 8)) + 'px';
  }
  // ── 每步耗时 + 总运行时间：直接按运行日志的时间戳算（运行中也能实时出，不必等后端）──
  var 行前缀 = /^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]/;
  function 行时刻(行) {
    var m = String(行 || '').match(行前缀);
    if (!m) return null;
    var t = new Date(m[1].replace(/-/g, '/')).getTime();
    return isNaN(t) ? null : t;
  }
  function 时长文本(秒) {
    var s = Math.max(0, Math.round(秒));
    return s < 60 ? (s + ' 秒') : (Math.floor(s / 60) + ' 分 ' + (s % 60) + ' 秒');
  }
  function 算时间线(行) {
    var i, 起 = -1;
    for (i = 行.length - 1; i >= 0; i--) { if (String(行[i]).indexOf(开始标记) >= 0) { 起 = i; break; } }
    if (起 < 0) return null;
    var 开始 = 行时刻(行[起]), 到达 = [], 末 = null, ptr = 0, 本轮 = [];
    for (i = 起; i < 行.length; i++) {          // [通知] 那几行是进程退出后发的，不算脚本运行时间
      if (String(行[i]).indexOf('[通知]') >= 0) break;
      本轮.push(行[i]);
    }
    for (i = 0; i < 本轮.length; i++) {
      var t = 行时刻(本轮[i]); if (t) 末 = t;
      for (var s = ptr; s < 判据.length; s++) {
        if (判据[s].成.some(function (m) { return String(本轮[i]).indexOf(m) >= 0; })) {
          if (!到达[s]) 到达[s] = t;
          ptr = s; break;
        }
      }
    }
    var 阶段时间 = [], 上一 = 开始;
    for (i = 0; i < 判据.length; i++) {
      if (!到达[i]) { 阶段时间.push(null); continue; }
      var 秒 = (到达[i] - 上一) / 1000;
      阶段时间.push({ 秒: Math.round(秒), 文本: 时长文本(秒) });
      上一 = 到达[i];
    }
    return { 阶段: 阶段时间, 总秒: 末 ? Math.round((末 - 开始) / 1000) : null,
             总文本: 末 ? 时长文本((末 - 开始) / 1000) : '' };
  }
  function renderNodes(行, 运行中) {
    var host = $('boxNodes'); if (!host) return;
    var 态 = 阶段态(行 || [], !!运行中);
    var 时间线 = 算时间线(行 || []);
    state.时间线 = 时间线;
    host.innerHTML = 阶段.map(function (st, i) {
      var s = 态[i], 文案 = { pending: '未执行', running: '运行中', success: '成功', failed: '失败' }[s] || s;
      var badge = s === 'success' ? '✓' : s === 'failed' ? '✕' : String(i + 1).padStart(2, '0');
      return '<div class="auto-node ' + s + '" data-ni="' + i + '">' +
        '<div class="an-core">' + badge + '</div><span class="flow-dot"></span>' +
        '<div class="an-body"><div class="an-name">' + esc(st.名) +
        '<span class="an-steps-badge" data-bs="' + i + '">' + st.步 + ' 步</span>' +
        '<span class="an-status">' + esc(文案) + '</span>' +
        ((时间线 && 时间线.阶段[i] && s !== 'pending') ? '<span class="an-time" title="本步耗时 ' + esc(时间线.阶段[i].文本) + '">' + esc(时间线.阶段[i].文本) + '</span>' : '') +
        '</div></div></div>';
    }).join('');
    var 成 = 态.filter(function (s) { return s === 'success'; }).length;
    var 败 = 态.filter(function (s) { return s === 'failed'; }).length;
    var 跑 = 态.filter(function (s) { return s === 'running'; }).length;
    var meta = $('boxStepsMeta');
    if (meta) {
      meta.innerHTML = '共 ' + 阶段.length + ' 步' +
        (跑 ? '　·　<span style="color:var(--color-primary)">运行中 ' + 跑 + '</span>' : '') +
        (败 ? '　·　<span style="color:#B02A37">失败 ' + 败 + '</span>' : '') +
        (成 ? '　·　<span style="color:#0F7A52">成功 ' + 成 + '</span>' : '') +
        ((时间线 && 时间线.总文本) ? '　·　总运行时间 <b class="box-total-time">' + esc(时间线.总文本) + '</b>' : '');
    }
  }

  // ═══════════ 受限影响面（受影响数据 chips） ═══════════
  function renderImpact() {
    var host = $('boxImpact'); if (!host) return;
    var 片 = [{ t: '物料维护流程页（ECS）', k: '源' }];
    if (通知总开关() && state.对象.length) {
      var 按 = {};
      state.对象.forEach(function (o) { (按[o.类型] = 按[o.类型] || []).push(o.值); });
      通道.forEach(function (t) { if (按[t]) 片.push({ t: t + '：' + 按[t].join('、'), k: '通知' }); });
    } else {
      片.push({ t: '结果 Excel（本机 output 目录）', k: '产出' });
    }
    host.style.display = 'block';
    host.innerHTML = '<span style="color:var(--color-text-tertiary)">受影响数据：</span>' + 片.map(function (x) {
      return '<span class="impact-link" title="' + esc(x.k + '：' + x.t) + '">📊 ' + esc(x.t) + '</span>';
    }).join(' ');
  }

  // ═══════════ 抬头：读 / 写 / 存盘 ═══════════
  // 通知总开关（用户 2026-10-24：三个通道开关合成一个；走哪条通道由「通知对象」各自的后缀决定）
  function 通知总开关() { var el = $('boxSwNotify'); return !!(el && el.getAttribute('aria-checked') === 'true'); }
  function 设总开关(on, 静默) {
    var el = $('boxSwNotify'); if (!el) return;
    el.setAttribute('aria-checked', on ? 'true' : 'false');
    var sw = el.querySelector('.trig-switch'); if (sw) sw.className = on ? 'trig-switch on' : 'trig-switch';
    if (!静默) { renderImpact(); 存盘(false); }
  }
  function fillBar() {
    var c = state.cfg || {}, l = c['登录'] || {};
    $('boxUser').value = l['账号'] || '';
    $('boxPass').value = l['密码'] || '';
    state.对象 = 从配置读对象(c);
    renderTags(); renderNselMenu('');
    设总开关(通道.some(function (t) { return !!(((c['通知'] || {})[t] || {})['启用']); }), true);
    renderImpact();
  }
  function 通知patch() {
    var out = { '通知': {} };
    通道.forEach(function (t) {
      out['通知'][t] = {
        '启用': 通知总开关(),
        '对象': state.对象.filter(function (o) { return o.类型 === t; }).map(function (o) { return o.键 || o.值; })
      };
    });
    return out;
  }
  function 存盘(立即) {
    clearTimeout(state.存盘句柄);
    state.存盘句柄 = setTimeout(function () {
      var patch = {
        '登录': { '账号': $('boxUser').value.trim(), '密码': $('boxPass').value },
        '通知': 通知patch()['通知']
      };
      var 群 = state.对象.filter(function (o) { return o.类型 === '群消息'; });
      var 旧群 = (state.cfg || {})['群消息'] || {};
      patch['群消息'] = {
        '启用': 通知总开关() && 群.length > 0,
        '群名': (群[0] && 群[0].群名) || 旧群['群名'] || '',
        '会话ID': (群[0] && 群[0].会话ID) || 旧群['会话ID'] || ''
      };
      fetch(API.config, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(patch) })
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (d && d.config) { state.cfg = d.config; renderImpact(); }
          tip('已保存 ' + new Date().toLocaleTimeString('zh-CN', { hour12: false }), true);
        })
        .catch(function () { tip('保存失败（看控制台）', false); });
    }, 立即 ? 0 : 600);
  }

  // ═══════════ 结果区 ═══════════
  function renderResult(d) {
    var host = $('boxResult'); if (!host) return;
    if (d) state.结果数据 = d;
    d = state.结果数据;                                  // 折叠切换时复用同一份数据
    var 表开 = !state.表格折叠;
    var 摘 = (d && d.summary) || {}, ex = d && d.excel, rows = (d && d.rows) || [], headers = (d && d.headers) || [];
    if (!ex && 摘['页面提取'] == null && 摘['筛选后'] == null) {
      host.innerHTML = '<div class="box-empty">还没有结果。点右上角「运行」跑一次；结束后这里显示「页面提取 N 条｜筛选后 M 条」和 Excel 超链接（悬浮看文件名、双击打开），下方表格可框选、可右键复制。</div>';
      return;
    }
    var html = '<div class="box-resline">' +
      '<span>页面提取 <b>' + esc(取值(摘['页面提取'])) + '</b> 条</span>' +
      '<span>筛选后 <b>' + esc(取值(摘['筛选后'])) + '</b> 条</span>' +
      (ex ? '<a class="box-xlink" href="' + esc(ex.url) + '" target="_blank" rel="noopener" title="' +
            esc(ex.name + '｜Excel 工作簿（.xlsx，' + (Math.round((ex.size || 0) / 102.4) / 10) + ' KB）｜双击打开 / 单击下载') +
            '">📊 ' + esc(ex.name) + '</a>'
          : '<span class="box-meta">本次没有生成 Excel（无待维护数据，或运行未完成）</span>') +
      通知chips(d && d.notify) +
      (rows.length ? '<button class="box-sec-toggle" id="boxTableToggle" type="button" title="折叠 / 展开下面的结果表格（默认展开）">' +
        (表开 ? '折叠 ▴' : '展开 ▾') + '</button>' : '') +
      '</div>' + 通知明细(d && d.notify);
    if (rows.length && !表开) {
      html += '<div class="box-resfoot">结果表格已折叠（' + rows.length + ' 行，点上方「展开」看明细）</div>';
    } else if (rows.length) {
      html += '<div class="box-table-wrap" id="boxTableWrap"><table class="box-table" id="boxTable"><thead><tr>' +
        headers.map(function (h) { return '<th>' + esc(h) + '</th>'; }).join('') + '</tr></thead><tbody>' +
        rows.map(function (r) {
          return '<tr>' + headers.map(function (h) { return '<td>' + esc(r[h]) + '</td>'; }).join('') + '</tr>';
        }).join('') + '</tbody></table></div>' +
        '<div class="box-resfoot">共 ' + rows.length + ' 行' + (d.truncated ? '（只显示前 ' + rows.length + ' 行，完整见 Excel）' : '') +
        '　·　表格可拖选 / 右键复制（复制表格、复制选中）</div>';
    }
    host.innerHTML = html;
    bindTableMenu();
  }
  // 通知结果回执（用户要求：发没发、发到哪、为什么没发，都必须在这块给出来）
  function 通知chips(通知) {
    if (!通知 || (!((通知.结果 || []).length) && !通知.跳过)) return '';
    var 片 = [];
    if ((通知.结果 || []).length) {
      (通知.结果 || []).forEach(function (r) {
        var 好 = r.状态 === '已发送', 中 = r.状态 === '部分成功';
        var 落盘 = r.状态 === '已落盘';
        var 号 = 好 ? '✓' : (中 || 落盘 ? '!' : '✗');
        片.push('<span class="box-nf ' + (好 ? 'ok' : (中 || 落盘 ? 'warn' : 'bad')) +
          '" title="' + esc(通知.时间 + '｜' + r.通道 + '｜对象：' + r.对象 + '｜说明：' + (r.说明 || '')) + '">' +
          esc(r.通道) + ' ' + 号 + ' ' + esc(r.状态) + (落盘 ? '（未发出）' : '') + '</span>');
      });
    } else {
      片.push('<span class="box-nf skip" title="' + esc((通知.时间 || '') + '｜' + (通知.跳过 || '')) + '">通知未发送</span>');
    }
    return '<span class="box-nfwrap" title="最近一次通知 ' + esc(通知.时间 || '') + '">' + 片.join('') + '</span>';
  }
  function 通知明细(通知) {
    if (!通知 || (!((通知.结果 || []).length) && !通知.跳过)) return '';
    var 细 = [];
    (通知.结果 || []).forEach(function (r) {
      if (r.状态 !== '已发送') 细.push('<b>' + esc(r.通道) + '</b> ' + esc(r.状态) + '：' + esc(r.说明 || '（无说明）') + '　→　' + esc(r.对象));
    });
    if (通知.跳过) 细.push(esc(通知.跳过));
    if ((通知.结果 || []).some(function (r) { return r.状态 === '已落盘'; })) {
      细.push('<span class="box-meta">「已落盘」＝ dws 没把邮件发出去，只落成一份本地文件（正文/附件都在，授权恢复后可重发）</span>');
    }
    return '<div class="box-notify">通知结果（' + esc(通知.时间 || '—') + '）：' +
      (细.length ? 细.join('<br>') : '钉钉消息 / 群消息 / 邮箱 全部已发送') + '</div>';
  }
  function 表格TSV() {
    var t = $('boxTable'); if (!t) return '';
    var out = [];
    t.querySelectorAll('tr').forEach(function (tr) {
      var cells = [];
      tr.querySelectorAll('th,td').forEach(function (c) { cells.push(c.textContent.replace(/[\t\n]/g, ' ')); });
      out.push(cells.join('\t'));
    });
    return out.join('\n');
  }
  function bindTableMenu() {
    var wrap = $('boxTableWrap'); if (!wrap) return;
    wrap.addEventListener('contextmenu', function (e) {
      e.preventDefault();
      关表格菜单();
      var m = document.createElement('div');
      m.className = 'box-tbl-menu'; m.id = 'boxTblMenu';
      m.innerHTML = '<div class="ctx-item" data-act="table"><span>复制表格</span><span>全表 TSV</span></div>' +
        '<div class="ctx-item" data-act="sel"><span>复制选中内容</span><span>Ctrl+C</span></div>' +
        '<div class="ctx-tip">拖选单元格后 Ctrl+C 也可直接复制</div>';
      document.body.appendChild(m);
      m.style.left = Math.min(e.clientX, window.innerWidth - 210) + 'px';
      m.style.top = Math.min(e.clientY, window.innerHeight - 130) + 'px';
      m.addEventListener('click', function (ev) {
        var it = ev.target.closest('.ctx-item'); if (!it) return;
        var act = it.getAttribute('data-act');
        关表格菜单();
        if (act === 'table') 复制文本(表格TSV(), '已复制表格（' + 表格TSV().split('\n').length + ' 行，可直接粘到 Excel）');
        else {
          var sel = window.getSelection();
          复制文本(sel ? String(sel) : '', '已复制选中内容');
        }
      });
      // 关闭只认「点在菜单外面」：早期写法在 mousedown 上无条件关菜单，会把菜单先删掉，
      // 后续 click 不再派发到它身上（自验实测：点了「复制表格」毫无反应）→ 用捕获阶段 + 归属判断
      document.addEventListener('mousedown', 表格菜单外点, true);
    });
    wrap.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 'c') {
        var sel = window.getSelection();
        if (sel && String(sel).length) { e.preventDefault(); 复制文本(String(sel), '已复制选中单元格'); }
      }
    });
  }
  function 表格菜单外点(ev) { if (!在里(ev, '#boxTblMenu')) 关表格菜单(); }
  function 关表格菜单() {
    var m = $('boxTblMenu');
    if (m && m.parentNode) m.parentNode.removeChild(m);
    document.removeEventListener('mousedown', 表格菜单外点, true);
  }

  function 刷新结果() {
    return fetch(API.latest + '?max_rows=300').then(function (r) { return r.json(); }).then(function (d) {
      renderResult(d);
      var 摘 = (d && d.summary) || {};
      var 时间 = 摘['开始时间'] || '';
      var 注 = '';
      if (摘['失败']) 注 = '　·　最近一次运行有异常（详见日志）';
      $('boxMeta').textContent = 时间 ? ('最近运行 ' + 时间 + 注) : '尚未运行';
      return d;
    }).catch(function () { });
  }

  // ═══════════ 运行日志 ═══════════
  function renderLog() {
    var 行 = state.日志;
    var 显示 = state.展开更多 ? 200 : state.行数;
    var 尾 = 行.slice(-显示);
    $('boxLogbox').textContent = 尾.join('\n');
    $('boxLogPreview').innerHTML = 行.length
      ? '<span class="blp-tag">最新</span><span class="blp-text">' + esc(行[行.length - 1]) + '</span>' : '';
    $('boxLogMore').style.display = state.展开更多 ? 'none' : 'inline-block';
    $('boxTermToggleText').textContent = $('boxTerminal').classList.contains('collapsed') ? '展开' : '折叠';
  }
  function 刷新日志() {
    return fetch(API.log + '?tail=200').then(function (r) { return r.json(); }).then(function (d) {
      state.日志 = d.lines || [];
      $('boxLogTotal').textContent = '　共 ' + (d.total || 0) + ' 行';
      renderNodes(state.日志, state.运行中);
      renderLog();
      return d;
    }).catch(function () { });
  }

  // ═══════════ 运行 / 停止 / 轮询 ═══════════
  function setRunning(on) {
    state.运行中 = !!on;
    $('boxRunBtn').disabled = !!on;
    $('boxStopBtn').disabled = !on;
    renderNodes(state.日志, !!on);
    $('boxTerminal').classList.add('collapsed');
    renderLog();
  }
  function 执行(dry) {
    if (state.运行中) { toast('已有运行中的任务，请先停止'); return; }
    state.日志 = []; state.cursor = 0;
    setRunning(true);
    $('boxRunBtn').textContent = '运行中…';
    fetch(API.run, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ dry_run: !!dry }) })
      .then(function (r) { return r.json().then(function (d) { return { s: r.status, d: d }; }); })
      .then(function (x) {
        if (x.s !== 200) { setRunning(false); $('boxRunBtn').textContent = '运行'; toast('启动失败：' + ((x.d && x.d.detail) || x.s)); return; }
        state.runId = x.d.run_id;
        toast(dry ? '试运行已启动（不发任何通知）' : '已启动运行');
        轮询();
      })
      .catch(function (e) { setRunning(false); $('boxRunBtn').textContent = '运行'; toast('启动异常：' + e); });
  }
  function 轮询() {
    clearTimeout(state.轮询句柄);
    fetch(API.status + '?run_id=' + encodeURIComponent(state.runId) + '&after=' + state.cursor)
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (d.lines && d.lines.length) state.日志 = state.日志.concat(d.lines);
        state.cursor = d.next || state.cursor;
        renderNodes(state.日志, d.running);
        renderLog();
        $('boxLogTotal').textContent = '　共 ' + state.日志.length + ' 行（本次运行）';
        if (d.running) { state.轮询句柄 = setTimeout(轮询, 1500); return; }
        收尾(d);
      })
      .catch(function () { state.轮询句柄 = setTimeout(轮询, 2500); });
  }
  function 收尾(d) {
    setRunning(false);
    $('boxRunBtn').textContent = '运行';
    var 码 = (d && d.exit_code);
    toast(码 === 0 ? '运行完成' : ('运行结束（退出码 ' + 码 + '），详见日志'));
    // 通知是进程结束后由后端异步发的（失败时又要等 dws 超时），所以要错峰多刷几次，
    // 否则「运行结果」旁的通知回执会一直是上一轮的（用户实测：发没发根本看不到反馈）
    [900, 3000, 7000, 12000, 20000].forEach(function (ms) {
      setTimeout(function () { 刷新结果(); 刷新日志(); }, ms);
    });
  }
  function 停止() {
    if (!state.运行中) return;
    fetch(API.stop, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ run_id: state.runId }) })
      .then(function () { toast('已请求停止'); });
  }
  // 重发通知：页面按钮已按用户要求撤掉，仅保留 JS 入口（控制台/其它模块可按需调用）
  function 重发通知(静默) {
    return fetch(API.notify, { method: 'POST' }).then(function (r) { return r.json(); }).then(function (d) {
      var 说明 = (d.已发 && d.已发.length) ? ('已发送 ' + d.已发.map(function (x) { return x.通道; }).join('/')) : (d.跳过 || '未发送');
      if (!静默) toast('重发通知：' + 说明);
      刷新结果(); 刷新日志();
      return d;
    }).catch(function () { if (!静默) toast('重发失败'); });
  }

  // ═══════════ 绑定 ═══════════
  function bind() {
    $('boxRunBtn').addEventListener('click', function () { 执行(false); });
    $('boxStopBtn').addEventListener('click', 停止);
    // 执行步骤：默认折叠（用户定规）；整行可点即折叠/展开（点右上按钮同一效果，只走这一个监听）
    function 切步骤() {
      var sec = $('boxStepsSec');
      sec.classList.toggle('collapsed');
      $('boxStepsToggle').textContent = sec.classList.contains('collapsed') ? '展开 ▾' : '折叠 ▴';
    }
    $('boxStepsHead').addEventListener('click', 切步骤);
    // 运行结果行：整行可点即折叠/展开表格（点 xlsx 超链接 / 按钮不算，各自有各自的行为）
    function 切表格() { state.表格折叠 = !state.表格折叠; renderResult(); }
    $('boxResult').addEventListener('click', function (e) {
      var t = e.target;
      if (!t) return;
      if (t.id === 'boxTableToggle') { 切表格(); return; }
      if (t.closest && (t.closest('a') || t.closest('button'))) return;
      if (!(t.closest && t.closest('.box-resline'))) return;
      if (!(((state.结果数据 || {}).rows) || []).length) return;
      切表格();
    });
    // 步数徽标悬浮 → 具体子步骤气泡
    $('boxNodes').addEventListener('mouseover', function (e) {
      var b = e.target.closest && e.target.closest('.an-steps-badge');
      if (b) 显步提示(b);
    });
    $('boxNodes').addEventListener('mouseout', function (e) {
      if (e.target.closest && e.target.closest('.an-steps-badge')) 关步提示();
    });
    window.addEventListener('scroll', 关步提示, true);
    var inp = $('boxUser'), pinp = $('boxPass');
    inp.addEventListener('input', function () { 存盘(false); });
    pinp.addEventListener('input', function () { 存盘(false); });
    inp.addEventListener('focus', function () { renderHistMenu('user'); 开菜单(false); 定位菜单($('boxUserMenu'), $('boxUser')); });
    pinp.addEventListener('focus', function () { renderHistMenu('pass'); 开菜单(false); 定位菜单($('boxPassMenu'), $('boxPass')); });
    $('boxUserCaret').addEventListener('click', function (e) {
      e.stopPropagation();
      var m = $('boxUserMenu');
      if (m.style.display === 'block') { m.style.display = 'none'; return; }
      renderHistMenu('user'); 开菜单(false); 定位菜单(m, $('boxUser'));
    });
    $('boxUserMenu').addEventListener('click', function (e) { var it = e.target.closest('.box-combo-item'); if (it) 选历史(+it.getAttribute('data-i')); });
    $('boxPassMenu').addEventListener('click', function (e) { var it = e.target.closest('.box-combo-item'); if (it) 选历史(+it.getAttribute('data-i')); });
    $('boxPassEye').addEventListener('click', function (e) {
      e.stopPropagation();
      pinp.type = pinp.type === 'password' ? 'text' : 'password';
      this.textContent = pinp.type === 'password' ? '👁' : '🙈';
    });
    $('boxSwNotify').addEventListener('click', function () { 设总开关(!通知总开关()); });
    $('boxNselTags').addEventListener('click', function (e) {
      var x = e.target.closest('.x'); if (!x) return;
      e.stopPropagation(); 删对象(state.对象[+x.getAttribute('data-i')].值);
    });
    $('boxNsel').addEventListener('click', function (e) {
      if (在里(e, '.box-nsel-menu')) return;
      $('boxNselInp').focus(); 开菜单(true);
    });
    $('boxNselInp').addEventListener('input', function () { renderNselMenu(this.value); 开菜单(true); });
    $('boxNselInp').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ',' || e.key === '，') {
        e.preventDefault();
        var v = this.value.trim(); if (!v) return;
        var it = 找候选(v);
        if (it) { if (有对象(it.值)) { 删对象(it.值); } else { 加对象(it.值); } }
        else if (!加对象(v)) return;          // 认不出类型：保留输入框内容，别清掉
        this.value = ''; renderNselMenu(''); 开菜单(true);
      } else if (e.key === 'Escape') { 开菜单(false); }
    });
    $('boxNselMenu').addEventListener('click', function (e) {
      var cu = e.target.closest('[data-custom]');
      if (cu) {
        var raw = $('boxNselInp').value.trim();
        if (raw && 加对象(raw)) { $('boxNselInp').value = ''; }
        renderNselMenu($('boxNselInp').value);
        return;
      }
      var it = e.target.closest('.box-nsel-opt');
      if (it) { var v = it.getAttribute('data-val'); if (有对象(v)) 删对象(v); else 加对象(v); renderNselMenu($('boxNselInp').value); }
    });
    document.addEventListener('click', function (e) {
      if (!在里(e, '#boxNsel') && !在里(e, '.box-nsel-menu')) 开菜单(false);
      if (!在里(e, '.box-combo') && !在里(e, '.box-combo-menu')) { $('boxUserMenu').style.display = 'none'; $('boxPassMenu').style.display = 'none'; }
    });
    // 注意：别在 .box-bar 上挂 scroll→关菜单——菜单是它的 DOM 后代，菜单内部滚动会冒泡上来
    // 把自己的菜单关掉（自验实测踩过）。关闭只走「点外部」与「窗口缩放」两条路。
    window.addEventListener('resize', function () { 关菜单(); 关表格菜单(); 关步提示(); });
    // 日志：折叠 / 查看更多 / 全选 / 复制
    $('boxTermHead').addEventListener('click', function (e) {
      if (e.target.closest('.box-log-tools')) return;
      var t = $('boxTerminal');
      t.classList.toggle('collapsed');
      $('boxTermToggleText').textContent = t.classList.contains('collapsed') ? '展开' : '折叠';
    });
    $('boxLogMore').addEventListener('click', function () { state.展开更多 = true; renderLog(); });
    $('boxLogSelAll').addEventListener('click', function (e) {
      e.stopPropagation();
      var el = $('boxLogbox'), r = document.createRange();
      r.selectNodeContents(el);
      var s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      toast('已全选日志（' + state.日志.slice(-(state.展开更多 ? 200 : state.行数)).length + ' 行），按 Ctrl+C 复制');
    });
    $('boxLogCopy').addEventListener('click', function (e) {
      e.stopPropagation();
      复制文本($('boxLogbox').innerText, '已复制当前显示的日志');
    });

  }

  function loadOptions() {
    return fetch(API.options).then(function (r) { return r.json(); }).then(function (d) {
      state.opts = d || {}; renderNselMenu(''); return d;
    }).catch(function () { state.opts = {}; });
  }
  function loadConfig() {
    return fetch(API.config).then(function (r) { return r.json(); }).then(function (d) {
      state.cfg = (d && d.config) || {};
      fillBar();
      return d;
    }).catch(function () { });
  }
  function boot() {
    if (!state.booted) {
      state.booted = true;
      bind();
      Promise.all([loadConfig(), loadOptions()]).then(function () {
        刷新日志(); 刷新结果();
        fetch(API.state).then(function (r) { return r.json(); }).then(function (d) {
          if (d && d.running) {
            state.runId = d.run_id; setRunning(true); $('boxRunBtn').textContent = '运行中…';
            state.日志 = []; state.cursor = 0; 轮询();
          } else { setRunning(false); $('boxRunBtn').textContent = '运行'; }
        });
      });
      setInterval(function () {
        if (state.运行中 || document.hidden) return;
        if (($('boxPage') || {}).style.display === 'none') return;
        刷新结果(); 刷新日志();
      }, 20000);
    } else {
      刷新结果(); 刷新日志();
    }
  }

  window.BoxBoard = { boot: boot, refresh: function () { 刷新结果(); 刷新日志(); }, run: function () { 执行(false); }, stop: 停止, notify: 重发通知 };
})();
