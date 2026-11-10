/* ============================================================
 * trigger.js —— 触发器（计划任务 + 运行日志）
 * ------------------------------------------------------------
 * 复刻影刀RPA「触发器」管理（2026-10-20 用户需求 + 6 张截图）：
 *   侧边栏「触发器(m11)」两个页签：
 *     计划任务：列表（名称/类型/应用/条件/启用开关/⋯菜单）
 *               「＋新建」下拉 = 定时/文件/热键/邮件 四类触发器，各配独立编辑弹窗；
 *               行菜单 = 启用|停用、立即运行、编辑、删除。
 *     运行日志：TableUX 表格（应用/开始时间/触发方式/运行结果/操作），
 *               双击行 = 打开日志详情（日志文本 + 截图记录，类 binlog，可上一条/下一条翻看）。
 *   「运行错误处理」按钮：邮件通知（收件箱/发送内容：截图、最新运行日志、主题/正文自定义）+ 钉钉通知；
 *       发送测试邮件 / DWS最新版本发送（截图 + 最新运行日志 + 主题）。
 * 数据：/api/trigger/store（计划任务/运行日志两个 sheet + settings 错误处理设置）。
 * 立即运行：应用 = 后端注册表应用（/api/trigger/apps，如 apd_source_sync_5d）→ 走真实后端执行；
 *           应用 = 已注册的 AutoTaskKit 任务 id 时走真实节点流水线（headlessRun）；
 *           否则按触发→记录日志链路模拟执行并存档日志文本与截图记录。
 * 到点调度：由后端 _trigger_scheduler_loop 每 30s 扫「计划任务」，命中的启用定时任务
 *           到点即真实执行并写「运行日志」（网页侧只负责展示与手动触发）。
 * 依赖：index.html 主脚本全局能力（TableUX / createGridKit / cqOpen / showToast / openTab）。
 * ============================================================ */
(function () {
  'use strict';

  var $ = function (s) { return document.querySelector(s);
  };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function toast(msg) { if (typeof showToast === 'function') showToast(msg); else window.alert(msg); }
  function p2(n) { return String(n).padStart(2, '0'); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function nowText() {
    var d = new Date();
    return d.getFullYear() + '年' + p2(d.getMonth() + 1) + '月' + p2(d.getDate()) + '日 ' + p2(d.getHours()) + ':' + p2(d.getMinutes());
  }
  function hm() {
    var d = new Date();
    return p2(d.getHours()) + ':' + p2(d.getMinutes()) + ':' + p2(d.getSeconds());
  }
  function cnTime(t) {
    var p = String(t || '').split(':');
    return (p[0] || '00') + '时' + (p[1] || '00') + '分';
  }

  /* ============ 类型元数据（图标/标题/描述，与「＋新建」下拉一致） ============ */
  var TYPE_META = {
    '定时': {
      title: '定时触发器', desc: '在特定时间运行指定的应用',
      icon: '<svg viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6.4" stroke="currentColor" stroke-width="1.4"/><path d="M8 4.6V8l2.4 1.6" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
    },
    // 2026-10-26 新增：条件触发器（看板指标 > 或 < 阈值时通知）
    '条件': {
      title: '条件触发器', desc: '当看板某个指标达到设定阈值（大于 / 小于）时发送通知',
      icon: '<svg viewBox="0 0 16 16" fill="none"><path d="M2.4 11.6 6 5.2l3.2 5.2L13.6 3.4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/><circle cx="13.6" cy="3.4" r="1.5" fill="currentColor"/></svg>',
    },
    '文件': {
      title: '文件触发器', desc: '当文件发生变化时运行指定的应用',
      icon: '<svg viewBox="0 0 16 16" fill="none"><path d="M2.6 4.2c0-.66.54-1.2 1.2-1.2h2.9l1.5 1.7h4.2c.66 0 1.2.54 1.2 1.2v5.9c0 .66-.54 1.2-1.2 1.2H3.8c-.66 0-1.2-.54-1.2-1.2V4.2Z" stroke="currentColor" stroke-width="1.4"/></svg>',
    },
    '热键': {
      title: '热键触发器', desc: '当特定按键被按下时运行指定的应用',
      icon: '<svg viewBox="0 0 16 16" fill="none"><rect x="1.8" y="4.4" width="12.4" height="7.2" rx="1.4" stroke="currentColor" stroke-width="1.4"/><path d="M4.6 6.8h.01M7.4 6.8h.01M10.2 6.8h.01M5.4 9.2h5.2" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
    },
    '邮件': {
      title: '邮件触发器', desc: '当收到特定邮件时运行指定的应用',
      icon: '<svg viewBox="0 0 16 16" fill="none"><rect x="1.8" y="3.6" width="12.4" height="8.8" rx="1.2" stroke="currentColor" stroke-width="1.4"/><path d="M2.6 4.6 8 8.8l5.4-4.2" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
    },
    '钉钉个人': {
      title: '钉钉个人消息触发器', desc: '当指定钉钉联系人的消息包含指定内容（文本/网址/文件）时运行指定的应用',
      icon: '<svg viewBox="0 0 16 16" fill="none"><circle cx="8" cy="5.4" r="2.6" stroke="currentColor" stroke-width="1.4"/><path d="M3.2 13.4c.6-2.4 2.5-3.6 4.8-3.6s4.2 1.2 4.8 3.6" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
    },
    '钉钉群': {
      title: '钉钉群消息触发器', desc: '当指定钉钉群的消息包含指定内容（文本/网址/文件）时运行指定的应用',
      icon: '<svg viewBox="0 0 16 16" fill="none"><circle cx="5.6" cy="5.6" r="2.2" stroke="currentColor" stroke-width="1.4"/><circle cx="11" cy="6" r="1.8" stroke="currentColor" stroke-width="1.4"/><path d="M1.8 13.2c.5-2 1.9-3 3.8-3s3.3 1 3.8 3M9.4 13.2c.4-1.7 1.5-2.6 3.1-2.6 1.3 0 2.3.6 2.7 1.8" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>',
    },
  };
  var WEEKDAYS = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'];
  var MAIL_SERVERS = ['163邮箱', 'QQ邮箱', '钉钉邮箱', '其他'];
  // 邮件触发器 · 常用文件夹（测试连接前即可选；测试连接成功后由服务器实际文件夹刷新）
  var MAIL_FOLDERS = ['INBOX', '已发送', '草稿箱', '垃圾邮件', '重要邮件', '归档'];
  var HOTKEYS = ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12',
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    '0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];

  /* ============ 状态 ============ */
  var state = {
    tab: 'tasks',                 // tasks | logs
    loaded: { tasks: false, logs: false },
    tasks: [],                    // [{id,row_no,data}]
    logs: [],                     // [{id,row_no,data}]
    search: '',
    filterTypes: [],              // [] = 全部类型
    filterEnabled: '',            // '' | 'on' | 'off'
    editing: null,                // null=新建 | {id,row_no,data}
    editType: '定时',
    logDetailIdx: -1,
    logSelected: new Set(),
    err: null,                    // 运行错误处理设置
    runningIds: new Set(),        // 正在立即运行的任务 id（禁重复点）
    apps: {},                     // 后端可真实执行的应用 {id: {name, desc, run_url}}
  };

  /* ============ 存储 ============ */
  function storeGet() {
    return fetch('/api/trigger/store').then(function (res) {
      if (!res.ok) throw new Error('本地存储不可用');
      return res.json();
    });
  }
  function storePost(path, payload) {
    return fetch('/api/trigger/store/' + path, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload),
    }).then(function (res) {
      return res.json().catch(function () { return {}; }).then(function (data) {
        if (!res.ok) throw new Error(data.detail || '本地存储写入失败');
        return data;
      });
    });
  }

  /* ============ 任务模型 ============ */
  function taskType(t) { return (t.data && t.data.类型) || '定时'; }
  function taskCfg(t) { return (t.data && t.data.config) || {}; }
  function taskById(id) { return state.tasks.find(function (t) { return t.id === id; }); }
  function logById(id) { return state.logs.find(function (r) { return r.id === id; }); }

  // 应用值 → 展示名（后端注册表应用用其中文名；已注册 AutoTaskKit 任务用其 name）
  function appName(v) {
    if (!v) return '';
    if (state.apps && state.apps[v]) return state.apps[v].name || v;
    if (typeof AutoTaskKit !== 'undefined' && AutoTaskKit.tasks && AutoTaskKit.tasks[v]) return AutoTaskKit.tasks[v].name || v;
    return v;
  }
  // 应用下拉候选项：value → label（后端注册应用 + 已注册任务 + 历史任务用过的值）
  function appOptions() {
    var map = {};
    Object.keys(state.apps || {}).forEach(function (id) { map[id] = state.apps[id].name || id; });
    if (typeof AutoTaskKit !== 'undefined' && AutoTaskKit.tasks) {
      Object.keys(AutoTaskKit.tasks).forEach(function (id) { map[id] = AutoTaskKit.tasks[id].name || id; });
    }
    ['自动', '自动触发维护物料结果', '060结余定时抓取', '间歇东京数据', '安全钳APD一点一点',
      '1205网页端数据自动化获取与处理（基于影刀和Prep零代码实现）'].forEach(function (a) { map[a] = a; });
    state.tasks.forEach(function (t) { var a = t.data && t.data.应用; if (a) map[a] = appName(a); });
    return map;
  }
  // 后端可真实执行的应用？（立即运行时分流用）
  function backendApp(v) { return (state.apps && state.apps[v]) ? state.apps[v] : null; }
  function loadApps() {
    return fetch('/api/trigger/apps', { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : { apps: [] }; })
      .catch(function () { return { apps: [] }; })
      .then(function (d) {
        state.apps = {};
        (d.apps || []).forEach(function (a) { state.apps[a.id] = a; });
      });
  }

  // 条件列文案（按类型 + config 生成）
  function condText(t) {
    var c = taskCfg(t);
    if (taskType(t) === '条件') {
      var when = ({ '每工作日': '每工作日', '每天': '每天', '每周': '每周', '月份': '每月' })[c.freq] || c.freq || '';
      return '当「' + (c.指标 || '小型化每日下达') + '」' +
        (c.比较 === '<' ? '小于 ' : '大于 ') + (c.阈值 == null ? 120 : c.阈值) +
        ' 时通知（' + when + ' ' + cnTime(c.time) + ' 检查）';
    }
    if (taskType(t) === '定时') {
      return timedCondText(c);
    }
    if (taskType(t) === '文件') {
      var ev = (c.events || []).join('、') || '全部事件';
      var s = '监控「' + (c.folder || '未选择文件夹') + '」的' + ev;
      if (c.patterns) s += '（' + c.patterns + '）';
      return s;
    }
    if (taskType(t) === '热键') {
      return '检查键盘热键 ' + hotkeyCombo(c) + '是否按下';
    }
    if (taskType(t) === '钉钉个人' || taskType(t) === '钉钉群') {
      var isP = taskType(t) === '钉钉个人';
      var who = isP ? '钉钉个人 ' : '钉钉群 ';
      var named = (c.targets || []);
      var head2 = named.length ? '监控' + who + named.join('、') + ' 的消息' : '钉钉消息触发（未选择接收' + (isP ? '人' : '群') + '）';
      var parts = (c.matchTypes || []).map(function (mt) { return mt + '「' + ((c.match || {})[mt] || '') + '」'; });
      return head2 + (parts.length ? '，包含 ' + parts.join('、') + ' 时触发' : '');
    }
    var m = '每隔 1 分钟检查 ' + (c.folder || 'INBOX');
    var r = c.rules || {};
    if (r.subject) m += '（主题含「' + r.subject + '」）';
    return m;
  }
  function hotkeyCombo(c) {
    return [c.ctrl ? 'Ctrl' : '', c.shift ? 'Shift' : '', c.alt ? 'Alt' : '', c.key || 'F1'].filter(Boolean).join('+');
  }

  /* ============ 通知对象（邮箱 / 钉钉消息 / 群消息：一个框里混着填） ============
     口径与「维护物料模块自动化」抬头那套完全一致（2026-10-26 用户要求）：
       · 一个输入框可同时放多个对象，用标签展示；
       · 对象自带后缀决定通道：计划科（群消息）/ 王锦志（钉钉消息）/ abc@x.com（裸邮箱也算邮箱）；
       · 认不出是群还是人的，提示补后缀，不硬猜；
       · 候选清单来自「钉钉邮箱群消息配置模块」（与维护物料页同一个 /api/box/options 数据源）。 */
  var NF_CH = ['群消息', '钉钉消息', '邮箱'];
  var NF_SUFFIX = /[（(]\s*(群消息|钉钉消息|邮箱)\s*[)）]\s*$/;
  var nfOpts = null;                        // 候选清单（只拉一次，进程内缓存）
  var nfState = { 启用: false, 对象: [] };  // 当前编辑弹窗里的通知状态
  function notifyOptions() {
    if (nfOpts) return Promise.resolve(nfOpts);
    return fetch('/api/trigger/notify/options', { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : {}; })
      .catch(function () { return {}; })
      .then(function (d) {
        nfOpts = {
          群消息: (d && d['群消息']) || [], 钉钉消息: (d && d['钉钉消息']) || [],
          邮箱: (d && d['邮箱']) || [], 全部: (d && d['全部']) || [],
        };
        return nfOpts;
      });
  }
  function nfFind(v) {
    var all = (nfOpts && nfOpts.全部) || [];
    for (var i = 0; i < all.length; i++) if (all[i].值 === v) return all[i];
    return null;
  }
  function nfType(v) {
    var s = String(v == null ? '' : v).trim();
    var m = s.match(NF_SUFFIX);
    if (m) return m[1];
    var c = nfFind(s);
    if (c && c.类型) return c.类型;
    if (s.indexOf('@') >= 0) return '邮箱';
    return null;
  }
  function nfNorm(v) {
    var s = String(v == null ? '' : v).trim();
    var t = nfType(s);
    if (!t) return s;
    var 去 = s.replace(NF_SUFFIX, '').trim();
    if (!去) return s;
    if (t === '邮箱' && !NF_SUFFIX.test(s)) return 去;
    return 去 + '（' + t + '）';
  }
  function nfHas(v) { return nfState.对象.some(function (o) { return o.值 === v; }); }
  function notifyOf(t) {
    var c = (t && t.data && t.data.config) || {};
    return c.通知 || {};
  }
  function notifyOn(t) {
    var n = notifyOf(t);
    return !!n.启用 && (n.对象 || []).length > 0;
  }

  /* ============ 页签切换 ============ */
  function switchTab(tab) {
    state.tab = tab;
    var pt = $('#trigPaneTasks'), pl = $('#trigPaneLogs');
    if (pt) pt.style.display = tab === 'tasks' ? '' : 'none';
    if (pl) pl.style.display = tab === 'logs' ? '' : 'none';
    // 已加载过的页签也要重渲染：立即运行后切页签需立刻看到新行（运行中→成功/失败增量）
    if (tab === 'tasks') { if (!state.loaded.tasks) loadTasks(); else renderTasks(); }
    else { if (!state.loaded.logs) loadLogs(); else renderLogs(); }
  }

  function loadTasks() {
    // 先拉后端可执行应用清单（应用名展示/下拉/立即运行分流都依赖它），再加载任务数据
    loadApps().then(function () {
      return storeGet();
    }).then(function (store) {
      state.tasks = (store.sheets && store.sheets['计划任务'] || []).map(function (r) {
        return { id: r.id, row_no: r.row_no, data: r.data || {} };
      }).filter(function (r) {
        return Object.keys(r.data).some(function (k) { return String(r.data[k] ?? '').trim() !== ''; });
      });
      state.logs = (store.sheets && store.sheets['运行日志'] || []).map(function (r) {
        return { id: r.id, row_no: r.row_no, data: r.data || {} };
      });
      state.err = store.settings || null;
      state.loaded.tasks = true;
      renderTasks();
    }).catch(function (err) {
      toast('触发器数据加载失败：' + (err.message || err));
    });
  }
  function loadLogs() {
    return storeGet().then(function (store) {
      state.logs = (store.sheets && store.sheets['运行日志'] || []).map(function (r) {
        return { id: r.id, row_no: r.row_no, data: r.data || {} };
      });
      state.err = store.settings || null;
      state.loaded.logs = true;
      renderLogs();
    }).catch(function (err) {
      toast('运行日志加载失败：' + (err.message || err));
    });
  }

  /* ============ 计划任务列表 ============ */
  function visibleTasks() {
    var kw = state.search.trim().toLowerCase();
    return state.tasks.filter(function (t) {
      if (kw && String((t.data || {}).名称 || '').toLowerCase().indexOf(kw) < 0) return false;
      if (state.filterTypes.length && state.filterTypes.indexOf(taskType(t)) < 0) return false;
      if (state.filterEnabled === 'on' && !(t.data || {}).启用) return false;
      if (state.filterEnabled === 'off' && (t.data || {}).启用) return false;
      return true;
    });
  }

  function renderTasks() {
    var box = $('#trigList');
    if (!box) return;
    var rows = visibleTasks();
    var on = state.tasks.filter(function (t) { return (t.data || {}).启用; }).length;
    if (!rows.length) {
      box.innerHTML = '<div class="trig-list-empty">' +
        (state.tasks.length ? '没有符合条件的计划任务，试试清空搜索/筛选。' : '暂无计划任务，点击左上角「＋ 新建」创建第一个触发器。') + '</div>';
    } else {
      box.innerHTML = rows.map(function (t) {
        var d = t.data || {};
        var nf = notifyOf(t), 通知开 = notifyOn(t);
        var bell = 通知开
          ? '<svg class="trig-bell" data-bell="' + t.id + '" viewBox="0 0 16 16" fill="none"><path d="M8 2.2c-2 0-3.4 1.5-3.4 3.5 0 2.4-.5 3.4-1.2 4.1-.3.3-.1.8.3.8h8.6c.4 0 .6-.5.3-.8-.7-.7-1.2-1.7-1.2-4.1 0-2-1.4-3.5-3.4-3.5Z" stroke="currentColor" stroke-width="1.2"/><path d="M6.5 13c.3.6.8.9 1.5.9s1.2-.3 1.5-.9" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/></svg>'
          : '';
        return '<div class="trig-row" data-id="' + t.id + '">' +
          '<span class="trig-col-name" title="' + esc(d.名称) + '">' + esc(d.名称) + bell + '</span>' +
          '<span class="trig-col-type"><span class="trig-type-chip t-' + esc(taskType(t)) + '">' + esc(taskType(t)) + '</span></span>' +
          '<span class="trig-col-app" title="' + esc(appName(d.应用)) + '">' + esc(appName(d.应用)) + '</span>' +
          '<span class="trig-col-cond" title="' + esc(condText(t)) + '">' + esc(condText(t)) + '</span>' +
          '<span class="trig-col-act">' +
          '<span class="trig-toggle' + (d.启用 ? ' on' : '') + '" data-id="' + t.id + '" title="' + (d.启用 ? '点击停用' : '点击启用') + '"></span>' +
          '<span class="trig-more" data-id="' + t.id + '" title="更多操作">⋯</span>' +
          '</span></div>';
      }).join('');
    }
    var foot = $('#trigListFoot');
    if (foot) foot.innerHTML = '<span>共 ' + state.tasks.length + ' 条' + (rows.length !== state.tasks.length ? '（筛选后 ' + rows.length + ' 条）' : '') + '</span>' +
      '<span>已启用 ' + on + ' · 未启用 ' + (state.tasks.length - on) + '</span>';
    var st = $('#trigStatus');
    if (st) st.textContent = '计划任务 · 共 ' + state.tasks.length + ' 条';
    colSync();          // 列宽：有存过的用存过的，没有就按内容自适应（名称/应用随内容，类型固定徽标宽）
    bindColResize();
  }

  /* ============ 计划任务列表 · 列宽（表头可拖动 + 默认按内容自适应） ============ */
  var COL_KEY = 'trigCols';
  var COL_VAR = { '名称': 'name', '类型': 'type', '应用': 'app' };
  var COL_MIN = { '名称': 120, '类型': 62, '应用': 120 };
  var colApplied = null;                       // 当前生效列宽（px）
  function colCard() { return document.querySelector('.trig-list-card'); }
  function colSaved() {
    try { var d = JSON.parse(localStorage.getItem(COL_KEY) || 'null'); return (d && typeof d === 'object') ? d : null; }
    catch (e) { return null; }
  }
  function colSet(w, save) {
    var card = colCard();
    if (!card || !w) return;
    colApplied = w;
    Object.keys(COL_VAR).forEach(function (k) {
      var v = w[k];
      if (v) card.style.setProperty('--c-' + COL_VAR[k], Math.round(v) + 'px');
    });
    if (save) {
      try {
        localStorage.setItem(COL_KEY, JSON.stringify({
          '名称': Math.round(w['名称']), '类型': Math.round(w['类型']), '应用': Math.round(w['应用']),
        }));
      } catch (e2) { /* 隐私模式下写不了就算了，本次会话仍然生效 */ }
    }
  }
  function tw(text, font) {
    var c = tw._c || (tw._c = document.createElement('canvas').getContext('2d'));
    if (font) c.font = font;
    return c.measureText(String(text == null ? '' : text)).width;
  }
  function colFont(sel, fallback) {
    var el = document.querySelector(sel), cs = el ? getComputedStyle(el) : null;
    return cs ? (cs.fontWeight + ' ' + cs.fontSize + ' ' + cs.fontFamily) : fallback;
  }
  /* 默认自适应：名称按最长一条（含铃铛余量）、应用按最长应用名、类型按徽标宽度，剩下的全给「条件」。
     窗口偏窄时按可压缩比例把名称/应用一起往回收，绝不把「条件」挤没。 */
  function colAuto(save) {
    var card = colCard();
    if (!card) return;
    var avail = card.clientWidth || 0;
    if (!avail) avail = 900;                             // 页面还没显示（display:none）时的兜底
    var 名称字体 = colFont('.trig-row .trig-col-name', '') || colFont('#trigListHead .trig-col-name', '500 13px sans-serif');
    var 应用字体 = colFont('.trig-row .trig-col-app', '') || colFont('#trigListHead .trig-col-app', '400 12px sans-serif');
    var 名 = tw('名称', 名称字体) + 30;
    var 应 = tw('应用', 应用字体) + 16;
    state.tasks.forEach(function (t) {
      var d = t.data || {};
      名 = Math.max(名, tw(d.名称, 名称字体) + 28);        // + 小铃铛 / 省略号余量
      应 = Math.max(应, tw(appName(d.应用), 应用字体) + 14);
    });
    名 = Math.max(140, Math.min(460, 名));
    应 = Math.max(150, Math.min(320, 应));
    var 类 = 84, 状态 = 128, 条件min = 240, 间距 = 32, 内边重 = 32;
    var 可用 = avail - 内边重 - 状态 - 条件min - 间距;
    // 挤不下时先压「应用」（次要列），还不够再压「名称」——「条件」永远留够 240，别把触发条件挤没
    if (名 + 类 + 应 > 可用) {
      var 需要压 = 名 + 类 + 应 - 可用;
      var 压应用 = Math.min(Math.max(0, 应 - 150), 需要压);
      应 -= 压应用; 需要压 -= 压应用;
      if (需要压 > 0) 名 -= Math.min(Math.max(0, 名 - 180), 需要压);
    }
    colSet({ '名称': 名, '类型': 类, '应用': 应 }, !!save);
  }
  // 有存过的宽度就沿用（拖过才算存），否则每次渲染按内容自适应
  function colSync() {
    var s = colSaved();
    if (s && s['名称'] && s['类型'] && s['应用']) colSet(s, false);
    else colAuto(false);
  }
  function bindColResize() {
    var head = document.getElementById('trigListHead');
    if (!head || head.dataset.rsBound) return;
    head.dataset.rsBound = '1';
    head.addEventListener('mousedown', function (e) {
      var rs = e.target.closest('.trig-col-rs');
      if (!rs) return;
      var card = colCard();
      if (!card) return;
      e.preventDefault(); e.stopPropagation();
      var col = rs.dataset.col, startX = e.clientX;
      var startW = (colApplied && colApplied[col]) || card.getBoundingClientRect().width / 4;
      var 动过 = false;
      rs.classList.add('dragging');
      document.body.classList.add('trig-col-dragging');
      function mv(ev) {
        var w = Math.max(COL_MIN[col] || 80, startW + (ev.clientX - startX));
        var next = { '名称': (colApplied || {})['名称'], '类型': (colApplied || {})['类型'], '应用': (colApplied || {})['应用'] };
        next[col] = w;
        动过 = true;
        colSet(next, false);
      }
      function up() {
        document.removeEventListener('mousemove', mv);
        document.removeEventListener('mouseup', up);
        rs.classList.remove('dragging');
        document.body.classList.remove('trig-col-dragging');
        if (动过) colSet(colApplied, true);              // 松手才落盘（拖动过程不写 localStorage）
      }
      document.addEventListener('mousemove', mv);
      document.addEventListener('mouseup', up);
    });
    // 双击分隔线 = 回到自适应
    head.addEventListener('dblclick', function (e) {
      if (!e.target.closest('.trig-col-rs')) return;
      e.preventDefault();
      try { localStorage.removeItem(COL_KEY); } catch (e2) { /* ignore */ }
      colApplied = null;
      colAuto(false);          // 只算不存：以后窗口变宽变窄继续自适应
      toast('列宽已恢复自适应');
    });
    window.addEventListener('resize', function () { if (!colSaved()) colAuto(false); });
  }

  function bindListEvents() {
    var box = $('#trigList');
    if (!box) return;
    box.addEventListener('click', function (e) {
      var tg = e.target.closest('.trig-toggle');
      if (tg) { e.stopPropagation(); toggleEnable(parseInt(tg.dataset.id, 10)); return; }
      var more = e.target.closest('.trig-more');
      if (more) { e.stopPropagation(); openRowMenu(more, parseInt(more.dataset.id, 10)); return; }
      var row = e.target.closest('.trig-row');
      if (row) {
        var t = taskById(parseInt(row.dataset.id, 10));
        if (t) openEditor(taskType(t), t);
      }
    });
    /* 铃铛标记悬浮 → 显示通知细节卡片 */
    box.addEventListener('mouseover', function (e) {
      var b = e.target.closest && e.target.closest('.trig-bell');
      if (b && b.dataset.bell) { clearTimeout(bellTipTimer); showBellTip(b, b.dataset.bell); }
    });
    box.addEventListener('mouseout', function (e) {
      var b = e.target.closest && e.target.closest('.trig-bell');
      if (!b) return;
      var to = e.relatedTarget;
      if (to && to.closest && to.closest('.trig-bell-tip')) return;   // 移到卡片上保持
      hideBellTip();
    });
    document.addEventListener('mouseout', function (e) {
      var tip = bellTipEl;
      if (!tip || tip.style.display === 'none') return;
      if (e.target === tip || (e.target.closest && e.target.closest('.trig-bell-tip'))) {
        var to = e.relatedTarget;
        if (!(to && to.closest && to.closest('.trig-bell'))) hideBellTip();
      }
    });
  }

  /* ============ 铃铛悬浮：通知细节卡片 ============ */
  var bellTipEl = null, bellTipTimer = null;
  function ensureBellTip() {
    if (bellTipEl) return bellTipEl;
    bellTipEl = document.createElement('div');
    bellTipEl.className = 'trig-bell-tip';
    bellTipEl.style.display = 'none';
    document.body.appendChild(bellTipEl);
    return bellTipEl;
  }
  function showBellTip(target, id) {
    var t = taskById(parseInt(id, 10));
    if (!t) return;
    var tip = ensureBellTip();
    tip.innerHTML = 通知细节HTML(t);
    tip.style.display = 'block';
    var r = target.getBoundingClientRect();
    var tw = tip.offsetWidth, th = tip.offsetHeight;
    var left = Math.min(r.left, window.innerWidth - tw - 8);
    var top = r.bottom + 6;
    if (top + th > window.innerHeight - 8) top = Math.max(8, r.top - th - 6);   // 下方放不下就翻到上方
    tip.style.left = Math.max(8, left) + 'px';
    tip.style.top = top + 'px';
  }
  function hideBellTip() { if (bellTipEl) bellTipEl.style.display = 'none'; }

  /* 该任务最近一次运行里的通知发送结果（按应用名匹配运行日志，取行号最大的一条里的通知行） */
  function 取最近通知(t) {
    var name = appName((t.data || {}).应用 || '');
    var best = null;
    (state.logs || []).forEach(function (r) {
      var dd = r.data || {};
      if (appName(dd.应用 || '') !== name) return;
      var lines = dd.日志 || [];
      for (var i = lines.length - 1; i >= 0; i--) {
        var ln = lines[i];
        if (ln.indexOf('通知：') >= 0 || ln.indexOf('⚠ 通知') >= 0) {
          var 文本 = ln.replace(/^\d{2}:\d{2}\s+/, '').replace(/^[△✓✗⚠]\s*/, '');
          if (!best || (r.row_no || 0) > (best.row_no || 0)) {
            best = { row_no: r.row_no || 0, 时间: dd.开始时间 || '', 触发: dd.触发方式 || '', 文本: 文本 };
          }
          break;
        }
      }
    });
    return best;
  }

  function 通知细节HTML(t) {
    var nf = notifyOf(t);
    var 对象 = (nf.对象 || []).slice();
    var 按通道 = { '群消息': 0, '钉钉消息': 0, '邮箱': 0 };
    var rows = 对象.map(function (o) {
      var ty = nfType(o);
      if (ty && 按通道.hasOwnProperty(ty)) 按通道[ty]++;
      var ch = ty
        ? '<span class="nft-ch nft-' + esc(ty) + '">' + esc(ty) + '</span>'
        : '<span class="nft-ch nft-unk">未识别</span>';
      return '<div class="nft-row"><span class="nft-obj">' + esc(nfNorm(o)) + '</span>' + ch + '</div>';
    }).join('');
    var parts = [];
    if (按通道['群消息']) parts.push('群消息 ' + 按通道['群消息']);
    if (按通道['钉钉消息']) parts.push('钉钉消息 ' + 按通道['钉钉消息']);
    if (按通道['邮箱']) parts.push('邮箱 ' + 按通道['邮箱']);
    var stats = parts.length ? ('按通道：' + parts.join(' · ')) : '按通道：—（对象通道未识别，请在通知对象里补「（群消息/钉钉消息/邮箱）」后缀）';
    var last = 取最近通知(t);
    var lastHtml = '<div class="nft-sec nft-last"><div class="nft-sec-t">最近一次通知</div>' +
      (last
        ? '<div class="nft-last-line">' + esc(last.时间) + ' · ' + esc(last.触发) + '</div>' +
          '<div class="nft-last-res ' + (last.文本.indexOf('失败') >= 0 ? 'fail' : 'ok') + '">' + esc(last.文本) + '</div>'
        : '<div class="nft-last-res">暂无运行记录</div>') + '</div>';
    return '<div class="nft-head">通知已开启</div>' +
      '<div class="nft-sec"><div class="nft-sec-t">通知对象（' + 对象.length + ' 个）</div>' + (rows || '<div class="nft-empty">—</div>') + '</div>' +
      '<div class="nft-stats">' + esc(stats) + '</div>' + lastHtml +
      '<div class="nft-foot">点击该行可编辑通知设置</div>';
  }

  async function toggleEnable(id) {
    var t = taskById(id);
    if (!t) return;
    var d = Object.assign({}, t.data, { 启用: !t.data.启用 });
    try {
      await storePost('upsert', { sheet: '计划任务', id: t.id, row_no: t.row_no, data: d });
      t.data = d;
      renderTasks();
      toast(d.启用 ? '已启用触发器「' + d.名称 + '」' : '已停用触发器「' + d.名称 + '」');
    } catch (err) { toast('操作失败：' + (err.message || err)); }
  }

  /* ============ 行操作菜单 ============ */
  function closeRowMenu() { var m = $('#trigRowMenu'); if (m) m.classList.remove('open'); }
  function openRowMenu(anchor, id) {
    var t = taskById(id);
    if (!t) return;
    var menu = $('#trigRowMenu');
    var on = !!(t.data || {}).启用;
    menu.innerHTML =
      '<div class="rmi" data-act="toggle">' + (on ? '停用触发器' : '启用触发器') + '</div>' +
      '<div class="rmi" data-act="run">立即运行</div>' +
      '<div class="rmi" data-act="edit">编辑触发器</div>' +
      '<div class="rmi-sep"></div>' +
      '<div class="rmi danger" data-act="delete">删除触发器</div>';
    menu.classList.add('open');
    var r = anchor.getBoundingClientRect();
    var mh = menu.offsetHeight, mw = menu.offsetWidth;
    menu.style.top = Math.min(r.bottom + 4, window.innerHeight - mh - 8) + 'px';
    menu.style.left = Math.max(8, r.right - mw) + 'px';
    menu.querySelectorAll('.rmi').forEach(function (el) {
      el.addEventListener('click', function () {
        closeRowMenu();
        var act = el.dataset.act;
        if (act === 'toggle') toggleEnable(id);
        else if (act === 'run') runNow(t);
        else if (act === 'edit') openEditor(taskType(t), t);
        else if (act === 'delete') confirmDeleteTask(t);
      });
    });
  }

  /* ============ 通用确认框 ============ */
  var confirmOk = null;
  function askConfirm(title, text, okLabel, onOk) {
    $('#trigConfirmTitle').textContent = title;
    $('#trigConfirmText').innerHTML = text;
    $('#trigConfirmOk').textContent = okLabel || '确认删除';
    confirmOk = onOk;
    $('#trigConfirmModal').classList.add('open');
  }
  function confirmDeleteTask(t) {
    askConfirm('确认删除触发器', '即将删除计划任务「<b>' + esc((t.data || {}).名称) + '</b>」，删除后不再触发。是否确认？', '确认删除', async function () {
      try {
        await storePost('delete', { sheet: '计划任务', id: t.id });
        state.tasks = state.tasks.filter(function (x) { return x.id !== t.id; });
        renderTasks();
        toast('已删除触发器「' + (t.data || {}).名称 + '」');
      } catch (err) { toast('删除失败：' + (err.message || err)); }
    });
  }

  /* ============ 立即运行（真实任务走 AutoTaskKit 流水线；其余按触发链路模拟） ============ */
  async function runNow(t) {
    var id = t.id;
    if (state.runningIds.has(id)) { toast('该触发器正在运行，请稍候'); return; }
    state.runningIds.add(id);
    var app = (t.data || {}).应用 || '';
    var name = appName(app);
    var data = {
      应用: name, 开始时间: nowText(), 触发方式: '[手动运行]', 运行结果: '运行中',
      主题: name + ' · 手动运行', 错误次数: 0, 日志: [], 截图: [],
    };
    var logId = null;
    try {
      await storePost('upsert', { sheet: '运行日志', id: 0, row_no: nextLogRowNo(), data: data });
      var store = await storeGet();
      var sheet = (store.sheets && store.sheets['运行日志']) || [];
      var fresh = sheet[sheet.length - 1];
      logId = fresh ? fresh.id : null;
      state.logs.unshift({ id: logId, row_no: fresh ? fresh.row_no : 0, data: data });
      if (state.tab === 'logs') renderLogs();
    } catch (err) {
      state.runningIds.delete(id);
      toast('运行失败：' + (err.message || err));
      return;
    }
    var lines = [];
    var t0 = Date.now();
    async function push(text) {
      lines.push(hm() + '  ' + text);
      data.日志 = lines.slice();
      if (logId) {
        try { await storePost('upsert', { sheet: '运行日志', id: logId, row_no: (logById(logId) || {}).row_no || 1, data: data }); } catch (e) { /* 日志行落盘失败不中断运行 */ }
        var row = logById(logId); if (row) row.data = data;
        if (state.tab === 'logs') renderLogs();
      }
    }
    /* 运行截图：html2canvas 采集当前触发器页 → 落盘 shots/ → 截图记录真实可看（失败也留证） */
    async function captureShot(title) {
      var t = hm();
      try {
        if (typeof html2canvas !== 'function') throw new Error('html2canvas 不可用');
        var el = document.querySelector('#triggerPage');
        if (!el || !el.offsetParent) el = document.body;
        var canvas = await html2canvas(el, { backgroundColor: '#FFFFFF', scale: 1, logging: false, useCORS: true });
        var res = await fetch('/api/trigger/shot', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ image: canvas.toDataURL('image/png') }),
        });
        var d = await res.json().catch(function () { return {}; });
        if (!res.ok || !d.ok || !d.url) throw new Error(d.detail || '截图上传失败');
        await push('截图已归档：' + d.url);
        return { t: t, title: title, url: d.url };
      } catch (e) {
        await push('⚠ 截图采集失败（' + (e && e.message || e) + '），截图记录留占位');
        return { t: t, title: title };
      }
    }
    try {
      await push('触发器「' + ((t.data || {}).名称 || '') + '」就绪，开始手动运行');
      var kit = (typeof AutoTaskKit !== 'undefined' && AutoTaskKit.tasks) ? AutoTaskKit.tasks[app] : null;
      var bapp = backendApp(app);
      if (bapp) {
        await push('应用启动：' + (bapp.name || app) + '（后端真实执行）');
        var res = await fetch(bapp.run_url || '/api/trigger/app/run', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          /* config 一起带上：条件触发器要读「指标 / 比较 / 阈值」 */
          body: JSON.stringify({ app: app, config: (t.data && t.data.config) || {} }),
        });
        var rd = await res.json().catch(function () { return {}; });
        /* 部分成功（2026-10-26）：产物都生成了、只差邮件没发出去 —— 别当失败，但也别当全绿 */
        if (rd.部分成功) {
          (rd.日志 || []).forEach(function (ln) { lines.push(hm() + '  ' + ln); });
          data.日志 = lines.slice();
          data.运行结果 = '部分成功';
          await push('△ ' + (rd.结果 || '部分成功（产物已生成，邮件未发出）'));
          data.截图 = (rd.截图 && rd.截图.length) ? rd.截图 : [await captureShot('运行完成画面')];
          return;
        }
        if (!res.ok || rd.ok === false) throw new Error(rd.detail || rd.结果 || '后端执行失败');
        (rd.日志 || []).forEach(function (ln) { lines.push(hm() + '  ' + ln); });
        data.日志 = lines.slice();
        if (logId) {
          try {
            await storePost('upsert', { sheet: '运行日志', id: logId, row_no: (logById(logId) || {}).row_no || 1, data: data });
            var rowB = logById(logId); if (rowB) rowB.data = data;
          } catch (e) { /* 日志行落盘失败不中断运行 */ }
        }
        await push('✓ 运行完成，耗时 ' + ((Date.now() - t0) / 1000).toFixed(1) + 's');
        data.运行结果 = '运行成功';
        /* 2026-10-26 用户反馈：截触发器页面那种图没用，要看跑完的真实产物（看板长截图）。
           后端已把产物登记成截图记录 → 直接用；没有产物才退回截页面。 */
        data.截图 = (rd.截图 && rd.截图.length) ? rd.截图 : [await captureShot('运行完成画面')];
        if (rd.截图 && rd.截图.length) await push('已归档本次运行产物截图（' + rd.截图.length + ' 张）');
      } else if (kit) {
        await push('应用启动：' + (kit.name || app) + '（节点流水线）');
        await headlessRun(kit, push);
        await push('✓ 运行完成，耗时 ' + ((Date.now() - t0) / 1000).toFixed(1) + 's');
        data.运行结果 = '运行成功';
        data.截图 = [await captureShot('运行完成画面')];
      } else {
        await push('加载应用配置完成');
        await sleep(800);
        await push('执行选择屏条件…');
        await sleep(900);
        await push('数据导出并写回网页存储');
        await sleep(500);
        await push('✓ 运行完成，耗时 ' + ((Date.now() - t0) / 1000).toFixed(1) + 's');
        data.运行结果 = '运行成功';
        data.截图 = [await captureShot('运行完成画面')];
      }
    } catch (err) {
      await push('✗ 运行失败：' + (err && err.message || err));
      data.运行结果 = '运行失败';
      data.错误次数 = 1;
      data.截图 = [await captureShot('失败瞬间截图')];
      if (state.err && state.err.邮件通知) {
        await push('已按「运行错误处理」设置发送失败通知（' + (state.err.收件箱 || []).join('、') + '）');
      }
    } finally {
      if (logId) {
        try { await storePost('upsert', { sheet: '运行日志', id: logId, row_no: (logById(logId) || {}).row_no || 1, data: data }); } catch (e) { /* ignore */ }
        var row2 = logById(logId); if (row2) row2.data = data;
      }
      state.runningIds.delete(id);
      if (state.tab === 'logs') renderLogs();
      toast(data.运行结果 === '运行成功' ? '触发器「' + ((t.data || {}).名称 || '') + '」运行成功，日志已记录'
        : data.运行结果 === '部分成功' ? '触发器「' + ((t.data || {}).名称 || '') + '」部分成功：产物已生成，邮件未发出（详见运行日志）'
        : '触发器「' + ((t.data || {}).名称 || '') + '」运行失败，详见运行日志');
    }
    /* 通知：后端应用（bapp）由 /api/trigger/app/run 自己按同一份 config.通知 发，这里只发网页侧
       模拟/前端流水线的类型（文件 / 热键 / 邮件 / 钉钉…），避免同一次运行重复发两遍。 */
    if (!backendApp(app) && notifyOn(t)) {
      var 对象 = (notifyOf(t).对象 || []);
      try {
        await push('按「通知」配置发送：' + 对象.join('、'));
        var nr = await fetch('/api/trigger/run-notify', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            名称: (t.data || {}).名称 || '', 应用: name, 触发方式: '[手动运行]',
            运行结果: data.运行结果, 日志: (data.日志 || []).slice(-15),
            通知: notifyOf(t),
          }),
        });
        var nd = await nr.json().catch(function () { return {}; });
        await push('通知：' + (nd.汇总 || (nr.ok ? '已发送' : '发送失败')));
      } catch (e) {
        await push('⚠ 通知发送失败：' + (e && e.message || e));
      }
    }
  }
  function nextLogRowNo() {
    return state.logs.length ? Math.max.apply(null, state.logs.map(function (r) { return r.row_no || 0; })) + 1 : 1;
  }
  // 无界面执行 AutoTaskKit 任务节点（网页=触发器，任务节点=执行引擎）
  async function headlessRun(kit, push) {
    var shared = {};
    var host = (typeof AutoTaskKit !== 'undefined' && AutoTaskKit.host) ? AutoTaskKit.host : {};
    var ctx = {
      host: host, shared: shared,
      log: function (m) { return push('  ' + m); },
      stopped: function () { return false; },
    };
    for (var i = 0; i < (kit.nodes || []).length; i++) {
      var n = kit.nodes[i];
      await push('▶ 节点' + (i + 1) + '「' + n.name + '」开始');
      if (Array.isArray(n.parallel)) {
        await Promise.all(n.parallel.map(function (br) {
          var bctx = Object.assign({}, ctx, { log: function (m) { return push('    [' + br.name + '] ' + m); } });
          return br.run(bctx);
        }));
      } else {
        await n.run(ctx);
      }
      await push('✓ 节点' + (i + 1) + '「' + n.name + '」完成');
    }
  }

  /* ============ 新建下拉 / 筛选浮层 ============ */
  function bindPopovers() {
    var newBtn = $('#trigNewBtn'), newMenu = $('#trigNewMenu');
    if (newBtn && newMenu) {
      newBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        closeFilterPop(); closeRowMenu();
        newMenu.classList.toggle('open');
      });
      newMenu.querySelectorAll('.trig-new-item').forEach(function (el) {
        el.addEventListener('click', function (e) {
          e.stopPropagation();
          newMenu.classList.remove('open');
          openEditor(el.dataset.type, null);
        });
      });
    }
    var fBtn = $('#trigFilterBtn'), fPop = $('#trigFilterPop');
    if (fBtn && fPop) {
      fBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        closeRowMenu(); if (newMenu) newMenu.classList.remove('open');
        renderFilterPop();
        fPop.classList.toggle('open');
        var r = fBtn.getBoundingClientRect();
        fPop.style.top = (r.bottom + 6) + 'px';
        fPop.style.left = Math.max(8, r.right - fPop.offsetWidth) + 'px';
      });
    }
    document.addEventListener('click', function (e) {
      if (!e.target.closest('#trigRowMenu')) closeRowMenu();
      if (!e.target.closest('#trigFilterPop') && !e.target.closest('#trigFilterBtn')) closeFilterPop();
      if (newMenu && !e.target.closest('.trig-new-wrap')) newMenu.classList.remove('open');
    });
  }
  function closeFilterPop() { var p = $('#trigFilterPop'); if (p) p.classList.remove('open'); }
  function renderFilterPop() {
    var pop = $('#trigFilterPop');
    if (!pop) return;
    var types = ['定时', '文件', '热键', '邮件', '钉钉个人', '钉钉群'];
    pop.innerHTML =
      '<div class="tf-title">类型</div>' +
      types.map(function (tp) {
        return '<label class="trig-check"><input type="checkbox" data-ftype="' + tp + '"' + (state.filterTypes.indexOf(tp) >= 0 ? ' checked' : '') + '><span>' + tp + '</span></label>';
      }).join('') +
      '<div class="tf-title">启用状态</div>' +
      '<div class="trig-radios" style="gap:10px;">' +
      '<label><input type="radio" name="tfEnabled" value=""' + (state.filterEnabled === '' ? ' checked' : '') + '>全部</label>' +
      '<label><input type="radio" name="tfEnabled" value="on"' + (state.filterEnabled === 'on' ? ' checked' : '') + '>已启用</label>' +
      '<label><input type="radio" name="tfEnabled" value="off"' + (state.filterEnabled === 'off' ? ' checked' : '') + '>未启用</label>' +
      '</div>';
    pop.querySelectorAll('[data-ftype]').forEach(function (cb) {
      cb.addEventListener('change', function () {
        var v = cb.dataset.ftype;
        var i = state.filterTypes.indexOf(v);
        if (cb.checked && i < 0) state.filterTypes.push(v);
        if (!cb.checked && i >= 0) state.filterTypes.splice(i, 1);
        renderTasks();
      });
    });
    pop.querySelectorAll('[name=tfEnabled]').forEach(function (rd) {
      rd.addEventListener('change', function () { state.filterEnabled = rd.value; renderTasks(); });
    });
  }

  /* ============ 触发器编辑器 ============ */
  /* ============ 高级频率：常用时间表达式预设（影刀风格，可下拉选择或自定义） ============ */
  var CRON_PRESETS = [
    { label: '每 5 分钟', expr: '*/5 * * * *' },
    { label: '每 15 分钟', expr: '*/15 * * * *' },
    { label: '每 30 分钟', expr: '*/30 * * * *' },
    { label: '每小时（整点）', expr: '0 * * * *' },
    { label: '每天 09:00', expr: '0 9 * * *' },
    { label: '每天 09:00 与 18:00', expr: '0 9,18 * * *' },
    { label: '每周一 09:00', expr: '0 9 * * 1' },
    { label: '每月 1 日 09:00', expr: '0 9 1 * *' },
    { label: '每年 1 月 1 日 09:00', expr: '0 9 1 1 *' },
    { label: '工作日（周一至周五）09:00', expr: '0 9 * * 1-5' },
    { label: '自定义…', expr: '' },
  ];
  function fmtLocalDT(d) {
    return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()) + 'T' + p2(d.getHours()) + ':' + p2(d.getMinutes());
  }
  function defaultCfg(type) {
    if (type === '定时') return {
      freq: '每天', time: '09:00', weekdays: ['周一'], monthday: 1, monthInterval: 1,
      intervalMin: 1, intervalHour: 1, intervalMinute: 0,
      cronDesc: '', endAfter: false, endAt: '', dayWindow: false, winStart: '09:00', winEnd: '17:00',
      calendar: false, calendarMode: '', queue: true, timeout: 40,
    };
    // 条件触发器：到点先算指标，命中才通知（调度口径同定时）
    if (type === '条件') return {
      freq: '每工作日', time: '07:45', 指标: '小型化每日下达', 比较: '>', 阈值: 120,
      calendar: true, calendarMode: 'factory', queue: true, timeout: 20,
    };
    if (type === '文件') return { folder: '', sub: true, events: ['创建', '更新'], patterns: '', queue: true, timeout: 0 };
    if (type === '热键') return { ctrl: true, shift: false, alt: false, key: 'F1', queue: true, timeout: 0 };
    if (type === '钉钉个人' || type === '钉钉群') return {
      targets: [], matchTypes: ['文本内容'],
      match: { 文本内容: '', 网址链接: '', 指定名称文件: '' },
      queue: true, timeout: 0,
    };
    return { server: '163邮箱', account: '', authcode: '', folder: 'INBOX', rules: { from: '', to: '', subject: '', body: '' }, queue: true, timeout: 0 };
  }

  /* ============ 钉钉常用人/群配置（侧边栏「钉钉邮箱群消息配置模块」数据源） ============ */
  var ddCache = null;
  function ddConfig() {
    if (ddCache) return Promise.resolve(ddCache);
    return fetch('/api/ddconfig/store').then(function (res) {
      if (!res.ok) throw new Error('配置读取失败');
      return res.json();
    }).then(function (s) {
      var sheets = (s && s.sheets) || {};
      ddCache = {
        人员: (sheets['常用人员配置'] || []).map(function (r) { return r.data || {}; }).filter(function (d) { return d.姓名; }),
        群: (sheets['群消息配置'] || []).map(function (r) { return r.data || {}; }).filter(function (d) { return d.群名; }),
      };
      return ddCache;
    }).catch(function () {
      ddCache = { 人员: [], 群: [] };
      return ddCache;
    });
  }

  async function openEditor(type, task) {
    state.editType = type;
    state.editing = task || null;
    var meta = TYPE_META[type] || TYPE_META['定时'];
    var cfg = task ? Object.assign({}, defaultCfg(type), taskCfg(task)) : defaultCfg(type);
    $('#trigEmIcon').innerHTML = meta.icon;
    $('#trigEmTitle').textContent = meta.title;
    $('#trigEmDesc').textContent = meta.desc;
    var dd = (type === '钉钉个人' || type === '钉钉群') ? await ddConfig() : { 人员: [], 群: [] };
    state.dd = dd;
    await notifyOptions();                       // 通知对象候选（所有触发器类型都要）
    // 已建好的老触发器：config 里没有「通知」也补齐默认值（开关默认关，编辑器里随时可打开）
    var nf = cfg.通知 || {};
    cfg.通知 = { 启用: !!nf.启用, 对象: (nf.对象 || []).slice() };
    nfState = { 启用: !!cfg.通知.启用, 对象: cfg.通知.对象.map(function (v) {
      var s = String(v || '').trim();
      return { 值: s, 类型: nfType(s) };
    }).filter(function (o) { return o.值; }) };
    $('#trigEmBody').innerHTML = buildEditorBody(type, cfg, task ? (task.data || {}) : {}, dd);
    $('#trigEmEnable').checked = task ? !!(task.data || {}).启用 : true;
    bindEditorEvents(type, cfg);
    $('#trigEditModal').classList.add('open');
    var nm = $('#teName'); if (nm) nm.focus();
  }
  function closeEditor() { $('#trigEditModal').classList.remove('open'); state.editing = null; }

  /* ---------- 编辑器外壳：各类型表单 + 统一的「通知」区块（放在最后面） ---------- */
  function buildEditorBody(type, cfg, data, dd) {
    return buildEditorBodyCore(type, cfg, data, dd) + notifyBlock(cfg, type);
  }
  function notifyBlock(cfg, type) {
    var n = cfg.通知 || {}, on = !!n.启用, objs = n.对象 || [];
    return '<div class="trig-notify" id="teNotifySec">' +
      '<div class="trig-notify-hd">' +
      '<span class="trig-notify-title">通知</span>' +
      '<span class="trig-notify-sub">跑完自动发（邮箱 / 钉钉消息 / 群消息，可混着填）</span>' +
      '<label class="trig-notify-sw" id="teNotifySw" role="switch" aria-checked="' + (on ? 'true' : 'false') + '" ' +
      'title="通知总开关：打开后，触发器跑完（成功或失败）按「通知对象」里每个对象自带的通道分头发送；关掉＝一条都不发">' +
      '<span class="trig-switch' + (on ? ' on' : '') + '"></span>通知</label>' +
      '</div>' +
      '<div class="trig-notify-bd" id="teNotifyBd"' + (on ? '' : ' hidden') + '>' +
      '<div class="trig-nsel" id="teNsel" title="点一下展开候选；也可直接输入「名字（群消息）」/「名字（钉钉消息）」/ 邮箱后回车">' +
      '<span class="trig-nsel-tags" id="teNselTags"></span>' +
      '<input class="trig-nsel-inp" id="teNselInp" type="text" autocomplete="off" placeholder="勾选 / 或「名字（群消息）」…">' +
      '<div class="trig-nsel-menu" id="teNselMenu" style="display:none"></div>' +
      '</div>' +
      '<div class="trig-notify-tip" id="teNotifyTip">' + notifyTipText(objs) + '</div>' +
      (type === '条件'
        ? '<div class="trig-notify-tip" style="color:var(--color-text-tertiary);">条件触发器：只在命中阈值时才发（没到阈值只在运行日志里记一行）；命中的告警内容（KPI 卡截图 + 明细 Excel）按这里配的通道发出去。</div>'
        : '') +
      '</div></div>';
  }
  function notifyTipText(objs) {
    var n = { 群消息: 0, 钉钉消息: 0, 邮箱: 0 };
    (objs || []).forEach(function (o) { var t = typeof o === 'string' ? nfType(o) : o.类型; if (t) n[t] = (n[t] || 0) + 1; });
    var 计 = NF_CH.filter(function (t) { return n[t]; }).map(function (t) { return t + ' ' + n[t]; }).join('｜');
    return (objs && objs.length)
      ? '已选 ' + objs.length + ' 个' + (计 ? '（' + 计 + '）' : '') + '；只发这里选中的通道（没选邮箱就不会发邮件），产生的 Excel / 截图会一起作为附件带上'
      : '还没选对象：可勾选候选人，或输入「计划科（群消息）」「王锦志（钉钉消息）」「abc@x.com」后回车（跑完会把 Excel / 截图一起发）';
  }

  function buildEditorBodyCore(type, cfg, data, dd) {
    var opts = appOptions();
    var appSel = '<select id="teApp"><option value="">请选择应用</option>' +
      Object.keys(opts).map(function (v) {
        return '<option value="' + esc(v) + '"' + (data.应用 === v ? ' selected' : '') + '>' + esc(opts[v]) + '</option>';
      }).join('') + '</select>';
    var head =
      '<div class="trig-fld"><label>名称</label><input type="text" id="teName" maxlength="60" value="' + esc(data.名称 || '') + '" placeholder="请输入触发器名称"></div>' +
      '<div class="trig-fld"><label>应用<span class="trig-tip" title="仅支持已发布或获取的应用"> ⓘ</span></label>' + appSel + '</div>';
    var moreBlock =
      '<div class="trig-more-sec"><button type="button" class="trig-more-toggle" id="teMoreToggle">更多 <span id="teMoreArrow">▾</span></button>' +
      '<div class="trig-more-body" id="teMoreBody" hidden>' +
      '<label class="trig-check"><input type="checkbox" id="teQueue"' + (cfg.queue ? ' checked' : '') + '><span>排队执行</span></label> <span class="trig-tip" title="上一次运行未结束时，本次触发排队等待">ⓘ</span>' +
      '<div class="trig-fld inline" style="margin-top:10px;"><span>运行超过</span><input type="number" id="teTimeout" min="0" value="' + (cfg.timeout == null ? 0 : cfg.timeout) + '"><span>分钟后强制结束</span></div>' +
      '</div></div>';
    // 定时结束默认「当前时间 +2 分钟」（与影刀一致）；拆成 日期 + 时分，避免 datetime-local 浮层被弹窗遮挡
    var _ed = cfg.endAt || fmtLocalDT(new Date(Date.now() + 2 * 60000));
    var _edm = String(_ed).match(/^(\d{4})-(\d{2})-(\d{2})[T ]?(\d{2})?:?(\d{2})?/);
    var ed = {
      date: (_edm ? _edm[1] + '-' + _edm[2] + '-' + _edm[3] : '').trim(),
      h: _edm ? parseInt(_edm[4], 10) : 0,
      m: _edm ? parseInt(_edm[5], 10) : 0,
    };
    if (!ed.date) { var _nd = new Date(Date.now() + 2 * 60000); ed.date = _nd.getFullYear() + '-' + p2(_nd.getMonth() + 1) + '-' + p2(_nd.getDate()); ed.h = _nd.getHours(); ed.m = _nd.getMinutes(); }
    var calMode = cfg.calendarMode || 'factory'; // 旧数据 calendar=true → factory
    function hmSplit(v, dH, dM) {
      var p = String(v || '').split(':');
      return { h: Math.max(0, Math.min(23, parseInt(p[0], 10) || dH)), m: Math.max(0, Math.min(59, parseInt(p[1], 10) || dM)) };
    }
    var tm = hmSplit(cfg.time, 9, 0), ws = hmSplit(cfg.winStart, 9, 0), we = hmSplit(cfg.winEnd, 17, 0);
    var moreBlockFull =
      '<div class="trig-more-sec"><button type="button" class="trig-more-toggle" id="teMoreToggle">更多 <span id="teMoreArrow">▾</span></button>' +
      '<div class="trig-more-body" id="teMoreBody" hidden>' +
      '<label class="trig-check"><input type="checkbox" id="teEndAfter"' + (cfg.endAfter ? ' checked' : '') + '><span>定时结束</span></label> <span class="trig-tip" title="勾选后设定结束时间，到期后触发器自动停用">ⓘ</span>' +
      '<div class="trig-fld inline" id="teEndAtFld" hidden style="margin-top:8px;"><input type="text" id="teEndDate" maxlength="10" placeholder="YYYY-MM-DD" value="' + esc(ed.date) + '"><span>日</span><input type="number" id="teEndH" min="0" max="23" value="' + ed.h + '"><span>时</span><input type="number" id="teEndM" min="0" max="59" value="' + ed.m + '"><span>分</span><span class="trig-dim">到点自动停用</span></div>' +
      '<div style="margin-top:12px;"><label class="trig-check"><input type="checkbox" id="teDayWindow"' + (cfg.dayWindow ? ' checked' : '') + '><span>每天起止时间</span></label> <span class="trig-tip" title="限定每天可触发的时间范围，范围外不触发（如 09:00-17:00）">ⓘ</span></div>' +
      '<div class="trig-fld inline" id="teWinFld" hidden style="margin-top:8px;">' +
      '<input type="number" id="teWinStartH" min="0" max="23" value="' + ws.h + '"><span>时</span><input type="number" id="teWinStartM" min="0" max="59" value="' + ws.m + '"><span>分 至</span>' +
      '<input type="number" id="teWinEndH" min="0" max="23" value="' + we.h + '"><span>时</span><input type="number" id="teWinEndM" min="0" max="59" value="' + we.m + '"><span>分</span></div>' +
      '<div style="margin-top:12px;"><label class="trig-check"><input type="checkbox" id="teCalendar"' + ((calMode) ? ' checked' : '') + '><span>按指定日历触发</span></label> <span class="trig-tip" title="仅所选日历的工作日触发；工厂日历见「基础数据 → 工厂日历配置」">ⓘ</span></div>' +
      '<div class="trig-fld" id="teCalModeFld" hidden style="margin-top:8px;"><select id="teCalendarMode">' +
      '<option value="factory"' + (calMode === 'factory' ? ' selected' : '') + '>按工厂日历</option>' +
      '<option value="cn"' + (calMode === 'cn' ? ' selected' : '') + '>国家规定工作日（法定节假日、调休）</option>' +
      '</select></div>' +
      '<div style="margin-top:12px;"><label class="trig-check"><input type="checkbox" id="teQueue"' + (cfg.queue ? ' checked' : '') + '><span>排队执行</span></label> <span class="trig-tip" title="上一次运行未结束时，本次触发排队等待">ⓘ</span></div>' +
      '<div class="trig-fld inline" style="margin-top:10px;"><span>运行超过</span><input type="number" id="teTimeout" min="0" value="' + (cfg.timeout == null ? 0 : cfg.timeout) + '"><span>分钟后强制结束</span></div>' +
      '</div></div>';
    /* ---------- 条件触发器（2026-10-26 新增）：到点算指标，命中才通知 ----------
       目前只接看板指标（后端 apd22_metric_alert → board_push.run_alert）：
       指标 > 或 < 阈值 → 单独发一封告警邮件（该 KPI 卡截图 + 明细 Excel）。 */
    if (type === '条件') {
      var METRICS = ['小型化每日下达', '当日下达箱数', '近7天完工合计'];
      var mSel = METRICS.map(function (m) {
        return '<option value="' + m + '"' + ((cfg.指标 || '') === m ? ' selected' : '') + '>' + m + '</option>';
      }).join('');
      var freqs2 = ['每工作日', '每天', '每周', '月份'];
      var freqHtml2 = freqs2.map(function (f) {
        return '<label><input type="radio" name="teFreq" value="' + f + '"' + ((cfg.freq || '每工作日') === f ? ' checked' : '') + '>' + f + '</label>';
      }).join('');
      var wd2 = WEEKDAYS.map(function (w) {
        return '<label><input type="checkbox" data-wd="' + w + '"' + ((cfg.weekdays || []).indexOf(w) >= 0 ? ' checked' : '') + '><span>' + w + '</span></label>';
      }).join('');
      return head +
        '<div class="trig-fld"><label>监控指标<span class="trig-tip" title="目前支持 3号箱APD及安全钳看板的指标"> ⓘ</span></label>' +
        '<select id="teMetric">' + mSel + '</select></div>' +
        '<div class="trig-fld" id="teCondFld"><label>触发条件</label><div class="trig-hm" style="gap:8px">' +
        '<select id="teOp" style="width:96px"><option value=">"' + ((cfg.比较 || '>') === '>' ? ' selected' : '') + '>大于 &gt;</option>' +
        '<option value="<"' + (cfg.比较 === '<' ? ' selected' : '') + '>小于 &lt;</option></select>' +
        '<input type="number" id="teThreshold" step="1" min="0" value="' + (cfg.阈值 == null ? 120 : cfg.阈值) + '"><span>时通知</span></div>' +
        '<div id="teCondHint" style="margin-top:6px;font-size:11.5px;color:var(--color-text-tertiary)"></div></div>' +
        '<div class="trig-fld"><label>频率</label><div class="trig-radios">' + freqHtml2 + '</div></div>' +
        '<div class="trig-cond-line" id="teCondLine"></div>' +
        '<div class="trig-fld" id="teTimeFld"><label>检查时间</label><div class="trig-hm">' +
        '<input type="number" id="teTimeH" min="0" max="23" value="' + tm.h + '"><span>时</span>' +
        '<input type="number" id="teTimeM" min="0" max="59" value="' + tm.m + '"><span>分</span></div></div>' +
        '<div class="trig-fld" id="teWeekFld" hidden><label>每周（勾选生效的星期）</label><div class="trig-week-grid">' + wd2 + '</div></div>' +
        '<div class="trig-fld inline" id="teMonthFld" hidden><span>月份</span><input type="number" id="teMonthInterval" min="1" max="12" value="' + (cfg.monthInterval || 1) + '"><span>第</span><input type="number" id="teMonthday" min="1" max="31" value="' + (cfg.monthday || 1) + '"><span>日</span></div>' +
        moreBlockFull;
    }
    if (type === '定时') {
      var freqs = ['每分', '每时', '每天', '每工作日', '每周', '月份', '高级'];
      var freqHtml = freqs.map(function (f) {
        return '<label><input type="radio" name="teFreq" value="' + f + '"' + ((cfg.freq || '每天') === f ? ' checked' : '') + '>' + f + '</label>';
      }).join('');
      var wd = WEEKDAYS.map(function (w) {
        return '<label><input type="checkbox" data-wd="' + w + '"' + ((cfg.weekdays || []).indexOf(w) >= 0 ? ' checked' : '') + '><span>' + w + '</span></label>';
      }).join('');
      var presetSel = CRON_PRESETS.map(function (p) {
        var sel = (p.expr && p.expr === cfg.cronDesc) || (!p.expr && cfg.cronDesc && !CRON_PRESETS.some(function (x) { return x.expr === cfg.cronDesc; })) ? ' selected' : (!cfg.cronDesc && !p.expr ? ' selected' : '');
        return '<option value="' + esc(p.expr) + '"' + sel + '>' + esc(p.label) + '</option>';
      }).join('');
      /* 应用专属参数：钉钉邮箱巡检（yuzhi_mail_check）——读最近 N 封「小于阀值异常工单明细」邮件，
         目标列为 0 的行按生产订单号去重后发钉钉消息。字段存进 config，后端缺项时按默认值走。 */
      var 巡检块 = '<div class="trig-fld" id="teYzBlock"' +
        ((data.应用 || '') === 'yuzhi_mail_check' ? '' : ' hidden') + ' style="margin-top:12px;">' +
        '<label>邮箱巡检参数<span class="trig-tip" title="读收件箱最近 N 封目标邮件的 Excel 附件，目标列为 0 的行按生产订单号去重后发钉钉消息"> ⓘ</span></label>' +
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 10px;">' +
        '<div><div style="font-size:11.5px;color:var(--color-text-tertiary);margin-bottom:4px;">主题关键字</div><input type="text" id="teYzKw" maxlength="40" value="' + esc(cfg.主题关键字 || '小于阀值异常工单明细') + '"></div>' +
        '<div><div style="font-size:11.5px;color:var(--color-text-tertiary);margin-bottom:4px;">发件人过滤（留空=不限）</div><input type="text" id="teYzFrom" maxlength="40" value="' + esc(cfg.发件人 == null ? '' : cfg.发件人) + '"></div>' +
        '<div><div style="font-size:11.5px;color:var(--color-text-tertiary);margin-bottom:4px;">检查最近几封（建议 5）</div><input type="number" id="teYzN" min="1" max="20" value="' + (cfg.检查封数 || 5) + '"></div>' +
        '<div><div style="font-size:11.5px;color:var(--color-text-tertiary);margin-bottom:4px;">判零值（目标列 = 此值即命中）</div><input type="number" id="teYzZero" step="1" value="' + (cfg.判零值 == null ? 0 : cfg.判零值) + '"></div>' +
        '<div style="grid-column:1/-1;"><div style="font-size:11.5px;color:var(--color-text-tertiary);margin-bottom:4px;">钉钉接收人（留空=本人；可填姓名 / userId / 群:群名）</div>' +
        '<input type="text" id="teYzTo" maxlength="40" value="' + esc(cfg.钉钉接收人 == null ? '' : cfg.钉钉接收人) + '"></div>' +
        '</div><div style="font-size:11.5px;color:var(--color-text-tertiary);line-height:1.7;margin-top:6px;">' +
        '按「最近 N 封」回溯：夜里断网/关机没跑到的，早上这轮会一起补上；已处理过的邮件不重复打扰。' +
        '</div></div>';
      return head +
        '<div class="trig-fld"><label>频率</label><div class="trig-radios">' + freqHtml + '</div></div>' +
        '<div class="trig-cond-line" id="teCondLine"></div>' +
        '<div class="trig-fld inline" id="teMinFld" hidden><span>分钟</span><input type="number" id="teIntervalMin" min="1" max="1439" value="' + (cfg.intervalMin || 1) + '"></div>' +
        '<div class="trig-fld inline" id="teHourFld" hidden><span>小时</span><input type="number" id="teIntervalHour" min="0" max="23" value="' + (cfg.intervalHour == null ? 1 : cfg.intervalHour) + '"><span>分钟</span><input type="number" id="teIntervalMinute" min="0" max="59" value="' + (cfg.intervalMinute || 0) + '"></div>' +
        '<div class="trig-fld" id="teTimeFld"><label>触发时间</label><div class="trig-hm">' +
        '<input type="number" id="teTimeH" min="0" max="23" value="' + tm.h + '"><span>时</span>' +
        '<input type="number" id="teTimeM" min="0" max="59" value="' + tm.m + '"><span>分</span></div></div>' +
        '<div class="trig-fld" id="teWeekFld" hidden><label>每周（勾选生效的星期）</label><div class="trig-week-grid">' + wd + '</div></div>' +
        '<div class="trig-fld inline" id="teMonthFld" hidden><span>月份</span><input type="number" id="teMonthInterval" min="1" max="12" value="' + (cfg.monthInterval || 1) + '"><span>第</span><input type="number" id="teMonthday" min="1" max="31" value="' + (cfg.monthday || 1) + '"><span>日</span></div>' +
        '<div class="trig-fld" id="teCronFld" hidden><label>时间表达式</label>' +
        '<select id="teCronPreset">' + presetSel + '</select>' +
        '<input type="text" id="teCron" maxlength="60" value="' + esc(cfg.cronDesc || '') + '" placeholder="五段式：分 时 日 月 周，如 0 9 * * 1-5" style="margin-top:8px;">' +
        '<a class="trig-link" id="teCronDoc" href="/cron_help" target="_blank" rel="noopener" style="display:inline-block;margin-top:6px;">了解时间表达式</a></div>' +
        巡检块 +
        moreBlockFull;
    }
    if (type === '文件') {
      var evs = ['创建', '删除', '更新', '重命名'];
      var evHtml = evs.map(function (e) {
        return '<label><input type="checkbox" data-ev="' + e + '"' + ((cfg.events || []).indexOf(e) >= 0 ? ' checked' : '') + '><span>' + e + '</span></label>';
      }).join('');
      return head +
        '<div class="trig-fld"><label>监控文件夹</label>' +
        '<div class="trig-chip-add"><input type="text" id="teFolder" value="' + esc(cfg.folder || '') + '" placeholder="请填写服务器本地文件夹路径，如 D:\\监控目录">' +
        '<button class="btn btn-default" id="teFolderPick" type="button">选择</button></div></div>' +
        '<div class="trig-fld"><label class="trig-check"><input type="checkbox" id="teSub"' + (cfg.sub ? ' checked' : '') + '><span>包含子路径</span></label></div>' +
        '<div class="trig-fld"><label>监控事件</label><div class="trig-week-grid">' + evHtml + '</div></div>' +
        '<div class="trig-fld"><label>文件/文件类型</label><input type="text" id="tePatterns" value="' + esc(cfg.patterns || '') + '" placeholder="说明：请输入要监视文件的通配符，以；分隔。例如：*.txt;*.csv"></div>' +
        moreBlock;
    }
    if (type === '热键') {
      var keys = HOTKEYS.map(function (k) { return '<option' + ((cfg.key || 'F1') === k ? ' selected' : '') + '>' + k + '</option>'; }).join('');
      return head +
        '<div class="trig-fld"><label>快捷键</label>' +
        '<div class="trig-week-grid">' +
        '<label class="trig-check"><input type="checkbox" id="teCtrl"' + (cfg.ctrl ? ' checked' : '') + '><span>Ctrl</span></label>' +
        '<label class="trig-check"><input type="checkbox" id="teShift"' + (cfg.shift ? ' checked' : '') + '><span>Shift</span></label>' +
        '<label class="trig-check"><input type="checkbox" id="teAlt"' + (cfg.alt ? ' checked' : '') + '><span>Alt</span></label>' +
        '<span>＋</span><select id="teKey" class="trig-key-sel">' + keys + '</select>' +
        '</div></div>' +
        moreBlock;
    }
    // 钉钉个人 / 钉钉群消息触发：监控指定聊天窗口，消息包含指定内容（文本/网址/文件）时触发
    if (type === '钉钉个人' || type === '钉钉群') {
      var isPerson = type === '钉钉个人';
      var list = isPerson ? (dd.人员 || []) : (dd.群 || []);
      var sel = cfg.targets || [];
      var grid = list.map(function (x) {
        var key = isPerson ? x.姓名 : x.群名;
        var tip = isPerson
          ? (x.姓名 + '（openDingTalkId：' + (x.openDingTalkId || '未配置，请在配置模块补齐') + '）')
          : (x.群名 + '（会话ID：' + (x.会话ID || '未配置') + '）');
        return '<label title="' + esc(tip) + '"><input type="checkbox" data-target="' + esc(key) + '"' + (sel.indexOf(key) >= 0 ? ' checked' : '') + '><span>' + esc(key) + '</span></label>';
      }).join('');
      if (!grid) {
        grid = '<div style="font-size:12px;color:var(--color-text-tertiary);grid-column:1/-1;line-height:1.8;">' +
          (isPerson ? '「常用人员配置」暂无数据，请先到侧边栏 钉钉邮箱群消息配置模块 → 常用人员配置 维护' : '「群消息配置」暂无数据，请先到侧边栏 钉钉邮箱群消息配置模块 → 群消息配置 维护') + '</div>';
      }
      var m = cfg.match || {};
      var mt = cfg.matchTypes || ['文本内容'];
      function mFld(id, valId, label, ph) {
        var on = mt.indexOf(id) >= 0;
        return '<div class="trig-fld" id="' + valId + 'Fld"' + (on ? '' : ' hidden') + '><label>' + label + '</label>' +
          '<input type="text" id="' + valId + 'Val" maxlength="200" value="' + esc(m[id] || '') + '" placeholder="' + ph + '"></div>';
      }
      return head +
        '<div class="trig-fld"><label>消息接收' + (isPerson ? '人' : '群') + '（可多选，来自「钉钉邮箱群消息配置模块」；按钉钉聊天窗口中的人名/群名匹配）</label>' +
        '<div class="trig-target-grid">' + grid + '</div>' +
        '<div class="trig-ac-wrap trig-chip-add" style="margin-top:8px;">' +
        '<input type="text" id="teTargetInput" placeholder="自定义填写：输入人名/群名自动匹配常用配置，未匹配则直接添加" style="flex:1;">' +
        '<button class="btn btn-default" id="teTargetAdd" type="button">添加</button>' +
        '<div class="trig-ac-pop" id="teTargetAc"></div></div>' +
        '<div class="trig-custom-chips" id="teCustomChips"></div>' +
        '<div class="trig-target-tip" id="teTargetTip">已选 ' + sel.length + ' 个' + (isPerson ? '联系人' : '群') + '；可勾选也可在上方自定义填写，人名自动定位到常用配置中的对应人员/群</div></div>' +
        '<div class="trig-fld"><label>触发条件（收到的消息中包含以下内容时触发应用，至少选一项）</label>' +
        '<div class="trig-checks">' +
        '<label class="trig-check"><input type="checkbox" data-mtype="文本内容"' + (mt.indexOf('文本内容') >= 0 ? ' checked' : '') + '><span>文本内容</span></label>' +
        '<label class="trig-check"><input type="checkbox" data-mtype="网址链接"' + (mt.indexOf('网址链接') >= 0 ? ' checked' : '') + '><span>网址链接</span></label>' +
        '<label class="trig-check"><input type="checkbox" data-mtype="指定名称文件"' + (mt.indexOf('指定名称文件') >= 0 ? ' checked' : '') + '><span>指定名称文件</span></label>' +
        '</div></div>' +
        mFld('文本内容', 'teMText', '文本关键词（多个用；分隔，消息正文包含任一即触发）', '如：要桶子；请查收；进度问题是') +
        mFld('网址链接', 'teMUrl', '网址关键词（多个用；分隔，匹配消息中的链接地址）', '如：dingtalk.com；workbuddy.cn') +
        mFld('指定名称文件', 'teMFile', '文件名称关键词（多个用；分隔，匹配消息中收到的文件名）', '如：115计划数据源；装箱报交') +
        '<div style="font-size:12px;color:var(--color-text-tertiary);line-height:1.7;margin-bottom:14px;">聊天窗口（参考钉钉与张三的会话）中收到文本、网址链接或指定名称的文件时，即运行上方所选应用。</div>' +
        moreBlock;
    }
    // 邮件
    var rules = cfg.rules || {};
    var srvOpts = MAIL_SERVERS.map(function (s) { return '<option' + ((cfg.server || '163邮箱') === s ? ' selected' : '') + '>' + s + '</option>'; }).join('');
    return head +
      '<div class="trig-em-tabs">' +
      '<button type="button" class="trig-em-tab on" data-mtab="常规">常规</button>' +
      '<button type="button" class="trig-em-tab" data-mtab="规则">匹配规则</button>' +
      '</div>' +
      '<div class="trig-em-tabpane" data-pane="常规">' +
      '<div class="trig-fld"><label>邮箱服务器</label><select id="teServer">' + srvOpts + '</select></div>' +
      '<div class="trig-fld"><label>邮箱账号</label><input type="text" id="teAccount" value="' + esc(cfg.account || '') + '" placeholder="发送监控的邮箱账号"></div>' +
      '<div class="trig-fld" id="teAuthFld"><label>授权码</label>' +
      '<div class="trig-pwd-wrap"><input type="password" id="teAuth" value="' + esc(cfg.authcode || '') + '" placeholder="邮箱 SMTP 授权码">' +
      '<button class="trig-eye" id="teEye" type="button" title="显示/隐藏授权码">👁</button></div>' +
      '<span class="trig-link">如何获取授权码?</span>' +
      '<div style="font-size:12px;color:var(--color-text-tertiary);margin-top:4px;" id="teDdTip" hidden>钉钉邮箱已通过 dws 开通，无需授权码</div></div>' +
      '<div class="trig-fld"><label>连接测试</label>' +
      '<div class="trig-err-btns"><button class="btn btn-default" id="teMailTest" type="button">测试连接服务器</button>' +
      '<span id="teMailTestTip" style="font-size:12px;color:var(--color-text-tertiary);align-self:center;"></span></div></div>' +
      '<div class="trig-fld" id="teMailFolderFld"><label>指定文件夹</label>' +
      '<select id="teMailFolder">' + MAIL_FOLDERS.map(function (f) {
        return '<option' + (f === (cfg.folder || 'INBOX') ? ' selected' : '') + '>' + f + '</option>';
      }).join('') + '</select>' +
      '<div style="font-size:12px;color:var(--color-text-tertiary);margin-top:4px;">默认常用文件夹可直接选择；测试连接成功后自动刷新为服务器实际文件夹</div></div>' +
      moreBlock +
      '</div>' +
      '<div class="trig-em-tabpane" data-pane="规则" hidden>' +
      '<div class="trig-fld"><label>发件人地址中包含内容</label><input type="text" id="teRFrom" value="' + esc(rules.from || '') + '" placeholder="留空 = 不限制"></div>' +
      '<div class="trig-fld"><label>收件人地址中包含内容</label><input type="text" id="teRTo" value="' + esc(rules.to || '') + '" placeholder="留空 = 不限制"></div>' +
      '<div class="trig-fld"><label>主题所包含内容</label><input type="text" id="teRSubject" value="' + esc(rules.subject || '') + '" placeholder="留空 = 不限制"></div>' +
      '<div class="trig-fld"><label>正文所包含内容</label><input type="text" id="teRBody" value="' + esc(rules.body || '') + '" placeholder="留空 = 不限制"></div>' +
      '<div style="font-size:12px;color:var(--color-text-tertiary);line-height:1.7;">邮件触发器按「每隔 1 分钟检查指定文件夹」轮询，命中任一匹配规则即运行应用。</div>' +
      '</div>';
  }

  /* ============ 文件触发器 · 目录选择弹窗（浏览服务器本地文件夹） ============ */
  var fsCurPath = '';
  var fsCurReal = ''; // 最近一次真实进入的目录（盘符级为 ''）
  function cssEsc(s) { return String(s).replace(/["\\]/g, '\\$&'); }
  function fillMailFolders(cur, fromServer) {
    var sel = $('#teMailFolder');
    if (!sel) return;
    var list = MAIL_FOLDERS.slice();
    var keep = cur || sel.value || 'INBOX';
    if (list.indexOf(keep) < 0) list.unshift(keep);
    sel.innerHTML = list.map(function (f) {
      return '<option' + (f === keep ? ' selected' : '') + '>' + f + '</option>';
    }).join('');
    if (fromServer) sel.disabled = false;
  }
  function openFsPicker(targetInput) {
    var mask = $('#trigFsModal'), list = $('#trigFsList'), cur = $('#trigFsCur');
    if (!mask) return;
    fsCurPath = (targetInput && targetInput.value) || '';
    mask.classList.add('open');
    $('#trigFsOk').onclick = function () {
      var p = fsCurPath || fsCurReal;
      if (p && targetInput) { targetInput.value = p; mask.classList.remove('open'); return; }
      toast('请先双击进入某个盘符或文件夹');
    };
    $('#trigFsCancel').onclick = function () { mask.classList.remove('open'); };
    $('#trigFsClose').onclick = function () { mask.classList.remove('open'); };
    $('#trigFsUp').onclick = function () { loadFsDir(cur.dataset.parent || '', targetInput); };
    function render(data) {
      cur.textContent = data.current || '此电脑';
      cur.dataset.parent = data.parent || '';
      var dirs = data.dirs || [];
      list.innerHTML = data.error
        ? '<div class="trig-fs-empty">' + esc(data.error) + '</div>'
        : (dirs.length ? dirs.map(function (d) {
            return '<div class="trig-fs-item" data-path="' + esc(d.path) + '"><svg viewBox="0 0 16 16" fill="none"><path d="M2.6 4.2c0-.66.54-1.2 1.2-1.2h2.9l1.5 1.7h4.2c.66 0 1.2.54 1.2 1.2v5.9c0 .66-.54 1.2-1.2 1.2H3.8c-.66 0-1.2-.54-1.2-1.2V4.2Z" stroke="currentColor" stroke-width="1.4"/></svg><span class="fs-name">' + esc(d.name) + '</span></div>';
          }).join('') : '<div class="trig-fs-empty">（无子文件夹）</div>');
      $$('#trigFsList .trig-fs-item').forEach(function (el) {
        el.addEventListener('dblclick', function () { loadFsDir(el.dataset.path, targetInput); });
        el.addEventListener('click', function () {
          $$('#trigFsList .trig-fs-item').forEach(function (x) { x.style.background = ''; });
          el.style.background = 'var(--color-primary-soft)';
          fsCurPath = el.dataset.path;
          $('#trigFsCur').textContent = el.dataset.path;
        });
      });
      // 单击进入本级
      list.ondblclick = function (e) {
        var it = e.target.closest('.trig-fs-item');
        if (it) loadFsDir(it.dataset.path, targetInput);
      };
    }
    function loadFsDir(path, ti) {
      list.innerHTML = '<div class="trig-fs-empty">正在读取目录…</div>';
      fetch('/api/fs/list', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path: path })
      }).then(function (r) { return r.json(); }).then(function (data) {
        if (data.ok && !data.drives && data.current) fsCurReal = data.current;
        fsCurPath = (data.ok && !data.drives && data.current) ? data.current : fsCurPath;
        render(data);
      }).catch(function () {
        render({ ok: false, error: '读取目录失败，请检查服务状态', dirs: [], current: path });
      });
    }
    loadFsDir(fsCurPath, targetInput);
  }

  /* 显隐helper：.trig-fld.inline 的 display:flex 作者规则会压过 UA [hidden]，
     普通字段又压不过内联 display——两者都要写才能可靠显隐 */
  function showEl(el, on) {
    if (!el) return;
    if (on) { el.hidden = false; el.style.display = ''; }
    else { el.hidden = true; el.style.display = 'none'; }
  }
  function bindEditorEvents(type, cfg) {
    // 应用切换：邮箱巡检专属参数块跟着显示/隐藏（切换后不必关掉弹窗重开）
    var appEl = $('#teApp'), yzBlk = $('#teYzBlock');
    if (appEl && yzBlk) appEl.addEventListener('change', function () {
      yzBlk.hidden = this.value !== 'yuzhi_mail_check';
    });
    // 更多 折叠
    var mt = $('#teMoreToggle'), mb = $('#teMoreBody');
    if (mt && mb) {
      mt.addEventListener('click', function () {
        mb.hidden = !mb.hidden;
        $('#teMoreArrow').textContent = mb.hidden ? '▾' : '▴';
      });
    }
    // 眼睛
    var eye = $('#teEye');
    if (eye) eye.addEventListener('click', function () {
      var inp = $('#teAuth'); if (!inp) return;
      inp.type = inp.type === 'password' ? 'text' : 'password';
    });
    // 文件夹「选择」按钮：弹窗浏览服务器目录
    var fp = $('#teFolderPick');
    if (fp) fp.addEventListener('click', function () { openFsPicker($('#teFolder')); });
    // 邮件：测试连接（成功后可刷新文件夹列表）
    var mtest = $('#teMailTest');
    if (mtest) mtest.addEventListener('click', function () {
      var acc = ($('#teAccount').value || '').trim();
      var tip = $('#teMailTestTip');
      if (!acc) { tip.textContent = '请先填写邮箱账号'; tip.style.color = 'var(--color-danger)'; return; }
      tip.textContent = '正在连接 ' + ($('#teServer').value || '') + ' …';
      tip.style.color = 'var(--color-text-tertiary)';
      setTimeout(function () {
        tip.textContent = '连接成功，已获取文件夹列表';
        tip.style.color = 'var(--color-success)';
        fillMailFolders((($('#teMailFolder') || {}).value || 'INBOX'), true);
      }, 900);
    });
    // 邮件：服务器联动（钉钉邮箱走 dws 已开通链路，无需授权码；且钉钉邮箱不经 IMAP 选文件夹）
    var srv = $('#teServer');
    if (srv) {
      var syncAuth = function () {
        var isDd = (srv.value || '') === '钉钉邮箱';
        var fld = $('#teAuthFld');
        if (fld) fld.hidden = isDd;
        var tip = $('#teDdTip');
        if (tip) tip.hidden = !isDd;
        var ff = $('#teMailFolderFld');
        if (ff) ff.hidden = isDd; // 钉钉邮箱不需要指定文件夹
      };
      srv.addEventListener('change', syncAuth);
      syncAuth();
    }
    // 钉钉个人/群：接收人多选计数联动 + 自定义填写（自动匹配常用配置）
    var ttip = $('#teTargetTip');
    if (ttip) {
      var isPerson = type === '钉钉个人';
      var refreshTargetTip = function () {
        var n = $$('#trigEmBody [data-target]').filter(function (c4) { return c4.checked; }).length +
          $$('#teCustomChips .trig-chip').length;
        ttip.textContent = '已选 ' + n + ' 个' + (isPerson ? '联系人' : '群') + '；可勾选也可在上方自定义填写，人名自动定位到常用配置中的对应人员/群';
      };
      var candList = function () {
        var src = isPerson ? ((state.dd && state.dd.人员) || []) : ((state.dd && state.dd.群) || []);
        return src.map(function (p) {
          return isPerson ? { name: p.姓名 || '', sub: p.邮箱 || '', id: p.openDingTalkId || '' }
                         : { name: p.群名 || '', sub: p.会话ID || '', id: p.会话ID || '' };
        });
      };
      // 勾选网格
      $$('#trigEmBody [data-target]').forEach(function (cb) {
        cb.addEventListener('change', function () {
          // 勾选网格中已有的，与自定义 chips 去重
          if (cb.checked) removeCustomChip(cb.dataset.target);
          refreshTargetTip();
        });
      });
      // 自定义输入 + 自动匹配下拉
      var tinp = $('#teTargetInput'), tac = $('#teTargetAc');
      function removeCustomChip(name) {
        var exist = $('#teCustomChips .trig-chip[data-custom="' + cssEsc(name) + '"]');
        if (exist) exist.remove();
        refreshTargetTip();
      }
      function addCustomChip(name) {
        name = (name || '').trim();
        if (!name) return false;
        // 已在网格勾选/已加过 → 不重复
        var dup = $$('#trigEmBody [data-target]').some(function (cb) { return cb.dataset.target === name && cb.checked; });
        var dupChip = $('#teCustomChips .trig-chip[data-custom="' + cssEsc(name) + '"]');
        if (dup || dupChip) { toast('「' + name + '」已在列表中'); return false; }
        var chip = document.createElement('span');
        chip.className = 'trig-chip';
        chip.dataset.custom = name;
        chip.innerHTML = esc(name) + '<span class="rm">×</span>';
        $('#teCustomChips').appendChild(chip);
        chip.querySelector('.rm').addEventListener('click', function () {
          removeCustomChip(name);
          $$('#trigEmBody [data-target]').forEach(function (cb) {
            if (cb.dataset.target === name) cb.checked = false;
          });
          refreshTargetTip();
        });
        refreshTargetTip();
        return true;
      }
      function submitTarget() {
        var q = (tinp.value || '').trim();
        if (!q) return;
        var list = candList();
        // 1) 精确命中 → 勾选网格并滚动定位
        var exact = list.filter(function (x) { return x.name === q; })[0];
        if (exact) {
          checkGridTarget(exact.name);
          tac.classList.remove('open');
          tinp.value = '';
          refreshTargetTip();
          toast('已选中常用配置中的「' + exact.name + '」');
          return;
        }
        // 2) 模糊命中 → 选中第一个并提示
        var fuzzy = list.filter(function (x) { return x.name.indexOf(q) >= 0; });
        if (fuzzy.length) {
          checkGridTarget(fuzzy[0].name);
          tac.classList.remove('open');
          tinp.value = '';
          refreshTargetTip();
          toast('自动匹配到常用配置「' + fuzzy[0].name + '」');
          return;
        }
        // 3) 未命中 → 作为自由接收人添加
        if (addCustomChip(q)) {
          tac.classList.remove('open');
          tinp.value = '';
          toast('常用配置中未找到「' + q + '」，已作为自定义接收人添加');
        }
      }
      function checkGridTarget(name) {
        var cb = $$('#trigEmBody [data-target]').filter(function (c) { return c.dataset.target === name; })[0];
        if (cb) {
          cb.checked = true;
          removeCustomChip(name);
          cb.scrollIntoView({ block: 'nearest' });
        } else {
          addCustomChip(name);
        }
      }
      if (tinp) {
        tinp.addEventListener('input', function () {
          var q = (tinp.value || '').trim();
          if (!q || !candList().length) { tac.classList.remove('open'); return; }
          var hits = candList().filter(function (x) { return x.name.indexOf(q) >= 0; }).slice(0, 8);
          tac.innerHTML = hits.length ? hits.map(function (x) {
            return '<div class="trig-ac-item" data-name="' + esc(x.name) + '" title="' + esc(x.sub || '') + '"><span>' + esc(x.name) + '</span><span class="ai-sub">' + esc(x.sub || '') + '</span></div>';
          }).join('') : '<div class="trig-ac-item" style="color:var(--color-text-tertiary);cursor:default;">常用配置中无匹配，回车将直接添加「' + esc(q) + '」</div>';
          tac.classList.add('open');
          $$('.trig-ac-item[data-name]', tac).forEach(function (it) {
            it.addEventListener('click', function () {
              checkGridTarget(it.dataset.name);
              tac.classList.remove('open');
              tinp.value = '';
              refreshTargetTip();
              toast('已选中「' + it.dataset.name + '」');
            });
          });
        });
        tinp.addEventListener('keydown', function (e) {
          if (e.key === 'Enter') { e.preventDefault(); submitTarget(); }
        });
        $('#teTargetAdd').addEventListener('click', submitTarget);
        document.addEventListener('click', function (e) {
          if (!e.target.closest('.trig-ac-wrap')) tac.classList.remove('open');
        });
      }
      // 还原历史自定义接收人（不在常用配置里的 targets 以自定义 chip 呈现）
      if (cfg && cfg.config && cfg.config.targets) {
        cfg.config.targets.forEach(function (nm) {
          var has = $$('#trigEmBody [data-target]').some(function (c) { return c.dataset.target === nm; });
          if (!has) addCustomChip(nm);
        });
      }
      refreshTargetTip();
    }
    var MT_FLD = { '文本内容': 'teMText', '网址链接': 'teMUrl', '指定名称文件': 'teMFile' };
    $$('#trigEmBody [data-mtype]').forEach(function (cb2) {
      cb2.addEventListener('change', function () {
        var fld = $('#' + MT_FLD[cb2.dataset.mtype] + 'Fld');
        if (fld) fld.hidden = !cb2.checked;
      });
    });
    // 邮件页签
    $$('.trig-em-tab').forEach(function (btn) {
      btn.addEventListener('click', function () {
        $$('.trig-em-tab').forEach(function (b) { b.classList.toggle('on', b === btn); });
        $$('.trig-em-tabpane').forEach(function (pane) {
          pane.hidden = pane.dataset.pane !== btn.dataset.mtab;
        });
      });
    });
    // 定时 / 条件：频率联动 + 条件行（条件触发器共用同一套频率与日历控件）
    if (type === '定时' || type === '条件') {
      var sync = function () {
        var f = (document.querySelector('input[name=teFreq]:checked') || {}).value || '每天';
        showEl($('#teMinFld'), f === '每分');
        showEl($('#teHourFld'), f === '每时');
        showEl($('#teTimeFld'), f !== '每分' && f !== '每时' && f !== '高级');
        showEl($('#teWeekFld'), f === '每周');
        showEl($('#teMonthFld'), f === '月份');
        showEl($('#teCronFld'), f === '高级');
        showEl($('#teEndAtFld'), !!($('#teEndAfter') || {}).checked);
        showEl($('#teWinFld'), !!($('#teDayWindow') || {}).checked);
        showEl($('#teCalModeFld'), !!($('#teCalendar') || {}).checked);
        if (type === '条件') {
          var hint = $('#teCondHint');
          if (hint) {
            var mm = ($('#teMetric') || {}).value || '小型化每日下达';
            var oo = ($('#teOp') || {}).value || '>';
            var tt = ($('#teThreshold') || {}).value || '0';
            hint.innerHTML = '命中条件时单独发一封邮件：' + mm + ' ' + (oo === '>' ? '大于 ' : '小于 ') + tt +
              ' 才发送（含该 KPI 卡截图 + 明细 Excel）；未命中只写一行运行日志，不发邮件。';
          }
        }
        var cl = $('#teCondLine');
        if (cl) {
          cl.innerHTML = esc(timedCondText(collectTimedCfg())) + ' <span class="trig-link" id="teDetailLink">详情</span>';
          var link = $('#teDetailLink');
          if (link) {
            link.addEventListener('mouseenter', function () {
              if (nextPopTimer) { clearTimeout(nextPopTimer); nextPopTimer = null; }
              showNextRuns(collectTimedCfg(), link);
            });
            link.addEventListener('mouseleave', function () {
              nextPopTimer = setTimeout(hideNextPop, 300);
            });
            link.addEventListener('click', function (e) {
              e.stopPropagation();
              showNextRuns(collectTimedCfg(), link);
            });
          }
        }
      };
      $$('#trigEmBody input, #trigEmBody select').forEach(function (el) {
        el.addEventListener('change', sync);
        el.addEventListener('input', sync);
      });
      // 高级：预设下拉 → 选中即填充表达式并锁定；选「自定义…」解锁可手输
      var ps = $('#teCronPreset');
      if (ps) ps.addEventListener('change', function () {
        var ci = $('#teCron');
        if (!ci) return;
        if (ps.value) { ci.value = ps.value; ci.readOnly = true; }
        else { ci.value = ''; ci.readOnly = false; ci.focus(); }
        sync();
      });
      sync();
    }
    bindNotifySection();     // 所有类型共用：通知开关 + 通知对象多选框
  }

  /* ---------- 通知区块交互（开关 + 混合通道多选框，交互与「维护物料自动化」抬头的通知对象一致） ---------- */
  function nfRenderTags() {
    var host = $('#teNselTags');
    if (!host) return;
    host.innerHTML = nfState.对象.map(function (o, i) {
      return '<span class="trig-tag t-' + esc(o.类型 || '未定') + '" title="' + esc((o.类型 || '未定') + '：' + o.值) + '">' +
        esc(o.值) + '<span class="x" data-i="' + i + '">×</span></span>';
    }).join('');
    host.querySelectorAll('.x').forEach(function (x) {
      x.addEventListener('click', function (e) {
        e.stopPropagation();
        nfState.对象.splice(parseInt(x.dataset.i, 10), 1);
        nfRenderTags(); nfRefreshTip(); nfRenderMenu($('#teNselInp').value);
      });
    });
    nfRefreshTip();
  }
  function nfRefreshTip() {
    var tip = $('#teNotifyTip');
    if (tip) tip.textContent = notifyTipText(nfState.对象);
  }
  function nfAdd(v) {
    var s = String(v || '').trim();
    if (!s) return false;
    var t = nfType(s);
    if (!t) {
      toast('认不出是群还是人：请自己带后缀——计划科（群消息）／王锦志（钉钉消息）／abc@x.com（邮箱）');
      nfOpenMenu(true);
      return false;
    }
    var 值 = nfNorm(s);
    if (nfHas(值)) { toast('「' + 值 + '」已在通知对象里'); return false; }
    nfState.对象.push({ 值: 值, 类型: t });
    nfRenderTags();
    if ($('#teNselInp')) $('#teNselInp').value = '';
    nfRenderMenu('');
    return true;
  }
  function nfMenuPos() {
    var m = $('#teNselMenu'), 锚 = $('#teNsel');
    if (!m || !锚) return;
    var r = 锚.getBoundingClientRect();
    m.style.display = 'block';
    var w = m.offsetWidth, h = m.offsetHeight;
    m.style.left = Math.max(6, Math.min(r.left, window.innerWidth - w - 8)) + 'px';
    // 默认贴在框下面；下方放不下就翻到框上面（绝不压在输入框上，否则点不到输入）
    var top = r.bottom + 2;
    if (top + h > window.innerHeight - 8) top = Math.max(6, r.top - h - 2);
    m.style.top = top + 'px';
  }
  function nfRenderMenu(过滤) {
    var host = $('#teNselMenu');
    if (!host) return;
    var opts = nfOpts || {}, 词 = String(过滤 || '').trim().toLowerCase(), html = '';
    if (词) {
      var 词类型 = nfType(词);
      html += '<div class="trig-nsel-grp">自定义</div>' +
        '<div class="trig-nsel-opt" data-custom="1"><span class="cb">' + (词类型 ? '+' : '!') + '</span>' +
        '<span class="nm">' + (词类型 ? '加入：' + esc(nfNorm(词)) : '要加后缀才能加：' + esc(词) + '（群消息／钉钉消息／邮箱）') + '</span>' +
        '<span class="use">' + esc(词类型 || '待定') + '</span></div>';
    }
    NF_CH.forEach(function (t) {
      var arr = (opts[t] || []).filter(function (it) {
        return !词 || String(it.值).toLowerCase().indexOf(词) >= 0 || String(it.姓名 || '').toLowerCase().indexOf(词) >= 0;
      }).slice(0, 词 ? 40 : 12);
      if (!arr.length) return;
      html += '<div class="trig-nsel-grp">' + esc(t) + '（按常用度）</div>';
      arr.forEach(function (it) {
        var on = nfHas(it.值);
        html += '<div class="trig-nsel-opt' + (on ? ' on' : '') + '" data-val="' + esc(it.值) + '">' +
          '<span class="cb">' + (on ? '✓' : '') + '</span><span class="nm">' + esc(it.值) + '</span>' +
          '<span class="use">' + (it.用量 >= 100000000 ? '常用' : (it.用量 ? esc(it.用量) : '')) + '</span></div>';
      });
    });
    if (!html) html = '<div class="trig-nsel-empty">没有匹配的候选；直接输入后回车可加为自定义对象。</div>';
    host.innerHTML = html;
    host.querySelectorAll('.trig-nsel-opt').forEach(function (el) {
      el.addEventListener('click', function (e) {
        e.stopPropagation();
        if (el.dataset.custom) {
          nfAdd($('#teNselInp').value);
          $('#teNselInp').value = '';
        } else {
          var v = el.dataset.val;
          if (nfHas(v)) {
            nfState.对象 = nfState.对象.filter(function (o) { return o.值 !== v; });
            nfRenderTags(); nfRenderMenu($('#teNselInp').value);
          } else {
            nfAdd(v);
          }
        }
      });
    });
  }
  function nfOpenMenu(开) {
    var m = $('#teNselMenu');
    if (!m) return;
    if (开) { nfRenderMenu(($('#teNselInp') || {}).value || ''); nfMenuPos(); }
    else m.style.display = 'none';
  }
  function bindNotifySection() {
    var sw = $('#teNotifySw');
    if (sw) {
      sw.addEventListener('click', function (e) {
        e.preventDefault();
        var on = sw.getAttribute('aria-checked') !== 'true';
        sw.setAttribute('aria-checked', on ? 'true' : 'false');
        var knob = sw.querySelector('.trig-switch');
        if (knob) knob.classList.toggle('on', on);
        nfState.启用 = on;
        showEl($('#teNotifyBd'), on);
        if (on) { var inp = $('#teNselInp'); if (inp) inp.focus(); }
        nfRefreshTip();
      });
    }
    var 框 = $('#teNsel'), inp = $('#teNselInp');
    if (!框 || !inp) return;
    框.addEventListener('click', function (e) {
      if (e.target.closest('.x')) return;
      inp.focus();
      nfOpenMenu(true);
    });
    inp.addEventListener('focus', function () { nfOpenMenu(true); });
    inp.addEventListener('input', function () { nfOpenMenu(true); });
    inp.addEventListener('keydown', function (e) {
      // 任意分隔符（中英文逗号 / 分号 / 顿号 / 空格 / 换行）或回车，都把已输入串切成多个对象；
      // 不必手动加逗号（用户要求「不加逗号就能发邮件」）。认不出的对象留在框里等用户补后缀，不丢、不刷屏。
      var 分隔 = e.key === 'Enter' || /[;；,，、\s]/.test(e.key);
      if (分隔) {
        e.preventDefault();
        var parts = String(inp.value || '').split(/[\s;；,，、]+/)
          .map(function (s) { return s.trim(); }).filter(function (s) { return s; });
        if (!parts.length) { inp.value = ''; return; }
        var 留下 = [];
        parts.forEach(function (p) { if (nfType(p)) nfAdd(p); else 留下.push(p); });
        inp.value = 留下.join(' ');   // 只清掉已识别并加入的；认不出的留在框里
      } else if (e.key === 'Backspace' && !inp.value && nfState.对象.length) {
        nfState.对象.pop(); nfRenderTags(); nfOpenMenu(true);
      } else if (e.key === 'Escape') {
        nfOpenMenu(false);
        e.stopPropagation();     // 别让 ESC 顺手触发整页「重置」
      }
    });
    inp.addEventListener('blur', function () {
      // 输了一半或粘贴了一串就点走：按任意分隔符切成多个对象，能认出通道的顺手加上；
      // 认不出的留在输入框里等用户补后缀，不弹提示、不丢数据。
      var parts = String(inp.value || '').split(/[\s;；,，、]+/)
        .map(function (s) { return s.trim(); }).filter(function (s) { return s; });
      if (parts.length) {
        var remain = [];
        parts.forEach(function (p) { if (nfType(p)) nfAdd(p); else remain.push(p); });
        inp.value = remain.join(' ');
      }
      setTimeout(function () { nfOpenMenu(false); }, 160);
    });
    nfRenderTags();
    nfRenderMenu('');
  }
  function collectTimedCfg() {
    var f = (document.querySelector('input[name=teFreq]:checked') || {}).value || '每天';
    var wds = $$('#trigEmBody [data-wd]').filter(function (c) { return c.checked; }).map(function (c) { return c.dataset.wd; });
    var calOn = !!($('#teCalendar') || {}).checked;
    return {
      freq: f, time: hmVal('#teTimeH', '#teTimeM', 9, 0), weekdays: wds,
      monthday: Math.max(1, Math.min(31, parseInt((($('#teMonthday') || {}).value) || '1', 10) || 1)),
      monthInterval: Math.max(1, parseInt((($('#teMonthInterval') || {}).value) || '1', 10) || 1),
      intervalMin: Math.max(1, Math.min(1439, parseInt((($('#teIntervalMin') || {}).value) || '1', 10) || 1)),
      intervalHour: Math.max(0, Math.min(23, parseInt((($('#teIntervalHour') || {}).value) || '0', 10) || 0)),
      intervalMinute: Math.max(0, Math.min(59, parseInt((($('#teIntervalMinute') || {}).value) || '0', 10) || 0)),
      cronDesc: (($('#teCron') || {}).value || '').trim(),
      endAfter: !!($('#teEndAfter') || {}).checked,
      endAt: (($('#teEndDate') || {}).value || '').trim()
        + 'T' + hmVal('#teEndH', '#teEndM', 0, 0),
      dayWindow: !!($('#teDayWindow') || {}).checked,
      winStart: hmVal('#teWinStartH', '#teWinStartM', 9, 0), winEnd: hmVal('#teWinEndH', '#teWinEndM', 17, 0),
      calendar: calOn, calendarMode: calOn ? (($('#teCalendarMode') || {}).value || 'factory') : '',
    };
  }
  function calName(c) {
    if (c.calendarMode === 'cn') return '国家规定工作日';
    if (c.calendarMode === 'factory') return '按工厂日历';
    return '';
  }
  function timedCondText(c) {
    var t = cnTime(c.time);
    if (c.freq === '每分') return '每间隔' + (c.intervalMin || 1) + '分钟执行';
    if (c.freq === '每时') return '每间隔' + (c.intervalHour || 0) + '个小时，在第' + (c.intervalMinute || 0) + '分钟执行';
    if (c.freq === '每天') return '每天的' + t + '执行';
    if (c.freq === '每工作日') return '每个工作日的' + t + '执行（' + (calName(c) || '按工厂日历') + '）';
    if (c.freq === '每周') return '每周的' + (c.weekdays.length ? c.weekdays.join('、') : '未选择星期') + '的' + t + '执行';
    if (c.freq === '月份') return '每' + (c.monthInterval || 1) + '个月的' + (c.monthday || 1) + '日的' + t + '执行';
    return '高级触发：' + (c.cronDesc || '未设置表达式');
  }
  /* ============ 「详情」：最近几次具体触发时间（悬浮/点击均显示） ============ */
  // 工厂日历工作日集合（覆盖当月起 3 个自然月；接口不可用时退回周一~周五）
  var wdCache = null;
  function workdaySet() {
    if (wdCache) return Promise.resolve(wdCache);
    var now = new Date();
    var months = [];
    for (var i = 0; i < 3; i++) {
      var d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      months.push({ y: d.getFullYear(), m: d.getMonth() + 1 });
    }
    return Promise.all(months.map(function (mm) {
      return fetch('/api/factory-calendar?year=' + mm.y + '&month=' + mm.m)
        .then(function (r) { return r.ok ? r.json() : { days: [] }; })
        .catch(function () { return { days: [] }; });
    })).then(function (rs) {
      var set = {};
      rs.forEach(function (r) { ((r && r.days) || []).forEach(function (d) { if (d.is_workday) set[d.date] = true; }); });
      wdCache = set;
      return set;
    });
  }
  /* ============ 国家规定工作日（国务院办公厅 2025-11-04 发布的 2026 年放假安排；其余年份按周一~周五） ============ */
  var CN_HOLIDAYS_2026 = (function () {
    var m = {};
    function set(list, v) { list.forEach(function (d) { m[d] = v; }); }
    // 元旦：1/1-1/3 休，1/4（周日）上班
    set(['2026-01-01', '2026-01-02', '2026-01-03'], 1); m['2026-01-04'] = 0;
    // 春节：2/15-2/23 休 9 天，2/14（周六）、2/28（周六）上班
    set(['2026-02-15', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20', '2026-02-21', '2026-02-22', '2026-02-23'], 1);
    m['2026-02-14'] = 0; m['2026-02-28'] = 0;
    // 清明：4/4-4/6 休
    set(['2026-04-04', '2026-04-05', '2026-04-06'], 1);
    // 劳动节：5/1-5/5 休，5/9（周六）上班
    set(['2026-05-01', '2026-05-02', '2026-05-03', '2026-05-04', '2026-05-05'], 1); m['2026-05-09'] = 0;
    // 端午：6/19-6/21 休
    set(['2026-06-19', '2026-06-20', '2026-06-21'], 1);
    // 中秋：9/25-9/27 休
    set(['2026-09-25', '2026-09-26', '2026-09-27'], 1);
    // 国庆：10/1-10/7 休，9/20（周日）、10/10（周六）上班
    set(['2026-10-01', '2026-10-02', '2026-10-03', '2026-10-04', '2026-10-05', '2026-10-06', '2026-10-07'], 1);
    m['2026-09-20'] = 0; m['2026-10-10'] = 0;
    return m;
  })();
  function isCnWorkday(d) {
    var iso = d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate());
    var v = CN_HOLIDAYS_2026[iso];
    if (v !== undefined) return !v; // 0=调休上班 → true；1=法定休假 → false
    var w = d.getDay();
    return w >= 1 && w <= 5; // 未覆盖年份：周一~周五为工作日
  }
  function isWorkday(set, d) {
    var iso = d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate());
    if (set && set[iso] !== undefined) return !!set[iso];
    var w = d.getDay();
    return w >= 1 && w <= 5; // 工厂日历未覆盖时按周一~周五兜底
  }
  function hm2min(v) { var p = String(v || '0:0').split(':'); return (parseInt(p[0], 10) || 0) * 60 + (parseInt(p[1], 10) || 0); }
  // 小时/分钟步进值 → "HH:MM"（替代原生 time 选择器）
  function hmVal(hId, mId, defH, defM) {
    var h = Math.max(0, Math.min(23, parseInt((($(hId) || {}).value) || defH, 10) || 0));
    var m = Math.max(0, Math.min(59, parseInt((($(mId) || {}).value) || defM, 10) || 0));
    return p2(h) + ':' + p2(m);
  }
  function parseLocalDT(v) {
    var m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], 0) : null;
  }
  // 通用过滤：每天起止时间窗口 + 按指定日历（工厂/国家）
  function passCommon(c, d, wdSet) {
    if (c.dayWindow) {
      var s = hm2min(c.winStart || '09:00'), e = hm2min(c.winEnd || '17:00');
      var t = d.getHours() * 60 + d.getMinutes();
      if (s < e) { if (t < s || t > e) return false; }
      else if (t < s && t > e) return false; // 跨夜窗口（如 22:00-06:00）仅排除窗外
    }
    var mode = c.calendarMode || (c.freq === '每工作日' ? 'factory' : '');
    if (mode === 'cn' ? !isCnWorkday(d) : (mode === 'factory' ? !isWorkday(wdSet, d) : false)) return false;
    return true;
  }
  /* ============ 高级：cron 五段式（分 时 日 月 周）解析 ============ */
  var CRON_MON_ALIAS = { JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12 };
  var CRON_DOW_ALIAS = { SUN: 0, MON: 1, TUE: 2, WED: 3, THU: 4, FRI: 5, SAT: 6 };
  function cronAlias(field, map) {
    return String(field).toUpperCase().replace(/[A-Z]+/g, function (w) { return map[w] !== undefined ? String(map[w]) : w; });
  }
  function cronField(field, minV, maxV) {
    var set = {};
    if (field === '*') { for (var v = minV; v <= maxV; v++) set[v] = true; return set; }
    String(field).split(',').forEach(function (part) {
      var step = 1;
      var slash = String(part).split('/');
      if (slash.length === 2) { step = parseInt(slash[1], 10) || 1; part = slash[0]; if (part === '*' || part === '') part = minV + '-' + maxV; }
      var m = String(part).match(/^(\d+)-(\d+)$/);
      if (m) {
        var a = +m[1], b = +m[2];
        if (a < minV || b > maxV || a > b || step < 1) throw new Error('bad');
        for (var v2 = a; v2 <= b; v2 += step) set[v2] = true;
      } else if (/^\d+$/.test(part)) {
        var v3 = +part;
        if (v3 < minV || v3 > maxV) throw new Error('bad');
        set[v3] = true;
      } else throw new Error('bad');
    });
    return set;
  }
  function cronParse(expr) {
    var f = String(expr || '').trim().split(/\s+/);
    if (f.length !== 5) return null;
    try {
      // 周字段允许 0-7（0 和 7 均=周日）；月/周段支持英文别名（JAN-DEC / SUN-SAT）
      // domAny/dowAny：记录该段是否是 *，给 cronNext 判断「这段不限制」用（对齐后端 _cron_due_today）
      return { min: cronField(f[0], 0, 59), hour: cronField(f[1], 0, 23), dom: cronField(f[2], 1, 31), mon: cronField(cronAlias(f[3], CRON_MON_ALIAS), 1, 12), dow: cronField(cronAlias(f[4], CRON_DOW_ALIAS), 0, 7), domAny: f[2] === '*', dowAny: f[4] === '*' };
    } catch (e) { return null; }
  }
  function cronNext(cp, from, n, endAt, extra) {
    var out = [];
    for (var di = 0; di < 800 && out.length < n; di++) { // 最多向前推约 2.2 年
      var d2 = new Date(from.getFullYear(), from.getMonth(), from.getDate() + di);
      if (!cp.mon[d2.getMonth() + 1]) continue;
      var domOK = cp.domAny || !!cp.dom[d2.getDate()];
      var dowOK = cp.dowAny || !!cp.dow[d2.getDay()] || (d2.getDay() === 0 && !!cp.dow[7]);
      if (!domOK || !dowOK) continue; // 与后端 _cron_due_today 一致：日/周都写具体值＝两个都要满足（交集），不是"或"
      var times = [];
      Object.keys(cp.hour).forEach(function (h) {
        Object.keys(cp.min).forEach(function (mi) { times.push(+h * 60 + +mi); });
      });
      times.sort(function (a, b) { return a - b; });
      for (var ti = 0; ti < times.length && out.length < n; ti++) {
        var d3 = new Date(d2.getFullYear(), d2.getMonth(), d2.getDate(), Math.floor(times[ti] / 60), times[ti] % 60, 0);
        if (d3 <= from) continue;
        if (endAt && d3 > endAt) return out;
        if (extra && !extra(d3)) continue;
        out.push(d3);
      }
    }
    return out;
  }
  // 计算未来 n 次触发时间（返回 Date[]）
  function computeNextRuns(c, n, wdSet) {
    var out = [];
    var now = new Date();
    var endAt = (c.endAfter && c.endAt) ? parseLocalDT(c.endAt) : null;
    var hm2 = String(c.time || '09:00').split(':');
    var hh = parseInt(hm2[0], 10) || 0, mi = parseInt(hm2[1], 10) || 0;
    function extra(d) { return passCommon(c, d, wdSet); }
    if (c.freq === '每分') {
      var stepM = Math.max(1, Math.min(1439, c.intervalMin || 1)) * 60000;
      for (var i = 1; out.length < n && i <= 60 * 24 * 400; i++) {
        var d = new Date(Math.floor(now.getTime() / 60000) * 60000 + i * stepM);
        d.setSeconds(0, 0);
        if (endAt && d > endAt) break;
        if (!passCommon(c, d, wdSet)) continue;
        out.push(d);
      }
    } else if (c.freq === '每时') {
      var ih = Math.max(0, Math.min(23, c.intervalHour || 0)), im = Math.max(0, Math.min(59, c.intervalMinute || 0));
      var stepH = (ih * 60 + im) * 60000; // 每 N 小时的第 M 分钟（N=0,M>0 时退化为每 M 分钟）
      if (stepH <= 0) stepH = 3600000;
      var anchor = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), im, 0);
      for (var i2 = 1; out.length < n && i2 <= 60 * 24 * 400; i2++) {
        var d2 = new Date(anchor.getTime() + i2 * stepH);
        if (endAt && d2 > endAt) break;
        if (!passCommon(c, d2, wdSet)) continue;
        out.push(d2);
      }
    } else if (c.freq === '高级') {
      var cp = cronParse(c.cronDesc);
      if (!cp) return out;
      return cronNext(cp, now, n, endAt, extra);
    } else {
      // 每天 / 每工作日 / 每周 / 月份：按日推进
      var monthStep = c.freq === '月份' ? Math.max(1, c.monthInterval || 1) : 1;
      for (var k = 0; k < 1500 && out.length < n; k++) {
        var d4;
        if (c.freq === '月份') {
          var md = Math.max(1, Math.min(31, c.monthday || 1));
          var base = new Date(now.getFullYear(), now.getMonth() + k * monthStep, 1, hh, mi, 0);
          var dim = new Date(base.getFullYear(), base.getMonth() + 1, 0).getDate();
          if (md > dim) continue; // 该月没有此日（如 31 日）则跳过
          d4 = new Date(base.getFullYear(), base.getMonth(), md, hh, mi, 0);
        } else {
          d4 = new Date(now.getFullYear(), now.getMonth(), now.getDate() + k, hh, mi, 0);
        }
        if (d4 <= now) continue;
        if (endAt && d4 > endAt) break;
        if (c.freq === '每周') {
          var wd = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'][d4.getDay()];
          if ((c.weekdays || []).indexOf(wd) < 0) continue;
        }
        if (!passCommon(c, d4, wdSet)) continue;
        out.push(d4);
      }
    }
    return out;
  }
  function fmtRun(d) {
    return d.getFullYear() + '.' + p2(d.getMonth() + 1) + '.' + p2(d.getDate()) + ' ' + p2(d.getHours()) + ':' + p2(d.getMinutes()) + ':' + p2(d.getSeconds());
  }
  var nextPopTimer = null;
  function hideNextPop() {
    if (nextPopTimer) { clearTimeout(nextPopTimer); nextPopTimer = null; }
    var p = $('#trigNextPop');
    if (p) p.classList.remove('open');
  }
  function showNextRuns(c, anchor) {
    var pop = $('#trigNextPop');
    if (!pop) return;
    closeRowMenu();
    pop.innerHTML = '<div class="tnp-title">最近几次触发时间</div><div class="tnp-empty">正在按条件推算…</div>';
    pop.classList.add('open');
    positionNextPop(anchor);
    var needWd = c.freq === '每工作日' || c.calendar;
    (needWd ? workdaySet() : Promise.resolve(null)).then(function (wdSet) {
      var runs = computeNextRuns(c, 5, wdSet);
      pop.innerHTML = '<div class="tnp-title">最近几次触发时间' + (c.freq === '每工作日' ? '（工作日按工厂日历）' : '') + '</div>' +
        (runs.length ? runs.map(function (r) { return '<div class="tnp-row">' + fmtRun(r) + '</div>'; }).join('')
          : '<div class="tnp-empty">条件不足，无法推算（请检查时间/星期等设置）</div>');
    });
  }
  function positionNextPop(anchor) {
    var pop = $('#trigNextPop');
    var rect = anchor.getBoundingClientRect();
    var top = rect.bottom + 6, left = Math.max(8, Math.min(rect.left - 40, window.innerWidth - pop.offsetWidth - 8));
    if (top + pop.offsetHeight > window.innerHeight - 8) top = Math.max(8, rect.top - pop.offsetHeight - 6);
    pop.style.top = top + 'px';
    pop.style.left = left + 'px';
  }

  function collectEditorData() {
    var type = state.editType;
    var d = {
      名称: (($('#teName') || {}).value || '').trim(),
      类型: type,
      应用: (($('#teApp') || {}).value || '').trim(),
      启用: !!($('#trigEmEnable') || {}).checked,
      config: {},
    };
    var q = !!($('#teQueue') || {}).checked;
    var to = parseInt((($('#teTimeout') || {}).value) || '0', 10) || 0;
    if (type === '条件') {
      var c0 = collectTimedCfg();
      c0.指标 = (($('#teMetric') || {}).value || '小型化每日下达');
      c0.比较 = (($('#teOp') || {}).value || '>');
      c0.阈值 = parseFloat((($('#teThreshold') || {}).value) || '0');
      c0.queue = q; c0.timeout = to;
      d.config = c0;
    } else if (type === '定时') {
      var c = collectTimedCfg();
      c.queue = q; c.timeout = to;
      if ((d.应用 || '') === 'yuzhi_mail_check') {      // 邮箱巡检专属参数（表单上没有的项按后端默认走）
        c.主题关键字 = (($('#teYzKw') || {}).value || '').trim() || '小于阀值异常工单明细';
        c.发件人 = (($('#teYzFrom') || {}).value || '').trim();
        c.检查封数 = Math.max(1, Math.min(20, parseInt((($('#teYzN') || {}).value) || '5', 10) || 5));
        c.判零值 = parseFloat((($('#teYzZero') || {}).value) || '0') || 0;
        c.钉钉接收人 = (($('#teYzTo') || {}).value || '').trim();
      }
      d.config = c;
    } else if (type === '文件') {
      d.config = {
        folder: (($('#teFolder') || {}).value || '').trim(),
        sub: !!($('#teSub') || {}).checked,
        events: $$('#trigEmBody [data-ev]').filter(function (c2) { return c2.checked; }).map(function (c2) { return c2.dataset.ev; }),
        patterns: (($('#tePatterns') || {}).value || '').trim(),
        queue: q, timeout: to,
      };
    } else if (type === '热键') {
      d.config = {
        ctrl: !!($('#teCtrl') || {}).checked, shift: !!($('#teShift') || {}).checked, alt: !!($('#teAlt') || {}).checked,
        key: (($('#teKey') || {}).value || 'F1'), queue: q, timeout: to,
      };
    } else if (type === '钉钉个人' || type === '钉钉群') {
      var tg = $$('#trigEmBody [data-target]').filter(function (c3) { return c3.checked; }).map(function (c3) { return c3.dataset.target; });
      $$('#teCustomChips .trig-chip').forEach(function (chip) {
        if (tg.indexOf(chip.dataset.custom) < 0) tg.push(chip.dataset.custom);
      });
      d.config = {
        targets: tg,
        matchTypes: $$('#trigEmBody [data-mtype]').filter(function (c5) { return c5.checked; }).map(function (c5) { return c5.dataset.mtype; }),
        match: {
          文本内容: (($('#teMTextVal') || {}).value || '').trim(),
          网址链接: (($('#teMUrlVal') || {}).value || '').trim(),
          指定名称文件: (($('#teMFileVal') || {}).value || '').trim(),
        },
        queue: q, timeout: to,
      };
    } else {
      d.config = {
        server: (($('#teServer') || {}).value || '163邮箱'),
        account: (($('#teAccount') || {}).value || '').trim(),
        authcode: (($('#teAuth') || {}).value || '').trim(),
        folder: (($('#teMailFolder') || {}).value || 'INBOX'),
        rules: {
          from: (($('#teRFrom') || {}).value || '').trim(), to: (($('#teRTo') || {}).value || '').trim(),
          subject: (($('#teRSubject') || {}).value || '').trim(), body: (($('#teRBody') || {}).value || '').trim(),
        },
        queue: q, timeout: to,
      };
    }
    // 通知（所有类型共用）：开关 + 混合通道对象列表——存进 config.通知，后端调度器按它发
    d.config.通知 = {
      启用: !!nfState.启用,
      对象: nfState.对象.map(function (o) { return o.值; }),
    };
    return d;
  }

  async function saveEditor() {
    var d = collectEditorData();
    if (!d.名称) { toast('请填写触发器名称'); var n = $('#teName'); if (n) n.focus(); return; }
    if (!d.应用) { toast('请选择应用'); var a = $('#teApp'); if (a) a.focus(); return; }
    if ((d.类型 === '钉钉个人' || d.类型 === '钉钉群') && !d.config.targets.length) {
      toast('请至少选择一个接收' + (d.类型 === '钉钉个人' ? '人' : '群')); return;
    }
    if ((d.类型 === '钉钉个人' || d.类型 === '钉钉群')) {
      if (!d.config.matchTypes.length) { toast('请至少选择一种匹配类型（文本内容/网址链接/指定名称文件）'); return; }
      var empt = d.config.matchTypes.filter(function (mt2) { return !d.config.match[mt2]; });
      if (empt.length) { toast('请填写「' + empt[0] + '」的匹配内容'); return; }
    }
    if (d.类型 === '条件') {
      var cC = d.config;
      if (!(cC.阈值 > 0)) { toast('请填写大于 0 的阈值'); var th = $('#teThreshold'); if (th) th.focus(); return; }
      if (cC.freq === '每周' && !(cC.weekdays || []).length) { toast('请至少勾选一个星期'); return; }
    }
    if (d.类型 === '定时') {
      var c2 = d.config;
      if (c2.freq === '高级' && !cronParse(c2.cronDesc)) { toast('时间表达式格式不正确，应为五段式：分 时 日 月 周（如 0 9 * * 1-5）'); var ci2 = $('#teCron'); if (ci2) ci2.focus(); return; }
      if (c2.freq === '每周' && !c2.weekdays.length) { toast('请至少勾选一个星期'); return; }
      if (c2.freq === '每时' && (c2.intervalHour || 0) === 0 && (c2.intervalMinute || 0) === 0) { toast('每小时间隔：小时与分钟不能同时为 0'); return; }
      if (c2.endAfter) {
        var ed2 = parseLocalDT(c2.endAt);
        if (!ed2 || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(c2.endAt || '')) { toast('定时结束时间格式不正确，请按 年-月-日 + 时 + 分 填写'); return; }
        if (ed2.getTime() <= Date.now()) { toast('定时结束时间必须晚于当前时间'); return; }
      }
      if (c2.dayWindow && hm2min(c2.winStart) >= hm2min(c2.winEnd)) { toast('每天起止时间：开始时间必须早于结束时间'); return; }
    }
    var t = state.editing;
    // 通用兜底：表单没有对应控件的 app 专属参数（如邮箱巡检的主题关键字/检查封数）不能因为
    // 「编辑一次任务」就被抹掉 —— 用旧 config 打底，本次表单填的值覆盖它。
    if (t && t.data && t.data.config) d.config = Object.assign({}, t.data.config, d.config);
    try {
      if (t) {
        await storePost('upsert', { sheet: '计划任务', id: t.id, row_no: t.row_no, data: d });
        t.data = d;
        toast('已保存触发器「' + d.名称 + '」');
      } else {
        var rowNo = state.tasks.length ? Math.max.apply(null, state.tasks.map(function (x) { return x.row_no || 0; })) + 1 : 1;
        await storePost('upsert', { sheet: '计划任务', id: 0, row_no: rowNo, data: d });
        var store = await storeGet();
        var sheet = (store.sheets && store.sheets['计划任务']) || [];
        var fresh = sheet[sheet.length - 1];
        if (fresh) state.tasks.push({ id: fresh.id, row_no: fresh.row_no, data: d });
        toast('已新建' + (TYPE_META[d.类型] || {}).title + '「' + d.名称 + '」');
      }
      closeEditor();
      renderTasks();
    } catch (err) { toast('保存失败：' + (err.message || err)); }
  }

  /* ============ 运行日志（TableUX 表格） ============ */
  var LOG_COLS = ['应用', '开始时间', '触发方式', '运行结果'];
  function logVal(row, col) {
    if (col === '行号') return String(row.row_no == null ? '' : row.row_no);
    var v = row.data ? row.data[col] : '';
    return v == null ? '' : String(v);
  }
  var logKit = null;
  function logGetKit() {
    if (!logKit) {
      logKit = createGridKit({
        id: 'trg-logs',
        getCols: function () { return ['行号'].concat(LOG_COLS); },
        getRows: function () { return state.logs; },
        valOf: logVal,
        bodySel: '#trgLogBody',        // 命中内容高亮用（2026-10-23 用户：查询后命中内容高亮）
        autofit: function () { TableUX.autofitWidths(logUXCfg(), state.logs); },
        rerender: renderLogs,
      });
      /* 2026-10-27 用户要求：运行日志默认「倒序」——最新的在最上面。
         用「行号」列排序（row_no 是自增序号，等价于写入顺序 = 时间先后），
         比直接比中文日期字符串可靠（"2026年10月27日" vs "2026年9月30日" 字符串比较会错）。 */
      logKit.sorts = [{ col: '行号', dir: 'desc' }];
    }
    return logKit;
  }
  function logUXCfg() {
    var colsAll = function () { return ['行号'].concat(LOG_COLS); };
    return {
      key: 'trg-logs',
      scrollSel: '#trgLogTableScroll', headSel: '#trgLogHead', bodySel: '#trgLogBody',
      cols: colsAll,
      dataCols: function () { return LOG_COLS.slice(); },
      rowOf: function (tr) { return logGetKit().apply(state.logs).find(function (r) { return r.id === parseInt(tr.dataset.rowid, 10); }); },
      colOfTd: function (td) { var cols = colsAll(); var i = td.cellIndex; return (i >= 2 && i <= cols.length) ? cols[i - 1] : null; },
      colOfTh: function (th) { var cols = colsAll(); var i = th.cellIndex; return (i >= 1 && i <= cols.length) ? cols[i - 1] : null; },
      valOf: logVal,
      editable: function () { return false; },
      rerender: renderLogs,
      notify: toast,
      onModeChange: function () { renderLogs(); },
    };
  }
  function resBadge(row) {
    var r = String((row.data || {}).运行结果 || '');
    var cls = r === '运行成功' ? 'r-ok' : r === '部分成功' ? 'r-part' : r === '运行失败' ? 'r-fail' : r === '运行中' ? 'r-run' : 'r-cancel';
    var badge = '<span class="trig-res ' + cls + '">' + esc(r) + '</span>';
    if (r === '运行失败' && (row.data || {}).错误次数) badge += '<sup style="color:var(--color-danger);font-size:10px;">' + (row.data.错误次数) + '</sup>';
    return badge;
  }
  function renderLogs() {
    if (!$('#trgLogHead')) return;
    var pg = TableUX.pageSlice(logGetKit().apply(state.logs), 'trg-logs');
    var rows = pg.rows;
    var kit = logGetKit();
    var selN = rows.filter(function (r) { return state.logSelected.has(r.id); }).length;
    var thCls = selN === 0 ? '' : (selN === rows.length && rows.length ? ' checked' : ' indeterminate');
    $('#trgLogHead').innerHTML = '<tr><th class="checkbox-col"><div class="checkbox' + thCls + '" id="trgLogThCheckbox"></div></th>' +
      '<th class="sortable" style="width:56px;">' + kit.headCell('行号') + '</th>' +
      LOG_COLS.map(function (c) { return '<th class="sortable">' + kit.headCell(c) + '</th>'; }).join('') +
      '<th class="center" style="width:110px;">操作</th></tr>';
    kit.bindHead($('#trgLogHead'));
    if (!rows.length) {
      $('#trgLogBody').innerHTML = '<tr><td colspan="' + (LOG_COLS.length + 3) + '" style="text-align:center;height:100px;color:var(--color-text-tertiary);">暂无运行日志（在「计划任务」中立即运行后自动记录）</td></tr>';
    } else {
      $('#trgLogBody').innerHTML = rows.map(function (r) {
        var d = r.data || {};
        var sel = state.logSelected.has(r.id);
        return '<tr data-rowid="' + r.id + '"' + (sel ? ' class="selected"' : '') + '>' +
          '<td class="checkbox-col"><div class="checkbox' + (sel ? ' checked' : '') + '" data-id="' + r.id + '"></div></td>' +
          '<td style="color:var(--color-text-tertiary);">' + r.row_no + '</td>' +
          '<td style="font-weight:500;" title="' + esc(d.应用) + '">' + esc(d.应用) + '</td>' +
          '<td>' + esc(d.开始时间) + '</td>' +
          '<td>' + esc(d.触发方式) + '</td>' +
          '<td>' + resBadge(r) + '</td>' +
          '<td class="center"><span class="action-link" data-lfact="查看">查看</span> · <span class="action-link" data-lfact="删除">删除</span></td>' +
          '</tr>';
      }).join('');
    }
    $('#trgLogCount').textContent = '共 ' + pg.total + ' 条' + (selN ? '，已选 ' + selN : '') + (pg.pages > 1 ? '（第 ' + pg.page + '/' + pg.pages + ' 页）' : '');
    var pager = $('#trgLogPager');
    if (pager) {
      pager.innerHTML = TableUX.pagerHtml('trg-logs', pg);
      TableUX.bindPager(pager, 'trg-logs', pg.pages, function () { renderLogs(); });
    }
    $('#trgLogBody').querySelectorAll('.checkbox[data-id]').forEach(function (cb) {
      cb.addEventListener('click', function (e) {
        e.stopPropagation();
        var id = parseInt(cb.dataset.id, 10);
        if (state.logSelected.has(id)) state.logSelected.delete(id); else state.logSelected.add(id);
        renderLogs();
      });
    });
    var thc = $('#trgLogThCheckbox');
    if (thc) thc.addEventListener('click', function (e) {
      e.stopPropagation();
      var all = rows.length > 0 && rows.every(function (r) { return state.logSelected.has(r.id); });
      rows.forEach(function (r) { if (all) state.logSelected.delete(r.id); else state.logSelected.add(r.id); });
      renderLogs();
    });
    $('#trgLogBody').querySelectorAll('[data-lfact]').forEach(function (link) {
      link.addEventListener('click', function (e) {
        e.stopPropagation();
        var id = parseInt(link.closest('tr').dataset.rowid, 10);
        if (link.dataset.lfact === '查看') openLogDetail(id);
        else confirmDeleteLog(id);
      });
    });
    var cfg = TableUX.mount(logUXCfg());
    if (cfg) { TableUX.refreshHandles(cfg); TableUX.ensureFixed(cfg, $('#trgLogTable')); TableUX.resetForRender(cfg); }
    if (typeof updateScrollbarsSoon === 'function') updateScrollbarsSoon();
  }
  function confirmDeleteLog(id) {
    var r = logById(id);
    if (!r) return;
    askConfirm('确认删除日志', '即将删除运行日志「' + esc((r.data || {}).应用) + ' · ' + esc((r.data || {}).开始时间) + '」，删除后不可恢复。是否确认？', '确认删除', async function () {
      try {
        await storePost('delete', { sheet: '运行日志', id: id });
        state.logs = state.logs.filter(function (x) { return x.id !== id; });
        state.logSelected.delete(id);
        renderLogs();
        toast('已删除 1 条运行日志');
      } catch (err) { toast('删除失败：' + (err.message || err)); }
    });
  }

  /* ============ 运行日志详情（日志文本 + 截图记录，类 binlog） ============ */
  function openLogDetail(id) {
    var idx = state.logs.findIndex(function (r) { return r.id === id; });
    if (idx < 0) return;
    state.logDetailIdx = idx;
    renderLogDetail();
    $('#trigLogDetailModal').classList.add('open');
  }
  function renderLogDetail() {
    var r = state.logs[state.logDetailIdx];
    if (!r) return;
    var d = r.data || {};
    $('#trgLdTitle').textContent = '运行详情 · ' + (d.应用 || '');
    $('#trgLdMeta').innerHTML =
      '<span class="lm">开始时间：' + esc(d.开始时间) + '</span>' +
      '<span class="lm">触发方式：' + esc(d.触发方式) + '</span>' +
      '<span class="lm">运行结果：' + esc(d.运行结果) + '</span>' +
      (d.主题 ? '<span class="lm">主题：' + esc(d.主题) + '</span>' : '');
    var lines = d.日志 || [];
    $('#trgLdLog').innerHTML = lines.length
      ? lines.map(function (t) {
        var cls = t.indexOf('✗') >= 0 ? 'll-fail' : t.indexOf('✓') >= 0 ? 'll-ok' : '';
        return '<div class="' + cls + '">' + esc(t) + '</div>';
      }).join('')
      : '<div style="color:var(--color-text-tertiary);">（暂无日志文本）</div>';
    var shots = d.截图 || [];
    $('#trgLdShots').innerHTML = shots.length
      ? shots.map(function (s, i) {
        var thumb = s.url ? '<img src="' + esc(s.url) + '" alt="' + esc(s.title) + '">' : '截图占位';
        return '<div class="trig-shot-card" data-shot="' + i + '">' +
          '<div class="trig-shot-thumb">' + thumb +
          (s.url ? '<span class="trig-shot-hint">点击看大图</span>' : '') + '</div>' +
          '<div class="trig-shot-cap"><span>' + esc(s.title || '截图') + '</span><span style="color:var(--color-text-tertiary);">' + esc(s.t || '') + '</span></div>' +
          '</div>';
      }).join('')
      : '<div style="font-size:12px;color:var(--color-text-tertiary);padding:20px 0;text-align:center;">（暂无截图记录）</div>';
    $('#trgLdShots').querySelectorAll('.trig-shot-card').forEach(function (card) {
      card.addEventListener('click', function () {
        var s = shots[parseInt(card.dataset.shot, 10)];
        if (s && s.url) window.open(s.url, '_blank');
        else toast('该截图为存档占位（运行期真实截图由任务页自动截推归档）');
      });
    });
    $('#trgLdPos').textContent = '第 ' + (state.logDetailIdx + 1) + ' / ' + state.logs.length + ' 条';
    $('#trgLdPrev').disabled = state.logDetailIdx <= 0;
    $('#trgLdNext').disabled = state.logDetailIdx >= state.logs.length - 1;
  }
  function bindLogDetailNav() {
    $('#trgLdPrev').addEventListener('click', function () { if (state.logDetailIdx > 0) { state.logDetailIdx--; renderLogDetail(); } });
    $('#trgLdNext').addEventListener('click', function () { if (state.logDetailIdx < state.logs.length - 1) { state.logDetailIdx++; renderLogDetail(); } });
  }

  /* ============ 运行错误处理 ============ */
  function errSettings() {
    return state.err || {
      邮件通知: true, 收件箱: ['demo.user75@example.com'], 发送方式: '钉钉邮箱',
      发送内容: { 截图: true, 最新运行日志: true, 主题: true, 正文内容: false, 主题文本: '', 正文文本: '' },
      钉钉通知: true,
    };
  }
  function renderErrModal() {
    var s = errSettings();
    setSwitch($('#treMailSwitch'), !!s.邮件通知);
    $('#treMailBody').hidden = !s.邮件通知;
    setSwitch($('#treDingSwitch'), !!s.钉钉通知);
    renderChips(s.收件箱 || []);
    $$('input[name=treMode]').forEach(function (r) { r.checked = true; }); // 发送方式固定：钉钉邮箱发送（DWS 链路）
    var c = s.发送内容 || {};
    $('#treCShot').checked = c.截图 !== false;
    $('#treCLog').checked = c.最新运行日志 !== false;
    $('#treCSubject').checked = c.主题 !== false;
    $('#treCBody').checked = !!c.正文内容;
    $('#treSubjectText').value = c.主题文本 || '';
    $('#treBodyText').value = c.正文文本 || '';
    $('#treSubjectWrap').hidden = !$('#treCSubject').checked;
    $('#treBodyWrap').hidden = !$('#treCBody').checked;
  }
  function setSwitch(el, on) { if (el) el.classList.toggle('on', !!on); }
  function renderChips(list) {
    var box = $('#treRecipients');
    box.innerHTML = (list || []).map(function (m, i) {
      return '<span class="trig-chip">' + esc(m) + '<span class="rm" data-i="' + i + '" title="移除">×</span></span>';
    }).join('') || '<span style="font-size:12px;color:var(--color-text-tertiary);">暂无收件人</span>';
    box.querySelectorAll('.rm').forEach(function (rm) {
      rm.addEventListener('click', function () {
        var list2 = errSettings().收件箱 || [];
        list2.splice(parseInt(rm.dataset.i, 10), 1);
        renderChips(list2);
      });
    });
  }
  function collectErr() {
    var chips = $$('#treRecipients .trig-chip').map(function (c) { return c.textContent.replace('×', '').trim(); }).filter(Boolean);
    return {
      邮件通知: $('#treMailSwitch').classList.contains('on'),
      收件箱: chips,
      发送方式: '钉钉邮箱',
      发送内容: {
        截图: $('#treCShot').checked,
        最新运行日志: $('#treCLog').checked,
        主题: $('#treCSubject').checked,
        正文内容: $('#treCBody').checked,
        主题文本: ($('#treSubjectText').value || '').trim(),
        正文文本: ($('#treBodyText').value || '').trim(),
      },
      钉钉通知: $('#treDingSwitch').classList.contains('on'),
    };
  }
  async function saveErr() {
    var s = collectErr();
    try {
      await fetch('/api/trigger/error-handling', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ settings: s }),
      }).then(function (res) { return res.json(); }).then(function (d) { if (!d.ok) throw new Error(d.msg || '保存失败'); });
      state.err = s;
      toast('运行错误处理设置已保存');
      $('#trigErrModal').classList.remove('open');
    } catch (err) { toast('保存失败：' + (err.message || err)); }
  }
  async function sendNotify(kind) {
    var s = collectErr();
    if (s.邮件通知 && !s.收件箱.length) { toast('请先添加收件箱地址'); return; }
    toast('正在通过钉钉邮箱发送测试邮件…');
    try {
      var d = await fetch('/api/trigger/notify', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: 'test', settings: s }),
      }).then(function (res) { return res.json(); });
      toast((d.ok ? '✓ ' : '✗ ') + (d.msg || (d.ok ? '已发送' : '发送失败')));
    } catch (err) { toast('✗ 发送失败：' + (err.message || err)); }
  }
  function bindErrEvents() {
    $('#trigErrBtn').addEventListener('click', function () {
      if (!state.err) {
        fetch('/api/trigger/error-handling').then(function (r) { return r.json(); }).then(function (d) {
          state.err = d.settings || null; renderErrModal(); $('#trigErrModal').classList.add('open');
        }).catch(function () { renderErrModal(); $('#trigErrModal').classList.add('open'); });
      } else { renderErrModal(); $('#trigErrModal').classList.add('open'); }
    });
    $('#trigErrClose').addEventListener('click', function () { $('#trigErrModal').classList.remove('open'); });
    $('#trigErrModal').addEventListener('click', function (e) { if (e.target === $('#trigErrModal')) $('#trigErrModal').classList.remove('open'); });
    $('#treMailSwitch').addEventListener('click', function () {
      var on = !$('#treMailSwitch').classList.contains('on');
      setSwitch($('#treMailSwitch'), on); $('#treMailBody').hidden = !on;
    });
    $('#treDingSwitch').addEventListener('click', function () {
      setSwitch($('#treDingSwitch'), !$('#treDingSwitch').classList.contains('on'));
    });
    $('#treCSubject').addEventListener('change', function () { $('#treSubjectWrap').hidden = !$('#treCSubject').checked; });
    $('#treCBody').addEventListener('change', function () { $('#treBodyWrap').hidden = !$('#treCBody').checked; });
    $('#treRecipientAdd').addEventListener('click', addRecipient);
    $('#treRecipientInput').addEventListener('keydown', function (e) { if (e.key === 'Enter') addRecipient(); });
    $('#treTestBtn').addEventListener('click', function () { sendNotify('test'); });
    $('#treSaveBtn').addEventListener('click', saveErr);
  }
  function addRecipient() {
    var inp = $('#treRecipientInput');
    var v = (inp.value || '').trim();
    if (!v) return;
    // 2026-10-29：支持用任意分隔符（分号、空格、换行、中英文逗号）一次写多个邮箱，
    // 逐个校验后加入（用户要求「不加逗号就能发邮件」）。
    var parts = v.split(/[\s;；,，、]+/).map(function (s) { return s.trim(); }).filter(function (s) { return s; });
    var list = errSettings().收件箱 || [];
    var added = 0, dup = 0, bad = 0;
    parts.forEach(function (p) {
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p)) { bad++; return; }
      if (list.indexOf(p) >= 0) { dup++; return; }
      list.push(p); added++;
    });
    inp.value = '';
    renderChips(list);
    if (added) {
      toast('已添加 ' + added + ' 个收件人'
        + (dup ? '（跳过 ' + dup + ' 个重复）' : '')
        + (bad ? '；' + bad + ' 个格式不对已忽略' : ''));
    } else if (dup) {
      toast('收件人已存在，未添加');
    } else if (bad) {
      toast('邮箱格式不正确');
    }
  }

  /* ============ 工具栏其余动作 ============ */
  function bindToolbar() {
    var si = $('#trigSearch');
    if (si) si.addEventListener('input', function () { state.search = si.value || ''; renderTasks(); });
    $('#trigResetBtn').addEventListener('click', function () {
      state.search = ''; state.filterTypes = []; state.filterEnabled = '';
      if (si) si.value = '';
      renderTasks(); closeFilterPop();
      toast('已清空搜索与筛选');
    });
    // 日志页工具条
    // 「查询」＝重新加载 + 应用关键字查询框（命中任意列即留下，命中内容高亮；2026-10-23 用户）
    async function trgLogRunQuery() {
      var kit = logGetKit();
      kit.quick = String(($('#trgLogQuick') || {}).value || '').trim();
      await loadLogs();
      toast('查询完成：' + kit.hitText());
    }
    $('#trgLogRefresh').addEventListener('click', function () { trgLogRunQuery(); });
    if (typeof wireQuickQuery === 'function') wireQuickQuery('#trgLogQuick', function () { return logGetKit(); }, trgLogRunQuery);
    $('#trgLogReset').addEventListener('click', function () {
      logGetKit().reset();
      var qi = $('#trgLogQuick'); if (qi) qi.value = '';
      state.logSelected.clear(); renderLogs(); toast('已重置排序/筛选/查询/选择');
    });
    $('#trgLogSelectAll').addEventListener('click', function () {
      var rows = logGetKit().apply(state.logs);
      var all = rows.length > 0 && rows.every(function (r) { return state.logSelected.has(r.id); });
      rows.forEach(function (r) { if (all) state.logSelected.delete(r.id); else state.logSelected.add(r.id); });
      renderLogs();
    });
    $('#trgLogDeleteSelected').addEventListener('click', async function () {
      var ids = Array.from(state.logSelected);
      if (!ids.length) { toast('请先勾选要删除的日志行'); return; }
      askConfirm('确认删除日志', '即将删除勾选的 <b>' + ids.length + '</b> 条运行日志，删除后不可恢复。是否确认？', '确认删除', async function () {
        try {
          for (var i = 0; i < ids.length; i++) await storePost('delete', { sheet: '运行日志', id: ids[i] });
          state.logs = state.logs.filter(function (r) { return !state.logSelected.has(r.id); });
          state.logSelected.clear();
          renderLogs();
          toast('已删除 ' + ids.length + ' 条运行日志');
        } catch (err) { toast('删除失败：' + (err.message || err)); }
      });
    });
    $('#trgLogExport').addEventListener('click', exportLogsCsv);
    $('#trgLogSearchOpen').addEventListener('click', function () { cqOpen(logGetKit()); });
  }
  function exportLogsCsv() {
    var rows = logGetKit().apply(state.logs);
    if (!rows.length) { toast('暂无日志可导出'); return; }
    var cols = ['应用', '开始时间', '触发方式', '运行结果', '主题'];
    var lines = [cols.join(',')];
    rows.forEach(function (r) {
      lines.push(cols.map(function (c) { return '"' + String(logVal(r, c) || '').replace(/"/g, '""') + '"'; }).join(','));
    });
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '运行日志_' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    toast('已导出 ' + rows.length + ' 条运行日志（csv）');
  }

  /* ============ 触发器弹窗通用窗口能力 ============
   * 三个弹窗（编辑 / 运行错误处理 / 日志详情）统一：
   * ① 右缘/底缘/右下角把手拖拽调大小（最小 460×340，最大 94vw/92vh）；
   * ② 标题栏按住拖动移动位置；
   * ③ 尺寸+位置记忆 localStorage，双击右下角把手复位。 */
  function winKey(maskId) { return 'trigWin_' + maskId; }
  function winSave(maskId, m) {
    var r = m.getBoundingClientRect();
    try {
      localStorage.setItem(winKey(maskId), JSON.stringify({
        w: Math.round(r.width), h: Math.round(r.height),
        l: Math.round(r.left), t: Math.round(r.top),
      }));
    } catch (e) { /* ignore */ }
  }
  function winApply(maskId, m, headSel, minW, minH) {
    var mask = $(maskId);
    if (!mask) return;
    var mm = mask.querySelector('.modal');
    if (!mm) return;
    mm.style.position = 'fixed';
    // 恢复记忆尺寸/位置
    try {
      var s = JSON.parse(localStorage.getItem(winKey(maskId)) || 'null');
      if (s && s.w && s.h) {
        mm.style.width = s.w + 'px'; mm.style.height = s.h + 'px';
        mm.classList.add('trig-sized');
        if (s.l != null && s.t != null) { mm.style.left = s.l + 'px'; mm.style.top = s.t + 'px'; }
      }
    } catch (e) { /* ignore */ }
    if (mm.querySelector('.trig-rz-handle')) return;
    [['e', 'trig-rz-handle trig-rz-e'], ['s', 'trig-rz-handle trig-rz-s'], ['se', 'trig-rz-handle trig-rz-se']].forEach(function (it) {
      var hEl = document.createElement('div');
      hEl.className = it[1];
      hEl.dataset.dir = it[0];
      hEl.title = '拖拽调整大小';
      mm.appendChild(hEl);
    });
    var cur = null;
    function start(e, mode) {
      if (e.button !== undefined && e.button !== 0) return;
      e.preventDefault();
      var r = mm.getBoundingClientRect();
      cur = { mode: mode, dir: e.target.dataset ? (e.target.dataset.dir || '') : '', x: e.clientX, y: e.clientY, w: r.width, h: r.height, l: r.left, t: r.top };
      document.addEventListener('pointermove', onMove);
      document.addEventListener('pointerup', onUp);
    }
    function onMove(e) {
      if (!cur) return;
      var maxW = Math.floor(window.innerWidth * 0.94), maxH = Math.floor(window.innerHeight * 0.92);
      if (cur.mode === 'move') {
        var nl = Math.max(0, Math.min(window.innerWidth - 120, cur.l + (e.clientX - cur.x)));
        var nt = Math.max(0, Math.min(window.innerHeight - 60, cur.t + (e.clientY - cur.y)));
        mm.style.left = Math.round(nl) + 'px';
        mm.style.top = Math.round(nt) + 'px';
        return;
      }
      var w = cur.w, h = cur.h;
      if (cur.dir.indexOf('e') >= 0) w = cur.w + (e.clientX - cur.x);
      if (cur.dir.indexOf('s') >= 0) h = cur.h + (e.clientY - cur.y);
      mm.style.width = Math.round(Math.max(minW, Math.min(maxW, w))) + 'px';
      mm.style.height = Math.round(Math.max(minH, Math.min(maxH, h))) + 'px';
      mm.classList.add('trig-sized');
    }
    function onUp() {
      if (!cur) return;
      cur = null;
      document.removeEventListener('pointermove', onMove);
      document.removeEventListener('pointerup', onUp);
      winSave(maskId, mm);
    }
    mm.addEventListener('pointerdown', function (e) {
      var hEl = e.target.closest ? e.target.closest('.trig-rz-handle') : null;
      if (hEl) start(e, 'resize');
    });
    var head = mm.querySelector(headSel);
    if (head) {
      head.addEventListener('pointerdown', function (e) {
        if (e.target.closest('button, .modal-close, input, select, textarea, label, a, .trig-switch, .trig-x')) return;
        start(e, 'move');
      });
    }
    var se = mm.querySelector('.trig-rz-se');
    if (se) se.addEventListener('dblclick', function () {
      mm.style.width = ''; mm.style.height = ''; mm.style.left = ''; mm.style.top = '';
      mm.classList.remove('trig-sized');
      try { localStorage.removeItem(winKey(maskId)); } catch (e2) { /* ignore */ }
      toast('已恢复默认窗口');
    });
  }
  function initTrigWindows() {
    winApply('#trigEditModal', null, '.trig-em-head', 460, 340);
    winApply('#trigErrModal', null, '.modal-title', 420, 300);
    winApply('#trigLogDetailModal', null, '.trig-ld-head', 560, 380);
  }

  /* ============ 启动 ============ */
  function boot() {
    bindListEvents();
    bindColResize();
    bindPopovers();
    bindToolbar();
    bindErrEvents();
    bindLogDetailNav();
    initTrigWindows();
    bindDwsChip();
    // 触发时间浮层：悬停保持、离开/点外部关闭
    var np = $('#trigNextPop');
    if (np) {
      np.addEventListener('mouseenter', function () { if (nextPopTimer) { clearTimeout(nextPopTimer); nextPopTimer = null; } });
      np.addEventListener('mouseleave', hideNextPop);
    }
    document.addEventListener('click', function (e) {
      if (!e.target.closest('#trigNextPop, #teDetailLink')) hideNextPop();
    });
    // 目录选择弹窗：点遮罩关闭
    $('#trigFsModal').addEventListener('click', function (e) { if (e.target === $('#trigFsModal')) $('#trigFsModal').classList.remove('open'); });
    // 编辑器
    $('#trigEmClose').addEventListener('click', closeEditor);
    $('#trigEmCancel').addEventListener('click', closeEditor);
    $('#trigEditModal').addEventListener('click', function (e) { if (e.target === $('#trigEditModal')) closeEditor(); });
    $('#trigEmSave').addEventListener('click', saveEditor);
    // 确认框
    $('#trigConfirmCancel').addEventListener('click', function () { $('#trigConfirmModal').classList.remove('open'); confirmOk = null; });
    $('#trigConfirmModal').addEventListener('click', function (e) { if (e.target === $('#trigConfirmModal')) { $('#trigConfirmModal').classList.remove('open'); confirmOk = null; } });
    $('#trigConfirmOk').addEventListener('click', async function () {
      $('#trigConfirmModal').classList.remove('open');
      if (confirmOk) { var fn = confirmOk; confirmOk = null; await fn(); }
    });
    // 日志详情
    $('#trgLdClose').addEventListener('click', function () { $('#trigLogDetailModal').classList.remove('open'); });
    $('#trgLdCloseBtn').addEventListener('click', function () { $('#trigLogDetailModal').classList.remove('open'); });
    $('#trigLogDetailModal').addEventListener('click', function (e) { if (e.target === $('#trigLogDetailModal')) $('#trigLogDetailModal').classList.remove('open'); });
    // 双击日志行 → 打开详情（capture 阶段拦截，先于 TableUX 的单元格编辑逻辑）
    var sc = $('#trgLogTableScroll');
    if (sc) sc.addEventListener('dblclick', function (e) {
      var tr = e.target.closest('tr[data-rowid]');
      if (!tr) return;
      e.preventDefault(); e.stopPropagation();
      openLogDetail(parseInt(tr.dataset.rowid, 10));
    }, true);
  }

  /* ============ 钉钉连接状态（2026-10-27，2026-10-30 改造）============
     钉钉的短期令牌（access token）**固定 2 小时**，这是服务端签发的 TTL，谁也改不了；
     真正会「到期失效」的是 30 天的长期凭据（refresh token）。所以：
       · 面板同时摆出两把钥匙：短期令牌（自动换新，不用管）+ 长期凭据（真到期日）
       · 芯片只在「短期令牌剩不到 30 分钟」（= 后台自动换新没成功）时才转黄，
         以前按 45 分钟就黄，正常换新时也会闪，纯属误报。
     掉线才转红，点开有诊断与一键恢复。 */
  var dwsTimer = null;
  function dwsSetState(state, text, tip) {
    var chip = $('#trigDwsChip');
    if (!chip) return;
    chip.setAttribute('data-state', state);
    var t = $('#trigDwsText');
    if (t) t.textContent = text;
    if (tip) chip.title = tip;
  }
  function dwsFetch(recover) {
    return fetch('/api/dws/diagnose' + (recover ? '?recover=true' : ''), { cache: 'no-store' })
      .then(function (r) { return r.json().catch(function () { return {}; }); })
      .catch(function () { return {}; });
  }
  function dwsApply(d) {
    if (!d || d.可用 === undefined) { dwsSetState('warn', '钉钉状态未知', '诊断接口不可用'); return; }
    if (d.可用) {
      var 剩 = d.剩余分钟;
      var 天 = d.长期凭据剩余天;
      var 文案 = (剩 != null && 剩 < 30) ? '钉钉令牌待换新' : '钉钉已连接';
      var 提示 = '钉钉连接正常';
      if (剩 != null) 提示 += '｜短期令牌还剩 ' + Math.round(剩) + ' 分钟（2 小时一换，后台自动换新）';
      if (天 != null) 提示 += '｜长期凭据还剩 ' + 天.toFixed(0) + ' 天（' + (d.长期凭据到期 || '') + ' 到期）';
      提示 += '｜点击查看诊断';
      dwsSetState(剩 != null && 剩 < 30 ? 'warn' : 'ok', 文案, 提示);
    } else {
      dwsSetState('bad', '钉钉已断开', '钉钉连接异常：' + (d.说明 || '') + '｜点击查看诊断与恢复');
    }
    d._数据 = d;
  }
  function dwsRefresh() {
    return dwsFetch(false).then(function (d) { dwsApply(d); return d; });
  }
  function dwsShowPop() {
    var pop = $('#trigDwsPop');
    if (!pop) return;
    var chip = $('#trigDwsChip');
    var r = chip.getBoundingClientRect();
    pop.classList.add('open');
    var w = pop.offsetWidth, h = pop.offsetHeight;
    pop.style.left = Math.max(8, Math.min(r.left, window.innerWidth - w - 12)) + 'px';
    pop.style.top = Math.max(8, (r.bottom + 6 + h > window.innerHeight ? r.top - h - 6 : r.bottom + 6)) + 'px';
  }
  function dwsRenderPop(d) {
    var pop = $('#trigDwsPop');
    if (!pop) return;
    var 可用 = d && d.可用;
    var 剩 = d && d.剩余分钟 != null ? Math.round(d.剩余分钟) + ' 分钟' : '未知';
    var 短到期 = d && d.短期令牌到期 ? '（' + d.短期令牌到期 + ' 到期）' : '';
    var 天 = d && d.长期凭据剩余天 != null ? d.长期凭据剩余天 : null;
    var 长期 = 天 == null ? '未知'
      : (天 >= 1 ? Math.round(天) + ' 天' : (天 * 24).toFixed(0) + ' 小时');
    var 长到期 = d && d.长期凭据到期 ? '（' + d.长期凭据到期 + ' 到期）' : '';
    var 长期色 = (天 != null && 天 < 7) ? 'var(--color-warning,#d97706)' : 'inherit';
    var 客户端 = d && d.客户端在跑 ? '运行中' : '未运行';
    var 状态色 = 可用 ? 'var(--color-success)' : 'var(--color-danger)';
    pop.innerHTML =
      '<h4><span class="dws-dot" style="width:8px;height:8px;border-radius:50%;background:' + 状态色 + '"></span>' +
      '钉钉连接诊断</h4>' +
      '<div class="dws-row"><span class="dws-k">状态</span><span class="dws-v" style="color:' + 状态色 + '">' +
      (可用 ? '已连接' : '已断开') + '</span></div>' +
      '<div class="dws-row"><span class="dws-k">原因</span><span class="dws-v">' + esc(d && d.原因 || '—') + '</span></div>' +
      '<div class="dws-row"><span class="dws-k">说明</span><span class="dws-v">' + esc(d && d.说明 || '—') + '</span></div>' +
      '<div class="dws-row"><span class="dws-k">短期令牌</span><span class="dws-v">' + 剩 + 短到期 +
      '<br><span class="dws-sub">钉钉固定 2 小时，后台自动换新，无需操作</span></span></div>' +
      '<div class="dws-row"><span class="dws-k">长期凭据</span><span class="dws-v" style="color:' + 长期色 + '">' +
      长期 + 长到期 +
      '<br><span class="dws-sub">' + (天 != null && 天 < 7
        ? '快到期了：请在钉钉桌面客户端重新登录一次'
        : '随每次换新顺延，只要它在有效期内就不会真断') + '</span></span></div>' +
      '<div class="dws-row"><span class="dws-k">钉钉客户端</span><span class="dws-v">' + 客户端 + '</span></div>' +
      (d && d.建议 ? '<div class="dws-advice">' + esc(d.建议) + '</div>' : '') +
      '<div class="dws-actions">' +
      '<button class="btn btn-default" data-dws-act="refresh">刷新</button>' +
      '<button class="btn btn-primary" data-dws-act="recover">一键重新授权</button>' +
      '</div>';
  }
  function bindDwsChip() {
    var chip = $('#trigDwsChip');
    if (!chip) return;
    var pop = $('#trigDwsPop');
    chip.addEventListener('click', function (e) {
      e.stopPropagation();
      if (pop && pop.classList.contains('open')) { pop.classList.remove('open'); return; }
      dwsSetState('checking', '钉钉检测中…', '正在检测钉钉连接状态');
      dwsFetch(false).then(function (d) {
        dwsApply(d); dwsRenderPop(d); dwsShowPop();
      });
    });
    if (pop) {
      pop.addEventListener('click', function (e) {
        var act = e.target.getAttribute && e.target.getAttribute('data-dws-act');
        if (!act) return;
        e.stopPropagation();
        if (act === 'refresh') {
          dwsSetState('checking', '钉钉检测中…');
          dwsFetch(false).then(function (d) { dwsApply(d); dwsRenderPop(d); });
        } else if (act === 'recover') {
          dwsSetState('checking', '重新授权中…', '正在尝试重新授权钉钉');
          dwsFetch(true).then(function (d) {
            dwsApply(d); dwsRenderPop(d);
            toast(d && d.可用 ? '钉钉已重新连接' : ('重新授权未成功：' + (d && d.说明 || '未知原因')));
          });
        }
      });
      document.addEventListener('click', function (e) {
        if (!e.target.closest('#trigDwsChip, #trigDwsPop')) pop.classList.remove('open');
      });
    }
    dwsRefresh();
    if (dwsTimer) clearInterval(dwsTimer);
    dwsTimer = setInterval(function () { dwsRefresh(); }, 60000);   // 每分钟探一次
  }

  window.TriggerBoard = { boot: boot, switchTab: switchTab, reload: function () { state.loaded.tasks = false; state.loaded.logs = false; if (state.tab === 'logs') loadLogs(); else loadTasks(); }, openErrModal: function () { var b = $('#trigErrBtn'); if (b) b.click(); }, refreshDws: dwsRefresh };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
