/* ==========================================================================
   Bom分层（常用功能 → Bom分层）
   数据源：zpp_cs12.exe（逆向还原）里的「展开层 → 根物料 + 层级物料」拍平逻辑，
          后端改由 xlwings 读 Excel（xls / xlsx 都支持），算法在 app/bom_layering.py。
   页面：表格走 TableUX（Excel 式选区/复制/编辑/列宽行高/冻结/分页/右键导出），
        排序筛选查询走 createGridKit，勾选走 rowPickWire（Shift 范围勾选）。
   ========================================================================== */
(function () {
  'use strict';

  const S = {
    path: '', file: '', sheets: [], sheet: '',
    srcHeaders: [], srcRows: [],
    outHeaders: [], outRows: [],
    stats: null,
    descCol: -1,            // 中文描述列：按表头自动定（命中「描述」类列＞第 11 列＞不输出），页面不放控件
    view: 'src',            // 'src' 源数据 | 'out' 分层结果
    headerHit: true, headerWhy: '', headerRow: 1,
    selected: new Set(),
    loading: false,
    bound: false,
    undo: [], redo: [],
  };
  let kit = null;
  let uxCfg = null;

  const $ = (sel) => document.querySelector(sel);

  // 表头行「没识别到」的红色高亮（自注入样式，避免改 index.html 里那块 1.6MB 的 <style>）
  function 注入样式() {
    if (document.getElementById('bomStyle')) return;
    const st = document.createElement('style');
    st.id = 'bomStyle';
    st.textContent = `
      #bomHeaderRow.is-bad {
        background: #FDE2E1 !important; border-color: #E5484D !important;
        color: #C4271F !important; font-weight: 700 !important;
        box-shadow: 0 0 0 3px rgba(229,72,77,.16) !important;
      }
      #bomHeaderRow.is-bad:hover { background: #FCCFCE !important; }
      .bom-hr-tip {
        position: absolute; z-index: 60; max-width: 340px; padding: 7px 10px;
        background: #2B1214; color: #FFE7E7; border-radius: 6px; font-size: 12px; line-height: 1.6;
        box-shadow: 0 10px 28px -8px rgba(0,0,0,.5); pointer-events: none;
        opacity: 0; transition: opacity .12s;
      }
      .bom-hr-tip.on { opacity: 1; }
      .bom-hr-tip b { color: #FFC9C9; }
      .tool-inline.bom-hr-wrap { position: relative; display: inline-flex; align-items: center; gap: 6px; }
    `;
    document.head.appendChild(st);
  }

  // 红色悬浮气泡：鼠标停在「表头行」上时显示「我们要找的是哪一行」
  let _tipEl = null, _tipTimer = 0;
  function 绑定表头悬浮() {
    const inp = $('#bomHeaderRow');
    if (!inp || inp.dataset.tipped) return;
    inp.dataset.tipped = '1';
    const wrap = inp.closest('.tool-inline') || inp.parentElement;
    if (wrap) wrap.classList.add('bom-hr-wrap');
    注入样式();
    const 显示 = () => {
      _tipEl = _tipEl || (() => {
        const d = document.createElement('div');
        d.className = 'bom-hr-tip';
        document.body.appendChild(d);
        return d;
      })();
      _tipEl.innerHTML = 当前表头提示();
      const r = inp.getBoundingClientRect();
      _tipEl.style.left = Math.min(r.left, window.innerWidth - 350) + 'px';
      _tipEl.style.top = (r.bottom + 8) + 'px';
      _tipEl.classList.add('on');
    };
    const 隐藏 = () => { clearTimeout(_tipTimer); _tipTimer = setTimeout(() => { if (_tipEl) _tipEl.classList.remove('on'); }, 80); };
    inp.addEventListener('mouseenter', 显示);
    inp.addEventListener('focus', 显示);
    inp.addEventListener('mouseleave', 隐藏);
    inp.addEventListener('blur', 隐藏);
  }

  function 当前表头提示() {
    const 词 = '展开层 / 项目 / 对象标识 / 组件数量';
    if (S.headerHit) {
      return `已自动定位到<b>第 ${S.headerRow} 行</b>（${escapeHtml(S.headerWhy || '')}）<br>识别依据：${词} 这几个字所在的那一行。`;
    }
    return `⚠ 没认出表头行，已填 <b>0</b>。<br>目标：<b>${词}</b> 所在的那一行。<br>请对照源数据把真实行号填进来（例：3），或点①重导让它再自动找一次。`;
  }
  const 当前头 = () => (S.view === 'out' ? S.outHeaders : S.srcHeaders);
  const 当前行 = () => (S.view === 'out' ? S.outRows : S.srcRows);

  function 状态文案(t) {
    const el = $('#bomStatus');
    if (el) el.textContent = t;
  }

  function 行对象(id) {
    const arr = 当前行()[id];
    return arr ? { id, arr } : null;
  }
  function 取值(row, col) {
    if (!row || !row.arr) return '';
    if (col === '行号') return String(row.id + 1);
    const i = 当前头().indexOf(col);
    if (i < 0) return '';
    const v = row.arr[i];
    return v == null ? '' : String(v);
  }

  function bomKit() {
    if (!kit) {
      kit = createGridKit({
        id: 'bom-table',
        bodySel: '#bomBody',
        getCols: () => ['行号'].concat(当前头()),
        getRows: () => 当前行().map((_, i) => ({ id: i, arr: 当前行()[i] })),
        valOf: (row, col) => 取值(row, col),
        autofit: () => TableUX.autofitWidths(uxCfg, 当前行().map((_, i) => ({ id: i, arr: 当前行()[i] }))),
        rerender: () => bomRender(),
      });
    }
    return kit;
  }
  function 视图行() { return bomKit().apply(当前行().map((_, i) => ({ id: i, arr: 当前行()[i] }))); }

  // ---------------------------------------------------------------- 渲染
  function bomRender() {
    if (!S.bound) return;
    const _sp = TableUX.scrollPos('#bomTableScroll');
    const heads = 当前头();
    const objs = 视图行();
    const pg = TableUX.pageSlice(objs, 'bom-table');
    const rows = pg.rows;
    const k = bomKit();

    const 已勾 = rows.filter(r => S.selected.has(r.id)).length;
    const thCls = 已勾 === 0 ? '' : (已勾 === rows.length && rows.length ? ' checked' : ' indeterminate');

    $('#bomHead').innerHTML = `<tr>`
      + `<th class="checkbox-col"><div class="checkbox${thCls}" id="bomThCheckbox"></div></th>`
      + `<th class="sortable" style="width:56px;">${k.headCell('行号')}</th>`
      + heads.map(h => `<th class="sortable">${k.headCell(h)}</th>`).join('')
      + `</tr>`;
    k.bindHead($('#bomHead'));

    const 空文案 = S.view === 'out'
      ? '还没有分层结果：先① 导入数据源，再点② Bom分层'
      : '还没有数据：点① zpp_cs12数据源导入（write格式） 选一个 .xlsx / .xls';

    if (S.loading) {
      $('#bomBody').innerHTML = `<tr><td colspan="${heads.length + 2}" style="text-align:center;height:120px;color:var(--color-text-tertiary);">正在读取 Excel…</td></tr>`;
    } else if (!rows.length) {
      $('#bomBody').innerHTML = `<tr><td colspan="${heads.length + 2}" style="text-align:center;height:120px;color:var(--color-text-tertiary);">${空文案}</td></tr>`;
    } else {
      // 层级多（可能到 20 层）时横向滚到右边就看不到第 1 列了 —— 每行挂个 title，悬停即知属于哪个根行
      const 组号 = new Map();
      if (S.view === 'out') {
        let 当前 = '';
        当前行().forEach((arr, i) => {
          const v = String((arr && arr[0]) || '').trim();
          if (v) 当前 = v;
          组号.set(i, 当前);
        });
      }
      $('#bomBody').innerHTML = rows.map(r => {
        const on = S.selected.has(r.id);
        const 组 = 组号.get(r.id) || '';
        const 行提示 = 组 ? `根物料：${组}` : '';
        return `<tr data-rowid="${r.id}"${on ? ' class="selected"' : ''}${行提示 ? ` title="${escapeHtml(行提示)}"` : ''}>`
          + `<td class="checkbox-col"><div class="checkbox${on ? ' checked' : ''}" data-id="${r.id}"></div></td>`
          + `<td style="color:var(--color-text-tertiary);">${r.id + 1}</td>`
          + heads.map(h => {
            const v = 取值(r, h);
            // 单元格 title 带上所属根行：层级多、横向滚到右边时也能知道这是哪一组的物料
            const tip = v ? (组 ? `${v}\n根物料：${组}` : v) : (组 ? `根物料：${组}` : '');
            return `<td title="${escapeHtml(tip)}">${escapeHtml(v)}</td>`;
          }).join('')
          + `</tr>`;
      }).join('');
    }

    const 已选 = rows.filter(r => S.selected.has(r.id)).length;
    $('#bomCount').textContent = `共 ${pg.total} 条` + (已选 ? `，已勾 ${已选}` : '')
      + (pg.pages > 1 ? `（第 ${pg.page}/${pg.pages} 页）` : '');
    const pager = $('#bomPager');
    if (pager) {
      pager.innerHTML = TableUX.pagerHtml('bom-table', pg);
      TableUX.bindPager(pager, 'bom-table', pg.pages, () => bomRender());
    }

    // Shift 范围勾选（全站统一口径）
    rowPickWire({
      key: 'bom', container: $('#bomBody'),
      order: () => rows.map(r => r.id),
      isChecked: (id) => S.selected.has(id),
      setChecked: (id, on) => { if (on) S.selected.add(id); else S.selected.delete(id); },
      after: () => bomRender(),
    });
    const thc = $('#bomThCheckbox');
    if (thc) thc.addEventListener('click', (e) => {
      e.stopPropagation();
      rowPickAll({
        order: () => rows.map(r => r.id),
        isChecked: (id) => S.selected.has(id),
        setChecked: (id, on) => { if (on) S.selected.add(id); else S.selected.delete(id); },
        after: () => bomRender(),
      });
    });

    uxCfg = TableUX.mount(bomUXCfg());
    if (uxCfg) {
      TableUX.refreshHandles(uxCfg);
      TableUX.ensureFixed(uxCfg, $('#bomTable'));
      TableUX.resetForRender(uxCfg);
    }
    TableUX.restoreScroll('#bomTableScroll', _sp);
    if (typeof updateScrollbarsSoon === 'function') updateScrollbarsSoon();

    // 页签与状态
    document.querySelectorAll('#bomTabs .tool-subtab').forEach(el => {
      el.classList.toggle('on', el.dataset.bomview === S.view);
    });
    // 状态条只放「口径 + 统计」（文件名/工作表已经在路径框和工作表下拉里，不再重复占宽）
    const 统计 = S.stats
      ? `｜ ${S.stats.合同数} 根物料 / ${S.stats.明细行数} 明细 / 最大 ${S.stats.最大层级} 层`
      : '';
    const 表头段 = S.headerHit ? `表头第 ${S.headerRow} 行` : '⚠ 表头行未识别';
    状态文案(表头段 + (S.view === 'out' ? 统计 : `｜ 源 ${S.srcRows.length} 行`));
  }

  function bomUXCfg() {
    const heads = () => ['行号'].concat(当前头());
    return {
      key: 'bom-' + S.view,
      scrollSel: '#bomTableScroll', headSel: '#bomHead', bodySel: '#bomBody',
      cols: heads,
      dataCols: () => 当前头().slice(),
      rowOf: (tr) => 行对象(parseInt(tr.dataset.rowid, 10)),
      colOfTd: (td) => { const c = heads(); const i = td.cellIndex; return (i >= 2 && i <= c.length) ? c[i - 1] : null; },
      colOfTh: (th) => { const c = heads(); const i = th.cellIndex; return (i >= 1 && i <= c.length) ? c[i - 1] : null; },
      valOf: (row, col) => 取值(row, col),
      editable: (row, col) => col !== '行号',
      rerender: () => bomRender(),
      notify: (msg) => showToast(msg),
      onModeChange: () => bomRender(),
      onSaveCell: (row, col, val) => {
        const i = 当前头().indexOf(col);
        if (!row || i < 0) return;
        压撤销();
        当前行()[row.id][i] = val;
        bomRender();
      },
      onPasteRow: (data) => {
        压撤销();
        当前行().push((data || []).slice(0, 当前头().length));
        bomRender();
      },
    };
  }

  // ---------------------------------------------------------------- 撤销（本地内存）
  function 快照() { return JSON.stringify(当前行()); }
  function 压撤销() {
    S.undo.push(快照());
    if (S.undo.length > 50) S.undo.shift();
    S.redo.length = 0;
  }
  function 应用快照(s) {
    try {
      const rows = JSON.parse(s);
      if (S.view === 'out') S.outRows = rows; else S.srcRows = rows;
      bomRender();
    } catch (e) { /* 快照坏了就不动 */ }
  }

  // ---------------------------------------------------------------- 网络
  async function 请求(url, body) {
    const res = await fetch(url, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {}),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.detail || ('请求失败（HTTP ' + res.status + '）'));
    return data;
  }

  // 把当前「口径」三个控件与刚读到的簿对齐（表头行＝后端识别出来的那一行）
  function 填口径(信息) {
    S.headerHit = 信息.headerHit !== false;
    S.headerWhy = 信息.headerWhy || '';
    S.headerRow = 信息.headerRow || 0;
    const hr = $('#bomHeaderRow');
    if (hr) {
      hr.value = S.headerHit ? (信息.headerRow || 1) : 0;
      hr.classList.toggle('is-bad', !S.headerHit);
      hr.title = 当前表头提示().replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
    }
    // 中文描述列不再放控件（用户 2026-10-29：口径在代码里已写死）——按表头自动定，存进 S.descCol
    const hs = 信息.headers || [];
    const 命中 = hs.findIndex(h => /中文描述|物料描述|品名|描述/.test(String(h)));
    S.descCol = 命中 >= 0 ? 命中 : (hs.length > 10 ? 10 : -1);
  }

  // auto=true：表头行交给后端自动识别（换文件时必须这样，否则会拿上一份文件的表头行去套新文件）
  function 读口径(auto) {
    const hr = parseInt(($('#bomHeaderRow') || {}).value, 10);
    return {
      表头行: (auto || isNaN(hr) || hr < 1) ? null : hr,
      中文描述列: S.descCol,
    };
  }

  function 装载(信息) {
    S.path = 信息.path || '';
    S.file = 信息.file || '';
    S.sheets = 信息.sheets || [];
    S.sheet = 信息.sheet || '';
    S.srcHeaders = 信息.headers || [];
    S.srcRows = 信息.rows || [];
    S.outHeaders = []; S.outRows = []; S.stats = null;
    S.selected.clear(); S.undo = []; S.redo = [];
    S.view = 'src';
    const sel = $('#bomSheet');
    if (sel) {
      sel.innerHTML = S.sheets.map(n => `<option${n === S.sheet ? ' selected' : ''}>${escapeHtml(n)}</option>`).join('');
      sel.disabled = !S.sheets.length;
    }
    $('#bomPathInput').value = S.path;
    填口径(信息);
    bomRender();
    showToast(`已导入：${S.file}（${S.srcRows.length} 行 × ${S.srcHeaders.length} 列）`, 10000);
  }

  // ---------------------------------------------------------------- 按钮
  async function 导入() {
    try {
      状态文案('正在等待选择文件…');
      const r = await 请求('/api/bom/pick');
      if (!r.ok && r.cancelled) { 状态文案('已取消选择'); return; }
      if (!r.ok) throw new Error('没有选中文件');
      装载(r);
    } catch (err) {
      showToast('导入失败：' + (err.message || err), 10000);
      bomRender();
    }
  }

  async function 按路径读取(手动表头) {
    const p = ($('#bomPathInput').value || '').trim();
    if (!p) { showToast('请先点①导入，或把文件完整路径粘到左边的输入框'); return; }
    try {
      S.loading = true; bomRender();
      const r = await 请求('/api/bom/load', {
        path: p, sheet: ($('#bomSheet').value || ''),
        表头行: 读口径(!手动表头).表头行,     // 换文件＝自动识别；手改了才用输入框的值
      });
      S.loading = false;
      装载(r);
    } catch (err) {
      S.loading = false; bomRender();
      showToast('读取失败：' + (err.message || err), 10000);
    }
  }

  async function 切表() {
    if (!S.path) { showToast('还没有数据源'); return; }
    try {
      S.loading = true; bomRender();
      const r = await 请求('/api/bom/load', { path: S.path, sheet: $('#bomSheet').value, 表头行: 读口径().表头行 });
      S.loading = false;
      装载(r);
    } catch (err) {
      S.loading = false; bomRender();
      showToast('切换工作表失败：' + (err.message || err), 10000);
    }
  }

  async function 分层() {
    if (!S.srcRows.length) { showToast('请先①导入数据源'); return; }
    // 表头行没识别到（红框）时挡一下，给的是"该填哪个数"的具体指引
    if (!S.headerHit && !读口径().表头行) {
      showToast('表头行没识别到（红框那格现在是 0）：目标是「展开层 / 项目 / 对象标识 / 组件数量」所在的那一行。'
        + '对照源数据把真实行号填进去（比如 3）再点②；或者点①重导一次让它自动再找。', 10000);
      const hr = $('#bomHeaderRow');
      if (hr) { hr.focus(); hr.classList.add('is-bad'); }
      return;
    }
    try {
      状态文案('正在分层…');
      const k = 读口径();
      const r = await 请求('/api/bom/run', {
        path: S.path, sheet: S.sheet,
        headers: S.srcHeaders, rows: S.srcRows,
        表头行: k.表头行 || 1,
        中文描述列: k.中文描述列,
      });
      S.outHeaders = r.headers || [];
      S.outRows = r.rows || [];
      S.stats = r.stats || null;
      S.view = 'out';
      S.selected.clear(); S.undo = []; S.redo = [];
      bomRender();
      const st = S.stats || {};
      const 备注 = (r.notes && r.notes.length) ? ('；注意：' + r.notes.join('；')) : '';
      showToast(`分层完成：${st.合同数 || 0} 个根物料 / ${st.明细行数 || 0} 条明细 / 最大 ${st.最大层级 || 0} 层${备注}`, 10000);
    } catch (err) {
      showToast('分层失败：' + (err.message || err), 10000);
      bomRender();
    }
  }

  // ---------------------------------------------------------------- ③④⑤ 只作用于②分层结果
  // 用户 2026-10-29 口径：③导出 / ④打开 / ⑤转化 处理的是「② Bom分层」出来的结果表，
  // 不是①导进来的数据源。没分层就点 → 直接提示，不动源数据；当前在源数据页签也照样切到结果。
  function 备好结果(动作) {
    if (!S.outRows.length) {
      showToast(`「${动作}」处理的是②Bom分层的结果：请先点② Bom分层`, 10000);
      const b = $('#bomRun');
      if (b) b.focus();
      return false;
    }
    if (S.view !== 'out') { S.view = 'out'; S.selected.clear(); bomRender(); }
    return true;
  }

  function 结果文件名() {
    const d = new Date();
    const p2 = (n) => String(n).padStart(2, '0');
    const 戳 = `${d.getFullYear()}${p2(d.getMonth() + 1)}${p2(d.getDate())}_${p2(d.getHours())}${p2(d.getMinutes())}${p2(d.getSeconds())}`;
    const 源 = (S.path || '').replace(/\\/g, '/').split('/').pop().replace(/\.[^.]+$/, '');
    return `${源 || 'Bom分层'}_分层结果_${戳}.xlsx`;
  }
  // 与源文件同目录（⑤「转化」原来的落点口径）；取不到就交给后端默认（桌面）
  function 结果目录() {
    const p = (S.path || '').replace(/\\/g, '/');
    const i = p.lastIndexOf('/');
    return i > 0 ? p.slice(0, i) : '';
  }

  // 把分层结果落成 xlsx：打开＝④，只生成＝⑤（③ 走全站导出弹窗，另算）
  async function 存结果(打开) {
    const 动作 = 打开 ? '④ 直接打开' : '⑤ 转化为xlsx';
    if (!备好结果(动作)) return;
    const headers = S.outHeaders.map(h => String(h));
    const rows = S.outRows.map(r => headers.map((_, i) => {
      const v = r[i];
      return v === null || v === undefined ? '' : String(v);
    }));
    try {
      状态文案(打开 ? '正在生成分层结果并打开…' : '正在把分层结果存成 xlsx…');
      const r = await 请求('/api/export/page', {
        filename: 结果文件名(), sheet: '分层结果',
        headers, rows, dir: 结果目录(), open_after: !!打开,
      });
      const 尾 = 打开 ? (r.opened ? '，已用 Excel 打开' : `；打开未成功：${r.openTip || '未知原因'}`) : '';
      showToast(`分层结果已生成：${r.name}（${r.rows} 行 × ${r.cols} 列）→ ${r.dir}${尾}`, 10000);
      bomRender();
    } catch (err) {
      showToast((打开 ? '生成/打开失败：' : '生成失败：') + (err.message || err), 10000);
      bomRender();
    }
  }

  // ③ 导出为Excel：仍是全站「导出当前页」弹窗（选路径/重名加(2)/可选打开），但数据固定取分层结果
  function 导出为Excel() {
    if (!备好结果('③ 导出为Excel')) return;
    if (typeof openPageExportDialog === 'function') openPageExportDialog();
    else showToast('导出组件未就绪');
  }

  // ---------------------------------------------------------------- 启动
  function bind() {
    if (S.bound) return;
    S.bound = true;

    $('#bomPick').addEventListener('click', 导入);
    $('#bomRun').addEventListener('click', 分层);
    $('#bomExport').addEventListener('click', 导出为Excel);
    $('#bomOpen').addEventListener('click', () => 存结果(true));
    $('#bomConvert').addEventListener('click', () => 存结果(false));
    $('#bomSheet').addEventListener('change', 切表);
    // 改「表头行」＝重新切片：直接回读（表头行变了，列名全变）
    $('#bomHeaderRow').addEventListener('change', () => { if (S.path) 按路径读取(true); });

    $('#bomPathGo').addEventListener('click', () => 按路径读取(false));   // 注意：直接传函数会把事件对象当参数
    $('#bomPathInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') 按路径读取(false); });
    $('#bomQuery').addEventListener('input', (e) => {
      bomKit().quick = e.target.value || '';
      bomRender();
    });
    // 重置：排序/筛选/查询/勾选全部归零（ESC 会自动点到它）
    $('#bomReset').addEventListener('click', () => {
      bomKit().reset();
      S.selected.clear();
      const q = $('#bomQuery'); if (q) q.value = '';
      bomRender();
    });
    document.querySelectorAll('#bomTabs .tool-subtab').forEach(el => {
      el.addEventListener('click', () => {
        const v = el.dataset.bomview;
        if (v === 'out' && !S.outRows.length) { showToast('还没有分层结果：先点② Bom分层'); return; }
        S.view = v; S.selected.clear(); bomRender();
      });
    });

    // Ctrl+Z / Ctrl+Y：本页没有接 UndoKit（数据在内存里），自己管撤销栈
    document.addEventListener('keydown', (e) => {
      const ctrl = e.ctrlKey || e.metaKey;
      if (!ctrl) return;
      const page = document.getElementById('bomPage');
      if (!page || page.offsetParent === null) return;      // 不在本页不抢键
      const t = e.target;
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA')) return;
      const k = (e.key || '').toLowerCase();
      if (k === 'z' && !e.shiftKey) {
        if (!S.undo.length) { showToast('没有可撤销的操作（单元格编辑/粘贴才能撤销）', 10000); return; }
        e.preventDefault(); S.redo.push(快照()); 应用快照(S.undo.pop());
        showToast('已撤销上一步', 10000);
      } else if (k === 'y' || (k === 'z' && e.shiftKey)) {
        if (!S.redo.length) { showToast('没有可重做的操作', 10000); return; }
        e.preventDefault(); S.undo.push(快照()); 应用快照(S.redo.pop());
        showToast('已重做', 10000);
      }
    }, true);
  }

  window.BomBoard = {
    boot() {
      bind();
      绑定表头悬浮();
      if (!S.stats && !S.srcRows.length) 状态文案('未加载数据');
      bomRender();
    },
    state: S,
  };
})();
