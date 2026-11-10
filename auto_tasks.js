/* ====================================================================
 * auto_tasks.js —— 自动化任务框架（模板文件）
 * ====================================================================
 * 用途：把原本只能在 xlsm 里跑的 VBA 自动化流程，做成网页工作台可运行的
 *       「节点流水线任务」：每个任务 = 一串节点，逐个执行，实时显示
 *       未执行 / 运行中 / 成功 / 失败 状态与运行日志。
 *
 * 【如何新增一个自动化任务】—— 复制下面的「任务模板」段落，改三处即可：
 *   1. id / name / desc：任务标识与展示文案
 *   2. nodes：节点数组，每个节点 { name, run(ctx) }，按顺序执行；
 *      run 内用 ctx.log() 打日志、ctx.shared 在节点间传递中间数据；
 *      抛异常即视为该节点失败，流水线停止。
 *   3. 在 index.html 的 AUTO_PAGES 映射里加一行：'菜单页名': '任务id'
 *
 * 本框架不直接触碰页面内部实现：所有读写能力由主程序在启动时通过
 * AutoTaskKit.bind(host) 注入（见 index.html 的 host 实现）。
 * ==================================================================== */

const AutoTaskKit = {
  tasks: {},                       // taskId → 任务定义
  host: null,                      // 主程序注入的运行环境
  register(task) { this.tasks[task.id] = task; },
  bind(host) { this.host = Object.assign(this.host || {}, host); },
  get(id) { return this.tasks[id]; },
};

/* --------------------------------------------------------------------
 * 任务模板（复制本块新增任务）
 * --------------------------------------------------------------------
AutoTaskKit.register({
  id: 'demo_task',
  name: '示例任务',
  desc: '这个模板复制粘贴即可注册新的自动化任务。',
  nodes: [
    { name: '读取数据',   run: async (ctx) => { ctx.log('读取…'); ctx.shared.rows = await ctx.host.loadSheetRows('需求汇总'); } },
    { name: '业务计算',   run: async (ctx) => { ctx.log('计算…'); } },
    { name: '写回结果',   run: async (ctx) => { ctx.log('写回…'); } },
    { name: '刷新页面',   run: async (ctx) => { ctx.host.refreshSheet('需求汇总'); } },
  ],
});
-------------------------------------------------------------------- */

/* ====================================================================
 * 任务 1：需求预测（移植自 xlsm VBA 模块「ab_需求预测」）
 * --------------------------------------------------------------------
 * 原 VBA：Sub 需求预测() —— 读取"需求汇总"各期间列（YYYY.MM），对每个
 * 物料号用加权平均法预测下一个月需求，四舍五入取整、负数取 0，写入
 * 最后期间列右侧的"下一个月预测"列（每次运行覆盖）。
 * 依赖说明：原代码依赖 vbayyds 模块的 LogStart/LogPrint/LogEnd，在本
 * 框架里由 ctx.log / 节点状态承担同等职责。
 * ==================================================================== */

/* 任务 1：Resb 需求调取与预测（直调 xlsm VBA，节点/日志按 VBA 备注分步细化）
 * --------------------------------------------------------------------
 * 设计原则（用户 2026-10-15 深夜明确要求）：不重写业务算法。网页端 = 触发器，
 * xlsm 里的 VBA 模块代码 = 执行引擎；**每个 VBA 主要功能对应一个节点**
 * （其他任务同样遵守此规则）。
 * z1_需求组合 模块头备注的调用顺序（即节点顺序，逐个 Run 子宏）：
 *   1. a0_resb需求调取.resb调取
 *      （SAP SE16/RESB：从 qqqq N8/O8 取日期 + Sheet2 C/D 物料清单，
 *       导出 Resb_YYYYMMDDHHMMSS.xls 到 D:\ai_project\forecast\数据源\）
 *   2. aa_resb需求汇总计算.resb需求汇总计算
 *      （透视：Resb_*.xls T/AA/AB 列 → 一料多码映射 → 「需求汇总」工作表）
 *   3. ab_需求预测.需求预测
 *      （加权平均 i/(n*(n+1)/2) 权重 → 写入「需求汇总」最后一列「下一个月预测」）
 *   4. ac_数量回填.数量回填
 *      （按 Sheet2 C 列匹配「需求汇总」A 列 → 预测值写入 Sheet2 H 列）
 *   5. 同步 xlsm Sheet2/需求汇总 → 网页
 *   6. 刷新页面数据
 * 前提：本机已登录 SAP 且 Excel 可用。
 * ==================================================================== */

AutoTaskKit.register({
  id: 'demand_forecast',
  name: 'Resb需求调取与预测（运行 VBA z1_需求组合）',
  desc: '按 z1_需求组合 模块头备注的调用顺序逐节点执行 VBA 子宏：a0_resb需求调取（SAP SE16 RESB 导出 Resb_时间戳.xls）→ aa_resb需求汇总计算（透视到「需求汇总」表）→ ab_需求预测（加权平均下一个月）→ ac_数量回填（按 Sheet2 C列匹配写到 H列），跑完后把 xlsm 的 Sheet2/需求汇总 同步到网页。',
  impactSheets: ['Sheet2', '需求汇总'],   // 声明受影响工作表 → 运行前自动快照，供「回滚恢复」
  nodes: [
    {
      name: 'SAP调取RESB数据（a0）',
      // steps = VBA a0 模块头备注的主要流程（1~7）
      steps: [
        'SAP连接：获取 SAP GUI Session 对象（GetObject("SAPGUI")）',
        '数据读取：qqqq 工作表 N8/O8 取日期范围，Sheet2 C/D 列取物料编码清单',
        '数据准备：物料编码去空格整理复制到剪贴板',
        'SAP操作：打开 SE16 事务进入 RESB 表查询界面',
        '条件设置：在 SAP 中设置物料编码和项目日期的筛选条件',
        '数据导出：导出为 Resb_YYYYMMDDHHMMSS.xls 到 D:\\ai_project\\forecast\\数据源\\',
        '文件处理：自动定位最新生成的文件并打开',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('a0_resb需求调取.resb调取', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），Resb_*.xls 已导出到 D:\\ai_project\\forecast\\数据源`);
      },
    },
    {
      name: '需求汇总计算（aa）',
      // steps = VBA aa 模块头备注的主要流程（0/1a/1b/2/2b/3/4）
      steps: [
        '0. 检查工作簿（yutouyouhuaa.xlsm + Resb_*.xls 均需已打开）',
        '1a. 读 Sheet2 C/D 列映射 + 1b. 读 resb*.xls T/AA/AB 列汇总源数据',
        '2. 用 C/D 列配对构建一料多码映射字典（DMP开头 > P开头 > 默认C列）',
        '2b. 提取主键物料号顺序（去重），用于结果行排序',
        '3. 调用 BuildPivotData 构建「物料号 × 期间」透视汇总',
        '4. 将结果写入「需求汇总」工作表（行=主键物料号，列=期间升序）',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('aa_resb需求汇总计算.resb需求汇总计算', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），需求透视已写入「需求汇总」表`);
      },
    },
    {
      name: '需求预测（ab）',
      // steps = VBA ab 模块头备注（加权平均法 + 写入预测列）
      steps: [
        '1. 获取「需求汇总」工作表（lastRow/lastCol 数据范围）',
        '2. 识别最后有效期间列（YYYY.MM 格式）→ 计算下一期间（跨年→01）',
        '3. 加权平均法：第 i 点权重 = i / (n*(n+1)/2) 预测下一个月',
        '4. 负数取 0、四舍五入取整，写入最后列「下一个月预测」（覆盖式）',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('ab_需求预测.需求预测', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），下一个月预测已写入「需求汇总」最后一列`);
      },
    },
    {
      name: '数量回填（ac）',
      // steps = VBA ac 模块头备注（1~5）
      steps: [
        '1. 获取工作表（需求汇总 + Sheet2）',
        '2. 获取需求汇总数据范围（lastRow/lastCol）',
        '3. 查找「下一个月预测」列位置',
        '4. 读取预测数据到字典（key=物料号，value=预测值）',
        '5. Sheet2 C 列匹配需求汇总 A 列 → 预测值写入 Sheet2 H 列（未匹配=0）',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('ac_数量回填.数量回填', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），预测数量已回填到 Sheet2 H 列`);
      },
    },
    {
      name: '同步VBA结果到网页',
      steps: [
        '读 xlsm Sheet2 + 需求汇总 UsedRange（COM）',
        '按 row_no 保留行 id 写回网页「预投物料号」与「需求汇总」表',
      ],
      run: async (ctx) => {
        const res = await ctx.host.syncVbaSheets(['Sheet2', '需求汇总']);
        const n2 = (res.synced && res.synced['Sheet2']) || 0;
        const n5 = (res.synced && res.synced['需求汇总']) || 0;
        if (!n2 && !n5) throw new Error('同步完成但未返回 Sheet2/需求汇总 行数');
        ctx.log(`xlsm 已同步到网页：Sheet2 ${n2} 行（H 列下一个月预测）+ 需求汇总 ${n5} 行（透视 + 预测列）`);
      },
    },
    {
      name: '刷新页面数据',
      steps: [
        '重载「预投·预投物料号」工作表页（H 列新预测）',
        '重载「需求汇总」工作表页（透视与下一期间列）',
      ],
      run: async (ctx) => {
        ctx.host.refreshSheet('Sheet2');
        ctx.host.refreshSheet('需求汇总');
        ctx.log('「预投·预投物料号」与「需求汇总」页面已刷新');
      },
    },
  ],
});

/* ====================================================================
 * 任务 2：预投物料账面结余计算（直调 xlsm VBA，节点/日志按 VBA 备注分步细化）
 * --------------------------------------------------------------------
 * 设计原则（用户明确要求）：不重写业务算法。网页端 = 触发器，
 * xlsm 里的 VBA 模块代码 = 执行引擎；**每个 VBA 主要功能对应一个节点**
 * （其他任务同样遵守此规则）。
 * z2_账面结余 模块头备注的调用顺序（即节点顺序，逐个 Run 子宏）：
 *   1. ba_zppmd04n调取.ba_zppmd04n调M15后及当年年底账面结余
 *      （SAP /nzpp_md04n：qqqq M 列物料 → 当天+15 与 当年12.31 各导一份）
 *   2. bb_M15账面结余计算.M15账面结余计算（解析 M15 文件 → Sheet2 I 列）
 *   3. bc_年底前账面结余计算.年底账面结余计算（解析年底文件 → Sheet2 J 列）
 *   4. 同步 xlsm Sheet2 → 网页「预投物料号」表
 *   5. 刷新页面数据
 * ba/bb/bc 内部备注的分步（---1 查找文件 → ---2 建映射 → ---3 读数据 →
 * ---4 筛选计算 → ---5 关文件 → ---6 写结果列）无法单独 Run（同一 Sub），
 * 因此在节点定义 steps 数组（悬浮节点名弹出子步骤气泡），节点本身仍只 Run VBA。
 * 前提：本机已登录 SAP 且 Excel 可用。
 * ==================================================================== */

AutoTaskKit.register({
  id: 'zpp_md04n_balance',
  name: '账面结余计算（运行 VBA z2_账面结余）',
  desc: '按 z2_账面结余 模块头备注的调用顺序逐节点执行 VBA 子宏：ba_zppmd04n调取（SAP /nzpp_md04n 导出 M15 与年底数据）→ bb_M15账面结余计算（Sheet2 I 列）→ bc_年底前账面结余计算（Sheet2 J 列），跑完后把 xlsm 的 Sheet2 结果同步回「预投物料号」表。',
  impactSheets: ['Sheet2'],   // 声明受影响工作表 → 运行前自动快照，供「回滚恢复」
  nodes: [
    {
      name: 'SAP调取zpp_md04n数据（ba）',
      // steps = VBA ba 模块备注的分步（悬浮节点名可查看子步骤）
      steps: [
        '调取M15数据：SAP /nzpp_md04n，qqqq M 列物料粘贴，C_BDTER=当天+15，导出 XLS',
        '调取年底数据：C_BDTER=当年12.31，导出 XLS',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('ba_zppmd04n调取.ba_zppmd04n调M15后及当年年底账面结余', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），两份 zpp_md04n_*.XLS 已导出到 D:\\ai_project\\forecast\\数据源`);
      },
    },
    {
      name: 'M15账面结余计算（bb）',
      // steps = VBA bb 模块备注 ---1~---6 分步
      steps: [
        '查找打开 zpp_md04n_M15日期*.xls（多文件取最新时分秒）',
        '读取 Sheet2 C/D 列构建一料多码映射字典',
        '读取导出数据（第5行起，B物料/D范围/F库存/G预留/H可用）',
        '逐行计算：筛选MRP范围=1 → 前缀P/DMP/G取期间可用；Y/DMY按当前库存取可用/预留 → 归并主键',
        '关闭数据文件（自动打开的才关）',
        '按 C 列物料号匹配写入 Sheet2 I 列，未匹配留空',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('bb_M15账面结余计算.M15账面结余计算', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），M15账面结余已写入 Sheet2 I 列`);
      },
    },
    {
      name: '年底账面结余计算（bc）',
      // steps = VBA bc 模块备注 ---1~---6 分步（文件查找规则与 bb 不同）
      steps: [
        '查找打开 zpp_md04n_年底日期*.xls（仅取修改时间为当天的最新文件）',
        '读取 Sheet2 C/D 列构建一料多码映射字典',
        '读取导出数据（第5行起，B物料/D范围/F库存/G预留/H可用）',
        '逐行计算：筛选MRP范围=1 → 前缀P/DMP/G取期间可用；Y/DMY按当前库存取可用/预留 → 归并主键',
        '关闭数据文件（自动打开的才关）',
        '按 C 列物料号匹配写入 Sheet2 J 列，未匹配留空',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('bc_年底前账面结余计算.年底账面结余计算', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），年底账面结余已写入 Sheet2 J 列`);
      },
    },
    {
      name: '同步VBA结果到网页',
      steps: [
        '读 xlsm Sheet2 UsedRange（COM）',
        '按 row_no 保留行 id 写回网页「预投物料号」表',
      ],
      run: async (ctx) => {
        const res = await ctx.host.syncVbaSheets(['Sheet2']);
        const n = res.synced && res.synced['Sheet2'];
        if (!n) throw new Error('同步完成但未返回 Sheet2 行数');
        ctx.log(`xlsm Sheet2 已同步到网页：${n} 行（含 M15账面结余/年底账面结余 最新计算结果）`);
      },
    },
    {
      name: '刷新页面数据',
      steps: ['重载「预投·预投物料号」工作表页显示'],
      run: async (ctx) => {
        ctx.host.refreshSheet('Sheet2');
        ctx.log('「预投·预投物料号」页面已刷新');
      },
    },
  ],
});

/* 任务 3：库存水平分类（混合模式：VBA 在 xlsm 内写 C 列颜色，网页展示层镜像）
 * --------------------------------------------------------------------
 * ca_库存水平分类 模块只有一个主 Sub「库存水平分类填充颜色」，其内 ---1~---7
 * 分步（见 VBA 头备注）：获取工作表 → 取数据范围 → 清 C 列原色 → 读数据到数组 →
 * 初始化统计 → 遍历分类填充 → 输出统计。Sub 内无法单独 Run，因此 1 VBA 节点 +
 * steps[] 忠实搬运 ---1~---7（与 z2 ba/bb/bc 同款处理）。
 *
 * 网页侧动作：VBA 完成后从 xlsm 同步 Sheet2（H/I/J 三列）；后端 ReadSheet 同步值
 * 不含单元格填充色——故网页「展示层」用 localStorage 缓存本轮 JS 重算结果
 * （与 VBA 分类规则逐行一致），供工作表页的预投物料号单元格套色+悬浮批注。
 *
 * 【待办：完整迁移到直调模式】需在 _vba_同步工作表 增加 Interior.Color 读取，
 * 把 fill 一起同步回网页，前端消费 fill 即可删 fc_库存水平分类行/计算 函数。
 * 依赖：vbayyds(LogPrint)。前提：本机 Excel 可用。
 * ==================================================================== */

/* 库存水平五级分类元数据（颜色值 = VBA RGB 值） */
const INVENTORY_LEVELS = [
  { key: 'sev',  name: '严重不足', color: '#C00000', range: 'M15可用 < 1月' },
  { key: 'jit',  name: '基本JIT',  color: '#FF9999', range: '1 ≤ M15可用 < 1.5月' },
  { key: 'more', name: '较多',     color: '#FFFF99', range: '1.5 ≤ M15可用 ≤ 12月' },
  { key: 'over', name: '过多',     color: '#FFFF00', range: '12 < M15可用 ≤ 24月' },
  { key: 'very', name: '非常多',   color: '#FFC000', range: 'M15可用 > 24月 或 H列需求为0' },
];

/* 移植自 VBA 主过程循环体：单行分类判定（展示层镜像，规则与 VBA If/ElseIf 链一致）。
 * h=下一个月预测，i=M15账面结余，j=年底账面结余。
 * 返回 { level:'sev|jit|more|over|very|none', m15m, ym, comment } */
function fc_库存水平分类行(h, i, j) {
  const H = fc_数字(h), I = fc_数字(i), J = fc_数字(j);
  let m15m = 0, ym = 0;
  if (H > 0) { m15m = I / H; ym = J / H; }          // VBA: 需求预测>0 才计算可用月数
  let level = 'none';
  if (m15m > 0 && m15m < 1) level = 'sev';           // RGB(192,0,0)
  else if (m15m >= 1 && m15m < 1.5) level = 'jit';   // RGB(255,153,153)
  else if (m15m >= 1.5 && m15m <= 12) level = 'more';// RGB(255,255,153)
  else if (m15m > 12 && m15m <= 24) level = 'over';  // RGB(255,255,0)
  else if (m15m > 24 || H === 0) level = 'very';     // RGB(255,192,0)
  // VBA 批注文本（H=0 时「可用很久」）
  const comment = H > 0
    ? `M15后：可用${m15m.toFixed(2)}月\n当年年底：可用${ym.toFixed(2)}月`
    : 'M15后：可用很久\n当年年底：可用很久';
  return { level, m15m, ym, comment };
}

/* VBA CDbl 容错：非数字 → 0 */
function fc_数字(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

/* 整表分类 + 统计（与 VBA 第5-7步同口径：C列空行跳过、严重不足记录物料号清单）。
 * rows：loadSheetRows 返回的行数组；返回 { items, stats, sevList } */
function fc_库存水平分类计算(rows) {
  const items = [], stats = { sev: 0, jit: 0, more: 0, over: 0, very: 0, none: 0, total: 0 };
  const sevList = [];
  for (const row of rows) {
    const mat = String(row.data['预投物料号'] ?? '').trim();
    if (!mat) continue;                              // VBA: C列为空 → GoTo 下一行（不填色不加批注）
    const r = fc_库存水平分类行(row.data['下一个月预测'], row.data['M15账面结余'], row.data['年底账面结余']);
    if (r.level !== 'none') stats[r.level] += 1; else stats.none += 1;
    if (r.level === 'sev') sevList.push(mat);        // VBA: 记录严重不足物料号
    stats.total += 1;
    items.push({ id: row.id, row_no: row.row_no, mat, ...r });
  }
  return { items, stats, sevList };
}

AutoTaskKit.register({
  id: 'inventory_level',
  name: '库存水平分类（运行 VBA ca_库存水平分类）',
  desc: '按 ca_库存水平分类 模块头备注：VBA 在 xlsm Sheet2 C 列预投物料号单元格按 M15/月数 填五级颜色并写批注（H>0 才计算可用月数）；跑完后同步 Sheet2 到网页供展示层镜像渲染（pre-migration 阶段尚有 JS 展示层重算，详见任务头注释【待办】）。',
  impactSheets: ['Sheet2'],   // 声明受影响工作表 → 运行前自动快照，供「回滚恢复」
  nodes: [
    {
      name: '读取预投物料号工作表',
      // 前置数据准备：不在 xlsm 内运行 VBA 也不写值，仅从网页存储加载 Sheet2
      steps: [
        '0. 检查工作簿（运行前必须先有数据：依赖 Resb→需求预测→zpp_md04n_balance 任一前置任务已在 Sheet2 写入 H/I/J）',
        '1. ctx.host.loadSheetRows("Sheet2") 读网页 forecast_store 全量行（含物料号/下一个月预测/M15账面结余/年底账面结余）',
        '2. 若 rows.length===0 → 抛错「预投物料号（Sheet2）没有物料数据！」中断流水线',
        '3. ctx.shared.rows 留给下一节点使用',
      ],
      run: async (ctx) => {
        const loaded = await ctx.host.loadSheetRows('Sheet2');
        if (!loaded.rows.length) throw new Error('「预投物料号」（Sheet2）没有物料数据！');
        ctx.shared.rows = loaded.rows;
        ctx.log(`读取完成：${loaded.rows.length} 行（工作表 Sheet2 / 网页显示名：预投物料号）`);
      },
    },
    {
      name: '计算可用月数与展示层分类',
      // 展示层镜像：与 VBA ca_库存水平分类 主过程 ---4~---6 同口径（规则逐行一致）
      steps: [
        '4. 遍历 rows：每行读 C列预投物料号 + H列下一个月预测 + I列M15账面结余 + J列年底账面结余',
        '5. 计算 M15可用月数 = I/H（H>0 才算），年底可用月数 = J/H；H=0 → 显示「可用很久」',
        '6. 按 M15可用月数判定五级：<1 严重不足 / 1-1.5 基本JIT / 1.5-12 较多 / 12-24 过多 / >24 或 H=0 非常多 / 其余不填充；统计每级数与严重不足物料号清单',
        '7. 把 { id, row_no, mat, level, m15m, ym, comment } 装入 ctx.shared.items 给下一节点',
      ],
      run: async (ctx) => {
        const { items, stats, sevList } = fc_库存水平分类计算(ctx.shared.rows);
        ctx.shared.items = items; ctx.shared.stats = stats; ctx.shared.sevList = sevList;
        ctx.log(`分类规则：M15可用月数 = M15账面结余 ÷ 下一个月预测（H列>0 才计算）`);
        ctx.log(`分类完成：严重不足 ${stats.sev} / 基本JIT ${stats.jit} / 较多 ${stats.more} / 过多 ${stats.over} / 非常多 ${stats.very} / 不填充 ${stats.none}，共 ${stats.total} 个物料`);
        if (sevList.length) ctx.log(`严重不足（<1月）物料号：${sevList.join(', ')}`);
      },
    },
    {
      name: '套用展示层颜色与批注',
      // 写 localStorage（in-memory 数据）与广播（不写源数据；展示层渲染）
      steps: [
        '0. ctx.stopped() 检查：被用户停止时返回（不写存储）',
        '1. 组装 payload：{ at, stats, sevList, items } 时间戳+统计+逐行结果',
        '2. localStorage.setItem("inventoryLevelResult", JSON.stringify(payload)) —— 缓存本轮结果供 fcLoadSheet 渲染',
        '3. window.fcApplyInventoryLevel(payload) —— 通知预投·预投物料号 页立即套色+注；不存在则静默跳过',
        '4. ctx.log 输出已套色物料数与可见页面',
      ],
      run: async (ctx) => {
        if (ctx.stopped()) { ctx.log('已停止：未套用颜色（可回滚恢复）'); return; }
        // 与 demand_forecast 的写回不同：颜色/批注是网页展示层结果，
        // 存 localStorage 并广播给工作表页（fcLoadSheet 渲染时读取），
        // Excel 源数据不变 —— 与 VBA 只填 C 列颜色/批注不改数据一致。
        const payload = {
          at: new Date().toLocaleString('zh-CN'),
          stats: ctx.shared.stats,
          sevList: ctx.shared.sevList,
          items: ctx.shared.items,   // { id, mat, level, m15m, ym, comment }
        };
        try { localStorage.setItem('inventoryLevelResult', JSON.stringify(payload)); }
        catch (e) { throw new Error('结果保存失败：' + (e.message || e)); }
        if (typeof window !== 'undefined' && typeof window.fcApplyInventoryLevel === 'function') {
          window.fcApplyInventoryLevel(payload);   // 通知工作表页立即套色
        }
        ctx.log(`已为 ${payload.items.length} 个物料的预投物料号单元格套用五级颜色与批注（预投·预投物料号 页可见）`);
      },
    },
    {
      name: '刷新预投物料号页面',
      // 网页侧动作：与 z2 节点 5 同款
      steps: [
        'ctx.host.refreshSheet("Sheet2") —— 通知 fcLoadSheet 重新拉取数据并应用 inventoryLevelResult 缓存',
        '「预投·预投物料号」页面数据与展示层颜色/批注均已刷新',
      ],
      run: async (ctx) => {
        ctx.host.refreshSheet('Sheet2');
        ctx.log('「预投·预投物料号」页面数据与颜色已刷新');
      },
    },
  ],
});

/* 任务 4：分类预投计算（直调 xlsm VBA da/db 两模块并行 + 同步 + 刷新）
 * --------------------------------------------------------------------
 * 设计原则（用户 2026-10-17 明令）：da_分类2预投计算逻辑1 与 db_分类1预投计算逻辑2
 * 两模块相互独立、并行执行——节点 2 为并行组（框架 Promise.all 并发下发，
 * 后端 VBA_RUN_LOCK 排队保证同工作簿 COM 串行安全），两分支写 Sheet2 K 列
 * 的行集不重叠（da 只处理 A 列分类="2"、db 只处理分类="1"），数据无冲突。
 * 节点拆细（同 inventory_level 先例：VBA 前置校验/网页侧动作各自成节点）：
 *   1. 读取并校验预投物料号数据（网页侧前置：分类1/2 行数统计 + H/I/J 就绪校验）
 *   2. 运行前提醒：Excel 弹窗需人工处理（纯提醒，不代处理——用户 2026-10-17 明令）
 *   3. 并行计算：da × db（并行组）
 *   4. 同步VBA结果到网页
 *   5. 刷新页面数据
 *
 * da_分类2预投计算逻辑1 模块头备注（分类="2" 物料）：
 *   【数据列】A:分类(="2") B:阈值 F:配套性 G:批量性 H:下月预测 I:M15结余 J:年底结余 K:输出
 *   【条件判断】I>0 且 J>0 为共同前提：① I,J 均>基准值→待确认；② I,J 均≤→预投
 *   日期=I/H×30；③ I>且J<→预投日期=J/H×30；④ I<且J>→订单倒挂需提前；⑤ 其他→跳过
 *   【预投计算】日期=min(当天+日期因子/H×30天, 当天+35天) 提前到周五；数量=B-J
 *   批量性处理，=0 输出"无需预投"；【配套性-就高原则】组内共享结果；【基准值】
 *   H<B 用 B 值 / H>B 弹窗输入 / H=B 用 B 值
 *
 * db_分类1预投计算逻辑2 模块头备注（分类="1" 物料）：
 *   【数据列】A:分类(="1") B:阈值 C:预投物料号 D:代用物料号 E:配额 F:配套性
 *   G:批量性(倍数) H:下月预测(基准值) I:M15账面结余 J:年底账面结余 K:输出结果 L:倍率系数
 *   【核心计算】前提 I>0 且 J>0，阈值=H×L：① I,J 均>阈值→无需预投；② I>且J<→
 *   预投数量=阈值-J 日期按 J/H；③ I<且J>→订单倒挂需提前；④ I,J 均≤→双笔/单笔子条件
 *   （H*1.5-maxIJ>=0 分两笔：第一笔 Ceil(H*1.5-maxIJ,G) 日期=当天+minIJ/(H*1.5)*30
 *   提前到周五、第二笔 Ceil(H*(L-1.5),G) 日期=30天后+(H*L-maxIJ)/H*30 提前到周五；
 *   <0 单笔：Ceil(H*L-maxIJ,G) 日期=30天后+(H*L-minIJ)/H*30 提前到周五）；⑤ 其他→跳过
 *   【配套组】F 列数字前缀+字母（"2da"/"fa"）同字母同组；配套数量=主数量×(配F/主F)；
 *   配套物料先按自身 I/J 判断是否需预投，非直接继承
 *
 * 前提：本机 Excel 可用；⚠️ da 在 H>B 时会弹 Excel InputBox 要基准值——网页端
 * **只提醒、不代处理**（用户明令），弹窗需用户在 Excel 里人工输入；VBA 只改内存
 * 不 Save，同步必须走 COM。
 * ==================================================================== */

AutoTaskKit.register({
  id: 'classify_precast',
  name: '分类预投计算（运行 VBA da_分类2预投计算逻辑1 ∥ db_分类1预投计算逻辑2）',
  desc: '按两 VBA 模块头备注并行计算：da（分类=2，基准值+条件①~⑤+配套就高）∥ db（分类=1，阈值=H×L+条件①~⑤含双笔/单笔+配套按比例），各自把预投结果写入 Sheet2 K 列（行集不重叠、可安全并行），跑完后同步 Sheet2 到网页并刷新。运行中 Excel 若弹 InputBox（da 的 H>B 场景）请人工输入基准值——网页端只提醒不代处理。',
  impactSheets: ['Sheet2'],   // 声明受影响工作表 → 运行前自动快照，供「回滚恢复」
  nodes: [
    {
      name: '读取并校验预投物料号数据',
      // 网页侧前置校验（同 inventory_level 节点1 先例）：分类行数统计 + H/I/J 就绪校验
      steps: [
        '1. ctx.host.loadSheetRows("Sheet2") 读网页 forecast_store 全量行',
        '2. 统计 A 列分类="1" 与分类="2" 的物料行数（两并行分支各自的处理对象）',
        '3. 校验 H 列（下一个月预测）就绪：分类行存在但 H 全为空/0 → 提示先跑「Resb需求调取与预测」（K 列结果会失真）',
        '4. rows 存入 ctx.shared，供并行分支行数与后续节点使用',
      ],
      run: async (ctx) => {
        const loaded = await ctx.host.loadSheetRows('Sheet2');
        if (!loaded.rows.length) throw new Error('「预投物料号」（Sheet2）没有物料数据！');
        let n1 = 0, n2 = 0, hEmpty = 0;
        for (const r of loaded.rows) {
          const c = String(r.data['分类'] ?? '').trim();
          if (c === '1') n1++;
          else if (c === '2') n2++;
          else continue;
          const h = parseFloat(r.data['下一个月预测']);
          if (isNaN(h) || h <= 0) hEmpty++;
        }
        const nClass = n1 + n2;
        if (!nClass) throw new Error('Sheet2 没有分类="1"/"2" 的物料行，无法分类预投计算！');
        ctx.shared.class1Count = n1;
        ctx.shared.class2Count = n2;
        ctx.shared.rows = loaded.rows;
        ctx.log(`读取完成：${loaded.rows.length} 行；分类=1 共 ${n1} 行（db 处理）、分类=2 共 ${n2} 行（da 处理）`);
        if (hEmpty > 0) {
          ctx.log(`⚠ ${hEmpty}/${nClass} 个分类物料 H 列（下一个月预测）为空或 0 —— K 列结果将按 VBA 内部规则跳过/输出"无需预投"；如需准确结果请先跑「Resb需求调取及预测」`);
        }
      },
    },
    {
      name: '运行前提醒：Excel 弹窗需人工处理',
      // 纯提醒节点（用户 2026-10-17 明令：不要处理弹窗，只做提醒作用）
      steps: [
        '1. da_分类2预投计算逻辑1 在 H列值 > B列值 时会弹 Excel InputBox 要求输入基准值（默认值=B 列值）',
        '2. 网页端只做提醒、不代处理：弹窗出现时请在 Excel 窗口人工输入/确认（直接回车=采用默认值）',
        '3. db_分类1预投计算逻辑2 无弹窗；两个 VBA 结束时各弹一个自动关闭统计框（3 秒，无需处理）',
        '4. 确认 Excel 已打开 yutouyouhuaa-ribbon.xlsm 且无残留对话框后，进入并行计算节点',
      ],
      run: async (ctx) => {
        ctx.log('提醒：da 分支运行中若 Excel 弹出「输入基准值」InputBox，请人工在 Excel 里输入/回车（网页端不代处理）');
        ctx.log('提醒：两个 VBA 结束时各有一个 3 秒自动关闭的统计弹框，无需处理');
        if (typeof showToast === 'function') {
          showToast('运行提醒：Excel 弹「输入基准值」请人工输入/回车（da 分支 H>B 时）');
        }
      },
    },
    {
      name: '并行计算：分类2预投计算（da）× 分类1预投计算（db）',
      // 框架并发能力：parallel 分支 Promise.all 并发下发；后端 VBA_RUN_LOCK 排队执行
      parallel: [
        {
          name: '分类2预投计算（da）',
          // steps = da 模块逐 Sub 过程映射（模块头备注 + 主过程内分节），忠实搬运零业务代码
          steps: [
            '1.（Sub 分类2预投计算·主过程）取 Sheet2 数据范围（A 列末行）→ 读 A1:L 到数组 → 建 dic配套/dic结果/dic物料行号/dic待确认 字典',
            '2.（主过程·第一遍扫描）识别 A 列分类="2" 的物料：物料号进 dic物料行号、F 列配套值建 dic配套 组映射',
            '3.（主过程·第二遍扫描）逐物料判定基准值：H<B 用 B 值 / H>B 弹 InputBox 人工输入（默认=B）/ H=B 用 B 值 → 存 dic基准值',
            '4.（Sub 处理条件判断）前提 I>0 且 J>0：① I,J 均>基准值→标记待确认；② I≤且J≤→调计算（日期=I/H×30）；③ I>且J<→调计算（日期=J/H×30）；④ I<且J>→"订单倒挂需提前"；⑤ 其他→跳过',
            '5.（Sub 计算预投日期＋计算预投数量）日期=min(当天+日期因子/H×30天, 当天+35天) 提前到周五；数量=B−J 按 G 列批量性 Ceiling_Math 取整（≤G 取 G）',
            '6.（Sub 输出结果并处理配套＋处理配套物料）组"数量(yyyy.mm.dd)"文本（=0→"无需预投"）写 dic结果；配套组就高原则——配套物料先按自身 I/J 判定，确实无需预投→"无需预投"，否则继承主物料结果',
            '7.（主过程·写回）dic结果 逐物料写 Sheet2 K 列（第11列），dic待确认→"无需预投"；AutoCloseMsgBox 弹自动关闭统计框（处理数/跳过数）',
          ],
          run: async (ctx) => {
            ctx.log(`开始执行（处理 ${ctx.shared.class2Count ?? '?'} 个分类=2 物料）…`);
            const res = await ctx.host.runVba('da_分类2预投计算逻辑1.分类2预投计算', []);
            ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），分类2 预投结果已写入 Sheet2 K 列`);
          },
        },
        {
          name: '分类1预投计算（db）',
          // steps = db 模块逐 Sub/Function 过程映射（模块头备注 + 主过程内分节），忠实搬运零业务代码
          steps: [
            '1.（Sub 分类1预投计算·主过程）取 Sheet2 数据范围 → 读 A1:L 到数组 → 建 dic配套组/dicF系数/dicFH字母/dic结果/dic物料行号 字典',
            '2.（主过程·第一遍扫描＋Sub 解析F值建立配套/解析F系数）识别分类="1" 物料；F 列"数字前缀+字母"解析：数字=F 系数、UCase(字母)=配套组键，同字母归 dic配套组',
            '3.（主过程·第二遍扫描）逐物料读 H/I/J/G/L（已有配套结果跳过；L≤0 按 1；H≤0 跳过计数）→ 基准值=H',
            '4.（Sub 处理条件判断）阈值=H×L，前提 I>0 且 J>0：① I,J 均>阈值→"无需预投"；② I>且J<→数量=阈值−J；③ I<且J>→"订单倒挂需提前"（同步配套文本）；④ I≤且J≤→子条件；⑤ 其他→跳过',
            '5.（Sub 处理条件4_1_1 / 处理条件4_1_2）判断用 maxIJ、日期用 minIJ：H*1.5−maxIJ≥0 → 双笔（Sub 生成双笔输出文本：第一笔 Ceil(H*1.5−maxIJ,G) 日=当天+minIJ/(H*1.5)*30、第二笔 Ceil(H*(L−1.5),G) 日=30天后+(H*L−maxIJ)/H*30）；<0 → 单笔（日=30天后+(H*L−minIJ)/H*30）',
            '6.（Sub 计算预投日期＋提前到周五＋批量取整）日期超 35/65 天封顶、非周五回退到周五；数量按 G 列 Ceiling_Math 批量取整',
            '7.（Sub 处理配套 / 处理配套4_1_1 / 输出配套单笔结果）配套数量=主数量×(配F系数/主F系数) 取整；配套物料先按自身 I/J 与阈值=配H×配L 判定，确实无需预投→跳过或加"***物料配套预投"提示，否则写结果',
            '8.（主过程·写回）dic结果 写 Sheet2 K 列（第11列，WrapText 自动换行）；AutoCloseMsgBox 弹自动关闭统计框（处理数/跳过数）',
          ],
          run: async (ctx) => {
            ctx.log(`开始执行（处理 ${ctx.shared.class1Count ?? '?'} 个分类=1 物料）…`);
            const res = await ctx.host.runVba('db_分类1预投计算逻辑2.分类1预投计算', []);
            ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），分类1 预投结果已写入 Sheet2 K 列`);
          },
        },
      ],
    },
    {
      name: '同步VBA结果到网页',
      // 网页侧动作：da/db 均已跑完后一次性同步 Sheet2（K 列结果）
      steps: [
        '读 xlsm Sheet2 UsedRange（COM）—— K 列已含 da/db 两模块的分类预投计算结果',
        '按 row_no 保留行 id 写回网页「预投物料号」表',
      ],
      run: async (ctx) => {
        const res = await ctx.host.syncVbaSheets(['Sheet2']);
        const n = res.synced && res.synced['Sheet2'];
        if (!n) throw new Error('同步完成但未返回 Sheet2 行数');
        ctx.log(`xlsm Sheet2 已同步到网页：${n} 行（含 K 列分类预投计算结果）`);
      },
    },
    {
      name: '刷新页面数据',
      // 网页侧动作：与 z2/inventory 节点同款
      steps: [
        '重载「预投·预投物料号」工作表页显示（K 列预投结果）',
      ],
      run: async (ctx) => {
        ctx.host.refreshSheet('Sheet2');
        ctx.log('「预投·预投物料号」页面已刷新');
      },
    },
  ],
});

/* 任务 5：R3标准格式文本构建及脚本创建订单（直调 xlsm VBA fz 总按钮）
 * --------------------------------------------------------------------
 * 设计原则（节点拆细规则，用户 2026-10-16 明令）：VBA 有几个主要功能，
 * 节点就拆几个。fz_创建采购申请计划订单总按钮 主过程只有两步大调用：
 *   1. fa_构建ME51B标准格式文本.构建ME51B标准格式文本
 *      （Sheet2 C/E/G/K 列筛选+拆单 → "根据配额拆单" G 列标准格式文本）
 *   2. fb_标准格式输出到txt.标准格式输出到txt
 *      （G 列按首字符分类写两个 txt → 格式A 跑 SAP ME51B 脚本回填
 *       "预投订单汇总" A/B/C 列；格式B 调 Python create_PPorder_from_XHPOtxt.py
 *       回填"预投订单汇总" A/B/C 列）
 *   3. 同步 xlsm 工作表到网页（根据配额拆单 G 列 + 预投订单汇总 A/B/C 列）
 *   4. 刷新页面数据
 * ⚠️ 运行前提（fb 内部会真连 SAP GUI）：
 *   - 已登录 SAP 且 SAP GUI Scripting 可用（格式A 才会触发 ME51B 导入）；
 *   - 本机 python 可用（格式B 计划订单走 create_PPorder_from_XHPOtxt.py）；
 *   - fb 结束时弹一个 MsgBox 完成框——网页端**只提醒、不代处理**（用户明令）。
 * ==================================================================== */

AutoTaskKit.register({
  id: 'r3_create_orders',
  name: 'R3标准格式文本构建及脚本创建订单（运行 VBA fz_创建采购申请计划订单总按钮）',
  desc: '按 fz 总按钮的调用顺序逐节点执行 VBA：fa 构建ME51B标准格式文本（Sheet2 筛选+按批量性拆单 → 根据配额拆单 G 列）→ fb 标准格式输出到txt（分类写 采购me51b文本格式.txt / X创建计划订单.txt → 格式A 跑 SAP ME51B 脚本、格式B 调 Python 脚本，结果回填 预投订单汇总 A/B/C 列）→ 同步到网页 → 刷新。运行中 SAP/Python/完成提示弹窗请人工确认——网页端只提醒不代处理。',
  impactSheets: ['根据配额拆单', '预投订单汇总'],   // 声明受影响工作表 → 运行前自动快照，供「回滚恢复」
  nodes: [
    {
      name: '构建ME51B标准格式文本（fa）',
      // steps = fa 模块逐 Sub/Function 过程映射（模块头备注 + 主过程内分节），忠实搬运零业务代码
      steps: [
        '1.（Sub 构建ME51B标准格式文本·主过程）取 Sheet2（C物料/E配额/G批量性/K日期数量）与 qqqq（L前缀/M物料号）数据范围；建「物料号→前缀(0或3)」Scripting.Dictionary',
        '2.（Function 是否包含日期数量格式）逐行筛选：物料号与 K 列非空且 K 列含"数量(yyyy.mm.dd)"括号格式才进入构建（支持多行 600(2026.06.12)/1200(2026.07.10)）',
        '3.（Function 判断是否特殊格式）按 E 列配额值分流：含 X(22*)/X(24*)/X(29*) 且 * 后紧跟 ")" → 格式B（计划订单）；否则 → 格式A（采购申请）',
        '4.（Function 解析日期数量文本＋提取完整表达式）按括号配对逐个拆出日期数量表达式，向前回溯数字起点',
        '5.（Function 提取数字部分＋提取日期部分）从表达式分离数量与 yyyy.mm.dd 日期',
        '6.（主过程·拆单）数量÷G列批量性 ≥2 → 拆成 2 单各取半、非整数倍向上取整到批量性倍数；<2 或批量性=0 → 不拆',
        '7.（Function 左填充空格＋格式化日期＋右填充零＋日期减七天）格式A=前缀+18位物料号+13位数量+8位日期+000219AYUTOU；格式B=18位物料号+8位日期+2空格+8位数量+8位(日期-7天)+2空格+YUTOU+2空格',
        '8.（主过程·输出）格式A 在前格式B 在后合并，写"根据配额拆单"工作表 G2 起的 G 列（先 Clear 旧值，AutoFit 列宽）',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('fa_构建ME51B标准格式文本.构建ME51B标准格式文本', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），标准格式文本已写入「根据配额拆单」G 列`);
      },
    },
    {
      name: '标准格式输出到txt并执行脚本（fb）',
      // steps = fb 模块逐 Sub/Function 过程映射（模块头备注 + 主过程内分节），忠实搬运零业务代码
      steps: [
        '1.（Sub 标准格式输出到txt·分类）读"根据配额拆单" G 列（跳过标题行）：首字符 0/3 → 格式A 数组，其他 → 格式B 数组（ReDim Preserve 动态扩容）',
        '2.（主过程·写txt）格式A → D:\\ai_project\\forecast\\脚本文件\\采购me51b文本格式.txt，格式B → X创建计划订单.txt（FSO CreateTextFile 覆盖式，无数据则清空文件）',
        '3.（Function 执行SAP脚本·连接）格式A 文件非空才触发：GetObject("SAPGUI")→ScriptingEngine→Connection→Session 四级判空，窗口存在() 校验 wnd[1]/[2]/[3]',
        '4.（Function 执行SAP脚本·ME51B 操作）最大化窗口 → /nzme51b → wnd[1] 设路径 D:\\ai_project\\forecast\\脚本文件 → wnd[3] 设文件名 采购me51b文本格式.txt → 逐层 btn[0] 确认导入 → 菜单选导出 radSPOPLI-SELFLAG[4,0]（SAP 自动创建采购申请并把结果复制到剪切板）',
        '5.（Sub 从剪切板粘贴数据）DataObject GetFromClipboard 取 SAP 回执 → 按行分割：Function 提取物料号（"I:"与"采购申请号"之间）→ A 列，Function 提取采购申请号（"采购申请号"与"被创建"之间纯数字）→ B 列',
        '6.（Sub 调用创建PI的Python脚本并回填结果）格式B 文件非空才触发：先删旧 计划订单号结果.txt → WScript.Shell 隐藏窗口执行 create_PPorder_from_XHPOtxt.py → 读结果文件，物料号=输入行前 18 位 → A 列，计划订单号 → B 列',
        '7.（两个回填 Sub 收尾）"预投订单汇总" A 列末行续写，C 列对有效行填当天 yyyy.mm.dd；主过程末尾 MsgBox 弹完成统计（采购申请 N 条/计划订单 N 条）',
      ],
      run: async (ctx) => {
        ctx.log('提醒：格式A 将真连 SAP GUI 执行 /nzme51b 导入（需已登录 SAP）；格式B 将隐藏窗口调用 Python 创建计划订单');
        ctx.log('提醒：fb 结束会弹完成统计 MsgBox，请在 Excel 里人工关闭（网页端不代处理）');
        if (typeof showToast === 'function') {
          showToast('运行提醒：SAP ME51B 脚本与 Python 计划订单执行中，结束弹窗请人工关闭');
        }
        const res = await ctx.host.runVba('fb_标准格式输出到txt.标准格式输出到txt', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），采购申请/计划订单结果已回填「预投订单汇总」A/B/C 列`);
      },
    },
    {
      name: '同步VBA结果到网页',
      // 网页侧动作：fa 写 G 列（根据配额拆单）+ fb 回填 A/B/C 列（预投订单汇总），一次同步
      steps: [
        '读 xlsm「根据配额拆单」「预投订单汇总」UsedRange（COM）—— G 列标准格式文本 + A/B/C 列回填结果',
        '按 row_no 保留行 id 写回网页对应工作表存储',
      ],
      run: async (ctx) => {
        const res = await ctx.host.syncVbaSheets(['根据配额拆单', '预投订单汇总']);
        const n1 = res.synced && res.synced['根据配额拆单'];
        const n2 = res.synced && res.synced['预投订单汇总'];
        if (!n1 && !n2) throw new Error('同步完成但未返回任何工作表行数');
        ctx.log(`xlsm 已同步到网页：根据配额拆单 ${n1 ?? 0} 行（G 列标准格式）、预投订单汇总 ${n2 ?? 0} 行（A/B/C 列创建结果）`);
      },
    },
    {
      name: '刷新页面数据',
      // 网页侧动作：与 z2/inventory/classify 节点同款
      steps: [
        '重载「预投·配额拆单」「预投·预投订单汇总」工作表页显示（G 列文本 + 订单回填结果）',
      ],
      run: async (ctx) => {
        ctx.host.refreshSheet('根据配额拆单');
        ctx.host.refreshSheet('预投订单汇总');
        ctx.log('「预投·配额拆单」「预投·预投订单汇总」页面已刷新');
      },
    },
  ],
});

/* 任务 6：备份汇总（直调 xlsm VBA ga_备份汇总 + 网页侧新增功能）
 * --------------------------------------------------------------------
 * 设计原则（节点拆细规则，用户 2026-10-16 明令）：VBA 有几个主要功能，
 * 节点就拆几个。ga_备份汇总 的总入口「执行全部备份」依次调用 4 个备份 Sub：
 *   1. ga_备份汇总.备份Sheet2到总表（Sheet2 全量 → yutouyouhuaa_汇总总表.xlsm
 *      Sheet2，Q 列写时间戳 yyyy.mm.dd_hh:mm:ss）
 *   2. ga_备份汇总.备份配额拆单到总表（根据配额拆单 A:H 列 → 总表同名表，H 列时间戳）
 *   3. ga_备份汇总.备份需求汇总到总表（需求汇总 → 总表同名表，H 列时间戳）
 *   4. ga_备份汇总.备份预投订单汇总到总表（预投订单汇总 → 总表同名表，H 列时间戳）
 *   （时间戳取自 Function 获取时间戳；备份目录/总表路径为模块头常量：
 *    D:\ai_project\forecast\预投备份\yutouyouhuaa_汇总总表.xlsm）
 * 【新增节点（VBA 没有的，用户 2026-10-17 明令）】运行备份汇总以后，把「预投
 * 自动化操作」下面所有模块的表格，以表格形式另存到
 * D:\ai_project\forecast\web——beifen_zong\<时间戳目录>\（新建目录），每表一个
 * xlsx 文件。数据源=网页侧双通道存储（loadSheetRows），后端
 * POST /api/backup/save_tables 落盘 openpyxl。
 * ⚠️ 运行前提：Excel 已打开 yutouyouhuaa-ribbon.xlsm 且
 * D:\ai_project\forecast\预投备份\yutouyouhuaa_汇总总表.xlsm 存在（总表含
 * Sheet2/根据配额拆单/需求汇总/预投订单汇总 四张同名工作表）；VBA 内部
 * Workbooks.Open 会打开总表窗口，运行期间请勿人工关闭。
 * ==================================================================== */

AutoTaskKit.register({
  id: 'backup_summary',
  name: '备份汇总（运行 VBA ga_备份汇总 + 网页表格另存）',
  desc: '按 ga_备份汇总 模块「执行全部备份」的调用顺序逐节点执行 VBA：备份Sheet2到总表（Q列时间戳）→ 备份配额拆单到总表（H列时间戳）→ 备份需求汇总到总表（H列时间戳）→ 备份预投订单汇总到总表（H列时间戳），汇总总表=D:\\ai_project\\forecast\\预投备份\\yutouyouhuaa_汇总总表.xlsm；跑完后执行网页侧新增功能——把预投自动化操作下所有模块的工作表（Sheet2/根据配额拆单/qqqq/需求汇总/合同池数量/预投订单汇总）以表格形式另存到 D:\\ai_project\\forecast\\web——beifen_zong\\<时间戳目录>\\ 并刷新。',
  // 受影响工作表（用户 2026-10-15 明令「受影响数据直达链接」所有任务页遵守）：
  // 备份虽只读网页工作表，但读取/另存的就是这 4 张表的数据 → 声明后顶栏显示直达链接
  //（双击跳工作表页，行灰显对比运行前后差异）+ 运行前自动快照供「回滚恢复」。
  impactSheets: ['Sheet2', '根据配额拆单', '需求汇总', '预投订单汇总'],
  nodes: [
    {
      name: '备份Sheet2到总表（ga）',
      // steps = VBA 备份Sheet2到总表 过程逐段映射（模块头常量 + 主过程内分节），忠实搬运零业务代码
      steps: [
        '1.（Function 获取时间戳）生成 yyyy.mm.dd_hh:mm:ss 时间戳',
        '2.（主过程·打开总表）源=ThisWorkbook.Sheets("Sheet2")；打开 D:\\ai_project\\forecast\\预投备份\\yutouyouhuaa_汇总总表.xlsm 的 Sheet2',
        '3.（主过程·行数对齐）源 UsedRange 行数与目标现有行数取 Max（空表判定兜底为 1）',
        '4.（主过程·复制）源数据 UsedRange.Value 整体写目标 A1 起（等维 Resize）',
        '5.（主过程·时间戳）Q 列（第17列）按最大行数逐行写时间戳',
        '6.（主过程·收尾）wb目标.Save 后 Close SaveChanges:=False；出错时兜底关闭不保存',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('ga_备份汇总.备份Sheet2到总表', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），Sheet2 已备份到汇总总表（Q 列时间戳）`);
      },
    },
    {
      name: '备份配额拆单到总表（ga）',
      // steps = VBA 备份配额拆单到总表 过程逐段映射
      steps: [
        '1.（Function 获取时间戳）生成 yyyy.mm.dd_hh:mm:ss 时间戳',
        '2.（主过程·表检查）源=根据配额拆单（找不到即退出）；目标=总表内同名工作表（找不到则关总表退出）',
        '3.（主过程·行数对齐）A 列 End(xlUp) 取源/目标行数取 Max（空表兜底 1）',
        '4.（主过程·复制）源 A1:H<源行数> 整体写目标 A1 起（只备 A:H 八列）',
        '5.（主过程·时间戳）H 列（第8列）按最大行数逐行写时间戳',
        '6.（主过程·收尾）wb目标.Save 后 Close SaveChanges:=False',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('ga_备份汇总.备份配额拆单到总表', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），根据配额拆单 A:H 已备份到汇总总表（H 列时间戳）`);
      },
    },
    {
      name: '备份需求汇总到总表（ga）',
      // steps = VBA 备份需求汇总到总表 过程逐段映射
      steps: [
        '1.（Function 获取时间戳）生成 yyyy.mm.dd_hh:mm:ss 时间戳',
        '2.（主过程·表检查）源=需求汇总（找不到即退出）；目标=总表内同名工作表（找不到则关总表退出）',
        '3.（主过程·行数对齐）A 列 End(xlUp) 取源/目标行数取 Max（空表兜底 1）',
        '4.（主过程·复制）源 UsedRange.Value 整体写目标 A1 起',
        '5.（主过程·时间戳）H 列（第8列）按最大行数逐行写时间戳',
        '6.（主过程·收尾）wb目标.Save 后 Close SaveChanges:=False',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('ga_备份汇总.备份需求汇总到总表', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），需求汇总已备份到汇总总表（H 列时间戳）`);
      },
    },
    {
      name: '备份预投订单汇总到总表（ga）',
      // steps = VBA 备份预投订单汇总到总表 过程逐段映射
      steps: [
        '1.（Function 获取时间戳）生成 yyyy.mm.dd_hh:mm:ss 时间戳',
        '2.（主过程·表检查）源=预投订单汇总（找不到即退出）；目标=总表内同名工作表（找不到则关总表退出）',
        '3.（主过程·行数对齐）A 列 End(xlUp) 取源/目标行数取 Max（空表兜底 1）',
        '4.（主过程·复制）源 UsedRange.Value 整体写目标 A1 起',
        '5.（主过程·时间戳）H 列（第8列）按最大行数逐行写时间戳',
        '6.（主过程·收尾）wb目标.Save 后 Close SaveChanges:=False（四个备份 Sub 至此构成「执行全部备份」总链）',
      ],
      run: async (ctx) => {
        const res = await ctx.host.runVba('ga_备份汇总.备份预投订单汇总到总表', []);
        ctx.log(`VBA 执行完成（${res.started} → ${res.ended}），预投订单汇总已备份到汇总总表（H 列时间戳）——VBA 四表备份全部完成`);
      },
    },
    {
      name: '网页表格另存到时间戳目录（新增功能）',
      // 【VBA 没有的新增功能，用户 2026-10-17 明令】预投自动化操作下所有模块的工作表
      // 以表格形式另存到 D:\ai_project\forecast\web——beifen_zong\<时间戳目录>\
      steps: [
        '1. 依次 ctx.host.loadSheetRows 读预投自动化操作全部 6 张工作表：Sheet2/根据配额拆单/qqqq/需求汇总/合同池数量/预投订单汇总（网页双通道存储，headers+rows）',
        '2. POST /api/backup/save_tables：后端新建 D:\\ai_project\\forecast\\web——beifen_zong\\<yyyymmdd_hhmmss>\\ 时间戳目录',
        '3. 每张工作表写成一个 xlsx（openpyxl，表头行+数据行，纯数字还原数值），文件名=物理表名',
        '4. 返回 {dir, saved}，逐表日志「表名 → 完整路径（N 行）」；任一表读取失败中断流水线',
      ],
      run: async (ctx) => {
        const sheets = ['Sheet2', '根据配额拆单', 'qqqq', '需求汇总', '合同池数量', '预投订单汇总'];
        const payload = [];
        for (const sh of sheets) {
          const loaded = await ctx.host.loadSheetRows(sh);
          payload.push({ sheet: sh, headers: loaded.headers || [], rows: loaded.rows });
          ctx.shared[`n_${sh}`] = loaded.rows.length;
        }
        const res = await fetch('/api/backup/save_tables', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sheets: payload }),
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(data.detail || '网页表格另存失败');
        ctx.log(`已创建时间戳目录：${data.dir}`);
        for (const [sh, info] of Object.entries(data.saved || {})) {
          ctx.log(`${sh} → ${info.file}（${info.rows} 行）`);
        }
        ctx.shared.savedDir = data.dir;
        ctx.shared.savedCount = Object.keys(data.saved || {}).length;
      },
    },
    {
      name: '刷新页面数据',
      // 网页侧动作：备份不改动工作表数据，刷新仅同步展示（与其他任务收尾节点同款）
      steps: [
        '重载「预投·预投物料号」等受影响工作表页显示（备份为只读操作，数据不变）',
      ],
      run: async (ctx) => {
        ctx.host.refreshSheet('Sheet2');
        ctx.host.refreshSheet('根据配额拆单');
        ctx.host.refreshSheet('需求汇总');
        ctx.host.refreshSheet('预投订单汇总');
        ctx.log(`备份汇总完成：VBA 四表备份到汇总总表 + 网页 ${ctx.shared.savedCount ?? '?'} 张表格另存到 ${ctx.shared.savedDir ?? '时间戳目录'}`);
      },
    },
  ],
});

/* ====================================================================
 * 任务 7：APD-F30-34#箱自动化（运行 RPA 脚本 34_f30计划，2026-10-18）
 * --------------------------------------------------------------------
 * 设计原则（与预投自动化操作各任务一致）：不重写业务算法。网页端 = 触发器 +
 * 节点流水线状态看板，RPA Python 脚本 = 执行引擎；**每个脚本对应一个节点**，
 * 节点 steps 忠实搬运脚本头注释的流程与异常分支（零业务代码改写）。
 *
 * 执行引擎位置：工作区 web_auto/计划平台操作/34_f30计划/（源码复刻自
 * D:\ai_project\aps_mes_dingding_SAP；Playwright + CDP 驱动专用 Chrome，
 * 目标平台 app.example.com 制造 APS 计划平台）。后端 /api/rpa/* 负责
 * 启动/状态/停止，节点 run() 只做触发与日志跟踪。
 *
 * 六步编排（AA → AB → AC → AD → AE → AF，任一节点失败=流水线中断，
 * 与脚本内部「截图+钉钉邮件+停止后续」的异常语义一致）：
 *   AA 登录：CDP attach 127.0.0.1:9222 专用 Chrome（cookie 持久化免扫码）→
 *      连不上则 DETACHED 拉起专用 Chrome 再 attach → 已登录直达 /
 *      平台登录页点真实登录控件（白屏自愈刷新） / 钉钉扫码页等人工（默认 300s）
 *      → 登录失败 = 整页截图 + 钉钉邮件告知，流程终止
 *   AB 接受任务清单+BOM解析：点【APD/独立需求-334-F30-发布-生效】正则提取
 *      右下角「共 X 条」→ T=1.5×X 秒 → X=0：点「重置」后结束全部操作 /
 *      X≠0：BOM解析→查看解析日志轮询到「完成」→ 刷新 2 次复查 X →
 *      X 未归零 = 截图邮件 + best-effort 自动恢复（批量查询物料号→重新解析/退回）
 *   AC 关键计划项评估：先「重置」→ ZZZ=0 直跳跟踪页 / 全选→标记1→SL01（只这一段策略）→
 *      → 全选→计算要求交货日（loading 基线+列值变化双判据）→ 策略指定完毕 →
 *      ZZZ 每 20s 刷新轮询归零（15 轮）→ 切「计划评估动作跟踪」
 *      （2026-10-27 用户明令：删除原第二段「SF09 + 需录入的部件供方代码 200038」及其「重置→确认」闸门，
 *        SL01 确定成功后直接进入计算要求交货日 → 策略指定完毕）
 *   AD 计划评估动作跟踪：进度 100% 读「共 c 条」→ c=0 直跳下单前确认 /
 *      全部消警 → c 仍≠0：退至关键计划评估 → 按「关键物料号 批量查询」只选本次
 *      退回行 → SF08 SMEC不下单 → 策略指定完毕 → 回跟踪页读 Z：
 *      Z=0 收尾 / Z≠0 = 截图邮件「Z 未归零，等指令」，流程停止
 *   AE 下单前确认：切页+「重置」+记 B0 → 全选→确认下单（浮层同名项）→
 *      B=0 切「下单动作跟踪」收尾 / B≠0 每 10s 轮询 10 轮 → 仍≠0 截图邮件停止
 *   AF 下单动作跟踪：读 B0 → 等 B0/100 分钟 → 进度 100% 且右下角=0：
 *      回「接收任务清单」收尾 / 进度 100% 且≠0：截图邮件+退至下单前确认 /
 *      进度≠100%：每 30s 轮询 50 次 → 仍≠100% 截图邮件停止
 *
 * 平台约定（RPA 项目记忆2，改节点前必读）：
 *   - Surely Table 虚拟表格：行=DIV.surely-table-cell 按 top 归并，无 table/tr/td；
 *   - 自绘组件（antd 下拉/Select）只认真实鼠标：统一 page.mouse（CDP，isTrusted=true）；
 *   - 「重 置」DOM 文本带空格，匹配须去空白；冷加载禁 page.reload()，用页签刷新图标；
 *   - 所有分支共用同一个已登录 Chrome：严禁原版与拷贝版、或多节点并行同时跑。
 * ==================================================================== */

/* 运行一个 RPA 脚本并跟踪到结束（POST /api/rpa/run → 轮询 /api/rpa/status 增量日志）。
 * 返回 = 脚本 exit 0；非 0 退出或启动失败 → 抛错（流水线中断，符合脚本异常语义）。
 * 例外：exit 5 = 脚本自报「本轮无待处理数据，按需求提前收工」——不抛失败，抛 { earlyFinish:true }
 *       让流水线优雅收尾（后续节点一律不再执行），既不算失败也不发失败通知。
 * 2026-10-20：branch 由入参传入（f30_34 / f30_03 共用同一实现）。 */
const RPA退出码_提前收工 = 5;   // 与 AB 脚本的「退出码_无数据收工」一致（X 首次=0 ⇒ 重置后结束全部操作）

async function apd运行脚本(ctx, branch, 脚本名) {
  // 2026-10-25：把「邮箱通知」开关状态 + 一级标题（邮件主题前缀）+ 流水线批次号传给后端。
  // 同一轮「运行」的所有节点共用同一个 pipeline_id，后端把各节点截图归档到该批次，
  // 流水线结束时由 autoRunTask 调 /api/rpa/pipeline/finish 统一发一封邮件（替代钉钉实时截图推送）。
  const emailChk = document.getElementById('autoEmailNotify');
  const emailNotify = !!(emailChk && emailChk.checked);
  const res = await fetch('/api/rpa/run', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      branch: branch, script: 脚本名,
      email_notify: emailNotify,
      email_to: 'demo.user1@example.com',
      task_title: (typeof autoState !== 'undefined' && autoState.pageName) || '',
      pipeline_id: (typeof autoState !== 'undefined' && autoState.pipelineId) || '',
    }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.detail || `RPA 启动失败（${脚本名}）`);
  const runId = data.run_id;
  ctx.log(`已启动 ${脚本名}（run_id=${runId}，解释器=${data.interpreter}）`);
  ctx.log('提示：脚本操作的是真实 Chrome 里的计划平台页面，期间请勿人工操作同一页面');
  let seen = 0;
  let 已提示通知 = false;
  for (;;) {
    if (ctx.stopped()) {
      try {
        await fetch('/api/rpa/stop', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ run_id: runId }),
        });
      } catch (e) { /* 停止指令发送失败也不阻断停止流程 */ }
      throw Object.assign(new Error('已停止：已向 RPA 脚本发送终止指令（进程树）'), { skipped: true });
    }
    const st = await fetch(`/api/rpa/status?run_id=${encodeURIComponent(runId)}&after=${seen}`)
      .then(r => r.json().catch(() => null))
      .catch(() => null);
    if (!st) {
      ctx.log('⚠ 状态查询失败（网络波动），2s 后重试…');
    } else {
      if (st.truncated) ctx.log('…（早期日志已滚出窗口，以下为最近日志）');
      if (st.notifying && !已提示通知) {
        已提示通知 = true;
        ctx.log('【邮箱】节点失败：网页端正在把本次运行结果（含各步截图）打包发邮件到 demo.user2@example.com…');
      }
      for (const line of st.lines || []) ctx.log(line);
      seen = st.next || seen;
      if (!st.running) {
        // 2026-10-25：节点结束不再实时截发钉钉（用户要求改用邮件通知）；
        // 各步截图由脚本产生，运行结束后由后端「邮箱通知」统一打包发邮件（附件一/二…）。
        // exit 5 = 无待处理数据提前收工：不算失败，交给流水线优雅停止后续节点
        if (st.exit_code === RPA退出码_提前收工) {
          ctx.log(`⏹ ${脚本名} 提前收工（exit ${RPA退出码_提前收工}：本轮无待处理数据，按需求结束全部操作）`
            + '——后续节点不再执行');
          throw Object.assign(
            new Error(`无待处理数据，流程提前结束（${脚本名} 已点「重置」并收工）`),
            { earlyFinish: true });
        }
        if (st.exit_code !== 0) {
          throw new Error(`脚本退出码 ${st.exit_code}（失败详情见上方日志；钉钉通知由脚本内部或网页端补发，请查收）`);
        }
        ctx.log(`✓ ${脚本名} 正常结束（exit 0）`);
        return;
      }
    }
    await new Promise(r => setTimeout(r, 2000));
  }
}

AutoTaskKit.register({
  id: 'apd_f30_34',
  name: 'APD-F30-34#箱自动化（运行 RPA 脚本 34_f30计划）',
  desc: '按 34_f30计划 脚本组的六步编排逐节点触发浏览器自动化：AA 登录（CDP attach 专用 Chrome，扫码等人工）→ AB 接受任务清单+BOM解析（334-F30 页签，X=0 重置收尾 / X 未归零异常恢复）→ AC 关键计划项评估（标记1→SL01→计算要求交货日→策略指定完毕，ZZZ 轮询归零）→ AD 计划评估动作跟踪（消警/退至关键计划评估/SF08，Z 未归零邮件停流程）→ AE 下单前确认（确认下单，B 轮询）→ AF 下单动作跟踪（等 B0/100 分钟，进度 100% 收尾/退回）。运行环境：web_auto/计划平台操作/34_f30计划/（Playwright+CDP 真实鼠标）；任一节点失败即中断流水线，与脚本「截图+钉钉邮件+停止后续」语义一致——脚本内部未及发邮件时由网页端自动补发（日志尾部+最新截图，防重），保证失败必达钉钉。',
  rpaNotify: true,            // 网页侧失败补发钉钉通知 + 任务页显示「测试通知」按钮
  rpaBranch: 'f30_34',
  nodes: [
    {
      name: '登录计划平台（AA）',
      steps: [
        '1. 连接外部 Chrome：优先 CDP attach http://127.0.0.1:9222（专用 Chrome，user-data-dir=%LOCALAPPDATA%\\aps_plan_chrome，cookie 永久保存，一般免扫码）',
        '2. 连不上则 subprocess DETACHED 拉起专用 Chrome（--remote-debugging-port=9222 + 目标地址）再 attach',
        '3. 打开任务清单页：已登录（页面出现任务清单特征）→ 直接进第 6 步',
        '4. 平台 /login 页：点击真实登录控件（文案含"登录"）；连续 2 轮无控件则刷新页面（白屏自愈）；同时扫描钉钉 OAuth 入口并跳转',
        '5. 钉钉扫码页（login.dingtalk.com）：等待人工扫码，默认 300s（APS_MANUAL_LOGIN_WAIT 可调）',
        '6. 登录成功：reload → 等页面渲染 → 点「接收任务清单」后的刷新按钮（[aria-label="reload"]，CDP 真实鼠标点击）',
        '异常：登录失败不继续后续流程 —— 整页截图存 公共_钉钉通知/附件截图/ + 钉钉邮件通知 demo.user3@example.com「计划平台无法登录，后续步骤无法继续！！！」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_34', 'AA_登录.py'); },
    },
    {
      name: '接受任务清单+BOM解析（AB）',
      steps: [
        '1. 复用 AA 的 CDP attach 逻辑连外部 Chrome（本脚本不做登录）',
        '2. 点侧边栏「接收任务清单」→ 点【APD/独立需求-334-F30-发布-生效】，正则提取右下角「共 X 条」的 X',
        '3. 按 X 换算默认等待时间 T=1.5×X 秒（BOM 解析轮询预算）',
        '4. X 首次=0：本轮无待解析数据 → 点工具栏「重置」→ 结束全部操作，不继续后面任何操作',
        '5. X≠0：点【BOM解析】split button → hover 展开下拉 → 点【查看解析日志】→ 等弹窗；最上一行状态 ∈ {已完成, 完成} → 关弹窗 → 表头全选 → 展开 BOM解析下拉 → 点【BOM解析】触发解析',
        '6. 等 T 秒后每 10s 轮询【查看解析日志】，直到最上一行状态为「完成',
        '7. 「完成」后点刷新按钮 2 次 → 复查右下角 X；X 未归零 ⇒ 判异常：整页截图 + 钉钉邮件（合同号/梯号/箱号明细）；best-effort 自动恢复：取 APD 表三列 → 切「关键计划项评估」→ 自定义查询（物料料号·批量查询）→ 共 X 条=0 不操作 / ≠0：全选→重新BOM解析→BOM退回→重置',
        '平台约定：Surely Table 虚拟表格（DIV 行按 top 归并）；自绘组件只认真实鼠标（page.mouse，isTrusted=true）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_34', 'AB_34箱_接受任务清单_BOM解析.py'); },
    },
    {
      name: '关键计划项评估（AC）',
      steps: [
        '0. 上来先点工具栏「重置」（所有点击的最前面，清上一轮残留筛选/搜索，睡 2s 等重渲染）',
        '1. 读右下角「共 ZZZ 条」：ZZZ=0 → 直接切侧边栏「计划评估动作跟踪」，不做后续任何操作',
        '2. ZZZ≠0：全选打✔ → 点「标记」两级菜单（标记→数字 1-8）选「1」→ 再全选 → 「指定策略」弹窗「指定计划」→ 策略编号选「SL01」→ 确定',
        '3. SL01 确定成功后不再指定第二段策略（2026-10-27 用户明令：AC 只指定 SL01 一段，不做 SF09 / 部件供方代码，也不要「重置确认闸门」）→ 全选 → 点「计算要求交货日」浮层同名菜单项 → 等计算完成（判据=绿色提示「所有批次交货日计算完成!」∪ 要求交货日列值变化，≤45s 无迹象重选，最多 5 次）',
        '4. 再全选 → 「策略指定完毕」浮层同名项 → 弹确认框点「确定」→ 等 10s',
        '5. 每 20s 刷新读 ZZZ 直到 0（最多 15 轮；仍≠0 ⇒ 整页截图+邮件告知，停止后续所有操作）',
        '6. ZZZ=0 → 点侧边栏「计划评估动作跟踪」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_34', 'AC_34箱_关键计划向评估.py'); },
    },
    {
      name: '计划评估动作跟踪（AD）',
      steps: [
        '1. 确保进入「计划评估动作跟踪」页 → 等进度条 100% → 读右下角「共 c 条」（进度未达 100% 轮询等，默认最长 3 分钟，超时邮件告知并停止）',
        '2. c=0：直接点侧边栏「下单前确认」（流程结束，AD 退出 0）',
        '3. c≠0：点工具栏「全部消警」→ 等 3s → 再读 c；c 此时=0 → 按 c=0 分支收尾',
        '4. c 仍≠0：点「退至关键计划评估」（本页无勾选列，直接退回当前全部行）→ 切「关键计划项评估」→ 容忍冷加载等退回数据渲染',
        '5. 只选中本次退回行：从跟踪页记「关键物料号」→ 关键计划项评估页「自定义查询」（.anticon-file-search）按「关键物料号·批量查询」过滤 → 全选 → 指定策略「SF08 SMEC不下单」→ 确定 →「策略指定完毕」→ 确认后右下角≠0 则停止；=0 才继续',
        '6. 切回「计划评估动作跟踪」→ 等进度 100% → 读「共 Z 条」：Z=0 → 按 c=0 分支点「下单前确认」收尾',
        '7. Z≠0：不再继续任何操作 → 整页截图 + 钉钉邮件告知「Z 未归零，等指令」，流程停止',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_34', 'AD_34箱_计划评估动作跟踪.py'); },
    },
    {
      name: '下单前确认（AE）',
      steps: [
        '0. 先切侧边栏「下单前确认」（AD 末尾虽切过，但页签可能被带走；切不到则邮件并停止）',
        '1. 点「重置」清残留 → 等 2S → 先读右下角 B0（仅记录，落模块级 AE.B0 供 AF 开头调用）',
        '2. 全选（表头那 1 个复选框）→ 点「确认下单」dropdown-trigger → 点浮层同名菜单项「确认下单」→ 点「确认」',
        '3. 等 3S 读「共 B 条」：B=0 → 点侧边栏「下单动作跟踪」，流程结束（AE 退出 0）',
        '4. B≠0：每 10S 轮询（刷新+重读 B）最多 10 轮；任一轮 B=0 → 切「下单动作跟踪」收尾',
        '5. 10 轮后 B 仍≠0：整页截图 + 钉钉邮件告知，流程停止（AE 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_34', 'AE_34箱_下单前确认.py'); },
    },
    {
      name: '下单动作跟踪（AF）',
      steps: [
        '0. 先读「下单动作跟踪」页右下角 B0（调用 AE.读取B0，同一「共 B0 条」口径）',
        '1. 按 B0 数量等待 B0/100 分钟，然后检查进度条是否 100%',
        '2. 进度=100% 且右下角=0：直接点侧边栏「接收任务清单」，流程结束（AF 退出 0）',
        '3. 进度=100% 且右下角≠0：先整页截图+邮件告知；再全选 → 点「退至下单前确认」→ 弹确认框则点它（AF 退出 0）',
        '4. 进度≠100%：每 30S 轮询检查，最多 50 次；仍≠100% ⇒ 整页截图+邮件告知，停止后续操作（AF 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_34', 'AF_34箱_下单动作跟踪.py'); },
    },
  ],
});

/* ====================================================================
 * 任务 8：APD-F30-03#箱自动化（运行 RPA 脚本 03_f30计划，2026-10-20）
 * --------------------------------------------------------------------
 * 与任务 7（apd_f30_34）同源同构：工作区 web_auto/计划平台操作/03_f30计划/
 * 的 5 个脚本（AB_3号箱/AC_3号箱/AD_3号箱/AE_3号箱/AF_3号箱）是 34 箱脚本的
 * 3 号箱同款副本（MD5 与源目录 D:\ai_project\aps_mes_dingding_SAP\计划平台操作\03_f30计划
 * 逐一一致；算法零转写）。执行引擎复用同一套 AA_登录.py + CDP attach 9222 专用 Chrome，
 * 差异仅在平台页签号：**3号箱走【APD/独立需求-303-F30-发布-生效】，34箱走 334**。
 *
 * 共享 Chrome 约束（沿用 RPA 项目记忆2·六实证）：03 与 34 两个分支共用同一个
 * 已登录 Chrome，**严禁两个分支同时运行**（后端按 branch 单实例栅栏 + 前端流水线
 * 串行编排，任一节点失败即中断）。
 *
 * 六步编排（AA → AB → AC → AD → AE → AF，逻辑与 34 箱完全一致，仅页签文案 303 化）：
 *   AA 登录：CDP attach 127.0.0.1:9222 专用 Chrome（cookie 持久化免扫码）→
 *      连不上则 DETACHED 拉起专用 Chrome 再 attach → 已登录直达 / 平台登录页点
 *      真实登录控件 / 钉钉扫码页等人工（默认 300s）→ 登录失败=整页截图+钉钉邮件，终止
 *   AB 接受任务清单+BOM解析：点【APD/独立需求-303-F30-发布-生效】正则提取右下角
 *      「共 X 条」→ T=1.5×X → X=0 重置收尾 / X≠0：BOM解析→查看解析日志轮询到「完成」→
 *      刷新 2 次复查 X → X 未归零=截图邮件+best-effort 自动恢复
 *   AC 关键计划项评估：重置→ZZZ=0 直跳跟踪页 / 全选→标记1→SL01（只这一段策略）→
 *      计算要求交货日→策略指定完毕→ZZZ 轮询归零
 *      （2026-10-27 用户明令：删除原第二段「SF09 + 需录入的部件供方代码 200038」及其「重置→确认」闸门）
 *   AD 计划评估动作跟踪：进度 100% 读 c 条→c=0 直跳 / 消警 / 退至关键计划评估按批量查询
 *      只选本次退回行→SF08 SMEC不下单→回跟踪页读 Z：Z=0 收尾 / Z≠0 邮件停止
 *   AE 下单前确认：重置→记 B0→全选→确认下单→B=0 切下单动作跟踪 / B≠0 轮询 10 轮→仍≠0 邮件停止
 *   AF 下单动作跟踪：读 B0→等 B0/100 分钟→进度 100% 且=0 收尾 / 进度 100% 且≠0 邮件+退至
 *      下单前确认 / 进度≠100% 每 30S 轮询 50 次→仍≠100% 邮件停止
 * ==================================================================== */
AutoTaskKit.register({
  id: 'apd_f30_03',
  name: 'APD-F30-03#箱自动化（运行 RPA 脚本 03_f30计划）',
  desc: '按 03_f30计划 脚本组的六步编排逐节点触发浏览器自动化：AA 登录（CDP attach 专用 Chrome，扫码等人工）→ AB 接受任务清单+BOM解析（303-F30 页签，X=0 重置收尾 / X 未归零异常恢复）→ AC 关键计划项评估（标记1→SL01→计算要求交货日→策略指定完毕，ZZZ 轮询归零）→ AD 计划评估动作跟踪（消警/退至关键计划评估/SF08，Z 未归零邮件停流程）→ AE 下单前确认（确认下单，B 轮询）→ AF 下单动作跟踪（等 B0/100 分钟，进度 100% 收尾/退回）。运行环境：web_auto/计划平台操作/03_f30计划/（Playwright+CDP 真实鼠标）；任一节点失败即中断流水线，与脚本「截图+钉钉邮件+停止后续」语义一致——脚本内部未及发邮件时由网页端自动补发（日志尾部+最新截图，防重），保证失败必达钉钉。注意：03 与 34 两个分支共用同一个已登录 Chrome（CDP 9222），严禁同时运行。',
  rpaNotify: true,            // 网页侧失败补发钉钉通知 + 任务页显示「测试通知」按钮
  rpaBranch: 'f30_03',
  nodes: [
    {
      name: '登录计划平台（AA）',
      steps: [
        '1. 连接外部 Chrome：优先 CDP attach http://127.0.0.1:9222（专用 Chrome，user-data-dir=%LOCALAPPDATA%\\aps_plan_chrome，cookie 永久保存，一般免扫码）',
        '2. 连不上则 subprocess DETACHED 拉起专用 Chrome（--remote-debugging-port=9222 + 目标地址）再 attach',
        '3. 打开任务清单页：已登录（页面出现任务清单特征）→ 直接进第 6 步',
        '4. 平台 /login 页：点击真实登录控件（文案含"登录"）；连续 2 轮无控件则刷新页面（白屏自愈）；同时扫描钉钉 OAuth 入口并跳转',
        '5. 钉钉扫码页（login.dingtalk.com）：等待人工扫码，默认 300s（APS_MANUAL_LOGIN_WAIT 可调）',
        '6. 登录成功：reload → 等页面渲染 → 点「接收任务清单」后的刷新按钮（[aria-label="reload"]，CDP 真实鼠标点击）',
        '异常：登录失败不继续后续流程 —— 整页截图存 公共_钉钉通知/附件截图/ + 钉钉邮件通知 demo.user4@example.com「计划平台无法登录，后续步骤无法继续！！！」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_03', 'AA_登录.py'); },
    },
    {
      name: '接受任务清单+BOM解析（AB）',
      steps: [
        '1. 复用 AA 的 CDP attach 逻辑连外部 Chrome（本脚本不做登录）',
        '2. 点侧边栏「接收任务清单」→ 点【APD/独立需求-303-F30-发布-生效】，正则提取右下角「共 X 条」的 X',
        '3. 按 X 换算默认等待时间 T=1.5×X 秒（BOM 解析轮询预算）',
        '4. X 首次=0：本轮无待解析数据 → 点工具栏「重置」→ 结束全部操作，不继续后面任何操作',
        '5. X≠0：点【BOM解析】split button → hover 展开下拉 → 点【查看解析日志】→ 等弹窗；最上一行状态 ∈ {已完成, 完成} → 关弹窗 → 表头全选 → 展开 BOM解析下拉 → 点【BOM解析】触发解析',
        '6. 等 T 秒后每 10s 轮询【查看解析日志】，直到最上一行状态为「完成',
        '7. 「完成」后点刷新按钮 2 次 → 复查右下角 X；X 未归零 ⇒ 判异常：整页截图 + 钉钉邮件（合同号/梯号/箱号明细）；best-effort 自动恢复：取 APD 表三列 → 切「关键计划项评估」→ 自定义查询（物料料号·批量查询）→ 共 X 条=0 不操作 / ≠0：全选→重新BOM解析→BOM退回→重置',
        '平台约定：Surely Table 虚拟表格（DIV 行按 top 归并）；自绘组件只认真实鼠标（page.mouse，isTrusted=true）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_03', 'AB_3号箱_接受任务清单_BOM解析.py'); },
    },
    {
      name: '关键计划项评估（AC）',
      steps: [
        '0. 上来先点工具栏「重置」（所有点击的最前面，清上一轮残留筛选/搜索，睡 2s 等重渲染）',
        '1. 读右下角「共 ZZZ 条」：ZZZ=0 → 直接切侧边栏「计划评估动作跟踪」，不做后续任何操作',
        '2. ZZZ≠0：全选打✔ → 点「标记」两级菜单（标记→数字 1-8）选「1」→ 再全选 → 「指定策略」弹窗「指定计划」→ 策略编号选「SL01」→ 确定',
        '3. SL01 确定成功后不再指定第二段策略（2026-10-27 用户明令：AC 只指定 SL01 一段，不做 SF09 / 部件供方代码，也不要「重置确认闸门」）→ 全选 → 点「计算要求交货日」浮层同名菜单项 → 等计算完成（判据=绿色提示「所有批次交货日计算完成!」∪ 要求交货日列值变化，≤45s 无迹象重选，最多 5 次）',
        '4. 再全选 → 「策略指定完毕」浮层同名项 → 弹确认框点「确定」→ 等 10s',
        '5. 每 20s 刷新读 ZZZ 直到 0（最多 15 轮；仍≠0 ⇒ 整页截图+邮件告知，停止后续所有操作）',
        '6. ZZZ=0 → 点侧边栏「计划评估动作跟踪」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_03', 'AC_3号箱_关键计划向评估.py'); },
    },
    {
      name: '计划评估动作跟踪（AD）',
      steps: [
        '1. 确保进入「计划评估动作跟踪」页 → 等进度条 100% → 读右下角「共 c 条」（进度未达 100% 轮询等，默认最长 3 分钟，超时邮件告知并停止）',
        '2. c=0：直接点侧边栏「下单前确认」（流程结束，AD 退出 0）',
        '3. c≠0：点工具栏「全部消警」→ 等 3s → 再读 c；c 此时=0 → 按 c=0 分支收尾',
        '4. c 仍≠0：点「退至关键计划评估」（本页无勾选列，直接退回当前全部行）→ 切「关键计划项评估」→ 容忍冷加载等退回数据渲染',
        '5. 只选中本次退回行：从跟踪页记「关键物料号」→ 关键计划项评估页「自定义查询」（.anticon-file-search）按「关键物料号·批量查询」过滤 → 全选 → 指定策略「SF08 SMEC不下单」→ 确定 →「策略指定完毕」→ 确认后右下角≠0 则停止；=0 才继续',
        '6. 切回「计划评估动作跟踪」→ 等进度 100% → 读「共 Z 条」：Z=0 → 按 c=0 分支点「下单前确认」收尾',
        '7. Z≠0：不再继续任何操作 → 整页截图 + 钉钉邮件告知「Z 未归零，等指令」，流程停止',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_03', 'AD_3号箱_计划评估动作跟踪.py'); },
    },
    {
      name: '下单前确认（AE）',
      steps: [
        '0. 先切侧边栏「下单前确认」（AD 末尾虽切过，但页签可能被带走；切不到则邮件并停止）',
        '1. 点「重置」清残留 → 等 2S → 先读右下角 B0（仅记录，落模块级 AE.B0 供 AF 开头调用）',
        '2. 全选（表头那 1 个复选框）→ 点「确认下单」dropdown-trigger → 点浮层同名菜单项「确认下单」→ 点「确认」',
        '3. 等 3S 读「共 B 条」：B=0 → 点侧边栏「下单动作跟踪」，流程结束（AE 退出 0）',
        '4. B≠0：每 10S 轮询（刷新+重读 B）最多 10 轮；任一轮 B=0 → 切「下单动作跟踪」收尾',
        '5. 10 轮后 B 仍≠0：整页截图 + 钉钉邮件告知，流程停止（AE 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_03', 'AE_3号箱_下单前确认.py'); },
    },
    {
      name: '下单动作跟踪（AF）',
      steps: [
        '0. 先读「下单动作跟踪」页右下角 B0（调用 AE.读取B0，同一「共 B0 条」口径）',
        '1. 按 B0 数量等待 B0/100 分钟，然后检查进度条是否 100%',
        '2. 进度=100% 且右下角=0：直接点侧边栏「接收任务清单」，流程结束（AF 退出 0）',
        '3. 进度=100% 且右下角≠0：先整页截图+邮件告知；再全选 → 点「退至下单前确认」→ 弹确认框则点它（AF 退出 0）',
        '4. 进度≠100%：每 30S 轮询检查，最多 50 次；仍≠100% ⇒ 整页截图+邮件告知，停止后续操作（AF 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'f30_03', 'AF_3号箱_下单动作跟踪.py'); },
    },
  ],
});

/* ====================================================================
 * 任务 9：APD-X-32#箱自动化（运行 RPA 脚本 32_x计划，2026-10-21）
 * --------------------------------------------------------------------
 * 与任务 8（apd_f30_03）同源同构：工作区 web_auto/计划平台操作/32_x计划/
 * 的 6 个脚本（AB_32箱/AC_32箱/AD_32箱/AE_32箱/AF_32箱/AG_邮件发送）是
 * 32 号箱的 RPA 编排副本（算法零转写，直接拷贝自源目录
 * D:\ai_project\aps_mes_dingding_SAP\计划平台操作\32_x计划）。AA_登录.py
 * 仍共用根级（各分支共享，CDP attach 9222 专用 Chrome）。差异仅在平台
 * 页签号：**32 号箱走【532-P0-32#箱】**，且比 3/34 号箱多一个收尾节点
 * AG 邮件发送（命中 32 号箱订单时给相关同事发钉钉邮件通知）。
 *
 * 共享 Chrome 约束：x_32 与 f30_03 / f30_34 三个分支共用同一个已登录
 * Chrome（CDP 9222），严禁同时运行（后端按 branch 单实例栅栏 + 前端串行
 * 编排，任一节点失败即中断）。32 号箱七步编排（AA → AB → AC → AD → AE →
 * AF → AG，逻辑与 3 号箱一致，仅页签 532-P0-32#箱 化 + 末步 AG 发信）：
 *   AA 登录 / AB 接受任务清单+BOM解析（532-P0-32#箱，X=0 重置收尾 / X 未
 *   归零异常恢复）/ AC 关键计划项评估 / AD 计划评估动作跟踪 / AE 下单前确认
 *   / AF 下单动作跟踪 / AG 订单下发邮件通知（构建按单树页命中 32 号箱→钉钉）
 * ==================================================================== */
AutoTaskKit.register({
  id: 'apd_x_32',
  name: 'APD-X-32#箱自动化（运行 RPA 脚本 32_x计划）',
  desc: '按 32_x计划 脚本组的七步编排逐节点触发浏览器自动化：AA 登录（CDP attach 专用 Chrome，扫码等人工）→ AB 接受任务清单+BOM解析（532-P0-32#箱 页签，X=0 重置收尾 / X 未归零异常恢复）→ AC 关键计划项评估（标记1→SL01→计算要求交货日→策略指定完毕，ZZZ 轮询归零）→ AD 计划评估动作跟踪（消警/退至关键计划评估/SF08，Z 未归零邮件停流程）→ AE 下单前确认（确认下单，B 轮询）→ AF 下单动作跟踪（等 B0/100 分钟，进度 100% 收尾/退回）→ AG 订单下发邮件通知（构建按单树页命中 32 号箱订单→钉钉邮件给相关同事）。运行环境：web_auto/计划平台操作/32_x计划/（Playwright+CDP 真实鼠标）；任一节点失败即中断流水线，脚本内部未及发邮件时由网页端自动补发（日志尾部+最新截图，防重），保证失败必达钉钉。注意：x_32 与 f30_03 / f30_34 三个分支共用同一个已登录 Chrome（CDP 9222），严禁同时运行。',
  rpaNotify: true,            // 网页侧失败补发钉钉通知 + 任务页显示「测试通知」按钮
  rpaBranch: 'x_32',
  nodes: [
    {
      name: '登录计划平台（AA）',
      steps: [
        '1. 连接外部 Chrome：优先 CDP attach http://127.0.0.1:9222（专用 Chrome，user-data-dir=%LOCALAPPDATA%\\aps_plan_chrome，cookie 永久保存，一般免扫码）',
        '2. 连不上则 subprocess DETACHED 拉起专用 Chrome（--remote-debugging-port=9222 + 目标地址）再 attach',
        '3. 打开任务清单页：已登录（页面出现任务清单特征）→ 直接进第 6 步',
        '4. 平台 /login 页：点击真实登录控件（文案含"登录"）；连续 2 轮无控件则刷新页面（白屏自愈）；同时扫描钉钉 OAuth 入口并跳转',
        '5. 钉钉扫码页（login.dingtalk.com）：等待人工扫码，默认 300s（APS_MANUAL_LOGIN_WAIT 可调）',
        '6. 登录成功：reload → 等页面渲染 → 点「接收任务清单」后的刷新按钮（[aria-label="reload"]，CDP 真实鼠标点击）',
        '异常：登录失败不继续后续流程 —— 整页截图存 公共_钉钉通知/附件截图/ + 钉钉邮件通知 demo.user5@example.com「计划平台无法登录，后续步骤无法继续！！！」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'x_32', 'AA_登录.py'); },
    },
    {
      name: '接受任务清单+BOM解析（AB）',
      steps: [
        '1. 复用 AA 的 CDP attach 逻辑连外部 Chrome（本脚本不做登录）',
        '2. 点侧边栏「接收任务清单」→ 点【532-P0-32#箱】，正则提取右下角「共 X 条」的 X',
        '3. 按 X 换算默认等待时间 T=1.5×X 秒（BOM 解析轮询预算）',
        '4. X 首次=0：本轮无待解析数据 → 点工具栏「重置」→ 结束全部操作，不继续后面任何操作',
        '5. X≠0：点【BOM解析】split button → hover 展开下拉 → 点【查看解析日志】→ 等弹窗；最上一行状态 ∈ {已完成, 完成} → 关弹窗 → 表头全选 → 展开 BOM解析下拉 → 点【BOM解析】触发解析',
        '6. 等 T 秒后每 10s 轮询【查看解析日志】，直到最上一行状态为「完成',
        '7. 「完成」后点刷新按钮 2 次 → 复查右下角 X；X 未归零 ⇒ 判异常：整页截图 + 钉钉邮件（合同号/梯号/箱号明细）；best-effort 自动恢复：取 APD 表三列 → 切「关键计划项评估」→ 自定义查询（物料料号·批量查询）→ 共 X 条=0 不操作 / ≠0：全选→重新BOM解析→BOM退回→重置',
        '平台约定：Surely Table 虚拟表格（DIV 行按 top 归并）；自绘组件只认真实鼠标（page.mouse，isTrusted=true）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'x_32', 'AB_32箱_接受任务清单_BOM解析.py'); },
    },
    {
      name: '关键计划项评估（AC）',
      steps: [
        '0. 上来先点工具栏「重置」（所有点击的最前面，清上一轮残留筛选/搜索，睡 2s 等重渲染）',
        '1. 读右下角「共 ZZZ 条」：ZZZ=0 → 直接切侧边栏「计划评估动作跟踪」，不做后续任何操作',
        '2. ZZZ≠0：全选打✔ → 点「标记」两级菜单（标记→数字 1-8）选「1」→ 再全选 → 「指定策略」弹窗「指定计划」→ 策略编号选「SL01」→ 确定',
        '3. SL01 确定成功后不再指定第二段策略（2026-10-27 用户明令：AC 只指定 SL01 一段，不做 SF09 / 部件供方代码，也不要「重置确认闸门」）→ 全选 → 点「计算要求交货日」浮层同名菜单项 → 等计算完成（判据=绿色提示「所有批次交货日计算完成!」∪ 要求交货日列值变化，≤45s 无迹象重选，最多 5 次）',
        '4. 再全选 → 「策略指定完毕」浮层同名项 → 弹确认框点「确定」→ 等 10s',
        '5. 每 20s 刷新读 ZZZ 直到 0（最多 15 轮；仍≠0 ⇒ 整页截图+邮件告知，停止后续所有操作）',
        '6. ZZZ=0 → 点侧边栏「计划评估动作跟踪」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'x_32', 'AC_32箱_关键计划向评估.py'); },
    },
    {
      name: '计划评估动作跟踪（AD）',
      steps: [
        '1. 确保进入「计划评估动作跟踪」页 → 等进度条 100% → 读右下角「共 c 条」（进度未达 100% 轮询等，默认最长 3 分钟，超时邮件告知并停止）',
        '2. c=0：直接点侧边栏「下单前确认」（流程结束，AD 退出 0）',
        '3. c≠0：点工具栏「全部消警」→ 等 3s → 再读 c；c 此时=0 → 按 c=0 分支收尾',
        '4. c 仍≠0：点「退至关键计划评估」（本页无勾选列，直接退回当前全部行）→ 切「关键计划项评估」→ 容忍冷加载等退回数据渲染',
        '5. 只选中本次退回行：从跟踪页记「关键物料号」→ 关键计划项评估页「自定义查询」（.anticon-file-search）按「关键物料号·批量查询」过滤 → 全选 → 指定策略「SF08 SMEC不下单」→ 确定 →「策略指定完毕」→ 确认后右下角≠0 则停止；=0 才继续',
        '6. 切回「计划评估动作跟踪」→ 等进度 100% → 读「共 Z 条」：Z=0 → 按 c=0 分支点「下单前确认」收尾',
        '7. Z≠0：不再继续任何操作 → 整页截图 + 钉钉邮件告知「Z 未归零，等指令」，流程停止',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'x_32', 'AD_32箱_计划评估动作跟踪.py'); },
    },
    {
      name: '下单前确认（AE）',
      steps: [
        '0. 先切侧边栏「下单前确认」（AD 末尾虽切过，但页签可能被带走；切不到则邮件并停止）',
        '1. 点「重置」清残留 → 等 2S → 先读右下角 B0（仅记录，落模块级 AE.B0 供 AF 开头调用）',
        '2. 全选（表头那 1 个复选框）→ 点「确认下单」dropdown-trigger → 点浮层同名菜单项「确认下单」→ 点「确认」',
        '3. 等 3S 读「共 B 条」：B=0 → 点侧边栏「下单动作跟踪」，流程结束（AE 退出 0）',
        '4. B≠0：每 10S 轮询（刷新+重读 B）最多 10 轮；任一轮 B=0 → 切「下单动作跟踪」收尾',
        '5. 10 轮后 B 仍≠0：整页截图 + 钉钉邮件告知，流程停止（AE 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'x_32', 'AE_32箱_下单前确认.py'); },
    },
    {
      name: '下单动作跟踪（AF）',
      steps: [
        '0. 先读「下单动作跟踪」页右下角 B0（调用 AE.读取B0，同一「共 B0 条」口径）',
        '1. 按 B0 数量等待 B0/100 分钟，然后检查进度条是否 100%',
        '2. 进度=100% 且右下角=0：直接点侧边栏「接收任务清单」，流程结束（AF 退出 0）',
        '3. 进度=100% 且右下角≠0：先整页截图+邮件告知；再全选 → 点「退至下单前确认」→ 弹确认框则点它（AF 退出 0）',
        '4. 进度≠100%：每 30S 轮询检查，最多 50 次；仍≠100% ⇒ 整页截图+邮件告知，停止后续操作（AF 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'x_32', 'AF_32箱_下单动作跟踪.py'); },
    },
    {
      name: '订单下发邮件通知（AG）',
      steps: [
        '0. 单实例守卫：与 AB~AF 共用「计划平台操作」锁（同一时刻只允许一个进程操作平台页面）',
        '1. 进入「构建按单树」页（页签→侧边栏→全页逐级找；已在则直接复用），进入判据=页面文本同时出现表头「根物料」+「根订单」',
        '2. 逐屏滚动读该页 Surely 虚拟表：表头文案 + 每行单元格（fix-left/center/fix-right 按 top 归并、left 还原列序）+ 右下角「共 X 条」。滚动容器取 .surely-table-vertical-scroll-viewport / body-viewport-container（实测 body/content 会被误判「已到底」，只采到首屏）；表头保留 48px 空 extra-cell 使列与数据行逐格对齐',
        '3. 任一「关键计划物料号」符合通配 *-*-*-32*（末段以 32 开头，26N3M16-111-9-32 与 26N3M16-111-9-321 都命中，5 段图号 26N3M16-111-9-8-32 同样命中；可用 APS_AG_ROOT_GLOB 调整）⇒ 用 DWS 最新版本命令发钉钉邮件',
        '4. 收件人 demo.user6@example.com,demo.user7@example.com,demo.user8@example.com,demo.user9@example.com；主题「32号箱已下达 {台数} 台」，台数=命中行「关键计划物料号」去重数；正文=HTML 表格，只含 6 列（关键计划物料号 / BOM数量 / 采购类型及特殊采购类 / 订单类型 / 生产订单号 / 生产订单基本完成时间），行只发命中的关键条目',
        '5. 未命中 ⇒ 仅记日志、不发邮件（退出 0），绝不无脑发信',
        '平台约定：复用 AC 基础工具（真实点击/整页截图/读取条数/滚动表格/重置）；切到侧边栏/页签后先「重置」（best-effort）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'x_32', 'AG_邮件发送.py'); },
    },
  ],
});

/* ====================================================================
 * 任务 10：GZT-X-03#箱自动化（运行 RPA 脚本 GZT-03-X，2026-10-23）
 * --------------------------------------------------------------------
 * 与任务 9（apd_x_32）同源同构：工作区 web_auto/计划平台操作/GZT-03-X/
 * 的 6 个脚本（AB_GZT_03_X/AC_GZT_03_X/AD_GZT_03_X/AE_GZT_03_X/AF_GZT_03_X/
 * AG_邮件发送）是 GZT-03-X 箱的 RPA 编排副本（算法零转写，直接拷贝自源目录
 * D:\ai_project\aps_mes_dingding_SAP\计划平台操作\GZT-03-X，md5 逐一一致）。
 * AA_登录.py 仍共用根级（各分支共享，CDP attach 9222 专用 Chrome）。
 * 差异点（相对 32箱）：
 *   1) 平台页签：**GZT 箱走【GZT-03-X】**（AB 第 113 行按钮_APD，
 *      可用 APS_BTN_APD 环境变量覆盖；源 32箱 为 532-P0-32#箱）；
 *   2) AC 关键计划项评估：计算完成后的勾选改为**排除式全选**（需求 2026-10-23）——
 *      先「表头全选」幂等补选（防计算后重渲染丢勾选态），再**按「根物料号」整组剔除**
 *      （需求 2026-10-23 晚，用户口径）：只看「关键物料号」符合 *GZT*-*-*-03* 的根行，
 *      其「要求交货日」为空 或 距今天>120 天 ⇒ 该根物料号下的所有行（根行+全部子行）
 *      一起取消勾选，其余行不受影响；复查残留⇒邮件中止。
 *      ⚠️ 2026-10-21 的「逐行剔除」已废弃：子行日期多半为空，会被大面积误取消（首跑即失败）。
 *      随后「策略指定完毕」不再全选（先全选=False，防把排除行勾回来）；
 *   3) AB 的 X 首次=0 分支 return 退出码 5（提前收工，与三箱同款修复）：
 *      网页端据此结束整条流水线，不再启动 AC/AD/AE/AF/AG；
 *   4) AG 邮件（用户 2026-10-23 两次改版）：收件人 demo.user10@example.com，
 *      主题「GZT已下达 {台数} 台」（台数=命中行「关键计划物料号」去重数）；
 *      正文 HTML 表格，只含 6 列（关键计划物料号/BOM数量/采购类型及特殊采购类/
 *      订单类型/生产订单号/生产订单基本完成时间），行只发命中的关键条目。
 *      判据为通配 *GZT*-*-*-03*（含 GZT 且末段以 03 开头，可用 APS_AG_ROOT_GLOB 调整）。
 *      注 1：2026-10-23 21:05 修复——此前误留 32 箱组的 -32 正则，导致真实根物料
 *      26GZT-C0NA-56-03 全部不匹配、「命中 0 条」邮件发不出。
 *      注 2：2026-10-23 21:20 改版——正文由「| 拼接整页文本」改为 6 列 HTML 表格，
 *      并修掉表头丢 extra-cell（列错位）与滚动容器选错（60 条只采 17 行）两处真因。
 *
 * 共享 Chrome 约束：gzt_03x 与 x_32 / f30_03 / f30_34 等分支共用同一个已登录
 * Chrome（CDP 9222），严禁同时运行（后端按 branch 单实例栅栏 + 前端串行
 * 编排，任一节点失败即中断）。七步编排 AA → AB → AC → AD → AE → AF → AG。
 * ==================================================================== */
AutoTaskKit.register({
  id: 'gzt_03x',
  name: 'GZT-X-03#箱自动化（运行 RPA 脚本 GZT-03-X）',
  desc: '按 GZT-03-X 脚本组的七步编排逐节点触发浏览器自动化：AA 登录（CDP attach 专用 Chrome，扫码等人工）→ AB 接受任务清单+BOM解析（【GZT-03-X】页签，X=0 提前收工 exit 5 / X 未归零异常恢复）→ AC 关键计划项评估（标记1→SL01→计算要求交货日→排除式全选：先表头全选，再按「根物料号」整组剔除（关键物料号符合 *GZT*-*-*-03* 的根行若要求交货日为空或距今天>120天，则该根物料号下所有行一起取消勾选）→策略指定完毕，ZZZ 轮询归零）→ AD 计划评估动作跟踪（消警/退至关键计划评估/SF08，Z 未归零邮件停流程）→ AE 下单前确认（确认下单，B 轮询）→ AF 下单动作跟踪（等 B0/100 分钟，进度 100% 收尾/退回）→ AG 订单下发邮件通知（构建按单树页命中根物料 *GZT*-*-*-03* → 钉钉邮件 demo_user）。运行环境：web_auto/计划平台操作/GZT-03-X/（Playwright+CDP 真实鼠标）；任一节点失败即中断流水线，脚本内部未及发邮件时由网页端自动补发（日志尾部+最新截图，防重），保证失败必达钉钉。注意：gzt_03x 与 x_32 / f30_03 / f30_34 等分支共用同一个已登录 Chrome（CDP 9222），严禁同时运行。',
  rpaNotify: true,            // 网页侧失败补发钉钉通知 + 任务页显示「测试通知」按钮
  rpaBranch: 'gzt_03x',
  nodes: [
    {
      name: '登录计划平台（AA）',
      steps: [
        '1. 连接外部 Chrome：优先 CDP attach http://127.0.0.1:9222（专用 Chrome，user-data-dir=%LOCALAPPDATA%\\aps_plan_chrome，cookie 永久保存，一般免扫码）',
        '2. 连不上则 subprocess DETACHED 拉起专用 Chrome（--remote-debugging-port=9222 + 目标地址）再 attach',
        '3. 打开任务清单页：已登录（页面出现任务清单特征）→ 直接进第 6 步',
        '4. 平台 /login 页：点击真实登录控件（文案含"登录"）；连续 2 轮无控件则刷新页面（白屏自愈）；同时扫描钉钉 OAuth 入口并跳转',
        '5. 钉钉扫码页（login.dingtalk.com）：等待人工扫码，默认 300s（APS_MANUAL_LOGIN_WAIT 可调）',
        '6. 登录成功：reload → 等页面渲染 → 点「接收任务清单」后的刷新按钮（[aria-label="reload"]，CDP 真实鼠标点击）',
        '异常：登录失败不继续后续流程 —— 整页截图存 公共_钉钉通知/附件截图/ + 钉钉邮件通知 demo.user11@example.com「计划平台无法登录，后续步骤无法继续！！！」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'gzt_03x', 'AA_登录.py'); },
    },
    {
      name: '接受任务清单+BOM解析（AB）',
      steps: [
        '1. 复用 AA 的 CDP attach 逻辑连外部 Chrome（本脚本不做登录）',
        '2. 点侧边栏「接收任务清单」→ 点【GZT-03-X】，正则提取右下角「共 X 条」的 X',
        '3. 按 X 换算默认等待时间 T=1.5×X 秒（BOM 解析轮询预算）',
        '4. X 首次=0：本轮无待解析数据 → 点工具栏「重置」→ 以退出码 5 提前收工（网页端据此结束整条流水线，不再启动后续节点）',
        '5. X≠0：点【BOM解析】split button → hover 展开下拉 → 点【查看解析日志】→ 等弹窗；最上一行状态 ∈ {已完成, 完成} → 关弹窗 → 表头全选 → 展开 BOM解析下拉 → 点【BOM解析】触发解析',
        '6. 等 T 秒后每 10s 轮询【查看解析日志】，直到最上一行状态为「完成',
        '7. 「完成」后点刷新按钮 2 次 → 复查右下角 X；X 未归零 ⇒ 判异常：整页截图 + 钉钉邮件（合同号/梯号/箱号明细）；best-effort 自动恢复：取 GZT 表三列 → 切「关键计划项评估」→ 自定义查询（物料料号·批量查询）→ 共 X 条=0 不操作 / ≠0：全选→重新BOM解析→BOM退回→重置',
        '平台约定：Surely Table 虚拟表格（DIV 行按 top 归并）；自绘组件只认真实鼠标（page.mouse，isTrusted=true）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'gzt_03x', 'AB_GZT_03_X_接受任务清单_BOM解析.py'); },
    },
    {
      name: '关键计划项评估（AC）',
      steps: [
        '0. 上来先点工具栏「重置」（所有点击的最前面，清上一轮残留筛选/搜索，睡 2s 等重渲染）',
        '1. 读右下角「共 ZZZ 条」：ZZZ=0 → 直接切侧边栏「计划评估动作跟踪」，不做后续任何操作',
        '2. ZZZ≠0：全选打✔ → 点「标记」两级菜单（标记→数字 1-8）选「1」→ 再全选 → 「指定策略」弹窗「指定计划」→ 策略编号选「SL01」→ 确定',
        '3. SL01 确定成功后不再指定第二段策略（2026-10-27 用户明令：AC 只指定 SL01 一段，不做 SF09 / 部件供方代码，也不要「重置确认闸门」）→ 全选 → 点「计算要求交货日」浮层同名菜单项 → 等计算完成（判据=绿色提示「所有批次交货日计算完成!」∪ 要求交货日列值变化，≤45s 无迹象重选，最多 5 次）',
        '4. 计算完成后改走「排除式全选」（需求 2026-10-23，替代旧「再全选」）：第①步表头全选（幂等补选，防计算后重渲染丢勾选态；失败=邮件中止不抢跑）；第②步按「根物料号」整组剔除——只看「关键物料号」符合 *GZT*-*-*-03* 的根行，若其「要求交货日」为空 或 距今天>120 天 ⇒ 该根物料号下的所有行（根行+全部子行）一起取消勾选，其余行不受影响（旧口径按单行判断会因子行日期为空而大面积误取消，2026-10-23 首跑即失败）；复查仍残留 ⇒ 截图+邮件中止。**没有不合格的根物料号 ⇒ 保持全选、继续往下执行**；**若全部根物料号都不合格（一行勾选都没剩）⇒ 立即停止后续所有操作，弹窗提醒 + 邮件说明到底是「要求交货日为空」还是「超 120 天」（列出各自根物料号）**；随后「策略指定完毕」浮层同名项（不再全选，防把排除行勾回来）→ 弹确认框点「确定」→ 等 10s',
        '5. 每 20s 刷新读 ZZZ 直到 0（最多 15 轮；仍≠0 ⇒ 整页截图+邮件告知，停止后续所有操作）',
        '6. ZZZ=0 → 点侧边栏「计划评估动作跟踪」',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'gzt_03x', 'AC_GZT_03_X_关键计划向评估.py'); },
    },
    {
      name: '计划评估动作跟踪（AD）',
      steps: [
        '1. 确保进入「计划评估动作跟踪」页 → 等进度条 100% → 读右下角「共 c 条」（进度未达 100% 轮询等，默认最长 3 分钟，超时邮件告知并停止）',
        '2. c=0：直接点侧边栏「下单前确认」（流程结束，AD 退出 0）',
        '3. c≠0：点工具栏「全部消警」→ 等 3s → 再读 c；c 此时=0 → 按 c=0 分支收尾',
        '4. c 仍≠0：点「退至关键计划评估」（本页无勾选列，直接退回当前全部行）→ 切「关键计划项评估」→ 容忍冷加载等退回数据渲染',
        '5. 只选中本次退回行：从跟踪页记「关键物料号」→ 关键计划项评估页「自定义查询」（.anticon-file-search）按「关键物料号·批量查询」过滤 → 全选 → 指定策略「SF08 SMEC不下单」→ 确定 →「策略指定完毕」→ 确认后右下角≠0 则停止；=0 才继续',
        '6. 切回「计划评估动作跟踪」→ 等进度 100% → 读「共 Z 条」：Z=0 → 按 c=0 分支点「下单前确认」收尾',
        '7. Z≠0：不再继续任何操作 → 整页截图 + 钉钉邮件告知「Z 未归零，等指令」，流程停止',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'gzt_03x', 'AD_GZT_03_X_计划评估动作跟踪.py'); },
    },
    {
      name: '下单前确认（AE）',
      steps: [
        '0. 先切侧边栏「下单前确认」（AD 末尾虽切过，但页签可能被带走；切不到则邮件并停止）',
        '1. 点「重置」清残留 → 等 2S → 先读右下角 B0（仅记录，落模块级 AE.B0 供 AF 开头调用）',
        '2. 全选（表头那 1 个复选框）→ 点「确认下单」dropdown-trigger → 点浮层同名菜单项「确认下单」→ 点「确认」',
        '3. 等 3S 读「共 B 条」：B=0 → 点侧边栏「下单动作跟踪」，流程结束（AE 退出 0）',
        '4. B≠0：每 10S 轮询（刷新+重读 B）最多 10 轮；任一轮 B=0 → 切「下单动作跟踪」收尾',
        '5. 10 轮后 B 仍≠0：整页截图 + 钉钉邮件告知，流程停止（AE 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'gzt_03x', 'AE_GZT_03_X_下单前确认.py'); },
    },
    {
      name: '下单动作跟踪（AF）',
      steps: [
        '0. 先读「下单动作跟踪」页右下角 B0（调用 AE.读取B0，同一「共 B0 条」口径）',
        '1. 按 B0 数量等待 B0/100 分钟，然后检查进度条是否 100%',
        '2. 进度=100% 且右下角=0：直接点侧边栏「接收任务清单」，流程结束（AF 退出 0）',
        '3. 进度=100% 且右下角≠0：先整页截图+邮件告知；再全选 → 点「退至下单前确认」→ 弹确认框则点它（AF 退出 0）',
        '4. 进度≠100%：每 30S 轮询检查，最多 50 次；仍≠100% ⇒ 整页截图+邮件告知，停止后续操作（AF 退出 3）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'gzt_03x', 'AF_GZT_03_X_下单动作跟踪.py'); },
    },
    {
      name: '订单下发邮件通知（AG）',
      steps: [
        '0. 单实例守卫：与 AB~AF 共用「计划平台操作」锁（同一时刻只允许一个进程操作平台页面）',
        '1. 进入「构建按单树」页（页签→侧边栏→全页逐级找；已在则直接复用），进入判据=页面文本同时出现表头「根物料」+「根订单」',
        '2. 逐屏滚动读该页 Surely 虚拟表：表头文案 + 每行单元格（fix-left/center/fix-right 按 top 归并、left 还原列序）+ 右下角「共 X 条」。滚动容器取 .surely-table-vertical-scroll-viewport / body-viewport-container（实测 body/content 的 scrollHeight≈clientHeight 会被误判「已到底」，60 条只采到 17 行）；表头保留 48px 空 extra-cell 使列与数据行逐格对齐（早先丢空列导致「根物料」取到空串）',
        '3. 任一「关键计划物料号」符合通配 *GZT*-*-*-03*（含 GZT 且末段以 03 开头，平台真实形如 26GZT-C0NA-56-03 / 26GZT-C0MH-75-03；子行 DMYA…/YA… 件号与末段 04、32 的不会被抓走；可用 APS_AG_ROOT_GLOB 调整）⇒ 用 DWS 最新版本命令发钉钉邮件',
        '4. 收件人 demo.user12@example.com（APS_AG_MAIL_TO 可调）；主题「GZT已下达 {台数} 台」，台数=命中行「关键计划物料号」去重数；正文=HTML 表格，只含截图口径的 6 列（关键计划物料号 / BOM数量 / 采购类型及特殊采购类 / 订单类型 / 生产订单号 / 生产订单基本完成时间），行只发命中的关键条目',
        '5. 未命中 ⇒ 仅记日志、不发邮件（退出 0），绝不无脑发信',
        '平台约定：复用 AC 基础工具（真实点击/整页截图/读取条数/滚动表格/重置）；切到侧边栏/页签后先「重置」（best-effort）',
      ],
      run: async (ctx) => { await apd运行脚本(ctx, 'gzt_03x', 'AG_邮件发送.py'); },
    },
  ],
});

/* ====================================================================
 * 任务：电脑计划任务同步（2026-10-20；把本机 Windows 计划任务迁移进网页）
 * --------------------------------------------------------------------
 * 用途：侧边栏「触发器 → 计划任务」里「电脑计划任务同步」定时触发器的应用。
 *   节点1 采集本机计划任务：GET /api/system/win-tasks
 *         （后台枚举本机任务计划程序，中文状态/触发类型由服务端映射）
 *   节点2 写入网页触发器存储：POST /api/system/win-tasks/snapshot
 *         （全量覆盖「电脑计划任务」sheet + 同步时间 meta，即迁移落地）
 *   节点3 汇总登记：启用/未启用计数写运行日志。
 * 电脑现有计划任务 9/11 为时间触发（Daily/Time/Logon），故网页侧新建的
 * 计划任务类型取「定时」，每天 08:00 同步一次。
 * ==================================================================== */
AutoTaskKit.register({
  id: 'win_task_sync',
  name: '电脑计划任务同步（采集本机 Windows 计划任务写入网页）',
  desc: '采集本机任务计划程序中的全部计划任务，写入触发器存储「电脑计划任务」sheet。',
  nodes: [
    {
      name: '采集本机计划任务',
      run: async (ctx) => {
        ctx.log('调用 /api/system/win-tasks 采集本机 Windows 计划任务…');
        const res = await fetch('/api/system/win-tasks');
        const d = await res.json().catch(() => ({}));
        if (!res.ok || !d.ok) throw new Error(d.detail || '采集失败（HTTP ' + res.status + '）');
        ctx.shared.tasks = d.tasks || [];
        ctx.log('采集到 ' + ctx.shared.tasks.length + ' 个计划任务');
      },
    },
    {
      name: '写入网页触发器存储',
      run: async (ctx) => {
        ctx.log('POST /api/system/win-tasks/snapshot 写入「电脑计划任务」sheet…');
        const res = await fetch('/api/system/win-tasks/snapshot', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tasks: ctx.shared.tasks || [] }),
        });
        const d = await res.json().catch(() => ({}));
        if (!res.ok || !d.ok) throw new Error(d.detail || '写入失败（HTTP ' + res.status + '）');
        ctx.log('已写入 ' + (d.count || 0) + ' 行');
      },
    },
    {
      name: '汇总登记',
      run: async (ctx) => {
        const ts = ctx.shared.tasks || [];
        const on = ts.filter(t => String(t.状态 || '').indexOf('就绪') >= 0).length;
        ctx.log('同步完成：共 ' + ts.length + ' 个任务（启用 ' + on + ' / 未启用 ' + (ts.length - on) + '），时间 ' +
          new Date().toLocaleString('zh-CN', { hour12: false }));
      },
    },
  ],
});
