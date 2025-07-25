// M3c 交互重构 e2e suite（F2 多区域选择 / F3 拖拽移动 / F4 富剪贴板 / F5 公式编辑器；叙述版见 apps/demo/e2e/m3c.md）。
// 页面内经 __xcell + W helper 驱动合成鼠标/键盘事件；F5 公式栏走真实 React DOM。
import { evaluateJS } from '../lib/bridge.mjs'
import { freshPage } from '../lib/env.mjs'

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const ev = (code) => evaluateJS(code)

function assertIncludes(hay, needle, label) {
  if (typeof hay !== 'string' || !hay.includes(needle)) {
    throw new Error(`${label}\n  应含: ${needle}\n  实际: ${JSON.stringify(hay)}`)
  }
}

// 页面侧工具（freshPage 已注 W helper；此处补 alert stub / React 受控赋值 / 填充柄拖拽 / 公式栏按键）
async function installPageTools() {
  await ev(`
    window.__lastAlert = null
    window.alert = (m) => { window.__lastAlert = m }
    // React 受控 input 赋值：native setter + input 事件
    window.__setReactValue = (el, v) => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v)
      el.dispatchEvent(new Event('input', { bubbles: true }))
    }
    // 公式栏按键（补全 ↑↓/Tab/Enter/Esc 走 React onKeyDown）
    window.__formulaKey = (k) => {
      window.__formulaInput()
        .dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }))
    }
    // 填充柄拖拽：mousedown 活动区域右下角 6px 方块（fillhandle 命中区），拖至目标格
    window.__fillDrag = (er, ec, tr, tc) => {
      const V = window.__xcell
      const R = document.querySelector('.xcell-root').getBoundingClientRect()
      const cv = V.cellViewportRect(er, ec)
      const cx = R.left + cv.x + cv.w
      const cy = R.top + cv.y + cv.h
      const disp = (type, el, x, y) => el.dispatchEvent(new MouseEvent(type, {
        bubbles: true, cancelable: true, view: window, clientX: x, clientY: y,
        button: 0, buttons: type === 'mousedown' ? 1 : 0 }))
      disp('mousedown', V.stage.content, cx, cy)
      const p = W.cellXY(tr, tc)
      disp('mousemove', window, p.clientX, p.clientY)
      disp('mouseup', window, p.clientX, p.clientY)
    }
    return 'tools-ready'`)
}

const rg = (sr, sc, er, ec) => ({ sr, sc, er, ec })

export default async function run({ assertEq }) {
  // ================= F2 多区域选择 =================

  // ---- F2.1 Ctrl+click 多区域 + activeCell=末项 ----
  async function f2_1() {
    const s = await ev(`
      W.click(6, 0); W.click(7, 2, { ctrl: true }); W.click(8, 4, { ctrl: true })
      return W.sel()`)
    assertEq(s.ranges, [rg(6, 0, 6, 0), rg(7, 2, 7, 2), rg(8, 4, 8, 4)], 'F2.1 ranges=[A7,C8,E9]')
    assertEq(s.activeCell, { row: 8, col: 4 }, 'F2.1 activeCell=末加入者 E9')
  }

  // ---- F2.2 Ctrl+click 反选移除（承接 F2.1 选区）----
  async function f2_2() {
    const s = await ev(`
      W.click(7, 2, { ctrl: true })
      return W.sel()`)
    assertEq(s.ranges, [rg(6, 0, 6, 0), rg(8, 4, 8, 4)], 'F2.2 反选后 ranges=[A7,E9]')
    assertEq(s.activeCell, { row: 8, col: 4 }, 'F2.2 activeCell 仍 E9')
  }

  // ---- F2.3 Ctrl+click 重叠 LIFO 反选（toggleRange 不变式）----
  async function f2_3() {
    const mid = await ev(`
      W.click(6, 0)
      W.dragSelect(5, 0, 7, 1, { ctrl: true })
      return W.sel()`)
    assertEq(mid.ranges, [rg(6, 0, 6, 0), rg(5, 0, 7, 1)], 'F2.3 追加后 ranges=[A7,A6:B8]')
    const s = await ev(`
      W.click(6, 0, { ctrl: true })
      return W.sel()`)
    assertEq(s.ranges, [rg(6, 0, 6, 0)], 'F2.3 LIFO 移除活动区域 A6:B8')
    assertEq(s.activeCell, { row: 6, col: 0 }, 'F2.3 activeCell∈ranges[last]（§1.1 不变式）')
  }

  // ---- F2.4 Shift+Arrow 生长（不丢起始格）----
  async function f2_4() {
    const s1 = await ev(`
      W.click(6, 0)
      W.key('ArrowRight', { shift: true })
      return W.sel()`)
    assertEq(s1.ranges, [rg(6, 0, 6, 1)], 'F2.4 第一次 Shift+→ 得 A7:B7')
    const s2 = await ev(`
      W.key('ArrowRight', { shift: true })
      return W.sel()`)
    assertEq(s2.ranges, [rg(6, 0, 6, 2)], 'F2.4 第二次 Shift+→ 得 A7:C7（sr/sc 恒 6/0）')
    assertEq(s2.activeCell, { row: 6, col: 2 }, 'F2.4 activeCell=焦点端')
  }

  // ---- F2.5 格式作用于全部区域 ----
  async function f2_5() {
    await ev(`
      W.click(6, 0); W.click(6, 2, { ctrl: true }); W.click(6, 4, { ctrl: true })
      document.querySelector('button[aria-label="加粗"]').click()
      return 1`)
    await sleep(150) // 等 React/事务 flush
    assertEq(await ev(`return [W.read(6,0)?.style?.bold ?? null, W.read(6,2)?.style?.bold ?? null, W.read(6,4)?.style?.bold ?? null]`),
      [true, true, true], 'F2.5 加粗作用于 A7/C7/E7 三格')
  }

  // ---- F2.6 填充仅活动区域 ----
  async function f2_6() {
    await ev(`
      W.setCell(6, 4, '1')
      W.click(6, 0); W.click(6, 2, { ctrl: true }); W.click(6, 4, { ctrl: true })
      window.__fillDrag(6, 4, 8, 4)
      return 1`)
    assertEq(await ev(`return [W.read(7,4)?.raw ?? null, W.read(8,4)?.raw ?? null]`),
      ['1', '1'], 'F2.6 活动区域 E7 填充 E8=1,E9=1')
    assertEq(await ev(`return [(W.read(6,0)?.raw ?? '') === '1', (W.read(6,2)?.raw ?? '') === '1']`),
      [false, false], 'F2.6 非活动区域 A7/C7 不变')
  }

  // ---- F2.7 排序/筛选多区域禁用 ----
  async function f2_7() {
    await ev(`
      W.dragSelect(6, 0, 8, 0)
      W.dragSelect(6, 2, 8, 2, { ctrl: true })
      return 1`)
    await sleep(150) // 等 React 重渲染 disabled
    // M5 下拉归并：升序项在「排序」Dropdown 内（需先开触发器）；筛选仍是独立 IconButton
    assertEq(await ev(`return window.__menuItemDisabled('排序', '按选区首列升序')`),
      true, 'F2.7 多区域时排序升序项禁用')
    assertEq(await ev(`
      return document.querySelector('button[aria-label="自动筛选（对选区启用/清除全表筛选）"]').disabled`),
      true, 'F2.7 多区域时筛选禁用')
    await ev(`W.dragSelect(6, 0, 8, 0); return 1`)
    await sleep(150)
    assertEq(await ev(`return window.__menuItemDisabled('排序', '按选区首列升序')`),
      false, 'F2.7 单区域基线时排序可用')
    assertEq(await ev(`
      return document.querySelector('button[aria-label="自动筛选（对选区启用/清除全表筛选）"]').disabled`),
      false, 'F2.7 单区域基线时筛选可用')
  }

  // ---- F2.8 表头并集高亮 ----
  async function f2_8() {
    const r = await ev(`
      W.dragSelect(6, 0, 6, 1)
      W.dragSelect(8, 3, 8, 4, { ctrl: true })
      const V = window.__xcell
      V.render() // dispatch 的渲染走 rAF 调度，扫节点树前须显式同步渲染
      const expCol = (c) => { const cv = V.cellViewportRect(0, c); return Math.round(cv.x + cv.w / 2) }
      const expRow = (rr) => { const cv = V.cellViewportRect(rr, 0); return Math.round(cv.y + cv.h / 2) }
      const cols = [], rows = []
      for (const n of V.gridLayer.find('Rect')) {
        if (n.fill() !== '#e8f0fe') continue
        const b = n.getClientRect()
        if (Math.round(b.y) === 0) cols.push(Math.round(b.x + b.width / 2))
        else if (Math.round(b.x) === 0) rows.push(Math.round(b.y + b.height / 2))
      }
      return { cols, rows, expCols: [0, 1, 3, 4].map(expCol), expRows: [6, 8].map(expRow), c2: expCol(2), r7: expRow(7) }`)
    assertEq(r.cols.slice().sort((a, b) => a - b), r.expCols.slice().sort((a, b) => a - b), 'F2.8 列头 A/B/D/E 高亮（并集）')
    assertEq(r.rows.slice().sort((a, b) => a - b), r.expRows.slice().sort((a, b) => a - b), 'F2.8 行头 7/9 高亮（并集）')
    if (r.cols.includes(r.c2)) throw new Error('F2.8 列头 C 不应高亮')
    if (r.rows.includes(r.r7)) throw new Error('F2.8 行头 8 不应高亮')
  }

  // ================= F3 拖动移动 =================

  // ---- F3.1 边框拖动 cut 移动（raw+style 搬、公式不 shift）----
  async function f3_1() {
    await ev(`
      W.setCell(11, 1, '5'); W.setCell(11, 2, '10'); W.setCell(11, 3, '=B11+C11')
      W.setStyle(11, 1, { bold: true }); W.setStyle(11, 3, { bg: '#ffff00' })
      W.click(11, 3)
      W.dragMove(11, 3, 11, 3, 2, 0)
      return 1`)
    assertEq(await ev(`return W.read(13, 3)`),
      { raw: '=B11+C11', style: { bg: '#ffff00' } }, 'F3.1 D13 公式不 shift + style 跟搬')
    assertEq(await ev(`return W.read(11, 3)`), null, 'F3.1 D11 清源')
    assertEq(await ev(`return [W.read(11,1)?.raw ?? null, W.read(11,2)?.raw ?? null]`),
      ['5', '10'], 'F3.1 非活动区域 B11/C11 不动')
  }

  // ---- F3.2 相交/merge 拒绝 notice（M5：pluginNotice 走状态栏，替代 window.alert）----
  async function f3_2() {
    await ev(`
      W.dragSelect(11, 1, 12, 2)
      W.dragMove(11, 1, 12, 2, 1, 0)
      return 1`)
    assertEq(await ev(`return window.__notice()`),
      '目标区域与源相交、落在合并区或超出表格边界，无法移动', 'F3.2 相交拒绝 notice')
    assertEq(await ev(`return W.read(11, 1)?.raw ?? null`), '5', 'F3.2 相交拒绝后 B11 不变')
    // merge 路径：追加合并区（0-based）E13:F13，D15 拖到目标 (12,4) 落合并区
    await ev(`
      const V = window.__xcell
      V.dispatch(V.state.tr.setMerges([...V.state.activeSheet.merges, { sr: 12, sc: 4, er: 12, ec: 5 }]))
      W.setCell(14, 3, 'moveme')
      W.click(14, 3)
      W.dragMove(14, 3, 14, 3, -2, 1)
      return 1`)
    assertEq(await ev(`return window.__notice()`),
      '目标区域与源相交、落在合并区或超出表格边界，无法移动', 'F3.2 merge 拒绝 notice')
    assertEq(await ev(`return W.read(14, 3)?.raw ?? null`), 'moveme', 'F3.2 merge 拒绝后源不变')
  }

  // ---- F3.3 越界 clamp+reject（不扩表、不丢数据）----
  async function f3_3() {
    await ev(`
      const V = window.__xcell
      V.scrollY = 996 * 24
      V.render()
      W.setCell(998, 0, 'rowOob1'); W.setCell(999, 0, 'rowOob2')
      W.dragSelect(998, 0, 999, 0)
      W.dragMove(998, 0, 999, 0, 1, 0)
      return 1`)
    assertEq(await ev(`return window.__notice()`),
      '目标区域与源相交、落在合并区或超出表格边界，无法移动', 'F3.3 越界拒绝 notice')
    assertEq(await ev(`return window.__xcell.state.activeSheet.rowCount`), 1000, 'F3.3 不扩表')
    assertEq(await ev(`return [W.read(998,0)?.raw ?? null, W.read(999,0)?.raw ?? null]`),
      ['rowOob1', 'rowOob2'], 'F3.3 不丢数据')
    await ev(`const V = window.__xcell; V.scrollY = 0; V.render(); return 1`)
  }

  // ---- F3.4 undo ----
  async function f3_4() {
    await ev(`
      W.setCell(15, 1, 'undoSrc')
      W.setStyle(15, 1, { bold: true, italic: true })
      W.click(15, 1)
      W.dragMove(15, 1, 15, 1, 2, 0)
      return 1`)
    assertEq(await ev(`return W.read(17, 1)`),
      { raw: 'undoSrc', style: { bold: true, italic: true } }, 'F3.4 移动后 B18 raw+style')
    assertEq(await ev(`return W.read(15, 1)`), null, 'F3.4 移动后 B16 清源')
    await ev(`W.key('z', { ctrl: true }); return 1`)
    assertEq(await ev(`return W.read(15, 1)`),
      { raw: 'undoSrc', style: { bold: true, italic: true } }, 'F3.4 undo 还原源')
    assertEq(await ev(`return W.read(17, 1)`), null, 'F3.4 undo 清目标')
  }

  // ================= F4 富剪贴板 =================

  // ---- F4.1 内部 copy/paste 保留样式+公式 ----
  async function f4_1() {
    await ev(`
      W.setCell(16, 1, '42'); W.setCell(16, 2, '10'); W.setCell(16, 3, '=B16+C16')
      W.setStyle(16, 3, { bold: true })
      W.click(16, 3); W.copy()
      W.click(18, 3); W.paste()
      return 1`)
    assertEq(await ev(`return W.read(18, 3)`),
      { raw: '=B18+C18', style: { bold: true } }, 'F4.1 D19 公式偏移+2行 + style 保留')
    assertEq(await ev(`return W.read(16, 3)`),
      { raw: '=B16+C16', style: { bold: true } }, 'F4.1 源 D17 不变')
  }

  // ---- F4.2 多区域 copy 块间空行 ----
  async function f4_2() {
    await ev(`
      W.setCell(20, 1, 'first'); W.setStyle(20, 1, { bold: true, color: '#ff0000' })
      W.setCell(20, 3, 'second'); W.setStyle(20, 3, { italic: true })
      W.click(20, 1); W.click(20, 3, { ctrl: true })
      W.copy()
      return 1`)
    assertEq(await ev(`return window.W.__clipText`), 'first\n\nsecond', 'F4.2 多区域 TSV 块间空行')
  }

  // ---- F4.3 出站带样式（HTML 内联 CSS；承接 F4.2 的 copy）----
  async function f4_3() {
    const html = await ev(`return window.W.__clipHtml`)
    assertIncludes(html, '<table', 'F4.3 每 area 一段 <table>')
    assertIncludes(html, 'font-weight:bold', 'F4.3 bold→font-weight:bold')
    assertIncludes(html, 'color:#ff0000', 'F4.3 color 内联')
    assertIncludes(html, 'font-style:italic', 'F4.3 italic→font-style:italic')
  }

  // ---- F4.4 从外部粘入样式丢失（兜底 cell:{raw}）----
  async function f4_4() {
    await ev(`
      W.click(22, 0); W.paste('alpha\tbeta')
      return 1`)
    assertEq(await ev(`return [W.read(22, 0), W.read(22, 1)]`),
      [{ raw: 'alpha', style: null }, { raw: 'beta', style: null }], 'F4.4 外部 TSV 兜底 raw、style 丢失')
    await ev(`W.click(23, 0); W.paste('=1+2'); return 1`)
    assertEq(await ev(`return W.read(23, 0)`),
      { raw: '=1+2', style: null }, 'F4.4 外部公式 normalizedCell 保公式')
  }

  // ---- F4.5 undo paste ----
  async function f4_5() {
    await ev(`
      W.setCell(16, 5, 'tag'); W.setStyle(16, 5, { bold: true, bg: '#00ff00' })
      W.click(16, 5); W.copy()
      W.click(26, 5); W.paste()
      return 1`)
    assertEq(await ev(`return W.read(26, 5)`),
      { raw: 'tag', style: { bold: true, bg: '#00ff00' } }, 'F4.5 paste 后 F27 raw+style')
    await ev(`W.key('z', { ctrl: true }); return 1`)
    assertEq(await ev(`return W.read(26, 5)`), null, 'F4.5 undo 一次清 raw+style')
    assertEq(await ev(`return W.read(16, 5)`),
      { raw: 'tag', style: { bold: true, bg: '#00ff00' } }, 'F4.5 源 F17 不变')
  }

  // ================= F5 公式编辑增强 =================

  // ---- F5.1 编辑公式画布对被引区域彩色框（每域一色）----
  async function f5_1() {
    const r = await ev(`
      W.click(16, 1)
      const inp = window.__formulaInput()
      window.__setReactValue(inp, '=B16+C16')
      window.__xcell.render()
      const V = window.__xcell
      const PAL = ['#1a73e8', '#ea4335', '#fbbc04', '#34a853', '#9334e6', '#ff6d00', '#00897b', '#7cb342']
      const hits = V.overlayLayer.find('Rect')
        .filter((n) => PAL.includes(n.stroke()) && JSON.stringify(n.dash()) === '[3,2]')
        .map((n) => { const b = n.getClientRect(); return { x: b.x, y: b.y, w: b.width, h: b.height, stroke: n.stroke() } })
        .sort((a, b) => a.x - b.x)
      const cvB = V.cellViewportRect(15, 1)
      const cvC = V.cellViewportRect(15, 2)
      return { hits, cvB, cvC }`)
    assertEq(r.hits.map((h) => h.stroke), ['#1a73e8', '#ea4335'], 'F5.1 被引区域彩色框每域一色（palette[0]/[1]）')
    if (r.hits.length !== 2) throw new Error(`F5.1 预期 2 个引用框，实际 ${r.hits.length}`)
    if (Math.abs(r.hits[0].x - r.cvB.x) > 3 || Math.abs(r.hits[0].y - r.cvB.y) > 3) {
      throw new Error(`F5.1 第一框应盖 B16: ${JSON.stringify(r.hits[0])} vs ${JSON.stringify(r.cvB)}`)
    }
    if (Math.abs(r.hits[1].x - r.cvC.x) > 3 || Math.abs(r.hits[1].y - r.cvC.y) > 3) {
      throw new Error(`F5.1 第二框应盖 C16: ${JSON.stringify(r.hits[1])} vs ${JSON.stringify(r.cvC)}`)
    }
  }

  // ---- F5.2 函数名补全下拉（↑↓/Tab/Enter/Esc）----
  async function f5_2() {
    const setInput = (v) => ev(`
      window.__setReactValue(window.__formulaInput(), ${JSON.stringify(v)})
      return 1`)
    // M5：补全下拉项为公式栏内 .cursor-pointer 行
    const readDD = () => ev(`
      return [...document.querySelectorAll('div.h-8 .cursor-pointer')].map((d) => d.textContent)`)
    const readVal = () => ev(`return window.__formulaInput().value`)
    // 前缀匹配出下拉
    await setInput('=SU')
    await sleep(150)
    assertEq(await readDD(), ['SUM', 'SUMIF'], 'F5.2 =SU 候选 [SUM,SUMIF]')
    // ↓↓（clamp 末项 SUMIF）→ Tab 接受；React 18 同 tick 批量更新会让后续按键读到旧 selIndex，
    // 每个 keydown 后须让出事件循环（await setTimeout）等 state flush
    await ev(`
      window.__formulaKey('ArrowDown'); await new Promise((r) => setTimeout(r, 50))
      window.__formulaKey('ArrowDown'); await new Promise((r) => setTimeout(r, 50))
      window.__formulaKey('Tab'); return 1`)
    await sleep(150)
    assertEq(await readVal(), '=SUMIF(', 'F5.2 ↓↓+Tab 补全 =SUMIF(')
    assertEq(await readDD(), [], 'F5.2 Tab 后下拉关')
    // Esc：关下拉不补全
    await setInput('=SU')
    await sleep(150)
    await ev(`window.__formulaKey('Escape'); return 1`)
    await sleep(150)
    assertEq(await readDD(), [], 'F5.2 Esc 后下拉关')
    assertEq(await readVal(), '=SU', 'F5.2 Esc 不补全')
    // Enter：接受首候选（前值同为 '=SU'，React 值跟踪会去重跳过 onChange，先重置再设）
    await setInput('')
    await sleep(100)
    await setInput('=SU')
    await sleep(150)
    await ev(`window.__formulaKey('Enter'); return 1`)
    await sleep(150)
    assertEq(await readVal(), '=SUM(', 'F5.2 Enter 接受 index 0 =SUM(')
    // ↑ clamp 0
    await setInput('=CO')
    await sleep(150)
    await ev(`
      window.__formulaKey('ArrowDown'); await new Promise((r) => setTimeout(r, 50))
      window.__formulaKey('ArrowUp'); await new Promise((r) => setTimeout(r, 50))
      window.__formulaKey('Enter'); return 1`)
    await sleep(150)
    assertEq(await readVal(), '=COUNT(', 'F5.2 ↓↑(clamp 0)+Enter 得 =COUNT(')
    // 清理：Esc 还原公式栏文本 + 清画布引用高亮
    await ev(`window.__formulaKey('Escape'); window.__formulaKey('Escape'); return 1`)
    await sleep(150)
  }

  await freshPage()
  await installPageTools()
  const scenarios = [
    ['F2.1 Ctrl+click 多区域+activeCell=末项', f2_1],
    ['F2.2 Ctrl+click 反选移除', f2_2],
    ['F2.3 重叠 LIFO 反选（toggleRange 不变式）', f2_3],
    ['F2.4 Shift+Arrow 生长', f2_4],
    ['F2.5 格式作用于全部区域', f2_5],
    ['F2.6 填充仅活动区域', f2_6],
    ['F2.7 排序/筛选多区域禁用', f2_7],
    ['F2.8 表头并集高亮', f2_8],
    ['F3.1 边框拖动 cut 移动', f3_1],
    ['F3.2 相交/merge 拒绝 alert', f3_2],
    ['F3.3 越界 clamp+reject', f3_3],
    ['F3.4 undo', f3_4],
    ['F4.1 内部 copy/paste 保留样式+公式', f4_1],
    ['F4.2 多区域 copy 块间空行', f4_2],
    ['F4.3 出站带样式 HTML', f4_3],
    ['F4.4 外部粘入样式丢失', f4_4],
    ['F4.5 undo paste', f4_5],
    ['F5.1 被引区域彩色框', f5_1],
    ['F5.2 函数名补全下拉', f5_2],
  ]
  for (const [name, fn] of scenarios) {
    try {
      await fn()
      console.log(`  ✓ ${name}`)
    } catch (e) {
      console.error(`  ✗ ${name}`)
      throw e // 快速失败：断言失败即终止 suite
    }
  }
}
