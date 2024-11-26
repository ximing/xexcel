// M4c 数据验证 e2e suite（E4 六场景；叙述版见 tests/e2e/m4c.md）。
// 页面内经 __xcell + W helper 驱动；对话框走真实 DOM；导出字节回 runner 用 exceljs 校验。
import ExcelJS from 'exceljs'
import { bringToFront, evaluateJS } from '../lib/bridge.mjs'
import { feedFile, freshPage } from '../lib/env.mjs'

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const ev = (code) => evaluateJS(code)

function assertIncludes(hay, needle, label) {
  if (typeof hay !== 'string' || !hay.includes(needle)) {
    throw new Error(`${label}\n  应含: ${needle}\n  实际: ${JSON.stringify(hay)}`)
  }
}

// 页面侧工具（每次 freshPage 后重注）：规则注入/notice/单元格输入/公式栏输入
async function installPageTools() {
  await ev(`
    window.__setRules = (rules) => window.__xcell.dispatch(window.__xcell.state.tr.setValidations(rules))
    window.__notice = () => document.querySelector('.status-notice')?.textContent ?? null
    window.__validations = () => window.__xcell.state.activeSheet.validations
    // React 受控 input 赋值：native setter + input 事件
    window.__setReactValue = (el, v) => {
      const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, v)
      el.dispatchEvent(new Event('input', { bubbles: true }))
    }
    // 单元格输入：click + 首字符 keydown 开编辑器 → 覆盖全文 → Enter。
    // 返回 'open'（被拒，编辑器保持打开）或 'closed'（提交成功已关闭）。
    window.__typeCell = (r, c, text) => {
      W.click(r, c)
      W.key(text[0])
      const ed = document.querySelector('textarea.xcell-editor')
      if (!ed) throw new Error('编辑器未打开')
      ed.value = text
      ed.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      return document.querySelector('textarea.xcell-editor') ? 'open' : 'closed'
    }
    // 被拒后继续：改当前编辑器文本再 Enter
    window.__typeContinue = (text) => {
      const ed = document.querySelector('textarea.xcell-editor')
      if (!ed) throw new Error('无打开中的编辑器')
      ed.value = text
      ed.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      return document.querySelector('textarea.xcell-editor') ? 'open' : 'closed'
    }
    // 公式栏路径：填值 + Enter 提交（拒绝时文本保留，可再次调用）
    window.__setFormula = (text) => {
      const inp = document.querySelector('input.formula-input')
      window.__setReactValue(inp, text)
      inp.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
    }
    // 对话框按钮按文本点击
    window.__clickDialogBtn = (text) => {
      const b = [...document.querySelectorAll('.dialog-actions button')].find((x) => x.textContent.includes(text))
      if (!b) throw new Error('对话框按钮未找到: ' + text)
      b.click()
    }
    return 'tools-ready'`)
}

// model validations 归一化（固定 key 序 + 排序，供 assertEq 比较）
const NORMALIZE_RULES = `
  window.__validations().map((r) => r.type === 'list'
    ? { range: { sr: r.range.sr, sc: r.range.sc, er: r.range.er, ec: r.range.ec }, type: 'list', items: r.items }
    : { range: { sr: r.range.sr, sc: r.range.sc, er: r.range.er, ec: r.range.ec }, type: r.type, op: r.op, v1: r.v1, v2: r.v2 })
    .sort((a, b) => (JSON.stringify(a) < JSON.stringify(b) ? -1 : 1))`

export default async function run({ assertEq }) {
  // ---- 场景 1：numRange 阻止+放行（editbox 与 FormulaBar 双路径）----
  async function s1() {
    await ev(`window.__setRules([{ id: 'v1', range: { sr: 19, sc: 0, er: 23, ec: 0 }, type: 'numRange', op: 'between', v1: '1', v2: '9' }]); return 1`)
    assertEq(await ev(`return window.__typeCell(19, 0, '10')`), 'open', 'S1 输入 10 应被拒（编辑器保持打开）')
    assertEq(await ev(`return W.read(19, 0)`), null, 'S1 拒绝后 A20 raw 仍空')
    assertIncludes(await ev(`return window.__notice()`), '介于 1 与 9', 'S1 拒绝 notice 文案')
    assertEq(await ev(`return window.__typeContinue('5')`), 'closed', 'S1 输入 5 应放行')
    assertEq(await ev(`return W.read(19, 0)?.raw`), '5', 'S1 放行后 A20 raw')
    // FormulaBar 路径（Enter 后选区已下移到 A21，仍在规则区内）
    assertEq(await ev(`const a = W.sel().activeCell; return { row: a.row, col: a.col }`), { row: 20, col: 0 }, 'S1 Enter 后选区下移 A21')
    await ev(`window.__setFormula('10'); return 1`)
    assertEq(await ev(`return W.read(20, 0)`), null, 'S1 公式栏拒绝后 A21 raw 仍空')
    assertIncludes(await ev(`return window.__notice()`), '介于 1 与 9', 'S1 公式栏拒绝 notice')
    await ev(`window.__setFormula('5'); return 1`)
    assertEq(await ev(`return W.read(20, 0)?.raw`), '5', 'S1 公式栏放行后 A21 raw')
  }

  // ---- 场景 2：textLen ----
  async function s2() {
    await ev(`window.__setRules([{ id: 'v1', range: { sr: 19, sc: 1, er: 23, ec: 1 }, type: 'textLen', op: 'lte', v1: '3' }]); return 1`)
    assertEq(await ev(`return window.__typeCell(19, 1, 'abcd')`), 'open', 'S2 长度 4 应被拒')
    assertEq(await ev(`return W.read(19, 1)`), null, 'S2 拒绝后 B20 raw 仍空')
    assertIncludes(await ev(`return window.__notice()`), '文本长度须小于等于 3', 'S2 拒绝 notice 文案')
    assertEq(await ev(`return window.__typeContinue('abc')`), 'closed', 'S2 长度 3 应放行')
    assertEq(await ev(`return W.read(19, 1)?.raw`), 'abc', 'S2 放行后 B20 raw')
  }

  // ---- 场景 3：list ----
  async function s3() {
    await ev(`window.__setRules([{ id: 'v1', range: { sr: 19, sc: 2, er: 23, ec: 2 }, type: 'list', items: ['Apple', 'Banana'] }]); return 1`)
    assertEq(await ev(`return window.__typeCell(19, 2, 'cherry')`), 'open', 'S3 序列外值应被拒')
    assertEq(await ev(`return W.read(19, 2)`), null, 'S3 拒绝后 C20 raw 仍空')
    assertIncludes(await ev(`return window.__notice()`), '输入值须在序列内：Apple, Banana', 'S3 拒绝 notice 文案')
    assertEq(await ev(`return window.__typeContinue('apple')`), 'closed', 'S3 小写命中应放行')
    assertEq(await ev(`return W.read(19, 2)?.raw`), 'apple', 'S3 放行后 C20 raw')
  }

  // ---- 场景 4：对话框配置 + 删除 + undo（清档刷新保证 undo 栈干净）----
  async function s4() {
    await freshPage()
    await installPageTools()
    assertEq(await ev(`W.click(19, 3); const a = W.sel().activeCell; return { row: a.row, col: a.col }`), { row: 19, col: 3 }, 'S4 选中 D20')
    await ev(`[...document.querySelectorAll('button.tool-btn')].find((b) => b.title === '数据验证').click(); return 1`)
    await sleep(300)
    await ev(`window.__clickDialogBtn('添加规则'); return 1`)
    await sleep(150)
    assertEq(await ev(`
      const vals = [...document.querySelectorAll('.cf-row .cf-value')]
      if (vals.length !== 2) throw new Error('预期两个数值输入框，实际 ' + vals.length)
      window.__setReactValue(vals[0], '1')
      window.__setReactValue(vals[1], '9')
      return vals.map((x) => x.placeholder)`), ['数值', '上界'], 'S4 默认 numRange between 两输入框')
    await ev(`window.__clickDialogBtn('确定'); return 1`)
    await sleep(150)
    assertEq(await ev(`return (${NORMALIZE_RULES})`),
      [{ range: { sr: 19, sc: 3, er: 19, ec: 3 }, type: 'numRange', op: 'between', v1: '1', v2: '9' }],
      'S4 对话框添加规则入 model')
    // 再开对话框删除
    await ev(`[...document.querySelectorAll('button.tool-btn')].find((b) => b.title === '数据验证').click(); return 1`)
    await sleep(300)
    assertEq(await ev(`document.querySelector('.cf-row button[title="删除规则"]').click(); return 1`), 1, 'S4 删除规则行')
    await sleep(150)
    await ev(`window.__clickDialogBtn('确定'); return 1`)
    await sleep(150)
    assertEq(await ev(`return window.__validations()`), [], 'S4 删除后 model 空')
    // undo：一次恢复规则，再一次回到空
    assertEq(await ev(`W.key('z', { ctrl: true }); return window.__validations().length`), 1, 'S4 undo 一次恢复规则')
    assertEq(await ev(`W.key('z', { ctrl: true }); return window.__validations()`), [], 'S4 undo 两次回到空')
  }

  // ---- 场景 5：xlsx 往返（导出字节 runner 侧 exceljs 校验 + 喂回导入保真）----
  async function s5() {
    await ev(`window.__setRules([
      { id: 'v1', range: { sr: 19, sc: 0, er: 23, ec: 0 }, type: 'numRange', op: 'between', v1: '1', v2: '9' },
      { id: 'v2', range: { sr: 19, sc: 2, er: 23, ec: 2 }, type: 'list', items: ['Apple', 'Banana'] },
    ]); return 1`)
    await bringToFront()
    const b64 = await ev(`
      window.__dl = null
      window.__oc = URL.createObjectURL.bind(URL)
      URL.createObjectURL = (b) => { window.__dl = b; return window.__oc(b) }
      ;[...document.querySelectorAll('.tool-btn')].find((b) => b.textContent === '文件').click()
      await new Promise((r) => setTimeout(r, 300))
      ;[...document.querySelectorAll('.file-menu-item')].find((b) => b.textContent === '导出 xlsx').click()
      await new Promise((r) => setTimeout(r, 1200))
      if (!window.__dl) throw new Error('未捕获导出 Blob')
      const bytes = new Uint8Array(await window.__dl.arrayBuffer())
      URL.createObjectURL = window.__oc
      let bin = ''
      for (let i = 0; i < bytes.length; i += 8192) bin += String.fromCharCode(...bytes.subarray(i, i + 8192))
      return btoa(bin)`)
    // runner 侧 exceljs 校验 dataValidations
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(Buffer.from(b64, 'base64'))
    const model = wb.worksheets[0].dataValidations.model
    for (const a of ['A20', 'A21', 'A22', 'A23', 'A24']) {
      const dv = model[a]
      assertEq({ t: dv?.type, o: dv?.operator, f: (dv?.formulae ?? []).map(String) },
        { t: 'decimal', o: 'between', f: ['1', '9'] }, `S5 导出 decimal 规则 ${a}`)
    }
    for (const a of ['C20', 'C21', 'C22', 'C23', 'C24']) {
      const dv = model[a]
      assertEq({ t: dv?.type, f: dv?.formulae ?? [] },
        { t: 'list', f: ['"Apple,Banana"'] }, `S5 导出 list 规则 ${a}`)
    }
    // 喂回导入，断言 validations 保真
    await bringToFront()
    await ev(`
      window.__lastInput = null
      ;[...document.querySelectorAll('.tool-btn')].find((b) => b.textContent === '文件').click()
      await new Promise((r) => setTimeout(r, 300))
      ;[...document.querySelectorAll('.file-menu-item')].find((b) => b.textContent.includes('打开 xlsx')).click()
      await new Promise((r) => setTimeout(r, 300))
      if (!window.__lastInput) throw new Error('未捕获 pickFile input')
      return true`)
    await feedFile(b64, 'm4c-roundtrip.xlsx')
    await bringToFront()
    await sleep(800)
    assertEq(await ev(`return (${NORMALIZE_RULES})`),
      [
        { range: { sr: 19, sc: 0, er: 23, ec: 0 }, type: 'numRange', op: 'between', v1: '1', v2: '9' },
        { range: { sr: 19, sc: 2, er: 23, ec: 2 }, type: 'list', items: ['Apple', 'Banana'] },
      ],
      'S5 导入后 validations 保真')
  }

  // ---- 场景 6：公式/清空跳过校验 ----
  async function s6() {
    await bringToFront()
    await ev(`window.__setRules([{ id: 'v1', range: { sr: 19, sc: 4, er: 23, ec: 4 }, type: 'numRange', op: 'between', v1: '1', v2: '9' }]); return 1`)
    assertEq(await ev(`return window.__typeCell(19, 4, '=1+1')`), 'closed', 'S6 公式应放行')
    assertEq(await ev(`return W.read(19, 4)?.raw`), '=1+1', 'S6 公式 raw 写入')
    // 不等 notice TTL（隐藏 tab 定时器密集节流，5s 等待不可靠；实测 visibilityState 恒 hidden）。
    // 改断言 Delete 后 notice 不含任何校验拒绝文案（拒绝串均以「请输入/文本长度/输入值须」开头）。
    await ev(`W.click(19, 4); W.key('Delete'); return 1`)
    assertEq(await ev(`return W.read(19, 4)`), null, 'S6 Delete 清空')
    const n6 = await ev(`return window.__notice()`)
    if (n6 && /请输入|文本长度|输入值须/.test(n6)) {
      throw new Error(`S6 Delete 不应触发校验拒绝 notice，实际: ${JSON.stringify(n6)}`)
    }
  }

  // ---- 场景 7：拒绝后 dblclick 另一格不建第二编辑器（I-1 僵尸编辑器回归）----
  async function s7() {
    await ev(`window.__setRules([{ id: 'v1', range: { sr: 15, sc: 0, er: 15, ec: 0 }, type: 'numRange', op: 'between', v1: '1', v2: '9' }]); return 1`)
    assertEq(await ev(`return window.__typeCell(15, 0, '99')`), 'open', 'S7 非法输入被拒（编辑器保持打开）')
    // dblclick 另一格 B16：走 openEditor 重开路径（closeEditor(true) 被拒 → 原会话存活 → 不得再开）
    await ev(`
      const p = W.cellXY(15, 1)
      window.__xcell.stage.content.dispatchEvent(new MouseEvent('dblclick', {
        bubbles: true, cancelable: true, view: window, clientX: p.clientX, clientY: p.clientY }))
      return 1`)
    assertEq(await ev(`return document.querySelectorAll('textarea.xcell-editor').length`), 1, 'S7 拒绝后 dblclick 不得建第二编辑器')
    assertEq(await ev(`return W.read(15, 1)`), null, 'S7 dblclick 目标格 raw 未写入')
    // 收尾：合法值提交，编辑器关闭
    assertEq(await ev(`return window.__typeContinue('5')`), 'closed', 'S7 合法值放行收尾')
    assertEq(await ev(`return W.read(15, 0)?.raw`), '5', 'S7 A16 raw 写入')
  }

  await freshPage()
  await installPageTools()
  const scenarios = [
    ['1 numRange 阻止+放行', s1],
    ['2 textLen', s2],
    ['3 list', s3],
    ['4 对话框配置+删除+undo', s4],
    ['5 xlsx 往返', s5],
    ['6 公式/清空跳过', s6],
    ['7 拒绝后 dblclick 无双编辑器', s7],
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
