// M4a 持久化 e2e suite（6 场景；叙述版见 apps/demo/e2e/m4a.md）。
// 页面内经 __xcell + W helper 驱动；CSV 导入走 harness feedFile（base64+DataTransfer 主路径）。
import { bringToFront, evaluateJS } from '../lib/bridge.mjs'
import { feedFile, freshPage, pollUntil, reinject, reload } from '../lib/env.mjs'

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const ev = (code) => evaluateJS(code)

// dev server root 为 apps/demo，包源码在 root 之外，须走 vite /@fs/ 绝对路径
const ENGINE_URL = '/@fs' + new URL('../../../../packages/excel-core/src/formula/engine.ts', import.meta.url).pathname

function assertIncludes(hay, needle, label) {
  if (typeof hay !== 'string' || !hay.includes(needle)) {
    throw new Error(`${label}\n  应含: ${needle}\n  实际: ${JSON.stringify(hay)}`)
  }
}

export default async function run({ assertEq }) {
  // ---- 场景 1：自动保存 + 刷新恢复 ----
  async function s1() {
    assertEq(await ev(`return W.read(0, 0)?.raw ?? null`), '产品', 'S1 初始 demo 表 A1')
    // 合成 keydown 不产生文本插入（仅首字符由 app 开编辑器带入），第二字符改走 editor-value（同 m4c __typeCell）
    await ev(`
      W.click(19, 1); W.key('9')
      const ed = document.querySelector('textarea.xcell-editor')
      if (!ed) throw new Error('编辑器未打开')
      ed.value = '99'
      ed.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
      return 1`)
    assertEq(await ev(`return __xcell.state.activeSheet.getCell(19, 1)?.raw ?? null`), '99', 'S1 B20 输入 99')
    // 防抖 1s 落盘；节流 tab 内定时器可能延迟，轮询而非固定 sleep
    await pollUntil(`(localStorage.getItem('xexcel.workbook.v2')?.length ?? 0) > 0`, 'S1 防抖落盘')
    await reload()
    assertEq(await ev(`return __xcell.state.activeSheet.getCell(19, 1)?.raw ?? null`), '99', 'S1 刷新后 B20 恢复')
    assertEq(await ev(`return __xcell.state.activeSheet.getCell(0, 0)?.raw ?? null`), '产品', 'S1 刷新后 A1 demo 数据也在存档')
  }

  // ---- 场景 2：清除存档（M5：React ConfirmDialog，菜单项点击后确认钮文案同为「清除」）----
  async function s2() {
    const r = await ev(`
      await window.__clickFileMenu('清除')
      window.__clickDialogBtn('清除')
      await new Promise((r) => setTimeout(r, 500))
      return [localStorage.getItem('xexcel.workbook.v2'), window.__notice()]`)
    assertEq(r[0], null, 'S2 存档已删')
    assertIncludes(r[1], '已清除浏览器存档', 'S2 StatusBar 提示')
    await reload()
    assertEq(await ev(`return __xcell.state.activeSheet.getCell(19, 1)?.raw ?? null`), null, 'S2 刷新后回默认 demo 表 B20 空')
  }

  // ---- 场景 3：CSV 导出（BOM 字节 + 转义 + 公式原文）----
  async function s3() {
    const r = await ev(`
      const st = __xcell.state
      const tr = st.tr.setCells(st.doc.active, [
        { row: 0, col: 5, cell: { raw: '含,逗号' } },
        { row: 1, col: 5, cell: { raw: '含"引号' } },
        { row: 2, col: 5, cell: { raw: '=B2*2' } },
      ])
      __xcell.dispatch(tr)
      window.__dl = null
      const oc = URL.createObjectURL.bind(URL)
      URL.createObjectURL = (b) => { window.__dl = b; return oc(b) }
      await window.__clickFileMenu('导出 CSV')
      if (!window.__dl) throw new Error('未捕获导出 Blob')
      // blob.text() 会剥离 UTF-8 BOM，BOM 须用 arrayBuffer 验字节
      const txt = await window.__dl.text()
      const buf = new Uint8Array(await window.__dl.arrayBuffer())
      URL.createObjectURL = oc
      return { b0: buf[0], b1: buf[1], b2: buf[2], txt }`)
    assertEq([r.b0, r.b1, r.b2], [0xEF, 0xBB, 0xBF], 'S3 CSV BOM 字节 EF BB BF')
    assertIncludes(r.txt, '"含,逗号"', 'S3 CSV 逗号转义')
    assertIncludes(r.txt, '"含""引号', 'S3 CSV 引号转义')
    assertIncludes(r.txt, '=B2*2', 'S3 CSV 公式原文')
  }

  // ---- 场景 4：CSV 导入（新 sheet、公式复活、undo 一步消失）----
  async function s4() {
    const before = await ev(`return __xcell.state.doc.order.length`)
    const b64 = Buffer.from('姓名,数量\r\n苹果,3\r\n=B2*2,"含,逗号"\r\n', 'utf8').toString('base64')
    await ev(`
      window.__lastInput = null
      await window.__clickFileMenu('打开 CSV')
      if (!window.__lastInput) throw new Error('未捕获 pickFile input')
      return true`)
    await feedFile(b64, '进货.csv')
    await sleep(800)
    const info = await ev(`
      const wb = __xcell.state.doc
      return { len: wb.order.length, name: wb.names.get(wb.active) }`)
    assertEq(info.len, before + 1, 'S4 导入新增 sheet')
    assertEq(info.name, '进货', 'S4 active 为新 sheet 名')
    assertEq(await ev(`
      const sh = __xcell.state.activeSheet
      return [sh.getCell(0, 0)?.raw ?? null, sh.getCell(1, 1)?.raw ?? null,
        sh.getCell(2, 0)?.raw ?? null, sh.getCell(2, 1)?.raw ?? null]`),
      ['姓名', '3', '=B2*2', '含,逗号'], 'S4 CSV 单元格内容')
    assertEq(await ev(`
      const { evaluatorFor } = await import('${ENGINE_URL}')
      return evaluatorFor(__xcell.state.doc).get(__xcell.state.doc.active, 2, 0)`),
      6, 'S4 公式复活计算值（B2=3 → =B2*2）')
    await ev(`document.querySelector('button[aria-label="撤销"]').click(); return 1`)
    assertEq(await ev(`return __xcell.state.doc.order.length`), before, 'S4 undo 一步新 sheet 消失')
  }

  // ---- 场景 5：损坏存档恢复 ----
  async function s5() {
    // 等 S4 undo 的防抖存档落盘（存档不再含「进货」sheet），避免 beforeunload flush 覆盖损坏档
    await pollUntil(`!localStorage.getItem('xexcel.workbook.v2')?.includes('进货')`, 'S5 undo 防抖落盘')
    await ev(`
      localStorage.setItem('xexcel.workbook.v2', '{broken')
      location.reload()
      'ok'`)
    await bringToFront()
    await pollUntil(`!!window.__xcell`, 'S5 启动完成', 20000)
    // 偏离说明：.md 预期主键为 null 是时效快照（人工 800ms 检查时启动后首个状态事件的
    // 防抖落盘尚未触发）；runner 观测时初始存档已写回。语义等价断言：主键不再是损坏串
    // （null 或有效初始存档均可）+ .corrupt 备份 + 正常启动。
    assertEq(await ev(`
      const main = localStorage.getItem('xexcel.workbook.v2')
      return [!!window.__xcell, localStorage.getItem('xexcel.workbook.v2.corrupt'), main !== '{broken']`),
      [true, '{broken', true], 'S5 正常启动 + 损坏档备份 .corrupt + 主键损坏串已清')
    await reinject()
    assertEq(await ev(`return __xcell.state.activeSheet.getCell(0, 0)?.raw ?? null`), '产品', 'S5 落默认 demo 表')
  }

  // ---- 场景 6：自动保存失败提示（配额模拟；轮询期间保持 setItem 抛错，命中后恢复防污染后续）----
  async function s6() {
    await ev(`
      window.__ose = Storage.prototype.setItem
      Storage.prototype.setItem = () => { throw new Error('QuotaExceeded') }
      W.click(0, 9); W.key('x'); W.key('Enter')
      return 1`)
    let errText = null
    for (let i = 0; i < 50; i++) {
      await sleep(400)
      errText = await ev(`return window.__statusError()`)
      if (errText) break
    }
    await ev(`Storage.prototype.setItem = window.__ose; return 1`)
    assertEq(errText, '自动保存失败', 'S6 配额失败 status-error 提示')
  }

  await freshPage()
  const scenarios = [
    ['1 自动保存+刷新恢复', s1],
    ['2 清除存档', s2],
    ['3 CSV 导出', s3],
    ['4 CSV 导入', s4],
    ['5 损坏存档恢复', s5],
    ['6 配额失败提示', s6],
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
