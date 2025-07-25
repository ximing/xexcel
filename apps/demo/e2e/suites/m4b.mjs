// M4b xlsx 导入导出 e2e suite（5 场景；叙述版见 apps/demo/e2e/m4b.md）。
// 导出字节回 runner 用 exceljs 回读校验（等价 m4b.md 附录 B 全断言）；
// 导入文件 runner 侧 exceljs 生成 /tmp/m4b-import.xlsx（附录 A）再读字节 feedFile。
import { readFileSync } from 'node:fs'
import ExcelJS from 'exceljs'
import { bringToFront, evaluateJS } from '../lib/bridge.mjs'
import { feedFile, freshPage, pollUntil, reload } from '../lib/env.mjs'

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const ev = (code) => evaluateJS(code)

// dev server root 为 apps/demo，包源码在 root 之外，须走 vite /@fs/ 绝对路径
const ENGINE_URL = '/@fs' + new URL('../../../../packages/excel-core/src/formula/engine.ts', import.meta.url).pathname

// 点「文件」→ 菜单项（M5：触发器 aria-label="文件"，项为 role="menuitem"；helper 内建 300ms 等待）
async function clickMenu(includes) {
  await ev(`return window.__clickFileMenu(${JSON.stringify(includes)})`)
}

export default async function run({ assertEq }) {
  // ---- 场景 1：导出 xlsx 结构校验（runner 侧 exceljs 回读，等价附录 B）----
  async function s1() {
    // 构造含样式/合并/筛选/CF/公式的区块（rows 20-23，避开 demo 数据区）
    assertEq(await ev(`
      const st = __xcell.state
      const sid = st.doc.active
      let tr = st.tr.setCells(sid, [
        { row: 20, col: 0, cell: { raw: '汇总区', style: { bold: true, bg: '#e8f0fe' } } },
        { row: 21, col: 0, cell: { raw: '品名', style: { bold: true } } },
        { row: 21, col: 1, cell: { raw: '数量', style: { bold: true } } },
        { row: 22, col: 0, cell: { raw: '苹果', style: { color: '#ff0000', border: { bottom: { style: 'thin' } } } } },
        { row: 22, col: 1, cell: { raw: '3', style: { numFmt: '0.00' } } },
        { row: 23, col: 1, cell: { raw: '=B23*2' } },
      ])
      tr = tr.setMerges([{ sr: 20, sc: 0, er: 20, ec: 1 }])
      tr = tr.setFilter({ range: { sr: 21, sc: 0, er: 23, ec: 1 }, criteria: {} })
      tr = tr.setCondFormats([
        { id: 'cf1', range: { sr: 22, sc: 1, er: 23, ec: 1 }, type: 'value', op: 'gt', v1: '2',
          style: { bold: true, bg: '#ffff00' } },
      ])
      __xcell.dispatch(tr)
      return __xcell.state.activeSheet.getCell(23, 1)?.raw ?? null`), '=B23*2', 'S1 区块构造 B24 公式')
    // stub 下载捕获 Blob，点「导出 xlsx」（exceljs writeBuffer 异步，前置 bringToFront）
    await bringToFront()
    const r = await ev(`
      window.__dl = null
      const oc = URL.createObjectURL.bind(URL)
      URL.createObjectURL = (b) => { window.__dl = b; return oc(b) }
      await window.__clickFileMenu('导出 xlsx')
      await new Promise((r) => setTimeout(r, 1200))
      if (!window.__dl) throw new Error('未捕获导出 Blob')
      const bytes = new Uint8Array(await window.__dl.arrayBuffer())
      URL.createObjectURL = oc
      let bin = ''
      for (let i = 0; i < bytes.length; i += 8192) bin += String.fromCharCode(...bytes.subarray(i, i + 8192))
      return { type: window.__dl.type, len: bytes.length, pk: bytes[0] === 0x50 && bytes[1] === 0x4B, b64: btoa(bin) }`)
    assertEq(r.type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'S1 导出 Blob MIME')
    assertEq(r.len > 2000, true, 'S1 导出字节数超 2KB')
    assertEq(r.pk, true, 'S1 zip 魔数 PK')
    // runner 侧 exceljs 回读校验（附录 B 全断言）
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(Buffer.from(r.b64, 'base64'))
    const ws = wb.worksheets[0]
    assertEq(ws.name, 'Sheet1', 'S1 sheet 名')
    const a21 = ws.getCell('A21')
    assertEq(a21.value, '汇总区', 'S1 A21 值')
    assertEq(a21.font?.bold ?? null, true, 'S1 A21 粗体')
    assertEq(a21.fill?.fgColor?.argb ?? null, 'FFE8F0FE', 'S1 A21 底色')
    assertEq(ws.model.merges.includes('A21:B21'), true, 'S1 merge A21:B21')
    const a23 = ws.getCell('A23')
    assertEq(a23.value, '苹果', 'S1 A23 值')
    assertEq(a23.font?.color?.argb ?? null, 'FFFF0000', 'S1 A23 红字')
    assertEq(a23.border?.bottom?.style ?? null, 'thin', 'S1 A23 下边框')
    const b23 = ws.getCell('B23')
    assertEq(b23.value, 3, 'S1 B23 数值')
    assertEq(b23.numFmt ?? null, '0.00', 'S1 B23 numFmt')
    assertEq(ws.getCell('B24').value?.formula ?? null, 'B23*2', 'S1 B24 公式')
    assertEq(ws.autoFilter ?? null, 'A22:B24', 'S1 autoFilter ref')
    const block = (ws.model.conditionalFormattings ?? []).find((b) => b.ref === 'B23:B24')
    assertEq(!!block, true, 'S1 CF 块 B23:B24')
    assertEq(block.rules[0].type ?? null, 'cellIs', 'S1 CF 类型')
    assertEq(block.rules[0].operator ?? null, 'greaterThan', 'S1 CF 操作符')
    assertEq((block.rules[0].formulae ?? []).map(String), ['2'], 'S1 CF formulae')
    assertEq(block.rules[0].style?.fill?.bgColor?.argb ?? null, 'FFFFFF00', 'S1 CF dxf 底色')
  }

  // ---- 场景 2：导入 xlsx（runner 侧生成附录 A 文件 → 字节喂入 → 替换断言）----
  async function s2() {
    // 附录 A 生成器（runner 侧 node + exceljs）
    const gen = new ExcelJS.Workbook()
    const gws = gen.addWorksheet('数据')
    const ga1 = gws.getCell('A1')
    ga1.value = '标题'
    ga1.font = { bold: true, color: { argb: 'FFFF0000' } }
    gws.getCell('B1').value = 3
    gws.getCell('B2').value = { formula: 'B1*2', result: 6 }
    gws.getCell('D1').value = '合并区'
    gws.mergeCells('D1:E2')
    gws.views = [{ state: 'frozen', ySplit: 1 }]
    gen.addWorksheet('空表')
    await gen.xlsx.writeFile('/tmp/m4b-import.xlsx')
    const b64 = readFileSync('/tmp/m4b-import.xlsx').toString('base64')
    // 导入前现场
    assertEq(await ev(`
      window.__archiveBefore = localStorage.getItem('xexcel.workbook')
      window.__namesBefore = [...__xcell.state.doc.order].map((id) => __xcell.state.doc.names.get(id))
      window.__lastInput = null
      return window.__namesBefore`), ['Sheet1'], 'S2 导入前 demo 单表')
    await bringToFront()
    await clickMenu('打开 xlsx')
    // M5：打开 xlsx 前置 ConfirmDialog（confirmLabel「打开」），确认后才创建 pickFile input
    await ev(`window.__clickDialogBtn('打开'); return 1`)
    await sleep(300)
    assertEq(await ev(`return !!window.__lastInput`), true, 'S2 pickFile input 已捕获')
    await feedFile(b64, 'm4b-import.xlsx')
    await bringToFront()
    // 导入完成标志：sheet 数变 2 或出现 notice（失败亦走 notice）；实测节流 tab 内解析 ~5s
    await pollUntil(
      `__xcell.state.doc.order.length === 2 || window.__notice()`,
      'S2 导入完成')
    // notice 断言须在完成后 5s TTL 内（隐藏 tab 定时器节流只会更晚过期，不会更早）
    assertEq(await ev(`
      const wb = __xcell.state.doc
      return [[...wb.order].map((id) => wb.names.get(id)), wb.order.length, wb.names.get(wb.active)]`),
      [['数据', '空表'], 2, '数据'], 'S2 旧 sheet 消失 + active=第一张')
    assertEq(await ev(`
      const sh = __xcell.state.activeSheet
      const a1 = sh.getCell(0, 0)
      return [a1?.raw ?? null, a1?.style?.bold ?? null, a1?.style?.color ?? null,
        sh.getCell(0, 1)?.raw ?? null, sh.getCell(1, 1)?.raw ?? null]`),
      ['标题', true, '#ff0000', '3', '=B1*2'], 'S2 粗体红字入 model + 公式复活为 raw')
    assertEq(await ev(`
      const { evaluatorFor } = await import('${ENGINE_URL}')
      return evaluatorFor(__xcell.state.doc).get(__xcell.state.doc.active, 1, 1)`),
      6, 'S2 公式计算值（B1=3 → =B1*2）')
    assertEq(await ev(`
      const sh = __xcell.state.activeSheet
      return [sh.merges.some((m) => m.sr === 0 && m.sc === 3 && m.er === 1 && m.ec === 4), sh.frozenRows,
        __xcell.state.selection.activeCell, window.__notice()]`),
      [true, 1, { row: 0, col: 0 }, '已打开 m4b-import.xlsx'], 'S2 合并 D1:E2 + 冻结首行 + 选区 A1 + notice')
  }

  // ---- 场景 3：替换语义与自动保存（接场景 2，中间不刷新）----
  async function s3() {
    // 导入成功立即落档（saveNow），存档此刻已是新 workbook
    assertEq(await ev(`
      const arch = localStorage.getItem('xexcel.workbook')
      return [arch !== window.__archiveBefore, arch.includes('数据'), arch.includes('空表')]`),
      [true, true, true], 'S3 导入成功瞬间存档已是新 workbook')
    await ev(`W.setCell(3, 0, '验收标记'); return 1`)
    // 防抖 1s 落盘；节流 tab 内定时器可能延迟，轮询而非固定 sleep
    await pollUntil(`localStorage.getItem('xexcel.workbook')?.includes('验收标记')`, 'S3 编辑防抖落盘')
    await reload() // 刷新后 helper/stub 已重注
    assertEq(await ev(`
      const wb = __xcell.state.doc
      return [[...wb.order].map((id) => wb.names.get(id)),
        wb.activeSheet.getCell(0, 0)?.raw ?? null, wb.activeSheet.getCell(3, 0)?.raw ?? null]`),
      [['数据', '空表'], '标题', '验收标记'], 'S3 刷新后新 workbook 恢复')
  }

  // ---- 场景 4：损坏文件（notice + 现场不动 + 自动保存恢复）----
  async function s4() {
    await bringToFront()
    await ev(`window.__lastInput = null; return 1`)
    await clickMenu('打开 xlsx')
    await ev(`window.__clickDialogBtn('打开'); return 1`)
    await sleep(300)
    assertEq(await ev(`return !!window.__lastInput`), true, 'S4 pickFile input 已捕获')
    await feedFile('Z2FyYmFnZQ==', 'm4b-bad.xlsx') // 'garbage' 7 字节
    await bringToFront()
    await pollUntil(`window.__notice()`, 'S4 失败 notice')
    assertEq(await ev(`
      return [window.__notice(),
        __xcell.state.doc.names.get(__xcell.state.doc.active),
        __xcell.state.activeSheet.getCell(0, 0)?.raw ?? null,
        __xcell.state.activeSheet.getCell(3, 0)?.raw ?? null]`),
      ['文件无法解析', '数据', '标题', '验收标记'], 'S4 notice 报无法解析 + 现场分毫未动')
    // 自动保存已恢复（导入失败后 resume 生效）
    await ev(`W.setCell(4, 0, '恢复标记'); return 1`)
    await pollUntil(`localStorage.getItem('xexcel.workbook')?.includes('恢复标记')`, 'S4 自动保存恢复落盘')
    assertEq(await ev(`return window.__statusError()`),
      null, 'S4 无自动保存失败提示')
  }

  // ---- 场景 5：双 ConfirmDialog 护栏（M5：window.confirm 已由 React askConfirm 替换）----
  async function s5() {
    // 先等上场景 notice 过期（5s TTL；隐藏 tab 定时器节流，固定 5s 等待不可靠，轮询至清空）
    let notice = await ev(`return window.__notice()`)
    for (let i = 0; i < 40 && notice !== null; i++) {
      await sleep(500)
      notice = await ev(`return window.__notice()`)
    }
    assertEq(notice, null, 'S5 上场景 notice 已过期')
    // 第一道护栏：ConfirmDialog 取消 → 不出现文件选择、内容不变
    assertEq(await ev(`
      window.__lastInput = null
      window.__docBefore = __xcell.state.doc
      await window.__clickFileMenu('打开 xlsx')
      const dlg = document.querySelector('[role="dialog"]')
      const hasDlg = !!dlg && dlg.textContent.includes('替换')
      window.__clickDialogBtn('取消')
      await new Promise((r) => setTimeout(r, 300))
      return [hasDlg, window.__lastInput, __xcell.state.doc === window.__docBefore]`),
      [true, null, true], 'S5 确认框取消：文案含「替换」+ input 未创建 + doc 同一引用')
    // 第二道护栏：确认放行 + 取消文件选择 → 无动作无报错
    assertEq(await ev(`
      await window.__clickFileMenu('打开 xlsx')
      window.__clickDialogBtn('打开')
      await new Promise((r) => setTimeout(r, 300))
      if (!window.__lastInput) throw new Error('未捕获 pickFile input')
      // 模拟用户在系统对话框点「取消」：pickFile 监听 cancel 事件（Chrome 113+）
      window.__lastInput.dispatchEvent(new Event('cancel'))
      await new Promise((r) => setTimeout(r, 500))
      return [__xcell.state.doc === window.__docBefore, window.__notice(), window.__statusError()]`),
      [true, null, null], 'S5 确认放行 + 取消选择：无动作无报错')
    assertEq(await ev(`
      document.createElement = window.__origCreateElement
      return 'ok'`), 'ok', 'S5 收尾恢复 stub')
  }

  await freshPage()
  const scenarios = [
    ['1 导出 xlsx 结构校验', s1],
    ['2 导入 xlsx 替换', s2],
    ['3 落档语义+自动保存', s3],
    ['4 损坏文件', s4],
    ['5 双 confirm 护栏', s5],
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
