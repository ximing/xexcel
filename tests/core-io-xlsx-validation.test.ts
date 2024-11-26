import ExcelJS from 'exceljs'
import { describe, expect, it, vi } from 'vitest'
import { SheetData, ValidationRule, Workbook } from '../src/core/model'
import { excelJSToWorkbook, parseXlsx, workbookToExcelJS } from '../src/core/io/xlsx'

const num: ValidationRule = { id: 'v1', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'numRange', op: 'between', v1: '1', v2: '9' }
const len: ValidationRule = { id: 'v2', range: { sr: 1, sc: 1, er: 5, ec: 1 }, type: 'textLen', op: 'lt', v1: '10' }
const list: ValidationRule = { id: 'v3', range: { sr: 0, sc: 2, er: 5, ec: 2 }, type: 'list', items: ['a', 'b', 'c'] }

// exceljs d.ts 未声明 Worksheet.dataValidations（运行时存在），结构化补型（同 dv-smoke 写法）
interface DataValidationsLike {
  add(address: string, def: { type: string; operator?: string; formulae?: string[] }): void
  model: Record<string, { type: string; operator?: string; formulae?: unknown[] }>
}
function dvs(ws: ExcelJS.Worksheet): DataValidationsLike {
  return (ws as never as { dataValidations: DataValidationsLike }).dataValidations
}

describe('xlsx dataValidations 映射', () => {
  it('导出：三类型写入正确 type/operator/formulae', () => {
    let wb = Workbook.create({ rowCount: 1, colCount: 1 })
    wb = wb.setSheet('s1', SheetData.create({ rowCount: 100, colCount: 26 }).setValidations([num, len, list]))
    const ws = workbookToExcelJS(wb).worksheets[0]
    const model = dvs(ws).model
    // 写盘前 model 以 range ref 为键（逐地址展开发生在 writeBuffer→load 后，见 dv-smoke 钉板）
    expect(model['A1:A10']).toMatchObject({ type: 'decimal', operator: 'between' })
    expect(model['A1:A10'].formulae).toEqual(['1', '9'])
    expect(model['B2:B6']).toMatchObject({ type: 'textLength', operator: 'lessThan' })
    expect(model['C1:C6']).toMatchObject({ type: 'list' })
    expect(model['C1:C6'].formulae).toEqual(['"a,b,c"'])
  })

  it('往返：导出→writeBuffer→load→导入 规则保真（range 矩形合并）', async () => {
    let wb = Workbook.create({ rowCount: 1, colCount: 1 })
    wb = wb.setSheet('s1', SheetData.create({ rowCount: 100, colCount: 26 }).setValidations([num, len, list]))
    const buf = await workbookToExcelJS(wb).xlsx.writeBuffer()
    const back = await parseXlsx(buf as unknown as Uint8Array)
    const rules = back.sheet('s1').validations
    expect(rules).toHaveLength(3)
    expect(rules[0]).toMatchObject({ type: 'numRange', op: 'between', v1: '1', v2: '9', range: num.range })
    expect(rules[1]).toMatchObject({ type: 'textLen', op: 'lt', v1: '10', range: len.range })
    expect(rules[2]).toMatchObject({ type: 'list', items: ['a', 'b', 'c'], range: list.range })
  })

  it('导入：whole→numRange；未知类型跳过；list 解析引号串', async () => {
    const ewb = new ExcelJS.Workbook()
    const ws = ewb.addWorksheet('S')
    ws.getCell('A1').value = 1
    dvs(ws).add('A1:A2', { type: 'whole', operator: 'greaterThan', formulae: ['0'] })
    dvs(ws).add('B1', { type: 'date', operator: 'equal', formulae: ['44927'] })
    const buf = await ewb.xlsx.writeBuffer()
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(buf)
    const wb = excelJSToWorkbook(wb2)
    const rules = wb.sheet('s1').validations
    expect(rules).toHaveLength(1)
    expect(rules[0]).toMatchObject({ type: 'numRange', op: 'gt', v1: '0', range: { sr: 0, sc: 0, er: 1, ec: 0 } })
  })

  it('导入：验证 range 超出内容边界时 clamp 到 sheet 边界（整列场景）', async () => {
    // 真实 Excel 整列验证（A:A）等价于展开后远超内容边界；exceljs API 写 A:A 会崩，用 A1:A500 模拟
    const ewb = new ExcelJS.Workbook()
    const ws = ewb.addWorksheet('S')
    ws.getCell('A1').value = 1
    dvs(ws).add('A1:A500', { type: 'whole', operator: 'greaterThan', formulae: ['0'] })
    const buf = await ewb.xlsx.writeBuffer()
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(buf)
    const wb = excelJSToWorkbook(wb2)
    const rules = wb.sheet('s1').validations
    expect(rules).toHaveLength(1)
    // 内容仅 1 行 → rowCount=IMPORT_MIN_ROWS(100)，er clamp 到 99
    expect(rules[0].range).toEqual({ sr: 0, sc: 0, er: 99, ec: 0 })
  })

  it('导入：between 缺第二个公式时警告并跳过（numRange/textLen）', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      const ewb = new ExcelJS.Workbook()
      const ws = ewb.addWorksheet('S')
      ws.getCell('A1').value = 1
      dvs(ws).add('A1:A2', { type: 'whole', operator: 'between', formulae: ['5'] })
      dvs(ws).add('B1:B2', { type: 'textLength', operator: 'between', formulae: ['3'] })
      // 对照组：完整 between 不受影响
      dvs(ws).add('C1:C2', { type: 'whole', operator: 'between', formulae: ['1', '9'] })
      const buf = await ewb.xlsx.writeBuffer()
      const wb2 = new ExcelJS.Workbook()
      await wb2.xlsx.load(buf)
      const rules = excelJSToWorkbook(wb2).sheet('s1').validations
      expect(rules).toHaveLength(1)
      expect(rules[0]).toMatchObject({ type: 'numRange', op: 'between', v1: '1', v2: '9' })
      expect(warn).toHaveBeenCalled()
    } finally {
      warn.mockRestore()
    }
  })
})
