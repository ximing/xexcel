// tests/core-io-xlsx-import.test.ts
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { Workbook } from '../src/core/model'
import { DAY_MS, EPOCH } from '../src/formula/date'
import { excelJSToWorkbook, parseXlsx, workbookToExcelJS } from '../src/core/io/xlsx'
import { SheetData } from '../src/core/model'

// 直接装配 exceljs workbook → 导入
async function importOf(build: (wb: ExcelJS.Workbook) => void): Promise<Workbook> {
  const ewb = new ExcelJS.Workbook()
  build(ewb)
  const buf = await ewb.xlsx.writeBuffer()
  const wb2 = new ExcelJS.Workbook()
  await wb2.xlsx.load(buf)
  return excelJSToWorkbook(wb2)
}

describe('excelJSToWorkbook 值映射', () => {
  it('字符串/数字/布尔/公式/日期/richText', async () => {
    const wb = await importOf((ewb) => {
      const ws = ewb.addWorksheet('S')
      ws.getCell('A1').value = 'hello'
      ws.getCell('B1').value = 42.5
      ws.getCell('C1').value = true
      ws.getCell('D1').value = { formula: 'B1*2' }
      ws.getCell('E1').value = new Date(Date.UTC(2026, 0, 15))
      ws.getCell('E1').numFmt = 'yyyy-mm-dd'
      ws.getCell('F1').value = { richText: [{ text: 'a' }, { text: 'b' }] }
    })
    const s = wb.sheet('s1')
    expect(s.getCell(0, 0)?.raw).toBe('hello')
    expect(s.getCell(0, 1)?.raw).toBe('42.5')
    expect(s.getCell(0, 2)?.raw).toBe('TRUE')
    expect(s.getCell(0, 3)?.raw).toBe('=B1*2')
    const serial = (Date.UTC(2026, 0, 15) - EPOCH) / DAY_MS
    expect(s.getCell(0, 4)?.raw).toBe(String(serial))
    expect(s.getCell(0, 4)?.style?.numFmt).toBe('yyyy-mm-dd')
    expect(s.getCell(0, 5)?.raw).toBe('ab')
  })
  it('sheet 尺寸下限 100×26；active=第一张；多 sheet 顺序命名', async () => {
    const wb = await importOf((ewb) => {
      ewb.addWorksheet('甲').getCell('B2').value = 1
      ewb.addWorksheet('乙')
    })
    expect(wb.order).toEqual(['s1', 's2'])
    expect(wb.names.get('s1')).toBe('甲')
    expect(wb.active).toBe('s1')
    expect(wb.sheet('s1').rowCount).toBe(100)
    expect(wb.sheet('s1').colCount).toBe(26)
  })
  it('无 sheet 抛错', () => {
    expect(() => excelJSToWorkbook(new ExcelJS.Workbook())).toThrow()
  })
})

describe('excelJSToWorkbook 结构映射', () => {
  it('合并/冻结/隐藏/尺寸/筛选(range-only)/样式', async () => {
    const wb = await importOf((ewb) => {
      const ws = ewb.addWorksheet('S')
      ws.getCell('A1').value = 'h'
      ws.getCell('A1').font = { bold: true, color: { argb: 'FFFF0000' } }
      ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } }
      ws.getCell('A6').value = 'row6'
      ws.getRow(6).hidden = true
      ws.getRow(7).height = 30
      ws.getColumn(2).hidden = true
      ws.getColumn(3).width = 20
      ws.mergeCells('A4:B5')
      ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }]
      ws.autoFilter = 'A1:D10' as never
    })
    const s = wb.sheet('s1')
    expect(s.getCell(0, 0)?.style).toMatchObject({ bold: true, color: '#ff0000', bg: '#e8f0fe' })
    expect(s.merges).toEqual([{ sr: 3, sc: 0, er: 4, ec: 1 }])
    expect(s.frozenCols).toBe(1)
    expect(s.frozenRows).toBe(2)
    expect(s.hiddenRows).toEqual([5])
    expect(s.hiddenCols).toEqual([1])
    expect(s.customRowHeights.get(6)).toBe(40) // 30pt / 0.75
    expect(s.customColWidths.get(2)).toBe(20 * 7 + 5)
    expect(s.filter).toEqual({ range: { sr: 0, sc: 0, er: 9, ec: 3 }, criteria: {} })
  })
  it('CF：cellIs（between 双值）与 containsText（formulae 提取 text）；未知类型跳过', async () => {
    const wb = await importOf((ewb) => {
      const ws = ewb.addWorksheet('S')
      ws.getCell('A1').value = 1
      ws.addConditionalFormatting({
        ref: 'A1:A5',
        rules: [
          { type: 'cellIs', operator: 'between', formulae: ['1', '9'], style: { font: { bold: true, color: { argb: 'FFFF0000' } } }, priority: 1 },
          { type: 'containsText', operator: 'containsText', formulae: ['NOT(ISERROR(SEARCH("说""明",A1)))'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' } } }, priority: 2 },
          // 注：brief 原文 colorScale 无 cfvo/color，exceljs 4.4.0 writeBuffer 会抛 TypeError；
          // 补最小必需字段使其可落盘（断言不变：导入仍跳过该未知类型）
          { type: 'colorScale', priority: 3, cfvo: [{ type: 'min' }, { type: 'max' }], color: [{ argb: 'FFFF0000' }, { argb: 'FF00FF00' }] },
        ] as never,
      })
    })
    const rules = wb.sheet('s1').condFormats
    expect(rules).toHaveLength(2)
    expect(rules[0]).toMatchObject({
      type: 'value', op: 'between', v1: '1', v2: '9',
      range: { sr: 0, sc: 0, er: 4, ec: 0 },
      style: { bold: true, color: '#ff0000' },
    })
    expect(rules[1]).toMatchObject({
      type: 'textContains', text: '说"明',
      range: { sr: 0, sc: 0, er: 4, ec: 0 },
      style: { bg: '#ffff00' },
    })
  })
})

describe('全量往返（export → writeBuffer → load → import）', () => {
  it('样式/numFmt/合并/冻结/筛选/CF/公式 保持', async () => {
    let s1 = SheetData.create({ rowCount: 100, colCount: 26 })
    s1 = s1.setCell(0, 0, { raw: '标题', style: { bold: true, bg: '#e8f0fe', align: 'center', vAlign: 'middle', wrap: true } })
    s1 = s1.setCell(1, 0, { raw: '=B2*2' })
    s1 = s1.setCell(1, 1, { raw: '42', style: { numFmt: '0.00%' } })
    s1 = s1.setCell(2, 0, {
      raw: ' bordered ',
      style: { border: { top: { style: 'thin' }, bottom: { style: 'double', color: '#0000ff' } } },
    })
    s1 = s1.setMerges([{ sr: 4, sc: 0, er: 5, ec: 1 }])
    s1 = s1.setFrozen(1, 1)
    s1 = s1.setRowHeight(2, 36)
    s1 = s1.setColWidth(1, 120)
    s1 = s1.setFilter({
      range: { sr: 0, sc: 0, er: 5, ec: 1 },
      criteria: { 0: { type: 'condition', field: 'num', op: 'gt', v1: '5' } },
    })
    s1 = s1.setCondFormats([
      { id: 'cf1', range: { sr: 1, sc: 1, er: 9, ec: 1 }, type: 'value', op: 'gte', v1: '10', style: { color: '#ff0000', bold: true } },
      { id: 'cf2', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'textContains', text: '标', style: { bg: '#ffff00' } },
    ])
    let wb = Workbook.create({ rowCount: 1, colCount: 1 })
    wb = wb.setSheet('s1', s1)
    wb = wb.renameSheet('s1', '数据')
    wb = wb.addSheet('s2', SheetData.create({ rowCount: 50, colCount: 10 }), undefined, 'Second')

    const buf = await workbookToExcelJS(wb).xlsx.writeBuffer()
    const back = await parseXlsx(buf as unknown as Uint8Array)

    expect(back.order).toEqual(['s1', 's2'])
    expect(back.names.get('s1')).toBe('数据')
    const s = back.sheet('s1')
    expect(s.getCell(0, 0)?.style).toMatchObject({
      bold: true, bg: '#e8f0fe', align: 'center', vAlign: 'middle', wrap: true,
    })
    expect(s.getCell(1, 0)?.raw).toBe('=B2*2')
    expect(s.getCell(1, 1)?.style?.numFmt).toBe('0.00%')
    expect(s.getCell(2, 0)?.style?.border).toEqual({
      top: { style: 'thin' },
      bottom: { style: 'double', color: '#0000ff' },
    })
    expect(s.merges).toEqual([{ sr: 4, sc: 0, er: 5, ec: 1 }])
    expect(s.frozenRows).toBe(1)
    expect(s.frozenCols).toBe(1)
    // 注：brief 原文断言 48 系笔误（36px→导出 27pt→回读 27pt→/0.75=36px，恒等往返；列宽断言同为恒等语义）
    expect(s.customRowHeights.get(2)).toBe(36)
    expect(s.customColWidths.get(1)).toBeCloseTo(Math.round(((120 - 5) / 7) * 7 + 5), 0)
    // 筛选 criteria 导出写入但回读不可得（exceljs 限制）→ range-only
    expect(s.filter).toEqual({ range: { sr: 0, sc: 0, er: 5, ec: 1 }, criteria: {} })
    const rules = s.condFormats
    expect(rules).toHaveLength(2)
    expect(rules[0]).toMatchObject({ type: 'value', op: 'gte', v1: '10', style: { color: '#ff0000', bold: true } })
    expect(rules[1]).toMatchObject({ type: 'textContains', text: '标', style: { bg: '#ffff00' } })
  })
})
