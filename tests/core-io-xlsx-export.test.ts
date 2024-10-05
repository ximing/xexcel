// tests/core-io-xlsx-export.test.ts
import { describe, expect, it } from 'vitest'
import { FilterState, SheetData, Workbook } from '../src/core/model'
import { buildAutoFilter, cellValueToExcel, cfRuleToExcel, workbookToExcelJS } from '../src/core/io/xlsx'

describe('cellValueToExcel', () => {
  it('公式去 = 前缀；TRUE/FALSE 转布尔；数字串转 number；其余原文', () => {
    expect(cellValueToExcel({ raw: '=A1+B1' })).toEqual({ formula: 'A1+B1' })
    expect(cellValueToExcel({ raw: 'TRUE' })).toBe(true)
    expect(cellValueToExcel({ raw: 'FALSE' })).toBe(false)
    expect(cellValueToExcel({ raw: '42.5' })).toBe(42.5)
    expect(cellValueToExcel({ raw: '-3' })).toBe(-3)
    expect(cellValueToExcel({ raw: 'hello' })).toBe('hello')
    expect(cellValueToExcel({ raw: '' })).toBe('')
  })
})

const sheetWith = (cells: [number, number, string][]): SheetData => {
  let s = SheetData.create({ rowCount: 100, colCount: 26 })
  for (const [r, c, raw] of cells) s = s.setCell(r, c, { raw })
  return s
}

describe('buildAutoFilter', () => {
  const range = { sr: 0, sc: 0, er: 5, ec: 2 }
  it('values 模式：包含集 = 数据区去重 − 排除集', () => {
    const sheet = sheetWith([[1, 0, 'a'], [2, 0, 'b'], [3, 0, 'a'], [4, 0, 'c']])
    const filter: FilterState = { range, criteria: { 0: { type: 'values', excluded: ['b'] } } }
    const af = buildAutoFilter(filter, sheet)
    expect(af.from).toBe('A1')
    expect(af.to).toBe('C6')
    expect(af.filters).toEqual([{ column: 0, filters: ['a', 'c'] }])
  })
  it('数值条件：op 直通；between → 两条（and）', () => {
    const sheet = sheetWith([[1, 1, '1']])
    const filter: FilterState = {
      range,
      criteria: {
        1: { type: 'condition', field: 'num', op: 'gt', v1: '5' },
        2: { type: 'condition', field: 'num', op: 'between', v1: '1', v2: '9' },
      },
    }
    const af = buildAutoFilter(filter, sheet)
    expect(af.filters).toEqual([
      { column: 1, customFilters: [{ operator: 'greaterThan', value: 5 }] },
      { column: 2, customFilters: [
        { operator: 'greaterThanOrEqual', value: 1 },
        { operator: 'lessThanOrEqual', value: 9 },
      ] },
    ])
  })
  it('文本条件 op 省略该列 + warn；criteria 为空 → 只有 from/to', () => {
    const sheet = sheetWith([[1, 0, 'a']])
    const filter: FilterState = { range, criteria: { 0: { type: 'condition', field: 'text', op: 'contains', v1: 'a' } } }
    const af = buildAutoFilter(filter, sheet)
    expect(af.filters).toBeUndefined()
    const only: FilterState = { range, criteria: {} }
    expect(buildAutoFilter(only, sheet)).toEqual({ from: 'A1', to: 'C6' })
  })
})

describe('cfRuleToExcel', () => {
  it('value → cellIs（between 双 formulae）', () => {
    const b = cfRuleToExcel(
      { id: 'cf1', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'value', op: 'between', v1: '1', v2: '9', style: { bold: true } },
      1,
    )
    expect(b!.ref).toBe('A1:A10')
    expect(b!.rules[0]).toMatchObject({ type: 'cellIs', operator: 'between', formulae: ['1', '9'], priority: 1 })
  })
  it('textContains → containsText（引号翻倍，锚点左上角）', () => {
    const b = cfRuleToExcel(
      { id: 'cf2', range: { sr: 2, sc: 1, er: 5, ec: 1 }, type: 'textContains', text: '说"明', style: { bg: '#ffff00' } },
      2,
    )
    expect(b!.ref).toBe('B3:B6')
    expect(b!.rules[0]).toMatchObject({
      type: 'containsText',
      text: '说"明',
      formulae: ['NOT(ISERROR(SEARCH("说""明",B3)))'],
      priority: 2,
    })
  })
  it('duplicate → null（exceljs 不支持）', () => {
    expect(
      cfRuleToExcel({ id: 'cf3', range: { sr: 0, sc: 0, er: 1, ec: 0 }, type: 'duplicate', style: {} }, 1),
    ).toBeNull()
  })
})

describe('workbookToExcelJS 装配', () => {
  it('多 sheet 顺序/命名/样式/合并/冻结/筛选/CF', () => {
    let s1 = SheetData.create({ rowCount: 100, colCount: 26 })
    s1 = s1.setCell(0, 0, { raw: '标题', style: { bold: true, bg: '#e8f0fe', align: 'center' } })
    s1 = s1.setCell(1, 0, { raw: '=1+1' })
    s1 = s1.setMerges([{ sr: 3, sc: 0, er: 4, ec: 1 }])
    s1 = s1.setFrozen(1, 2)
    s1 = s1.setFilter({ range: { sr: 0, sc: 0, er: 5, ec: 1 }, criteria: {} })
    s1 = s1.setCondFormats([
      { id: 'cf1', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'value', op: 'gt', v1: '10', style: { color: '#ff0000' } },
    ])
    s1 = s1.setRowHeight(2, 30)
    s1 = s1.setColWidth(1, 140)
    let wb = Workbook.create({ rowCount: 1, colCount: 1 })
    wb = wb.setSheet('s1', s1)
    wb = wb.renameSheet('s1', '数据')
    wb = wb.addSheet('s2', SheetData.create({ rowCount: 50, colCount: 10 }), undefined, 'Empty')

    const ewb = workbookToExcelJS(wb)
    expect(ewb.worksheets.map((w) => w.name)).toEqual(['数据', 'Empty'])
    const ws = ewb.worksheets[0]
    const a1 = ws.getCell('A1')
    expect(a1.value).toBe('标题')
    expect(a1.font).toMatchObject({ bold: true, name: undefined })
    expect(a1.fill).toEqual({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } })
    expect(ws.getCell('A2').value).toEqual({ formula: '1+1' })
    expect(ws.model.merges).toEqual(['A4:B5'])
    expect(ws.views![0]).toMatchObject({ state: 'frozen', xSplit: 2, ySplit: 1 })
    expect(ws.autoFilter).toEqual({ from: 'A1', to: 'B6' })
    expect(ws.getRow(3).height).toBeCloseTo(22.5)
    expect(ws.getColumn(2).width).toBeCloseTo((140 - 5) / 7)
    // exceljs 4.4.0 类型未声明 conditionalFormattings，整体强转后取值（同 smoke 测试写法）
    const cf = (ws.model as never as { conditionalFormattings: { ref: string; rules: { operator?: string }[] }[] }).conditionalFormattings
    expect(cf[0].ref).toBe('A1:A10')
    expect(cf[0].rules[0].operator).toBe('greaterThan')
  })
})
