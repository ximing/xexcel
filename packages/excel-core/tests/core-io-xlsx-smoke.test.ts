// tests/core-io-xlsx-smoke.test.ts
// exceljs 4.4.0 API 行为钉板：本测试失败 = exceljs 升级破坏了映射层假设，先修这里再动映射。
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'

async function roundtrip(build: (ws: ExcelJS.Worksheet) => void): Promise<ExcelJS.Worksheet> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('S')
  build(ws)
  const buf = await wb.xlsx.writeBuffer()
  const wb2 = new ExcelJS.Workbook()
  await wb2.xlsx.load(buf)
  return wb2.getWorksheet('S')!
}

describe('exceljs 4.4.0 API 冒烟', () => {
  it('merges 经 ws.model.merges 回读为 A1 串数组', async () => {
    const ws = await roundtrip((w) => {
      w.getCell('A1').value = 'x'
      w.mergeCells('A4:B5')
    })
    expect(ws.model.merges).toEqual(['A4:B5'])
  })

  it('CF cellIs（含 between/notEqual）与 containsText 往返；containsText 无 text 字段', async () => {
    const ws = await roundtrip((w) => {
      w.addConditionalFormatting({
        ref: 'A1:A5',
        rules: [
          { type: 'cellIs', operator: 'between', formulae: ['1', '9'], style: { font: { bold: true } }, priority: 1 },
          { type: 'cellIs', operator: 'notEqual', formulae: ['3'], style: { font: { italic: true } }, priority: 2 },
          { type: 'containsText', operator: 'containsText', text: 'x', formulae: ['NOT(ISERROR(SEARCH("x",A1)))'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } }, priority: 3 },
        ] as never,
      })
    })
    // 注：exceljs 4.4.0 类型未声明 conditionalFormattings，只能从 model 强转读取
    const cf = (ws.model as never as { conditionalFormattings: {
      ref: string
      rules: { type: string; operator?: string; formulae?: string[]; text?: string }[]
    }[] }).conditionalFormattings
    expect(cf).toHaveLength(1)
    expect(cf[0].ref).toBe('A1:A5')
    expect(cf[0].rules.map((r) => r.type)).toEqual(['cellIs', 'cellIs', 'containsText'])
    expect(cf[0].rules[0].operator).toBe('between')
    expect(cf[0].rules[0].formulae).toEqual(['1', '9'])
    expect(cf[0].rules[1].operator).toBe('notEqual')
    // containsText 回读无 text 字段，需从 formulae 提取
    expect(cf[0].rules[2].text).toBeUndefined()
    expect(cf[0].rules[2].formulae![0]).toContain('SEARCH("x",A1)')
  })

  it('CF duplicateValues 写入即丢（不支持，导出侧须跳过）', async () => {
    const ws = await roundtrip((w) => {
      w.addConditionalFormatting({
        ref: 'A1:A5',
        rules: [{ type: 'duplicateValues', style: { font: { color: { argb: 'FF0000FF' } } }, priority: 1 }] as never,
      })
    })
    const cf = (ws.model as never as { conditionalFormattings: { rules: unknown[] }[] }).conditionalFormattings
    expect(cf[0].rules).toHaveLength(0)
  })

  it('autoFilter 回读只有 ref 字符串（criteria 不回读）', async () => {
    const ws = await roundtrip((w) => {
      w.getCell('A10').value = 'h'
      w.autoFilter = {
        from: 'A10',
        to: 'D20',
        filters: [{ column: 0, filters: ['a', 'b'] }, { column: 1, customFilters: [{ operator: 'greaterThan', value: 5 }] }],
      } as never
    })
    expect(ws.autoFilter).toBe('A10:D20')
  })

  it('冻结 views / 有内容行 hidden / 列 hidden / 尺寸 往返', async () => {
    const ws = await roundtrip((w) => {
      w.getCell('A6').value = 'row6'
      w.getRow(6).hidden = true
      w.getRow(7).height = 30
      w.getColumn('C').hidden = true
      w.getColumn('D').width = 20
      w.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }]
    })
    expect(ws.getRow(6).hidden).toBe(true)
    expect(ws.getRow(7).height).toBe(30)
    expect(ws.getColumn('C').hidden).toBe(true)
    expect(ws.getColumn('D').width).toBe(20)
    // 注：视图联合类型中只有 frozen 分支有 xSplit/ySplit，强转后断言
    const v = ws.views![0] as { state?: string; xSplit?: number; ySplit?: number }
    expect(v.state).toBe('frozen')
    expect(v.xSplit).toBe(1)
    expect(v.ySplit).toBe(2)
  })

  it('值形状：公式 {formula} / 日期 Date / 布尔 / richText', async () => {
    const ws = await roundtrip((w) => {
      w.getCell('A1').value = { formula: 'B1*2' }
      w.getCell('A2').value = new Date(Date.UTC(2026, 0, 15))
      w.getCell('A3').value = true
      w.getCell('A4').value = { richText: [{ text: 'a' }, { text: 'b' }] }
    })
    expect(ws.getCell('A1').value).toEqual({ formula: 'B1*2' })
    expect(ws.getCell('A2').value instanceof Date).toBe(true)
    expect(ws.getCell('A3').value).toBe(true)
    expect(ws.getCell('A4').value).toEqual({ richText: [{ text: 'a' }, { text: 'b' }] })
  })
})
