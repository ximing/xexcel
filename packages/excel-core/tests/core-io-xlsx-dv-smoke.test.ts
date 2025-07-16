// exceljs 4.4.0 dataValidations 行为钉板：失败 = 升级破坏假设，先修这里再动映射。
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'

// 4.4.0 官方 d.ts 未声明 Worksheet.dataValidations（运行时存在），此处结构化补型
interface DataValidationsLike {
  add(
    address: string,
    def: { type: string; operator?: string; formulae?: string[]; allowBlank?: boolean }
  ): void
  model: Record<string, { type: string; operator?: string; formulae?: unknown[] }>
}

function dvs(ws: ExcelJS.Worksheet): DataValidationsLike {
  return (ws as never as { dataValidations: DataValidationsLike }).dataValidations
}

async function roundtrip(build: (ws: ExcelJS.Worksheet) => void): Promise<ExcelJS.Worksheet> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('S')
  build(ws)
  const buf = await wb.xlsx.writeBuffer()
  const wb2 = new ExcelJS.Workbook()
  await wb2.xlsx.load(buf)
  return wb2.getWorksheet('S')!
}

describe('exceljs dataValidations 冒烟', () => {
  it('decimal/textLength/list/whole 往返；读回为按地址展开的 model map', async () => {
    const ws = await roundtrip((w) => {
      w.getCell('A1').value = 1
      dvs(w).add('A1:A3', { type: 'decimal', operator: 'between', formulae: ['1', '9'], allowBlank: true })
      dvs(w).add('B1:B2', { type: 'textLength', operator: 'lessThan', formulae: ['10'] })
      dvs(w).add('C1:C2', { type: 'list', formulae: ['"a,b,c"'] })
      dvs(w).add('D1', { type: 'whole', operator: 'greaterThan', formulae: ['0'] })
    })
    const model = dvs(ws).model
    // 范围展开为逐地址
    expect(Object.keys(model)).toEqual(['A1', 'A2', 'A3', 'B1', 'B2', 'C1', 'C2', 'D1'])
    expect(model.A1).toMatchObject({ type: 'decimal', operator: 'between' })
    expect(model.A1.formulae).toEqual([1, 9]) // 数字回读为 number
    expect(model.B1).toMatchObject({ type: 'textLength', operator: 'lessThan' })
    expect(model.B1.formulae).toEqual([10])
    expect(model.C1).toMatchObject({ type: 'list' })
    expect(model.C1.formulae).toEqual(['"a,b,c"']) // list 保留带引号串
    expect(model.D1).toMatchObject({ type: 'whole', operator: 'greaterThan' })
  })
})
