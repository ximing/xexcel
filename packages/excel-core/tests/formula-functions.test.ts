import { describe, expect, it } from 'vitest'
import { fromA1 } from '../src/core/addr'
import { Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'
import { nowSerial } from '../src/formula/date'

// 与实现无关的 serial 推导（Excel 1900 系统：基准 1899-12-30），避免用被测函数验证被测函数
const serialOf = (y: number, m: number, d: number): number =>
  Math.floor((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000)

function wbWith(cells: Record<string, string>): Workbook {
  let wb = Workbook.create({ rowCount: 30, colCount: 10 })
  let data = wb.activeSheet
  for (const [a1, raw] of Object.entries(cells)) {
    const { row, col } = fromA1(a1)!
    data = data.setCell(row, col, { raw })
  }
  return wb.setSheet(wb.active, data)
}

function get(wb: Workbook, a1: string): unknown {
  const { row, col } = fromA1(a1)!
  return evaluatorFor(wb).get(wb.active, row, col)
}

describe('逻辑函数', () => {
  it('AND/OR/NOT', () => {
    const wb = wbWith({ A1: '=AND(TRUE,1)', A2: '=AND(TRUE,0)', A3: '=OR(FALSE,0)', A4: '=OR(FALSE,2)', A5: '=NOT(0)', A6: '=NOT(1)' })
    expect(get(wb, 'A1')).toBe(true)
    expect(get(wb, 'A2')).toBe(false)
    expect(get(wb, 'A3')).toBe(false)
    expect(get(wb, 'A4')).toBe(true)
    expect(get(wb, 'A5')).toBe(true)
    expect(get(wb, 'A6')).toBe(false)
  })

  it('AND 接受区域参数（空格跳过、布尔/数字参与）', () => {
    const wb = wbWith({ A1: '1', A2: '2', A3: '', B1: '=AND(A1:A3)', B2: '=AND(A1:A2,B1)' })
    expect(get(wb, 'B1')).toBe(true)
    expect(get(wb, 'B2')).toBe(true) // B1 求值为 true，与 A1:A2 一起全真
  })

  it('AND 无有效值 → #VALUE!；文本 → #VALUE!', () => {
    const wb = wbWith({ A1: '=AND("abc")', A2: '=AND()' })
    expect(get(wb, 'A1')).toEqual({ error: '#VALUE!' })
    expect(get(wb, 'A2')).toEqual({ error: '#VALUE!' })
  })

  it('IFERROR 兜住错误', () => {
    const wb = wbWith({ A1: '=IFERROR(1/0,"x")', A2: '=IFERROR(5,"x")', A3: '=IFERROR(NOPE,"name")' })
    expect(get(wb, 'A1')).toBe('x')
    expect(get(wb, 'A2')).toBe(5)
    expect(get(wb, 'A3')).toBe('name')
  })
})

describe('文本函数', () => {
  it('LEN/LEFT/RIGHT/MID', () => {
    const wb = wbWith({ A1: '=LEN("abc")', A2: '=LEFT("abc",2)', A3: '=LEFT("abc")', A4: '=RIGHT("abc",2)', A5: '=MID("abcdef",2,3)', A6: '=LEN(123)' })
    expect(get(wb, 'A1')).toBe(3)
    expect(get(wb, 'A2')).toBe('ab')
    expect(get(wb, 'A3')).toBe('a')
    expect(get(wb, 'A4')).toBe('bc')
    expect(get(wb, 'A5')).toBe('bcd')
    expect(get(wb, 'A6')).toBe(3)
  })

  it('UPPER/LOWER/TRIM', () => {
    const wb = wbWith({ A1: '=UPPER("aBc")', A2: '=LOWER("aBc")', A3: '=TRIM("  a   b  ")' })
    expect(get(wb, 'A1')).toBe('ABC')
    expect(get(wb, 'A2')).toBe('abc')
    expect(get(wb, 'A3')).toBe('a b')
  })

  it('CONCAT 标量与区域', () => {
    const wb = wbWith({ A1: 'x', A2: 'y', B1: '=CONCAT("a",1,"b")', B2: '=CONCAT(A1:A2,"z")' })
    expect(get(wb, 'B1')).toBe('a1b')
    expect(get(wb, 'B2')).toBe('xyz')
  })

  it('MID 起始小于 1 → #VALUE!', () => {
    const wb = wbWith({ A1: '=MID("abc",0,2)' })
    expect(get(wb, 'A1')).toEqual({ error: '#VALUE!' })
  })
})

describe('条件聚合', () => {
  it('SUMIF 数值条件（文本格不参与求和）', () => {
    const wb2 = wbWith({ A1: '1', A2: '2', A3: '3', A4: 'abc', C1: '=SUMIF(A1:A4,">1")' })
    expect(get(wb2, 'C1')).toBe(5)
  })

  it('SUMIF 带求和域（第三参）', () => {
    const w = wbWith({ A1: 'x', A2: 'y', A3: 'x', B1: '10', B2: '20', B3: '30', C1: '=SUMIF(A1:A3,"x",B1:B3)' })
    expect(get(w, 'C1')).toBe(40)
  })

  it('COUNTIF 通配与精确', () => {
    const w = wbWith({ A1: 'apple', A2: 'apricot', A3: 'banana', C1: '=COUNTIF(A1:A3,"a*")', C2: '=COUNTIF(A1:A3,"banana")', C3: '=COUNTIF(A1:A3,"BANANA")' })
    expect(get(w, 'C1')).toBe(2)
    expect(get(w, 'C2')).toBe(1)
    expect(get(w, 'C3')).toBe(1)
  })

  it('AVERAGEIF 无匹配 → #DIV/0!', () => {
    const w = wbWith({ A1: '1', A2: '2', C1: '=AVERAGEIF(A1:A2,">100")', C2: '=AVERAGEIF(A1:A2,">1")' })
    expect(get(w, 'C1')).toEqual({ error: '#DIV/0!' })
    expect(get(w, 'C2')).toBe(2)
  })

  it('COUNTIF 参数个数错误 → #VALUE!', () => {
    const w = wbWith({ A1: '=COUNTIF(A1:A2)' })
    expect(get(w, 'A1')).toEqual({ error: '#VALUE!' })
  })

  it('SUMIF 求和域越界格按 0/空处理，不传播 #REF!（对齐 Excel）', () => {
    // 30 行表：求和域锚 B29，A30 匹配时偏移到 B31（越界）→ 按 0
    const w = wbWith({ A28: 'x', A29: 'x', A30: 'x', B29: '10', C1: '=SUMIF(A28:A30,"x",B29)' })
    expect(get(w, 'C1')).toBe(10)
  })

  it('SUMIF 求和域内真实 #REF! 仍传播', () => {
    const w = wbWith({ A1: 'x', B1: '=Z99', C1: '=SUMIF(A1,"x",B1)' })
    expect(get(w, 'C1')).toEqual({ error: '#REF!' })
  })
})

describe('日期函数', () => {
  it('DATE/YEAR/MONTH/DAY 往返', () => {
    const wb = wbWith({ A1: '=DATE(2026,7,31)', A2: '=YEAR(A1)', A3: '=MONTH(A1)', A4: '=DAY(A1)' })
    expect(get(wb, 'A1')).toBe(serialOf(2026, 7, 31))
    expect(get(wb, 'A2')).toBe(2026)
    expect(get(wb, 'A3')).toBe(7)
    expect(get(wb, 'A4')).toBe(31)
  })

  it('DATE 溢出进位（Excel 语义）', () => {
    const wb = wbWith({ A1: '=YEAR(DATE(2026,13,1))', A2: '=MONTH(DATE(2026,13,1))' })
    expect(get(wb, 'A1')).toBe(2027)
    expect(get(wb, 'A2')).toBe(1)
  })

  it('TODAY/NOW 为当前 serial（易失，全量重算自然刷新）', () => {
    const wb = wbWith({ A1: '=TODAY()', A2: '=NOW()' })
    // 外部推导 TODAY：本地今日 serial；跨午夜边界时取前后两值之一
    const localSerial = (): number => {
      const n = new Date()
      return serialOf(n.getFullYear(), n.getMonth() + 1, n.getDate())
    }
    const before = localSerial()
    const today = get(wb, 'A1') as number
    expect(today === before || today === localSerial()).toBe(true)
    const n = get(wb, 'A2') as number
    expect(Math.abs(n - nowSerial())).toBeLessThan(2 / 86400) // 两秒容差
  })

  it('YEAR 负 serial → #VALUE!', () => {
    const wb = wbWith({ A1: '=YEAR(-1)' })
    expect(get(wb, 'A1')).toEqual({ error: '#VALUE!' })
  })

  it('带 numFmt 的 TODAY 显示日期', () => {
    let wb = Workbook.create({ rowCount: 5, colCount: 5 })
    const data = wb.activeSheet.setCell(0, 0, { raw: '=TODAY()', style: { numFmt: 'yyyy/m/d' } })
    wb = wb.setSheet(wb.active, data)
    const text = evaluatorFor(wb).displayText(wb.active, 0, 0)
    expect(text).toMatch(/^\d{4}\/\d{1,2}\/\d{1,2}$/)
  })
})
