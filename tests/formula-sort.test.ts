import { describe, expect, it } from 'vitest'
import { fromA1 } from '../src/core/addr'
import { Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'
import { computeSortEntries, sortBlockedByMerges } from '../src/formula/sort'

function wbWith(cells: Record<string, string>): Workbook {
  let wb = Workbook.create({ rowCount: 20, colCount: 10 })
  let data = wb.activeSheet
  for (const [a1, raw] of Object.entries(cells)) {
    const { row, col } = fromA1(a1)!
    data = data.setCell(row, col, { raw })
  }
  return wb.setSheet(wb.active, data)
}

// 应用排序条目后读指定列的 raw 序列
function sortedRaws(wb: ReturnType<typeof wbWith>, range: { sr: number; sc: number; er: number; ec: number }, keys: { col: number; asc: boolean }[], hasHeader: boolean, readCol: number): string[] {
  const data = wb.activeSheet
  const entries = computeSortEntries(data, wb.active, evaluatorFor(wb), range, keys, hasHeader)
  let next = data
  for (const e of entries) next = next.setCell(e.row, e.col, e.cell)
  const out: string[] = []
  for (let r = range.sr; r <= range.er; r++) out.push(next.getCell(r, readCol)?.raw ?? '')
  return out
}

const range = { sr: 0, sc: 0, er: 3, ec: 1 }

describe('computeSortEntries', () => {
  it('数值升序/降序', () => {
    const wb = wbWith({ A1: '3', A2: '1', A3: '4', A4: '2' })
    expect(sortedRaws(wb, range, [{ col: 0, asc: true }], false, 0)).toEqual(['1', '2', '3', '4'])
    expect(sortedRaws(wb, range, [{ col: 0, asc: false }], false, 0)).toEqual(['4', '3', '2', '1'])
  })

  it('比较序：数字 < 文本（不区分大小写） < FALSE < TRUE', () => {
    const wb = wbWith({ A1: 'banana', A2: '5', A3: 'Apple', A4: 'TRUE' })
    expect(sortedRaws(wb, range, [{ col: 0, asc: true }], false, 0)).toEqual(['5', 'Apple', 'banana', 'TRUE'])
  })

  it('空格恒排末尾（与方向无关）', () => {
    const wb = wbWith({ A1: '2', A3: '1' }) // A2/A4 空
    expect(sortedRaws(wb, range, [{ col: 0, asc: true }], false, 0)).toEqual(['1', '2', '', ''])
    expect(sortedRaws(wb, range, [{ col: 0, asc: false }], false, 0)).toEqual(['2', '1', '', ''])
  })

  it('稳定排序：同键保持原相对序', () => {
    const wb = wbWith({ A1: '1', B1: 'x', A2: '1', B2: 'y', A3: '1', B3: 'z', A4: '0', B4: 'w' })
    expect(sortedRaws(wb, range, [{ col: 0, asc: true }], false, 1)).toEqual(['w', 'x', 'y', 'z'])
  })

  it('多关键字：第一键相同比第二键', () => {
    const wb = wbWith({ A1: '1', B1: '2', A2: '1', B2: '1', A3: '0', B3: '9', A4: '1', B4: '3' })
    const keys = [{ col: 0, asc: true }, { col: 1, asc: false }]
    expect(sortedRaws(wb, range, keys, false, 1)).toEqual(['9', '3', '2', '1'])
  })

  it('hasHeader：首行不动', () => {
    const wb = wbWith({ A1: 'name', A2: '3', A3: '1', A4: '2' })
    expect(sortedRaws(wb, range, [{ col: 0, asc: true }], true, 0)).toEqual(['name', '1', '2', '3'])
  })

  it('整行移动：第二列随行走；公式引用不改写（移动语义）', () => {
    const wb = wbWith({ A1: '2', B1: '=A1*10', A2: '1', B2: 'b', A3: '3', B3: 'c', A4: '4', B4: 'd' })
    const raws = sortedRaws(wb, range, [{ col: 0, asc: true }], false, 1)
    expect(raws).toEqual(['b', '=A1*10', 'c', 'd']) // 公式原文随行移动，引用不变
  })

  it('目标格无来源内容 → 清空（cell null）', () => {
    const wb = wbWith({ A1: '1', A2: '2', A3: '3', A4: '4', B2: 'x' })
    const data = wb.activeSheet
    const entries = computeSortEntries(data, wb.active, evaluatorFor(wb), range, [{ col: 0, asc: false }], false)
    let next = data
    for (const e of entries) next = next.setCell(e.row, e.col, e.cell)
    expect(next.getCell(1, 1)).toBeUndefined() // B2 的 'x' 被 B3 来源（空）覆盖 → 清空
    expect(next.getCell(2, 1)?.raw).toBe('x') // 'x' 随 A2 行移动到新位置
  })
})

describe('sortBlockedByMerges', () => {
  it('range 与 merge 相交 → true', () => {
    const wb = wbWith({})
    const data = wb.activeSheet.setMerges([{ sr: 1, sc: 0, er: 2, ec: 1 }])
    expect(sortBlockedByMerges(data, { sr: 0, sc: 0, er: 3, ec: 1 })).toBe(true)
    expect(sortBlockedByMerges(data, { sr: 5, sc: 0, er: 8, ec: 1 })).toBe(false)
  })
})
