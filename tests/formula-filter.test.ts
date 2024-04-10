import { describe, expect, it } from 'vitest'
import { fromA1 } from '../src/core/addr'
import { FilterState, Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'
import { filterHiddenRows } from '../src/formula/filter'

function wbWith(cells: Record<string, string>, filter?: FilterState): Workbook {
  let wb = Workbook.create({ rowCount: 20, colCount: 10 })
  let data = wb.activeSheet
  for (const [a1, raw] of Object.entries(cells)) {
    const { row, col } = fromA1(a1)!
    data = data.setCell(row, col, { raw })
  }
  if (filter) data = data.setFilter(filter)
  return wb.setSheet(wb.active, data)
}

const base = { A1: 'name', B1: 'score', A2: 'alice', B2: '90', A3: 'bob', B3: '60', A4: 'carol', B4: '75', A5: 'dave', B5: '60' }
const range = { sr: 0, sc: 0, er: 4, ec: 1 }

function hidden(wb: Workbook): number[] {
  return [...filterHiddenRows(wb.active, wb.activeSheet, evaluatorFor(wb))].sort((a, b) => a - b)
}

describe('filterHiddenRows', () => {
  it('无筛选 → 空集', () => {
    expect(hidden(wbWith(base))).toEqual([])
  })

  it('值勾选：excluded 的行被隐藏', () => {
    const wb = wbWith(base, { range, criteria: { 0: { type: 'values', excluded: ['bob', 'dave'] } } })
    expect(hidden(wb)).toEqual([2, 4])
  })

  it('值勾选：excluded 为空 → 不隐藏（新值默认可见）', () => {
    const wb = wbWith(base, { range, criteria: { 0: { type: 'values', excluded: [] } } })
    expect(hidden(wb)).toEqual([])
  })

  it('数值条件：gt / between', () => {
    const gt = wbWith(base, { range, criteria: { 1: { type: 'condition', field: 'num', op: 'gt', v1: '70' } } })
    expect(hidden(gt)).toEqual([2, 4]) // bob/dave 60 不 >70
    const between = wbWith(base, { range, criteria: { 1: { type: 'condition', field: 'num', op: 'between', v1: '60', v2: '75' } } })
    expect(hidden(between)).toEqual([1]) // alice 90 不在区间
  })

  it('文本条件：contains / startsWith（不区分大小写）', () => {
    const contains = wbWith(base, { range, criteria: { 0: { type: 'condition', field: 'text', op: 'contains', v1: 'A' } } })
    expect(hidden(contains)).toEqual([2]) // bob 不含 a
    const starts = wbWith(base, { range, criteria: { 0: { type: 'condition', field: 'text', op: 'startsWith', v1: 'c' } } })
    expect(hidden(starts)).toEqual([1, 2, 4])
  })

  it('多列 AND', () => {
    const wb = wbWith(base, {
      range,
      criteria: {
        0: { type: 'condition', field: 'text', op: 'contains', v1: 'a' },
        1: { type: 'condition', field: 'num', op: 'gte', v1: '76' },
      },
    })
    expect(hidden(wb)).toEqual([2, 3, 4]) // 仅 alice 同时满足
  })

  it('公式格按显示值/求值结果参与', () => {
    const wb = wbWith(
      { A1: 'v', B1: 'n', A2: '=1+1', B2: '=B3*2', A3: 'x', B3: '5' },
      { range: { sr: 0, sc: 0, er: 2, ec: 1 }, criteria: { 1: { type: 'condition', field: 'num', op: 'gt', v1: '5' } } },
    )
    expect(hidden(wb)).toEqual([2]) // B3=5 不 >5；B2=10 保留
  })

  it('表头行永不隐藏', () => {
    const wb = wbWith({ A1: 'zzz', A2: 'aaa' }, { range: { sr: 0, sc: 0, er: 1, ec: 0 }, criteria: { 0: { type: 'values', excluded: ['zzz', 'aaa'] } } })
    expect(hidden(wb)).toEqual([1]) // 行 0 是表头
  })
})
