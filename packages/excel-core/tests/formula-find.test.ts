import { describe, expect, it } from 'vitest'
import { fromA1 } from '../src/core/addr'
import { Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'
import { findAll, FindMatch, FindQuery, indexAfterReplace, replaceInRaw } from '../src/formula/find'

function wbWith(cells: Record<string, string>, cells2?: Record<string, string>): Workbook {
  let wb = Workbook.create({ rowCount: 20, colCount: 10 })
  let data = wb.activeSheet
  for (const [a1, raw] of Object.entries(cells)) {
    const { row, col } = fromA1(a1)!
    data = data.setCell(row, col, { raw })
  }
  wb = wb.setSheet(wb.active, data)
  if (cells2) {
    wb = wb.addSheet('s2', Workbook.create({ rowCount: 20, colCount: 10 }).activeSheet, 1, 'Sheet2')
    let d2 = wb.sheet('s2')
    for (const [a1, raw] of Object.entries(cells2)) {
      const { row, col } = fromA1(a1)!
      d2 = d2.setCell(row, col, { raw })
    }
    wb = wb.setSheet('s2', d2)
  }
  return wb
}

const q = (text: string, extra?: Partial<FindQuery>): FindQuery => ({
  text,
  caseSensitive: false,
  wholeCell: false,
  workbook: false,
  ...extra,
})

describe('findAll', () => {
  it('匹配显示文本（含格式化显示）', () => {
    const wb = wbWith({ A1: 'hello world', A2: 'HELLO', A3: 'bye' })
    const ms = findAll(wb, evaluatorFor(wb), q('hello'))
    expect(ms).toEqual([
      { sheet: 's1', row: 0, col: 0 },
      { sheet: 's1', row: 1, col: 0 },
    ])
  })

  it('匹配公式原文；公式显示值也匹配（每格至多一条）', () => {
    const wb = wbWith({ A1: '5', B1: '=A1*2', C1: 'x' })
    expect(findAll(wb, evaluatorFor(wb), q('=A1'))).toEqual([{ sheet: 's1', row: 0, col: 1 }])
    expect(findAll(wb, evaluatorFor(wb), q('10'))).toEqual([{ sheet: 's1', row: 0, col: 1 }]) // 显示值 10
  })

  it('区分大小写与整格匹配', () => {
    const wb = wbWith({ A1: 'Hello', A2: 'hello', A3: 'hello world' })
    expect(findAll(wb, evaluatorFor(wb), q('Hello', { caseSensitive: true }))).toEqual([{ sheet: 's1', row: 0, col: 0 }])
    expect(findAll(wb, evaluatorFor(wb), q('hello', { wholeCell: true }))).toEqual([
      { sheet: 's1', row: 0, col: 0 },
      { sheet: 's1', row: 1, col: 0 },
    ])
  })

  it('全簿范围按表序行主序返回', () => {
    const wb = wbWith({ B2: 'x-foo' }, { A1: 'foo-y' })
    const ms = findAll(wb, evaluatorFor(wb), q('foo', { workbook: true }))
    expect(ms).toEqual([
      { sheet: 's1', row: 1, col: 1 },
      { sheet: 's2', row: 0, col: 0 },
    ])
  })

  it('当前表范围不含别表', () => {
    const wb = wbWith({}, { A1: 'foo' })
    expect(findAll(wb, evaluatorFor(wb), q('foo'))).toEqual([])
  })
})

describe('replaceInRaw', () => {
  it('不区分大小写全局替换', () => {
    expect(replaceInRaw('Foo foo FOO', q('foo'), 'bar')).toBe('bar bar bar')
  })

  it('区分大小写', () => {
    expect(replaceInRaw('Foo foo', q('Foo', { caseSensitive: true }), 'bar')).toBe('bar foo')
  })

  it('整格匹配替换为整体', () => {
    expect(replaceInRaw('hello', q('hello', { wholeCell: true }), 'x')).toBe('x')
    expect(replaceInRaw('hello world', q('hello', { wholeCell: true }), 'x')).toBeNull()
  })

  it('公式原文按文本片段替换（区域端点未归一由 eval 归一）', () => {
    expect(replaceInRaw('=SUM(A1:A3)', q('A1'), 'B1')).toBe('=SUM(B1:A3)')
  })

  it('无匹配 → null', () => {
    expect(replaceInRaw('abc', q('xyz'), 'r')).toBeNull()
  })
})

describe('indexAfterReplace（替换后前进）', () => {
  const m = (row: number, col = 0): FindMatch => ({ sheet: 's1', row, col })

  it('替换文本仍命中：游标前进到下一匹配（查 a 换 ab 不反复替换同格）', () => {
    const oldM = [m(0), m(1)]
    expect(indexAfterReplace(oldM, 0, [...oldM])).toBe(1)
  })

  it('替换后本格不再命中：指向原下一匹配（新序列已前移）', () => {
    const oldM = [m(0), m(1), m(2)]
    expect(indexAfterReplace(oldM, 0, [m(1), m(2)])).toBe(0)
  })

  it('末尾匹配循环回卷到首个', () => {
    const oldM = [m(0), m(1)]
    expect(indexAfterReplace(oldM, 1, [...oldM])).toBe(0)
  })

  it('新序列为空 → 0；oldIdx 越界钳制', () => {
    expect(indexAfterReplace([m(0)], 0, [])).toBe(0)
    expect(indexAfterReplace([m(0), m(1)], 99, [m(0), m(1)])).toBe(0)
  })

  it('doc 级回归：replaceInRaw 替换后 findAll 重算 + 前进到下一匹配', () => {
    let wb = wbWith({ A1: 'a', A2: 'a', A3: 'x' })
    const query = q('a')
    const oldM = findAll(wb, evaluatorFor(wb), query)
    expect(oldM).toEqual([
      { sheet: 's1', row: 0, col: 0 },
      { sheet: 's1', row: 1, col: 0 },
    ])
    // 替换 A1：a → ab（仍命中 'a'）
    const cell = wb.activeSheet.getCell(0, 0)!
    wb = wb.setSheet('s1', wb.activeSheet.setCell(0, 0, { ...cell, raw: replaceInRaw(cell.raw, query, 'ab')! }))
    const newM = findAll(wb, evaluatorFor(wb), query)
    expect(newM).toHaveLength(2)
    expect(indexAfterReplace(oldM, 0, newM)).toBe(1) // 指向 A2
  })
})
