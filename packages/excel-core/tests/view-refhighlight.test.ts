import { describe, expect, it } from 'vitest'
import { completionCandidates, extractCurrentSheetRanges } from '../src/formula/rangeRefs'

describe('extractCurrentSheetRanges', () => {
  it('抽取 range 与 ref 节点 → CellRange[]', () => {
    expect(extractCurrentSheetRanges('=SUM(A1:B2)+C3')).toEqual([
      { sr: 0, sc: 0, er: 1, ec: 1 }, { sr: 2, sc: 2, er: 2, ec: 2 },
    ])
  })
  it('跨表引用 Sheet2!A1 不返回（后续项）', () => {
    expect(extractCurrentSheetRanges('=Sheet2!A1+B1')).toEqual([{ sr: 0, sc: 1, er: 0, ec: 1 }])
  })
  it('非公式（无 = 前缀）→ 空', () => {
    expect(extractCurrentSheetRanges('hello')).toEqual([])
  })
  it('语法错误 → 空（不抛）', () => {
    expect(extractCurrentSheetRanges('=((')).toEqual([])
  })
})

// 纯函数补全候选：= 后末尾标识符 token 前缀匹配函数名表（≤8）。DOM/react 下拉接线由手动/集成覆盖。
describe('completionCandidates', () => {
  const names = ['SUM', 'SUMIF', 'SUMIFS', 'VLOOKUP', 'ABS', 'AVERAGE', 'AND', 'OR', 'NOT', 'MAX']
  it('前缀匹配并截断到 8 个', () => {
    expect(completionCandidates('=SU', names)).toEqual(['SUM', 'SUMIF', 'SUMIFS'])
  })
  it('超过 8 个时只取前 8', () => {
    // A 前缀：ABS,AVERAGE,AND（3 个）；用更长表构造 >8 场景
    const big = Array.from({ length: 12 }, (_, i) => `AB${i}`) // AB0..AB11
    expect(completionCandidates('=AB', big)).toHaveLength(8)
  })
  it('大小写不敏感（函数名大写，token 转大写匹配）', () => {
    expect(completionCandidates('=sum', names)).toEqual(['SUM', 'SUMIF', 'SUMIFS'])
  })
  it('无末尾标识符 token → 空', () => {
    expect(completionCandidates('=SUM(', names)).toEqual([])
  })
  it('非公式（无 =）→ 空', () => {
    expect(completionCandidates('hello', names)).toEqual([])
  })
})
