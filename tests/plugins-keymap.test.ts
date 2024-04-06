import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { navigateFocus } from '../src/plugins/keymap'

// 10×10 表，合并区 B2:C3（sr:1,sc:1,er:2,ec:2）
const sheetWithMerge = () => {
  const wb = Workbook.create({ rowCount: 10, colCount: 10 })
  return wb.activeSheet.setMerges([{ sr: 1, sc: 1, er: 2, ec: 2 }])
}

describe('navigateFocus', () => {
  it('focus 在合并区锚点：ArrowRight/ArrowDown 跳过整区到远侧之外', () => {
    const sheet = sheetWithMerge()
    expect(navigateFocus(sheet, { row: 1, col: 1 }, 0, 1)).toEqual({ row: 1, col: 3 })
    expect(navigateFocus(sheet, { row: 1, col: 1 }, 1, 0)).toEqual({ row: 3, col: 1 })
  })
  it('focus 在合并区外：移入合并区 → 落在锚点', () => {
    const sheet = sheetWithMerge()
    expect(navigateFocus(sheet, { row: 1, col: 0 }, 0, 1)).toEqual({ row: 1, col: 1 })
    expect(navigateFocus(sheet, { row: 0, col: 2 }, 1, 0)).toEqual({ row: 1, col: 1 })
  })
  it('focus 在锚点：向左/向上移出合并区 → 普通相邻格', () => {
    const sheet = sheetWithMerge()
    expect(navigateFocus(sheet, { row: 1, col: 1 }, 0, -1)).toEqual({ row: 1, col: 0 })
    expect(navigateFocus(sheet, { row: 1, col: 1 }, -1, 0)).toEqual({ row: 0, col: 1 })
  })
  it('无合并区：普通 clamp 移动（含边界钳制）', () => {
    const wb = Workbook.create({ rowCount: 10, colCount: 10 })
    const sheet = wb.activeSheet
    expect(navigateFocus(sheet, { row: 4, col: 4 }, 1, 1)).toEqual({ row: 5, col: 5 })
    expect(navigateFocus(sheet, { row: 0, col: 0 }, -1, -1)).toEqual({ row: 0, col: 0 })
    expect(navigateFocus(sheet, { row: 9, col: 9 }, 1, 1)).toEqual({ row: 9, col: 9 })
  })
})

describe('navigateFocus 隐藏行列', () => {
  const hiddenSheet = (hiddenRows: number[], hiddenCols: number[]) => ({
    mergeAt: () => null,
    rowCount: 10,
    colCount: 10,
    hiddenRows,
    hiddenCols,
  })
  const isHiddenOf = (s: { hiddenRows: number[]; hiddenCols: number[] }) => (r: number, c: number) =>
    s.hiddenRows.includes(r) || s.hiddenCols.includes(c)

  it('下移跳过隐藏行', () => {
    const s = hiddenSheet([2], [])
    expect(navigateFocus(s, { row: 1, col: 0 }, 1, 0, isHiddenOf(s))).toEqual({ row: 3, col: 0 })
  })

  it('上移跳过连续隐藏行', () => {
    const s = hiddenSheet([2, 3], [])
    expect(navigateFocus(s, { row: 4, col: 0 }, -1, 0, isHiddenOf(s))).toEqual({ row: 1, col: 0 })
  })

  it('右移跳过隐藏列；边界处停止', () => {
    const s = hiddenSheet([], [1, 2])
    expect(navigateFocus(s, { row: 0, col: 0 }, 0, 1, isHiddenOf(s))).toEqual({ row: 0, col: 3 })
    const edge = hiddenSheet([], [9])
    expect(navigateFocus(edge, { row: 0, col: 8 }, 0, 1, isHiddenOf(edge))).toEqual({ row: 0, col: 9 })
  })

  it('默认谓词（不传）行为不变', () => {
    const s = hiddenSheet([2], [])
    expect(navigateFocus(s, { row: 1, col: 0 }, 1, 0)).toEqual({ row: 2, col: 0 })
  })
})
