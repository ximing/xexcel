import { describe, it, expect } from 'vitest'
import { Workbook } from '@xexcel/core'
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
  it('锚点隐藏：吸附后再判隐藏，跳过整个合并区到远侧之外（不落回隐藏锚点）', () => {
    // 隐藏锚点列 1：从 (1,0) 右移，入区吸附锚点 (1,1) 仍隐藏 → 越过合并区到 (1,3)
    const byCol = sheetWithMerge().setHidden('col', [1], true)
    const hiddenCol = (r: number, c: number) => byCol.hiddenRows.includes(r) || byCol.hiddenCols.includes(c)
    expect(navigateFocus(byCol, { row: 1, col: 0 }, 0, 1, hiddenCol)).toEqual({ row: 1, col: 3 })
    // 隐藏锚点行 1：从 (0,1) 下移 → 越过合并区到 (3,1)
    const byRow = sheetWithMerge().setHidden('row', [1], true)
    const hiddenRow = (r: number, c: number) => byRow.hiddenRows.includes(r) || byRow.hiddenCols.includes(c)
    expect(navigateFocus(byRow, { row: 0, col: 1 }, 1, 0, hiddenRow)).toEqual({ row: 3, col: 1 })
  })
  it('锚点隐藏且远侧格也隐藏：继续步进到下一可见格', () => {
    const sheet = sheetWithMerge().setHidden('col', [1, 3], true)
    const isHidden = (r: number, c: number) => sheet.hiddenRows.includes(r) || sheet.hiddenCols.includes(c)
    expect(navigateFocus(sheet, { row: 1, col: 0 }, 0, 1, isHidden)).toEqual({ row: 1, col: 4 })
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
