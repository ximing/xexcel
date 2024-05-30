import { describe, expect, it } from 'vitest'
import { wholeRange } from '../src/core/addr'
import {
  activeRange, appendRange, extendActiveRange, forEachSelectionRange,
  rangeSelection, selectionRange, singleCell, toggleRange, Selection,
} from '../src/core/selection'

// 不变式：activeCell 必落在 ranges[last]
function invariantHolds(sel: Selection): boolean {
  const r = sel.ranges[sel.ranges.length - 1]
  return sel.activeCell.row >= r.sr && sel.activeCell.row <= r.er &&
    sel.activeCell.col >= r.sc && sel.activeCell.col <= r.ec
}

describe('Selection 模型', () => {
  it('singleCell：单区域且 activeCell 落其中', () => {
    const sel = singleCell(3, 4)
    expect(sel.ranges).toEqual([{ sr: 3, sc: 4, er: 3, ec: 4 }])
    expect(sel.activeCell).toEqual({ row: 3, col: 4 })
    expect(invariantHolds(sel)).toBe(true)
  })
  it('selectionRange 返回活动区域 = ranges[last]', () => {
    const sel = appendRange(singleCell(0, 0), { sr: 5, sc: 5, er: 7, ec: 9 })
    expect(selectionRange(sel)).toEqual({ sr: 5, sc: 5, er: 7, ec: 9 })
    expect(activeRange(sel)).toEqual(selectionRange(sel))
  })
  it('appendRange：新造数组、末项为新增区域', () => {
    const a = singleCell(0, 0)
    const b = appendRange(a, { sr: 1, sc: 1, er: 2, ec: 2 })
    expect(b.ranges.length).toBe(2)
    expect(a.ranges.length).toBe(1) // 原对象未被改
    expect(invariantHolds(b)).toBe(true)
  })
  it('toggleRange：移除最后加入的含该格 range（LIFO）', () => {
    // 两重叠 range 都含 (0,0)：ranges[0]={0,0..1,1}, ranges[1]={0,0..2,2}
    const sel = appendRange(appendRange(singleCell(0, 0), { sr: 0, sc: 0, er: 1, ec: 1 }), { sr: 0, sc: 0, er: 2, ec: 2 })
    const t = toggleRange(sel, 0, 0) // 移除最后加入者 = ranges[2]
    expect(t.ranges.length).toBe(2)
    expect(invariantHolds(t)).toBe(true)
  })
  it('toggleRange：移除最后一个 range → 回到该格单选', () => {
    const sel = singleCell(0, 0)
    const t = toggleRange(sel, 0, 0)
    expect(t.ranges.length).toBe(1)
    expect(invariantHolds(t)).toBe(true)
  })
  it('toggleRange：格不在任何 range → 不变', () => {
    const sel = singleCell(0, 0)
    expect(toggleRange(sel, 9, 9)).toBe(sel)
  })
  it('extendActiveRange：替换末项边界，activeCell=新焦点，禁就地改', () => {
    const sel = singleCell(0, 0)
    const ext = extendActiveRange(sel, { row: 3, col: 4 })
    expect(ext.ranges).toEqual([{ sr: 0, sc: 0, er: 3, ec: 4 }])
    expect(ext.activeCell).toEqual({ row: 3, col: 4 })
    expect(sel.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 0 }]) // 原未改
  })
  // 连续两步 Shift 生长（审查员 repro：旧实现滑动得 {5,6..5,7}，丢起始格）
  it('extendActiveRange：连续两步 Shift 生长不滑动（锚点固定为起始格）', () => {
    let sel = singleCell(5, 5)
    sel = extendActiveRange(sel, { row: 5, col: 6 })
    expect(sel.ranges).toEqual([{ sr: 5, sc: 5, er: 5, ec: 6 }])
    expect(sel.activeCell).toEqual({ row: 5, col: 6 })
    sel = extendActiveRange(sel, { row: 5, col: 7 })
    expect(sel.ranges).toEqual([{ sr: 5, sc: 5, er: 5, ec: 7 }]) // 生长，不滑动
    expect(sel.activeCell).toEqual({ row: 5, col: 7 })
  })
  // Shift+Left 回缩：锚点固定在起始格，右缘收回
  it('extendActiveRange：Shift+Left 回缩右缘（锚点固定）', () => {
    let sel = singleCell(5, 5)
    sel = extendActiveRange(sel, { row: 5, col: 7 })
    sel = extendActiveRange(sel, { row: 5, col: 6 })
    expect(sel.ranges).toEqual([{ sr: 5, sc: 5, er: 5, ec: 6 }])
    expect(sel.activeCell).toEqual({ row: 5, col: 6 })
  })
  // 2D 连续生长：先右扩再下扩，锚点始终固定在起始格
  it('extendActiveRange：2D 连续生长（先右后下）固定锚点', () => {
    let sel = singleCell(5, 5)
    sel = extendActiveRange(sel, { row: 5, col: 7 }) // 右扩
    sel = extendActiveRange(sel, { row: 8, col: 7 }) // 下扩
    expect(sel.ranges).toEqual([{ sr: 5, sc: 5, er: 8, ec: 7 }])
    expect(sel.activeCell).toEqual({ row: 8, col: 7 })
  })
  it('rangeSelection：normalize + activeCell 默认左上', () => {
    const sel = rangeSelection({ sr: 2, sc: 5, er: 0, ec: 3 })
    expect(sel.ranges).toEqual([{ sr: 0, sc: 3, er: 2, ec: 5 }])
    expect(sel.activeCell).toEqual({ row: 0, col: 3 })
  })
  it('forEachSelectionRange：遍历全部区域', () => {
    const sel = appendRange(appendRange(singleCell(0, 0), { sr: 1, sc: 1, er: 1, ec: 1 }), { sr: 2, sc: 2, er: 2, ec: 2 })
    const got: number[] = []
    forEachSelectionRange(sel, r => got.push(r.sr))
    expect(got).toEqual([0, 1, 2])
  })
})

describe('Selection 不可变性（守护 undo）', () => {
  it('appendRange 不就地改原 ranges 数组', () => {
    const a = singleCell(0, 0)
    const frozen = a.ranges
    appendRange(a, { sr: 1, sc: 1, er: 1, ec: 1 })
    expect(a.ranges).toBe(frozen) // 同一引用，未被 push
    expect(a.ranges.length).toBe(1)
  })
  it('extendActiveRange 不就地改末项 .er/.ec', () => {
    const sel = appendRange(singleCell(0, 0), { sr: 1, sc: 1, er: 1, ec: 1 })
    const last = sel.ranges[sel.ranges.length - 1]
    extendActiveRange(sel, { row: 5, col: 5 })
    expect(last).toEqual({ sr: 1, sc: 1, er: 1, ec: 1 }) // 末项对象未被就地改
  })
  it('wholeRange 用于 selectAll（活动区=全表，activeCell={0,0}）', () => {
    const sel = rangeSelection(wholeRange(10, 10), { row: 0, col: 0 })
    expect(selectionRange(sel)).toEqual({ sr: 0, sc: 0, er: 9, ec: 9 })
    expect(invariantHolds(sel)).toBe(true)
  })
})
