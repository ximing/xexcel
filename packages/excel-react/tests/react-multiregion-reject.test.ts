// T6：多区域选区下 sort/filter/condformat 必须（1）按钮禁用、（2）触发拒绝并提示单区域要求。
// 沿用既有 react 测试的纯函数模式（menuItems/toggleCFStyle 均为提取的纯判定，无 DOM 渲染）。
// 此处对按钮禁用判定（canSort/canFilter/canCondFormat）与拒绝消息解析（sort/filter/condFormat
// Rejection）做断言：多区域 → 禁用 + 返回拒绝消息；单区域 → 启用 + 放行（null）。零回归。
import { describe, expect, it } from 'vitest'
import { Workbook } from '@xexcel/core'
import { SheetState } from '@xexcel/core'
import { appendRange, rangeSelection } from '@xexcel/core'
import type { Selection } from '@xexcel/core'
import { canSort, canFilter, sortRejection, filterRejection } from '../src/Toolbar'
import { canCondFormat, condFormatRejection } from '../src/CondFormatDialog'

const mk = (sel: Selection): SheetState =>
  SheetState.create({ doc: Workbook.create({ rowCount: 10, colCount: 10 }), selection: sel })

// 单区域：两行一列（满足排序“多行”要求 er>sr）
const single = rangeSelection({ sr: 0, sc: 0, er: 1, ec: 0 })
// 多区域：Ctrl 追加一列选区 → ranges.length === 2
const multi = appendRange(single, { sr: 0, sc: 2, er: 1, ec: 2 })

describe('多区域 sort/filter/condformat 禁用 + 拒绝（单区域零回归）', () => {
  it('canSort：单区域可用，多区域禁用', () => {
    expect(canSort(mk(single))).toBe(true)
    expect(canSort(mk(multi))).toBe(false)
  })

  it('canFilter：单区域可用，多区域禁用', () => {
    expect(canFilter(mk(single))).toBe(true)
    expect(canFilter(mk(multi))).toBe(false)
  })

  it('canCondFormat：单区域可用，多区域禁用', () => {
    expect(canCondFormat(mk(single))).toBe(true)
    expect(canCondFormat(mk(multi))).toBe(false)
  })

  it('多区域触发排序/筛选/条件格式 → 返回拒绝消息；单区域 → null（放行）', () => {
    expect(sortRejection(mk(single))).toBe(null)
    expect(sortRejection(mk(multi))).toBe('排序仅支持单区域选择')
    expect(filterRejection(mk(single))).toBe(null)
    expect(filterRejection(mk(multi))).toBe('筛选仅支持单区域选择')
    expect(condFormatRejection(mk(single))).toBe(null)
    expect(condFormatRejection(mk(multi))).toBe('条件格式仅支持单区域选择')
  })

  it('单行选区排序仍禁用（er === sr，零回归）', () => {
    const oneRow = rangeSelection({ sr: 0, sc: 0, er: 0, ec: 0 })
    expect(canSort(mk(oneRow))).toBe(false)
  })

  it('多区域 ranges.length 校验：single=1，multi=2', () => {
    expect(mk(single).selection.ranges.length).toBe(1)
    expect(mk(multi).selection.ranges.length).toBe(2)
  })
})
