import { Plugin, PluginKey } from './plugin'
import { Selection } from './selection'
import { Step } from './steps'
import type { SheetState } from './state'
import type { Transaction } from './transaction'

export interface HistoryGroup {
  steps: Step[] // 反向 step（应用即撤销产生本组的事务）
  selection: Selection // 事务前选区
}

interface HistoryFieldValue {
  done: HistoryGroup[]
  undone: HistoryGroup[]
}

const MAX_GROUPS = 200

export const historyKey = new PluginKey('history')

// 对 tr.steps 逆序求逆：beforeDoc 首步用 tr.before.doc，其余用 tr.docs[i-1]
function invertTrSteps(tr: Transaction): Step[] {
  const inverse: Step[] = []
  for (let i = tr.steps.length - 1; i >= 0; i--) {
    const before = i === 0 ? tr.before.doc : tr.docs[i - 1]
    inverse.push(tr.steps[i].invert(before))
  }
  return inverse
}

export function history(): Plugin {
  return new Plugin({
    key: historyKey,
    state: {
      init: (): HistoryFieldValue => ({ done: [], undone: [] }),
      apply: (tr: Transaction, value: HistoryFieldValue): HistoryFieldValue => {
        // undo/redo 内部事务：直接按 meta 恢复双栈
        const restore = tr.getMeta('restoreHistory') as HistoryFieldValue | undefined
        if (restore) return restore
        if (tr.getMeta('addToHistory') === false || tr.steps.length === 0) return value
        const group: HistoryGroup = { steps: invertTrSteps(tr), selection: tr.before.selection }
        return { done: [...value.done, group].slice(-MAX_GROUPS), undone: [] }
      },
    },
  })
}

function fieldValue(state: SheetState): HistoryFieldValue | undefined {
  return historyKey.getState(state) as HistoryFieldValue | undefined
}

// 弹出一组并构造反向事务；mirror=弹出组的镜像（再 invert 回来）压入对侧栈
function applyGroup(
  state: SheetState,
  group: HistoryGroup,
  done: HistoryGroup[],
  undone: HistoryGroup[],
  dispatch?: (tr: Transaction) => void,
): boolean {
  const tr = state.tr
  for (const step of group.steps) tr._pushStep(step)
  const mirror: HistoryGroup = { steps: invertTrSteps(tr), selection: tr.before.selection }
  tr.setMeta('addToHistory', false)
  tr.setMeta('restoreHistory', { done, undone: [...undone, mirror] })
  tr.setSelection(group.selection)
  if (dispatch) dispatch(tr)
  return true
}

export function undo(state: SheetState, dispatch?: (tr: Transaction) => void): boolean {
  const h = fieldValue(state)
  if (!h || h.done.length === 0) return false
  const group = h.done[h.done.length - 1]
  return applyGroup(state, group, h.done.slice(0, -1), h.undone, dispatch)
}

export function redo(state: SheetState, dispatch?: (tr: Transaction) => void): boolean {
  const h = fieldValue(state)
  if (!h || h.undone.length === 0) return false
  const group = h.undone[h.undone.length - 1]
  // 对称：done 侧追加 mirror
  const tr = state.tr
  for (const step of group.steps) tr._pushStep(step)
  const mirror: HistoryGroup = { steps: invertTrSteps(tr), selection: tr.before.selection }
  tr.setMeta('addToHistory', false)
  tr.setMeta('restoreHistory', { done: [...h.done, mirror], undone: h.undone.slice(0, -1) })
  tr.setSelection(group.selection)
  if (dispatch) dispatch(tr)
  return true
}

export function undoDepth(state: SheetState): number {
  return fieldValue(state)?.done.length ?? 0
}

export function redoDepth(state: SheetState): number {
  return fieldValue(state)?.undone.length ?? 0
}
