// React bridge：把 EditorView 的 subscribe/state 快照接进 useSyncExternalStore。
import { useSyncExternalStore } from 'react'
import type { SheetState } from '@xexcel/core'
import type { EditorView } from '@xexcel/view'

export function useSheetState(view: EditorView): SheetState {
  return useSyncExternalStore(
    (cb) => view.subscribe(cb),
    () => view.state,
  )
}
