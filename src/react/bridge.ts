// React bridge：把 EditorView 的 subscribe/state 快照接进 useSyncExternalStore。
import { useSyncExternalStore } from 'react'
import type { SheetState } from '../core/state'
import type { EditorView } from '../view/editorview'

export function useSheetState(view: EditorView): SheetState {
  return useSyncExternalStore(
    (cb) => view.subscribe(cb),
    () => view.state,
  )
}
