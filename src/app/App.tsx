// demo 应用壳：存档恢复/自动保存/__xcell 调试暴露；通用编辑器壳在 @gmi/excel-react。
import { useEffect, useState } from 'react'
import type { EditorView } from '@gmi/excel-view'
import { ExcelEditor, createStateFromWorkbook, workbookStorage } from '@gmi/excel-react'
import type { SheetState } from '@gmi/excel-core'
import { createDemoState } from './demo'

// 启动：有存档以存档建初始 state（恢复非用户操作，不可撤销，不走事务）
function initialState() {
  const wb = workbookStorage.load()
  return wb ? createStateFromWorkbook(wb) : createDemoState()
}

export function App() {
  const [state] = useState<SheetState>(initialState)
  const [view, setView] = useState<EditorView | null>(null)

  useEffect(() => {
    if (!view) return
    if (import.meta.env.DEV) (window as unknown as { __xcell: EditorView }).__xcell = view
    // dispatch 后防抖自动保存；getter 延迟取值，只序列化防抖窗口末态
    const unsub = view.subscribe(() => workbookStorage.schedule(() => view.state.doc))
    const onUnload = () => workbookStorage.flush()
    window.addEventListener('beforeunload', onUnload)
    return () => {
      window.removeEventListener('beforeunload', onUnload)
      unsub()
      workbookStorage.flush()
    }
  }, [view])

  return <ExcelEditor state={state} onView={setView} />
}
