// 应用外壳：Toolbar → FormulaBar → grid（EditorView mount 点）→ StatusBar。
// EditorView 在 useEffect 内一次性创建，unmount 时 destroy。
import { useEffect, useRef, useState } from 'react'
import { ContextMenu } from '../react/ContextMenu'
import { FilterDropdown } from '../react/FilterDropdown'
import { FindBar } from '../react/FindBar'
import { FormulaBar } from '../react/FormulaBar'
import { SheetTabBar } from '../react/SheetTabBar'
import { StatusBar } from '../react/StatusBar'
import { Toolbar } from '../react/Toolbar'
import { ConfirmHost } from '../react/ui/ConfirmHost'
import { EditorView } from '../view/editorview'
import { workbookStorage } from './storage'
import { createDemoState, createStateFromWorkbook } from './demo'

// 启动：有存档以存档建初始 state（恢复非用户操作，不可撤销，不走事务）
function initialState() {
  const wb = workbookStorage.load()
  return wb ? createStateFromWorkbook(wb) : createDemoState()
}

export function App() {
  const mountRef = useRef<HTMLDivElement>(null)
  const [view, setView] = useState<EditorView | null>(null)

  useEffect(() => {
    // React 侧经 useSheetState(subscribe) 感知 state 变化，无需宿主 dispatch 回调
    const v = new EditorView(mountRef.current!, { state: initialState() })
    setView(v)
    v.focus()
    if (import.meta.env.DEV) (window as unknown as { __xcell: EditorView }).__xcell = v
    // dispatch 后防抖自动保存；getter 延迟取值，只序列化防抖窗口末态
    const unsub = v.subscribe(() => workbookStorage.schedule(() => v.state.doc))
    const onUnload = () => workbookStorage.flush()
    window.addEventListener('beforeunload', onUnload)
    return () => {
      window.removeEventListener('beforeunload', onUnload)
      unsub()
      workbookStorage.flush()
      setView(null)
      v.destroy()
    }
  }, [])

  // 固定槽位顺序渲染，保证 view 出现/消失时 grid 容器 DOM 节点不被重建
  return (
    <div className="flex h-full flex-col">
      {view ? <Toolbar view={view} /> : null}
      {view ? <FormulaBar view={view} /> : null}
      {view ? <FindBar view={view} /> : null}
      <div className="relative min-h-0 flex-1" ref={mountRef} />
      {view ? <SheetTabBar view={view} /> : null}
      {view ? <StatusBar view={view} /> : null}
      {view ? <FilterDropdown view={view} /> : null}
      {view ? <ContextMenu view={view} /> : null}
      <ConfirmHost />
    </div>
  )
}
