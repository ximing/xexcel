// 应用外壳：Toolbar → FormulaBar → grid（EditorView mount 点）→ StatusBar。
// EditorView 在 useEffect 内一次性创建，unmount 时 destroy。
import { useEffect, useRef, useState } from 'react'
import { FormulaBar } from '../react/FormulaBar'
import { StatusBar } from '../react/StatusBar'
import { Toolbar } from '../react/Toolbar'
import { EditorView } from '../view/editorview'
import { createDemoState } from './demo'

export function App() {
  const mountRef = useRef<HTMLDivElement>(null)
  const [view, setView] = useState<EditorView | null>(null)

  useEffect(() => {
    // React 侧经 useSheetState(subscribe) 感知 state 变化，无需宿主 dispatch 回调
    const v = new EditorView(mountRef.current!, { state: createDemoState() })
    setView(v)
    v.focus()
    return () => {
      setView(null)
      v.destroy()
    }
  }, [])

  // 固定槽位顺序渲染，保证 view 出现/消失时 grid 容器 DOM 节点不被重建
  return (
    <div className="app-shell">
      {view ? <Toolbar view={view} /> : null}
      {view ? <FormulaBar view={view} /> : null}
      <div className="grid-container" ref={mountRef} />
      {view ? <StatusBar view={view} /> : null}
    </div>
  )
}
