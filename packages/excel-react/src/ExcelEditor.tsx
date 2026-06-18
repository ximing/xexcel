// 通用编辑器壳：Toolbar → FormulaBar → FindBar → grid（EditorView mount 点）→ SheetTabBar → StatusBar。
// EditorView 在 useEffect 内一次性创建，unmount 时 destroy；state 仅挂载时取一次（同原 App 语义）。
import { useEffect, useRef, useState } from 'react'
import type { SheetState } from '@xexcel/core'
import { EditorView } from '@xexcel/view'
import { ContextMenu } from './ContextMenu'
import { FilterDropdown } from './FilterDropdown'
import { FindBar } from './FindBar'
import { FormulaBar } from './FormulaBar'
import { SheetTabBar } from './SheetTabBar'
import { StatusBar } from './StatusBar'
import { Toolbar } from './Toolbar'
import { ConfirmHost } from './ui/ConfirmHost'

export interface ExcelEditorProps {
  state: SheetState
  // 宿主拿 view 句柄接线自动保存/调试暴露；unmount 时回调 null
  onView?: (view: EditorView | null) => void
}

export function ExcelEditor({ state, onView }: ExcelEditorProps) {
  const mountRef = useRef<HTMLDivElement>(null)
  const [view, setView] = useState<EditorView | null>(null)

  useEffect(() => {
    // React 侧经 useSheetState(subscribe) 感知 state 变化，无需宿主 dispatch 回调
    const v = new EditorView(mountRef.current!, { state })
    setView(v)
    v.focus()
    onView?.(v)
    return () => {
      onView?.(null)
      setView(null)
      v.destroy()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps -- state 仅挂载时取一次
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
