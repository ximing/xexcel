// 触发器 + 下拉（Menu 或自绘面板）：开合状态、外点关闭、左/右缘对齐
import { useRef, useState, type ReactNode } from 'react'
import { Menu, type MenuEntry } from './Menu'

interface Props {
  trigger: (open: boolean, toggle: () => void) => ReactNode
  entries?: MenuEntry[]
  children?: ReactNode
  align?: 'left' | 'right'
  // 'up'：面板向触发器上方弹出（如状态栏缩放菜单）
  side?: 'down' | 'up'
}

export function Dropdown({ trigger, entries, children, align = 'left', side = 'down' }: Props) {
  const [open, setOpen] = useState(false)
  const wrap = useRef<HTMLSpanElement>(null)
  const close = (): void => setOpen(false)

  return (
    <span ref={wrap} className="relative inline-flex">
      {trigger(open, () => setOpen(!open))}
      {open && (
        <>
          <span className="fixed inset-0 z-40" onClick={close} />
          <span
            className={`absolute z-40 ${side === 'up' ? 'bottom-full mb-0.5' : 'top-full mt-0.5'} ${align === 'right' ? 'right-0' : 'left-0'}`}
          >
            {entries ? <Menu entries={entries} onClose={close} /> : children}
          </span>
        </>
      )}
    </span>
  )
}
