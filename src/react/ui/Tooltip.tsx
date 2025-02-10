// 自绘提示：hover 500ms 出现，离开即消；深底白字；auto 方位（下方不够则朝上）
import { useRef, useState, type ReactNode } from 'react'
import { resolvePlacement, type TipPlacement } from './tooltipPlacement'

export const TOOLTIP_DELAY_MS = 500

interface Props {
  tip: string
  kbd?: string
  placement?: 'auto' | TipPlacement
  children: ReactNode
}

export function Tooltip({ tip, kbd, placement = 'auto', children }: Props) {
  const [show, setShow] = useState(false)
  const [pos, setPos] = useState<{ x: number; y: number; p: TipPlacement } | null>(null)
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const anchor = useRef<HTMLSpanElement>(null)

  const open = (): void => {
    timer.current = setTimeout(() => {
      const r = anchor.current?.getBoundingClientRect()
      if (!r) return
      const p = resolvePlacement(r.bottom, r.top, 28, placement, window.innerHeight)
      setPos({ x: r.left + r.width / 2, y: p === 'bottom' ? r.bottom + 6 : r.top - 6, p })
      setShow(true)
    }, TOOLTIP_DELAY_MS)
  }
  const close = (): void => {
    if (timer.current) clearTimeout(timer.current)
    timer.current = null
    setShow(false)
  }

  return (
    <span ref={anchor} className="inline-flex" onMouseEnter={open} onMouseLeave={close}>
      {children}
      {show && pos && (
        <span
          role="tooltip"
          className="fixed z-120 -translate-x-1/2 whitespace-nowrap rounded-sm bg-ink/90 px-2 py-1 text-xs text-surface shadow-1"
          style={{ left: pos.x, top: pos.p === 'bottom' ? pos.y : undefined, bottom: pos.p === 'top' ? window.innerHeight - pos.y : undefined }}
        >
          {tip}
          {kbd && <span className="ml-1.5 text-surface/60">{kbd}</span>}
        </span>
      )}
    </span>
  )
}
