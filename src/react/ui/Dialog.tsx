import { useEffect, useRef, type ReactNode } from 'react'

interface Props {
  title: string
  onClose: () => void
  footer?: ReactNode
  width?: number
  children: ReactNode
}

export function Dialog({ title, onClose, footer, width = 360, children }: Props) {
  const panel = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') onClose()
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [onClose])

  useEffect(() => {
    // 面板兜底聚焦（接管方向键，避免穿透到底层画布）；
    // 内部控件已被 autoFocus 抢占时不夺焦
    const el = panel.current
    if (el && !el.contains(document.activeElement)) el.focus()
  }, [])

  return (
    <div className="fixed inset-0 z-100 flex items-center justify-center bg-black/15" role="dialog" aria-label={title}>
      <div ref={panel} tabIndex={-1} className="rounded-lg border border-line-strong bg-surface p-4 shadow-3 outline-none" style={{ width }}>
        <div className="mb-3 text-md font-semibold">{title}</div>
        {children}
        {footer && <div className="mt-3 flex justify-end gap-2">{footer}</div>}
      </div>
    </div>
  )
}
