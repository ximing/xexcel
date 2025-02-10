import { useEffect, type ReactNode } from 'react'

interface Props {
  title: string
  onClose: () => void
  footer?: ReactNode
  width?: number
  children: ReactNode
}

export function Dialog({ title, onClose, footer, width = 360, children }: Props) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') onClose()
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="fixed inset-0 z-100 flex items-center justify-center bg-black/15" role="dialog" aria-label={title}>
      <div className="rounded-lg border border-line-strong bg-surface p-4 shadow-3" style={{ width }}>
        <div className="mb-3 text-md font-semibold">{title}</div>
        {children}
        {footer && <div className="mt-3 flex justify-end gap-2">{footer}</div>}
      </div>
    </div>
  )
}
