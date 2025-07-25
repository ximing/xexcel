import type { LucideIcon } from 'lucide-react'
import { useEffect } from 'react'
import { Icon } from './Icon'

export interface MenuItemDef {
  id: string
  label: string
  icon?: LucideIcon
  kbd?: string
  danger?: boolean
  disabled?: boolean
  active?: boolean
  onSelect: () => void
}
export type MenuEntry = MenuItemDef | { sep: true }

export function Menu({ entries, onClose }: { entries: MenuEntry[]; onClose: () => void }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') onClose()
    }
    document.addEventListener('keydown', onKey)
    return () => document.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="flex min-w-40 flex-col rounded-md border border-line-strong bg-surface py-1 shadow-2" role="menu">
      {entries.map((it, i) =>
        'sep' in it ? (
          <div key={i} className="my-1 h-px bg-line" />
        ) : (
          <button
            key={it.id}
            type="button"
            role="menuitem"
            disabled={it.disabled}
            className={[
              'flex h-7 items-center gap-2 px-3 text-left text-xs',
              it.disabled ? 'cursor-default text-ink-disabled' : it.danger ? 'text-danger-deep hover:bg-primary-soft' : 'text-ink hover:bg-primary-soft',
              it.active ? 'text-primary' : '',
            ].join(' ')}
            onClick={() => {
              onClose()
              it.onSelect()
            }}
          >
            {it.icon && <Icon icon={it.icon} />}
            <span className="flex-1">{it.label}</span>
            {it.kbd && <span className="text-ink-3">{it.kbd}</span>}
          </button>
        ),
      )}
    </div>
  )
}
