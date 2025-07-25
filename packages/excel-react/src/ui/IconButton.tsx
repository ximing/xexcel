import type { LucideIcon } from 'lucide-react'
import type { MouseEvent } from 'react'
import { Icon } from './Icon'
import { Tooltip } from './Tooltip'

interface Props {
  icon: LucideIcon
  tip: string
  kbd?: string
  active?: boolean
  disabled?: boolean
  onClick: (e: MouseEvent<HTMLButtonElement>) => void
}

export function IconButton({ icon, tip, kbd, active, disabled, onClick }: Props) {
  const cls = [
    'flex h-7 w-7 flex-none items-center justify-center rounded-md transition-colors duration-120',
    disabled ? 'cursor-default text-ink-disabled' : active ? 'bg-primary-soft text-primary' : 'text-ink hover:bg-hover',
  ].join(' ')
  return (
    <Tooltip tip={tip} kbd={kbd}>
      <button type="button" aria-label={tip} className={cls} disabled={disabled} onClick={onClick}>
        <Icon icon={icon} />
      </button>
    </Tooltip>
  )
}
