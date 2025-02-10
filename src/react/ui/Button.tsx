import type { MouseEvent, ReactNode } from 'react'

interface Props {
  variant?: 'primary' | 'outline' | 'ghost' | 'danger'
  size?: 'sm' | 'md'
  disabled?: boolean
  onClick?: (e: MouseEvent<HTMLButtonElement>) => void
  children: ReactNode
}

const VARIANT: Record<string, string> = {
  primary: 'bg-primary text-surface hover:bg-primary-hover',
  danger: 'bg-danger text-surface hover:bg-danger-deep',
  outline: 'border border-line-strong text-ink hover:bg-hover',
  ghost: 'text-ink hover:bg-hover',
}

export function Button({ variant = 'outline', size = 'md', disabled, onClick, children }: Props) {
  const cls = [
    'flex items-center justify-center rounded-md px-3 text-sm transition-colors duration-120',
    size === 'sm' ? 'h-7' : 'h-8',
    disabled ? 'cursor-default text-ink-disabled' : VARIANT[variant],
    disabled && (variant === 'primary' || variant === 'danger') ? 'bg-line' : '',
  ].join(' ')
  return (
    <button type="button" className={cls} disabled={disabled} onClick={onClick}>
      {children}
    </button>
  )
}
