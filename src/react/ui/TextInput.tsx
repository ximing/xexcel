import type { KeyboardEvent } from 'react'

interface Props {
  value: string
  onChange: (v: string) => void
  placeholder?: string
  invalid?: boolean
  width?: number
  autoFocus?: boolean
  title?: string
  onEnter?: () => void
  onBlur?: () => void
  // 需要 Shift+Enter / Escape 等组合键时的完整事件透传（先于 onEnter 触发）
  onKeyDown?: (e: KeyboardEvent<HTMLInputElement>) => void
}

export function TextInput({ value, onChange, placeholder, invalid, width, autoFocus, title, onEnter, onBlur, onKeyDown }: Props) {
  return (
    <input
      className={[
        'h-7 rounded-md border px-2 text-sm text-ink outline-none',
        invalid ? 'border-danger bg-danger-soft' : 'border-line-strong focus:border-primary',
      ].join(' ')}
      style={width ? { width } : undefined}
      value={value}
      placeholder={placeholder}
      autoFocus={autoFocus}
      title={title}
      onChange={(e) => onChange(e.target.value)}
      onBlur={onBlur}
      onKeyDown={(e) => {
        onKeyDown?.(e)
        if (!e.defaultPrevented && onEnter && e.key === 'Enter') onEnter()
      }}
    />
  )
}
