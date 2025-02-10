interface Props {
  value: string
  onChange: (v: string) => void
  placeholder?: string
  invalid?: boolean
  width?: number
  autoFocus?: boolean
  onEnter?: () => void
}

export function TextInput({ value, onChange, placeholder, invalid, width, autoFocus, onEnter }: Props) {
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
      onChange={(e) => onChange(e.target.value)}
      onKeyDown={onEnter ? (e) => { if (e.key === 'Enter') onEnter() } : undefined}
    />
  )
}
