import { Tooltip } from './Tooltip'

interface Props {
  value: string | number
  options: { value: string | number; label: string }[]
  tip?: string
  width?: number
  onChange: (v: string) => void
}

export function Select({ value, options, tip, width, onChange }: Props) {
  const sel = (
    <select
      aria-label={tip}
      className="h-7 rounded-md bg-transparent px-1 text-sm text-ink outline-none hover:bg-hover"
      style={width !== undefined ? { width } : undefined}
      value={value}
      onChange={(e) => onChange(e.target.value)}
    >
      {options.map((o) => (
        <option key={o.value} value={o.value}>{o.label}</option>
      ))}
    </select>
  )
  return tip ? <Tooltip tip={tip}>{sel}</Tooltip> : sel
}
