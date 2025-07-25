import type { LucideIcon } from 'lucide-react'

export function Icon({ icon: Glyph, size = 16 }: { icon: LucideIcon; size?: number }) {
  return <Glyph size={size} strokeWidth={2} aria-hidden />
}
