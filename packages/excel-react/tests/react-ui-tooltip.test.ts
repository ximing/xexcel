import { describe, expect, it } from 'vitest'
import { resolvePlacement } from '../src/ui/tooltipPlacement'

describe('resolvePlacement', () => {
  const H = 800
  it('auto：下方空间够 → bottom', () => {
    expect(resolvePlacement(40, 12, 28, 'auto', H)).toBe('bottom')
  })
  it('auto：下方贴底（状态栏）→ top', () => {
    expect(resolvePlacement(790, 762, 28, 'auto', H)).toBe('top')
  })
  it('显式指定优先', () => {
    expect(resolvePlacement(40, 12, 28, 'top', H)).toBe('top')
    expect(resolvePlacement(790, 762, 28, 'bottom', H)).toBe('bottom')
  })
})
