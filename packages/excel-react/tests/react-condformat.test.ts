import { describe, expect, it } from 'vitest'
import { toggleCFStyle } from '../src/cfstyle'

describe('toggleCFStyle', () => {
  it('关态剔除键，不留显式 undefined own property', () => {
    const on = toggleCFStyle({}, 'bold')
    expect(on).toEqual({ bold: true })
    const off = toggleCFStyle(on, 'bold')
    expect(off).toEqual({})
    expect(Object.prototype.hasOwnProperty.call(off, 'bold')).toBe(false)
  })
  it('关态保留其他样式键', () => {
    const s = toggleCFStyle({ bg: '#ffc7ce', color: '#9c0006', underline: true }, 'underline')
    expect(s).toEqual({ bg: '#ffc7ce', color: '#9c0006' })
    expect('underline' in s).toBe(false)
  })
  it('开态写 true 且不改动原对象', () => {
    const src = { italic: true }
    const s = toggleCFStyle(src, 'strikethrough')
    expect(s).toEqual({ italic: true, strikethrough: true })
    expect(src).toEqual({ italic: true })
  })
})
