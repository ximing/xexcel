import { describe, expect, it } from 'vitest'
import { askConfirm, getConfirm, resolveConfirm } from '../src/ui/confirmStore'

describe('confirmStore', () => {
  it('ask 后挂起，resolve(true) 兑现 true 并清空', async () => {
    const p = askConfirm({ title: '确认删除？' })
    expect(getConfirm()?.title).toBe('确认删除？')
    resolveConfirm(true)
    expect(await p).toBe(true)
    expect(getConfirm()).toBeNull()
  })
  it('resolve(false) 兑现 false', async () => {
    const p = askConfirm({ title: '继续？' })
    resolveConfirm(false)
    expect(await p).toBe(false)
  })
  it('前问未决时再 ask：前问兑现 false，新问挂起', async () => {
    const p1 = askConfirm({ title: 'A' })
    const p2 = askConfirm({ title: 'B' })
    expect(await p1).toBe(false)
    expect(getConfirm()?.title).toBe('B')
    resolveConfirm(true)
    expect(await p2).toBe(true)
  })
})
