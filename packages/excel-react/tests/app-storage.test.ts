import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Workbook } from '@xexcel/core'
import { STORAGE_KEY, StorageLike, WorkbookStorage } from '../src/storage'

function memStorage(): StorageLike & { dump: Map<string, string> } {
  const dump = new Map<string, string>()
  return {
    dump,
    getItem: (k) => dump.get(k) ?? null,
    setItem: (k, v) => void dump.set(k, v),
    removeItem: (k) => void dump.delete(k),
  }
}

const mkWb = () => Workbook.create({ rowCount: 5, colCount: 5 })

describe('app/storage', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  it('saveNow 写入信封；load 还原等价 workbook', () => {
    const st = new WorkbookStorage(memStorage())
    const wb = mkWb()
    st.saveNow(wb)
    const back = st.load()
    expect(back?.toJSON()).toEqual(wb.toJSON())
    expect(st.getStatus().error).toBeNull()
    expect(st.getStatus().savedAt).not.toBeNull()
  })

  it('无存档 load 返回 null', () => {
    expect(new WorkbookStorage(memStorage()).load()).toBeNull()
  })

  it('损坏存档备份到 .corrupt 并删除原键，返回 null', () => {
    const s = memStorage()
    s.setItem(STORAGE_KEY, '{broken')
    const st = new WorkbookStorage(s)
    expect(st.load()).toBeNull()
    expect(s.dump.get(STORAGE_KEY + '.corrupt')).toBe('{broken')
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
  })

  it('version 不符同样走备份路径', () => {
    const s = memStorage()
    s.setItem(STORAGE_KEY, JSON.stringify({ version: 999, savedAt: 'x', workbook: {} }))
    expect(new WorkbookStorage(s).load()).toBeNull()
    expect(s.dump.has(STORAGE_KEY + '.corrupt')).toBe(true)
  })

  it('schedule 防抖：连续调度只写一次，取最新 getter', () => {
    const s = memStorage()
    const st = new WorkbookStorage(s, 1000)
    const wb1 = mkWb()
    let wb2 = mkWb().renameSheet('s1', '改名')
    st.schedule(() => wb1)
    st.schedule(() => wb2)
    vi.advanceTimersByTime(999)
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
    vi.advanceTimersByTime(1)
    expect(s.dump.has(STORAGE_KEY)).toBe(true)
    expect(st.load()!.names.get('s1')).toBe('改名')
  })

  it('flush 立即写并清 pending；重复 flush 幂等', () => {
    const s = memStorage()
    const st = new WorkbookStorage(s, 1000)
    st.schedule(() => mkWb())
    st.flush()
    expect(s.dump.has(STORAGE_KEY)).toBe(true)
    s.dump.clear()
    st.flush()
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
  })

  it('写入失败置 status.error 不抛，并通知订阅者', () => {
    const s = memStorage()
    s.setItem = () => { throw new Error('QuotaExceeded') }
    const st = new WorkbookStorage(s)
    const seen: string[] = []
    st.subscribeStatus(() => seen.push(st.getStatus().error ?? 'null'))
    expect(() => st.saveNow(mkWb())).not.toThrow()
    expect(st.getStatus().error).toContain('QuotaExceeded')
    expect(seen).toEqual(['QuotaExceeded'])
  })

  it('clear 删除存档键', () => {
    const s = memStorage()
    const st = new WorkbookStorage(s)
    st.saveNow(mkWb())
    st.clear()
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
  })

  it('suspend 删除存档后 schedule/saveNow 不再写入', () => {
    const s = memStorage()
    const st = new WorkbookStorage(s, 1000)
    st.saveNow(mkWb())
    st.suspend()
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
    st.saveNow(mkWb())
    st.schedule(() => mkWb())
    vi.advanceTimersByTime(2000)
    st.flush()
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
  })

  it('suspend 取消已挂起的防抖 timer：到点不写回', () => {
    const s = memStorage()
    const st = new WorkbookStorage(s, 1000)
    st.schedule(() => mkWb())
    st.suspend()
    vi.advanceTimersByTime(2000)
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
  })

  describe('suspend/resume', () => {
    it('suspend 后 schedule/saveNow 均不写；resume 后恢复', () => {
      const s = memStorage()
      const st = new WorkbookStorage(s, 1000)
      const wb = mkWb()
      st.suspend()
      st.schedule(() => wb)
      st.saveNow(wb)
      expect(s.dump.has(STORAGE_KEY)).toBe(false)
      st.resume()
      st.schedule(() => wb)
      st.flush()
      expect(s.dump.has(STORAGE_KEY)).toBe(true)
    })

    it('suspend 幂等；resume 不补保存（无 pending 时不写）', () => {
      const s = memStorage()
      const st = new WorkbookStorage(s, 1000)
      st.suspend()
      st.suspend()
      st.resume()
      st.flush()
      expect(s.dump.has(STORAGE_KEY)).toBe(false)
    })
  })

  describe('pause', () => {
    it('pause 不清存档：既有存档保留', () => {
      const s = memStorage()
      const st = new WorkbookStorage(s, 1000)
      st.saveNow(mkWb())
      st.pause()
      expect(s.dump.has(STORAGE_KEY)).toBe(true)
    })

    it('pause 期间 schedule/saveNow 不写，且取消已挂起的防抖 timer', () => {
      const s = memStorage()
      const st = new WorkbookStorage(s, 1000)
      const wb = mkWb().renameSheet('s1', '旧名')
      st.schedule(() => wb)
      st.pause()
      vi.advanceTimersByTime(2000)
      expect(s.dump.has(STORAGE_KEY)).toBe(false)
      st.schedule(() => wb)
      st.saveNow(wb)
      expect(s.dump.has(STORAGE_KEY)).toBe(false)
    })

    it('resume 后恢复自动保存', () => {
      const s = memStorage()
      const st = new WorkbookStorage(s, 1000)
      const wb = mkWb()
      st.pause()
      st.resume()
      st.schedule(() => wb)
      st.flush()
      expect(s.dump.has(STORAGE_KEY)).toBe(true)
    })
  })

  it('语义损坏存档（结构合法、语义空）走 .corrupt 备份路径返回 null', () => {
    const s = memStorage()
    const bad = JSON.stringify({
      version: 1,
      savedAt: 'x',
      workbook: { order: [], active: 's1', names: [], sheets: {} },
    })
    s.setItem(STORAGE_KEY, bad)
    const st = new WorkbookStorage(s)
    expect(st.load()).toBeNull()
    expect(s.dump.get(STORAGE_KEY + '.corrupt')).toBe(bad)
    expect(s.dump.has(STORAGE_KEY)).toBe(false)
  })
})
