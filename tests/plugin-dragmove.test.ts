import { describe, expect, it } from 'vitest'
import { SheetData } from '../src/core/model'
import { buildMove } from '../src/plugins/dragmove'
const mk = (r = 8, c = 8) => SheetData.create({ rowCount: r, colCount: c })
const set = (s: SheetData, r: number, c: number, raw: string, style?: any) =>
  s.setCell(r, c, style ? { raw, style } : { raw })

describe('buildMove cut 移动', () => {
  it('源 raw+style 搬到目标，清源；公式不 shift', () => {
    let s = mk(); s = set(s, 0, 0, '1', { bold: true }); s = set(s, 0, 1, '=A1')
    const { entries, clearSource, reject } = buildMove(s, { sr: 0, sc: 0, er: 0, ec: 1 }, { sr: 2, sc: 2, er: 2, ec: 3 })
    expect(reject).toBe(false); expect(clearSource).toBe(true)
    const e = (r: number, c: number) => entries.find(x => x.row === r && x.col === c)!.cell
    expect(e(2, 2)).toEqual({ raw: '1', style: { bold: true } })
    expect(e(2, 3)).toEqual({ raw: '=A1' }) // 公式不 shift（cut 语义）
  })
  it('源与目标相交 → reject', () => {
    let s = mk(); s = set(s, 0, 0, 'x')
    const { reject } = buildMove(s, { sr: 0, sc: 0, er: 1, ec: 1 }, { sr: 1, sc: 1, er: 2, ec: 2 })
    expect(reject).toBe(true)
  })
  it('目标落 merge → reject', () => {
    let s = mk().setMerges([{ sr: 3, sc: 3, er: 4, ec: 4 }]); s = set(s, 0, 0, 'x')
    const { reject } = buildMove(s, { sr: 0, sc: 0, er: 0, ec: 0 }, { sr: 3, sc: 3, er: 3, ec: 3 })
    expect(reject).toBe(true)
  })
  it('目标越界 → clampRange 不扩表', () => {
    let s = mk(4, 4); s = set(s, 0, 0, 'x')
    const { entries } = buildMove(s, { sr: 0, sc: 0, er: 0, ec: 0 }, { sr: 3, sc: 3, er: 4, ec: 4 })
    expect(entries.every(e => e.row <= 3 && e.col <= 3)).toBe(true)
  })
})
