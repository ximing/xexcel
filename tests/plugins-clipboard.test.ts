import { describe, it, expect } from 'vitest'
import { ClipboardPayload, planPaste } from '../src/plugins/clipboard'

const BOUNDS = { rowCount: 100, colCount: 26 }

const payload = (over: Partial<ClipboardPayload> = {}): ClipboardPayload => ({
  sheet: 's1',
  range: { sr: 0, sc: 0, er: 0, ec: 0 },
  tsv: '2',
  raws: [['=A1*2']],
  cut: false,
  ...over,
})

describe('planPaste', () => {
  it('内部 copy：公式按目标偏移', () => {
    const { entries, clearSource } = planPaste(payload(), '2', { sr: 5, sc: 3, er: 5, ec: 3 }, BOUNDS)
    expect(entries).toEqual([{ row: 5, col: 3, cell: { raw: '=D6*2' } }])
    expect(clearSource).toBe(false)
  })
  it('内部 copy 平铺：多格选区逐格偏移', () => {
    const { entries } = planPaste(payload(), '2', { sr: 1, sc: 0, er: 2, ec: 0 }, BOUNDS)
    expect(entries.map((e) => e.cell?.raw)).toEqual(['=A2*2', '=A3*2'])
  })
  it('内部 cut：公式不偏移（移动语义），且不相交 → clearSource', () => {
    const p = payload({ cut: true })
    const { entries, clearSource } = planPaste(p, '2', { sr: 5, sc: 5, er: 5, ec: 5 }, BOUNDS)
    expect(entries).toEqual([{ row: 5, col: 5, cell: { raw: '=A1*2' } }])
    expect(clearSource).toBe(true)
  })
  it('内部 cut 粘贴到相交区域 → 不清源', () => {
    const p = payload({ cut: true, range: { sr: 0, sc: 0, er: 1, ec: 1 }, raws: [['1', '2'], ['3', '4']], tsv: '1\t2\n3\t4' })
    const { clearSource } = planPaste(p, p.tsv, { sr: 1, sc: 1, er: 2, ec: 2 }, BOUNDS)
    expect(clearSource).toBe(false)
  })
  it('源空格 → 清目标格', () => {
    const p = payload({ raws: [[null]] })
    const { entries } = planPaste(p, '2', { sr: 3, sc: 3, er: 3, ec: 3 }, BOUNDS)
    expect(entries).toEqual([{ row: 3, col: 3, cell: null }])
  })
  it('外部 TSV（文本不匹配负载）：原样落格不偏移', () => {
    const { entries, clearSource } = planPaste(payload(), 'x\ty', { sr: 0, sc: 0, er: 0, ec: 1 }, BOUNDS)
    expect(entries).toEqual([
      { row: 0, col: 0, cell: { raw: 'x' } },
      { row: 0, col: 1, cell: { raw: 'y' } },
    ])
    expect(clearSource).toBe(false)
  })
  it('无负载纯 TSV：多行网格、越界裁剪、空串清格', () => {
    const { entries } = planPaste(null, 'a\t\nb\tc', { sr: 98, sc: 25, er: 99, ec: 26 }, BOUNDS)
    expect(entries).toEqual([
      { row: 98, col: 25, cell: { raw: 'a' } },
      { row: 99, col: 25, cell: { raw: 'b' } },
    ])
  })
})
