// planPaste 单测（M3c 富 payload areas 结构迁移后）。覆盖：内部 copy 公式偏移/平铺、
// cut 不偏移+清源、源空格清目标、外部 TSV 兜底、越界裁剪、CRLF 指纹、copy 负载复用。
import { describe, it, expect } from 'vitest'
import { ClipboardPayload, planPaste } from '../src/plugins/clipboard'

const BOUNDS = { rowCount: 100, colCount: 26 }

// M3c 富 payload：单 area，range+raws+styles。旧 range/raws 字段已迁至 areas[0]。
const payload = (over: Partial<ClipboardPayload> = {}): ClipboardPayload => ({
  sheet: 's1',
  areas: [{ range: { sr: 0, sc: 0, er: 0, ec: 0 }, raws: [['=A1*2']], styles: [[null]] }],
  tsv: '2',
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
    const p = payload({
      cut: true,
      areas: [{
        range: { sr: 0, sc: 0, er: 1, ec: 1 },
        raws: [['1', '2'], ['3', '4']],
        styles: [[null, null], [null, null]],
      }],
      tsv: '1\t2\n3\t4',
    })
    const { clearSource } = planPaste(p, p.tsv, { sr: 1, sc: 1, er: 2, ec: 2 }, BOUNDS)
    expect(clearSource).toBe(false)
  })
  it('源空格 → 清目标格', () => {
    const p = payload({ areas: [{ range: { sr: 0, sc: 0, er: 0, ec: 0 }, raws: [[null]], styles: [[null]] }] })
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
  it('内部 copy 多格：源内相对偏移不重复计入（F1）', () => {
    const p = payload({
      areas: [{ range: { sr: 0, sc: 0, er: 1, ec: 0 }, raws: [['=B1'], ['=B2']], styles: [[null], [null]] }],
      tsv: '2\n4',
    })
    const { entries } = planPaste(p, '2\n4', { sr: 4, sc: 0, er: 5, ec: 0 }, BOUNDS)
    expect(entries.map((e) => e.cell?.raw)).toEqual(['=B5', '=B6'])
  })
  it('内部 copy 平铺到更高选区：每个 tile 按自身起点偏移（F1）', () => {
    // 注：协调者原期望第二个 tile 复用同一 delta（'=B5','=B6' 重复），
    // 但与 F1 评审修复公式（tileRow = r - (r-target.sr)%h）、既有单格平铺用例
    // （'=A2*2','=A3*2'）及 Excel 行为矛盾；此处按 per-tile 语义断言
    const p = payload({
      areas: [{ range: { sr: 0, sc: 0, er: 1, ec: 0 }, raws: [['=B1'], ['=B2']], styles: [[null], [null]] }],
      tsv: '2\n4',
    })
    const { entries } = planPaste(p, '2\n4', { sr: 4, sc: 0, er: 7, ec: 0 }, BOUNDS)
    expect(entries.map((e) => e.cell?.raw)).toEqual(['=B5', '=B6', '=B7', '=B8'])
  })
  it('CRLF 指纹：粘贴文本带 \\r\\n 仍命中内部负载（F2）', () => {
    const p = payload({
      areas: [{ range: { sr: 0, sc: 0, er: 1, ec: 0 }, raws: [['=B1'], ['=B2']], styles: [[null], [null]] }],
      tsv: '2\n4',
    })
    const { entries } = planPaste(p, '2\r\n4', { sr: 4, sc: 0, er: 5, ec: 0 }, BOUNDS)
    expect(entries.map((e) => e.cell?.raw)).toEqual(['=B5', '=B6'])
  })
  it('copy 负载可复用：同一 payload 两次粘贴各自按目标偏移（F3）', () => {
    // planPaste 是纯函数：负载生命周期由插件闭包管理（copy 保留、cut 一次性），
    // 这里验证同一 copy payload 连续两次调用互不污染、各自正确偏移
    const p = payload()
    const first = planPaste(p, '2', { sr: 5, sc: 3, er: 5, ec: 3 }, BOUNDS)
    const second = planPaste(p, '2', { sr: 6, sc: 3, er: 6, ec: 3 }, BOUNDS)
    expect(first.entries).toEqual([{ row: 5, col: 3, cell: { raw: '=D6*2' } }])
    expect(second.entries).toEqual([{ row: 6, col: 3, cell: { raw: '=D7*2' } }])
    // cut 负载（一次性移动语义由插件侧清 payload）：planPaste 层面 clearSource 为 true
    const cut = planPaste(payload({ cut: true }), '2', { sr: 6, sc: 3, er: 6, ec: 3 }, BOUNDS)
    expect(cut.clearSource).toBe(true)
  })
})
