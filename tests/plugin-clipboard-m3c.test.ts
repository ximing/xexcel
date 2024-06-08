// F4 富剪贴板单测：富 payload（areas+styles）/ planPaste（raw entries + styleEntries 整体替换）/
// copy 多区域 tsv 块间空行 + text/html table / 外部 TSV cell:{raw} 兜底 / cut 不相交清源 / clampRange 不扩表。
import { describe, expect, it } from 'vitest'
import { Workbook } from '../src/core/model'
import { appendRange, rangeSelection } from '../src/core/selection'
import { buildCopyPayload, ClipboardPayload, planPaste } from '../src/plugins/clipboard'

const mk = (rows = 6, cols = 6) => Workbook.create({ rowCount: rows, colCount: cols })

describe('planPaste 富 payload', () => {
  it('内部 copy：raw 落 entries、style 整体替换落 styleEntries（RestoreStyleStep 语义）', () => {
    const payload: ClipboardPayload = {
      sheet: 's1', cut: false, tsv: '1\t2',
      areas: [{ range: { sr: 0, sc: 0, er: 0, ec: 1 }, raws: [['1', '2']], styles: [[{ bold: true }, null]] }],
    }
    const { entries, styleEntries } = planPaste(payload, '1\t2', { sr: 2, sc: 2, er: 2, ec: 3 }, { rowCount: 6, colCount: 6 })
    // planPaste 在 T2 只产 raw entries；style 应用由 handlePaste 用 setCellStyles（整体替换）
    expect(entries.find(e => e.row === 2 && e.col === 2)!.cell).toEqual({ raw: '1' })
    expect(styleEntries.find(e => e.row === 2 && e.col === 2)!.style).toEqual({ bold: true })
    expect(styleEntries.find(e => e.row === 2 && e.col === 3)!.style).toBeNull()
  })

  it('内部 cut 移动：不相交 → clearSource=true；相交 → false', () => {
    const payload: ClipboardPayload = {
      sheet: 's1', cut: true, tsv: 'x',
      areas: [{ range: { sr: 0, sc: 0, er: 0, ec: 0 }, raws: [['x']], styles: [[null]] }],
    }
    const a = planPaste(payload, 'x', { sr: 3, sc: 3, er: 3, ec: 3 }, { rowCount: 6, colCount: 6 })
    expect(a.clearSource).toBe(true)
    const b = planPaste(payload, 'x', { sr: 0, sc: 0, er: 0, ec: 0 }, { rowCount: 6, colCount: 6 })
    expect(b.clearSource).toBe(false)
  })

  it('外部 TSV：未命中指纹 → cell:{raw} 兜底，样式丢失', () => {
    const { entries, clearSource, styleEntries } = planPaste(null, 'a\tb\n1\t2', { sr: 0, sc: 0, er: 1, ec: 1 }, { rowCount: 6, colCount: 6 })
    expect(clearSource).toBe(false)
    expect(entries.find(e => e.row === 0 && e.col === 0)!.cell).toEqual({ raw: 'a' })
    expect(styleEntries).toEqual([])
  })

  it('多区域 copy：tsv 块间空行分隔 + areas 多条', () => {
    // 两 area：area1 = 2×2（"1\t2\n3\t4"），area2 = 1×2（"5\t6"）；块间一空行
    let wb = mk()
    let sheet = wb.activeSheet
      .setCell(0, 0, { raw: '1' }).setCell(0, 1, { raw: '2' })
      .setCell(1, 0, { raw: '3' }).setCell(1, 1, { raw: '4' })
      .setCell(0, 3, { raw: '5' }).setCell(0, 4, { raw: '6' })
    wb = wb.setSheet('s1', sheet)
    // 选区：ranges[0]={0,0..1,1}；ranges[1]={0,3..0,4}
    const sel = appendRange(rangeSelection({ sr: 0, sc: 0, er: 1, ec: 1 }), { sr: 0, sc: 3, er: 0, ec: 4 })
    const { tsv, payload } = buildCopyPayload(wb, sel, false)
    expect(tsv).toBe('1\t2\n3\t4\n\n5\t6')
    expect(payload.areas.length).toBe(2)
    expect(payload.areas[0].raws).toEqual([['1', '2'], ['3', '4']])
    expect(payload.areas[1].raws).toEqual([['5', '6']])
  })

  it('目标越出表边界 → clampRange 不扩表', () => {
    const payload: ClipboardPayload = {
      sheet: 's1', cut: false, tsv: '1\t2',
      areas: [{ range: { sr: 0, sc: 0, er: 0, ec: 1 }, raws: [['1', '2']], styles: [[null, null]] }],
    }
    const { entries } = planPaste(payload, '1\t2', { sr: 5, sc: 5, er: 5, ec: 6 }, { rowCount: 6, colCount: 6 })
    // 越界列 6 → clamp 到 5；entries 不含 col 6
    expect(entries.every(e => e.col <= 5)).toBe(true)
  })

  it('内部 copy 公式按目标偏移（单区域 tile 平铺逐格偏移）', () => {
    const payload: ClipboardPayload = {
      sheet: 's1', cut: false, tsv: '2',
      areas: [{ range: { sr: 0, sc: 0, er: 0, ec: 0 }, raws: [['=A1*2']], styles: [[null]] }],
    }
    // 单格 payload 平铺到 2 行选区：每格按自身起点偏移
    const { entries } = planPaste(payload, '2', { sr: 1, sc: 0, er: 2, ec: 0 }, { rowCount: 100, colCount: 26 })
    expect(entries.map(e => e.cell?.raw)).toEqual(['=A2*2', '=A3*2'])
  })
})

describe('buildCopyPayload 富导出', () => {
  it('单区域：tsv=text/plain，html 含 <table>，payload.areas=1', () => {
    let wb = mk()
    wb = wb.setSheet('s1', wb.activeSheet.setCell(0, 0, { raw: '1', style: { bold: true } }).setCell(0, 1, { raw: '2' }))
    const sel = rangeSelection({ sr: 0, sc: 0, er: 0, ec: 1 })
    const { tsv, html, payload } = buildCopyPayload(wb, sel, false)
    expect(tsv).toBe('1\t2')
    expect(html).toContain('<table>')
    expect(html).toContain('font-weight:bold')
    expect(payload.areas.length).toBe(1)
    expect(payload.areas[0].styles[0][0]).toEqual({ bold: true })
    expect(payload.cut).toBe(false)
  })

  it('cut 标记透传到 payload', () => {
    const wb = mk()
    const sel = rangeSelection({ sr: 0, sc: 0, er: 0, ec: 0 })
    const { payload } = buildCopyPayload(wb, sel, true)
    expect(payload.cut).toBe(true)
  })
})
