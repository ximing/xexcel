// 富剪贴板插件。职责：
// - copy/cut（handleCopy）：选区每个 range → ClipboardArea（range+raws+styles）；
//   text/plain = 各 area TSV 块、块间空行分隔（多区域语义）；text/html = 每 area 一段
//   <table>、单元格内联 style→CSS；闭包记 ClipboardPayload（指纹=tsv）；cut 标记移动语义
// - paste（handlePaste）：剪贴板文本与负载指纹一致 → 富落格
//   （raw 落 setCells、style 整体替换落 setCellStyles；copy 偏移公式引用、cut 不偏移=移动语义，
//   不相交则清源）；不一致（外部内容）→ TSV 兜底：文本原样落 cell:{raw}、空串清格、越界裁剪、
//   样式丢失（已知限制）。多区域 paste：首 area 锚 activeCell，余 area 按源内相对偏移落格。
// 全部经 dispatch transaction，不直接改 doc。
import { CellRange, clampRange, normalizeRange, rangeCellCount, rangesEqual, rangesIntersect } from '../core/addr'
import { Cell, CellStyle, Workbook } from '../core/model'
import { EditorViewLike, Plugin } from '../core/plugin'
import { forEachSelectionRange, rangeSelection, selectionRange } from '../core/selection'
import type { Selection } from '../core/selection'
import { evaluatorFor } from '../formula/engine'
import { normalizedCell } from '../formula/input'
import { shiftFormula } from '../formula/transform'

export interface ClipboardArea {
  range: CellRange
  raws: (string | null)[][]       // 源区域 raw 网格（null=空格）
  styles: (CellStyle | null)[][] // 源区域完整 style 快照（null=无 style）
}
export interface ClipboardPayload {
  sheet: string
  areas: ClipboardArea[]
  tsv: string  // 指纹：多区域块间空行分隔，与 text/plain 一致
  cut: boolean
}

// CellStyle → 内联 CSS（出站 text/html 简单映射；复杂样式可能简化）
function styleToCss(st: CellStyle | null): string {
  if (!st) return ''
  const p: string[] = []
  if (st.bold) p.push('font-weight:bold')
  if (st.italic) p.push('font-style:italic')
  if (st.underline) p.push('text-decoration:underline')
  if (st.strikethrough) p.push('text-decoration:line-through')
  if (st.color) p.push(`color:${st.color}`)
  if (st.bg) p.push(`background-color:${st.bg}`)
  if (st.align) p.push(`text-align:${st.align}`)
  if (st.fontSize) p.push(`font-size:${st.fontSize}px`)
  if (st.fontFamily) p.push(`font-family:${st.fontFamily}`)
  return p.join(';')
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

// 单格写入（raw 落 entries、style 落 styleEntries；copy 偏移公式引用、cut 不偏移）。
// planPaste 在 T2 只产 raw entries；style 应用由 handlePaste 用 setCellStyles（整体替换）。
function writeEntry(
  entries: { row: number; col: number; cell: Cell | null }[],
  styleEntries: { row: number; col: number; style: CellStyle | null }[],
  area: ClipboardArea, r: number, c: number, i: number, j: number, cut: boolean,
): void {
  const src = area.raws[i][j]
  const style = area.styles[i][j]
  if (src === null || src === '') {
    entries.push({ row: r, col: c, cell: null })
    styleEntries.push({ row: r, col: c, style })
    return
  }
  // copy 偏移公式引用：delta 按目标格减源内偏移(i,j)再减源区域起点（tile 平铺时每格按自身起点）；
  // cut 移动语义不 shift
  const raw = !cut && src.startsWith('=')
    ? shiftFormula(src, r - i - area.range.sr, c - j - area.range.sc)
    : src
  entries.push({ row: r, col: c, cell: { raw } })
  styleEntries.push({ row: r, col: c, style })
}

// 计划粘贴：内部命中 → 富（raw+style 整体替换，copy 偏移公式、cut 不 shift）；外部 → TSV 落 cell:{raw}
export function planPaste(
  payload: ClipboardPayload | null,
  text: string,
  target: CellRange,
  bounds: { rowCount: number; colCount: number },
): { entries: { row: number; col: number; cell: Cell | null }[]; styleEntries: { row: number; col: number; style: CellStyle | null }[]; clearSource: boolean } {
  // 防御性规范化：系统剪贴板可能把 \n 变成 \r\n，不先归一会让指纹比对假阴性
  const normText = text.replace(/\r\n?/g, '\n')
  const er = Math.min(target.er, bounds.rowCount - 1)
  const ec = Math.min(target.ec, bounds.colCount - 1)
  const entries: { row: number; col: number; cell: Cell | null }[] = []
  const styleEntries: { row: number; col: number; style: CellStyle | null }[] = []

  if (payload && normText === payload.tsv) {
    if (payload.areas.length === 1) {
      // 单区域：tile 平铺到 target（保留 F1 逐格偏移语义——单格 payload 落多格选区逐格偏移）
      const area = payload.areas[0]
      const h = area.raws.length
      const w = area.raws[0]?.length ?? 1
      for (let r = target.sr; r <= er; r++) {
        const i = (r - target.sr) % h
        for (let c = target.sc; c <= ec; c++) {
          const j = (c - target.sc) % w
          writeEntry(entries, styleEntries, area, r, c, i, j, payload.cut)
        }
      }
    } else {
      // 多区域：首 area 锚 target.sr/sc；余 area 按源内相对偏移落格（无平铺）
      const a0 = payload.areas[0]
      for (let ai = 0; ai < payload.areas.length; ai++) {
        const area = payload.areas[ai]
        const baseR = target.sr + (ai === 0 ? 0 : area.range.sr - a0.range.sr)
        const baseC = target.sc + (ai === 0 ? 0 : area.range.sc - a0.range.sc)
        const h = area.raws.length
        const w = area.raws[0]?.length ?? 1
        const aEr = Math.min(baseR + h - 1, er)
        const aEc = Math.min(baseC + w - 1, ec)
        for (let r = Math.max(target.sr, baseR); r <= aEr; r++) {
          const i = r - baseR
          if (i < 0 || i >= h) continue
          for (let c = Math.max(target.sc, baseC); c <= aEc; c++) {
            const j = c - baseC
            if (j < 0 || j >= w) continue
            writeEntry(entries, styleEntries, area, r, c, i, j, payload.cut)
          }
        }
      }
    }
    // cut 清源：仅当所有 area 都不等于且不相交于 target（避免覆盖自身时清空）
    const clearSource = payload.cut &&
      payload.areas.every(a => !rangesEqual(a.range, target)) &&
      !payload.areas.some(a => rangesIntersect(a.range, target))
    return { entries, styleEntries, clearSource }
  }

  // 外部 TSV：文本原样落格（含 '=' 开头按新公式处理），空串清格，平铺到 target，样式丢失
  const lines = normText.split('\n')
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop()
  const grid = lines.map((l) => l.split('\t'))
  for (let r = target.sr; r <= er; r++) {
    const rowArr = grid[(r - target.sr) % grid.length]
    for (let c = target.sc; c <= ec; c++) {
      const t = rowArr[(c - target.sc) % rowArr.length]
      entries.push({ row: r, col: c, cell: t === '' ? null : normalizedCell(t, undefined) })
    }
  }
  return { entries, styleEntries, clearSource: false }
}

// 纯函数：从 doc + 选区构造 copy 负载 + text/plain + text/html（供 handleCopy 与单测共用，保持 handleCopy 薄）
export function buildCopyPayload(
  doc: Workbook, selection: Selection, cut: boolean,
): { payload: ClipboardPayload; tsv: string; html: string } {
  const ev = evaluatorFor(doc)
  const sheetId = doc.active
  const sheet = doc.activeSheet
  const areas: ClipboardArea[] = []
  const displayGrids: string[][][] = [] // [area][row][col] 显示文本，供 html td 内容
  const tsvBlocks: string[] = []
  forEachSelectionRange(selection, (r) => {
    const raws: (string | null)[][] = []
    const styles: (CellStyle | null)[][] = []
    const displayGrid: string[][] = []
    const displayLines: string[] = []
    for (let row = r.sr; row <= r.er; row++) {
      const rawRow: (string | null)[] = []
      const styleRow: (CellStyle | null)[] = []
      const displayRow: string[] = []
      for (let col = r.sc; col <= r.ec; col++) {
        const cell = sheet.getCell(row, col)
        displayRow.push(ev.displayText(sheetId, row, col))
        rawRow.push(cell?.raw ?? null)
        styleRow.push(cell?.style ?? null)
      }
      raws.push(rawRow)
      styles.push(styleRow)
      displayGrid.push(displayRow)
      displayLines.push(displayRow.join('\t'))
    }
    areas.push({ range: normalizeRange(r), raws, styles })
    displayGrids.push(displayGrid)
    tsvBlocks.push(displayLines.join('\n'))
  })
  const tsv = tsvBlocks.join('\n\n') // 多区域块间空行（Excel 语义）
  // text/html：每 area 一段 <table>，单元格内联 style→CSS + 显示文本
  const html = areas
    .map((a, ai) => {
      const rows = a.raws
        .map((rowArr, ri) =>
          '<tr>' +
          rowArr
            .map((_, ci) => `<td style="${styleToCss(a.styles[ri][ci])}">${escapeHtml(displayGrids[ai][ri][ci] ?? '')}</td>`)
            .join('') +
          '</tr>',
        )
        .join('')
      return `<table>${rows}</table>`
    })
    .join('')
  return { payload: { sheet: sheetId, areas, tsv, cut }, tsv, html }
}

export function clipboard(): Plugin {
  let payload: ClipboardPayload | null = null

  return new Plugin({
    props: {
      handleCopy(view: EditorViewLike, cut: boolean, event: ClipboardEvent): boolean {
        const state = view.state
        const { payload: p, tsv, html } = buildCopyPayload(state.doc, state.selection, cut)
        event.clipboardData?.setData('text/plain', tsv)
        event.clipboardData?.setData('text/html', html)
        payload = p
        return true // EditorView 侧 preventDefault
      },
      handlePaste(view: EditorViewLike, text: string): boolean {
        if (text === '') return true // 空剪贴板：吞掉事件即可
        const state = view.state
        const sheet = state.activeSheet
        const sheetId = state.doc.active
        // 规范化一次：指纹比对（planPaste 内）与网格解析统一用 \n 文本
        const normText = text.replace(/\r\n?/g, '\n')
        const lines = normText.split('\n')
        if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop()
        const grid = lines.map((l) => l.split('\t'))
        const oneCell = grid.length === 1 && grid[0].length === 1
        const selRange = selectionRange(state.selection)
        const start = state.selection.activeCell
        // 目标区域：单格剪贴板 + 多格选区 → 平铺整个选区；否则从 activeCell 按剪贴板尺寸展开
        const target: CellRange =
          oneCell && rangeCellCount(selRange) > 1
            ? selRange
            : {
                sr: start.row,
                sc: start.col,
                er: start.row + grid.length - 1,
                ec: start.col + grid[0].length - 1,
              }
        const { entries, styleEntries, clearSource } = planPaste(payload, normText, target, {
          rowCount: sheet.rowCount,
          colCount: sheet.colCount,
        })
        const tr = state.tr
        if (entries.length) tr.setCells(sheetId, entries)
        // style 整体替换（RestoreStyleStep 语义：逐格覆盖，null=删 style 键）
        if (styleEntries.length) tr.setCellStyles(styleEntries)
        // cut 移动语义：仅当源在本表才清源（跨表粘贴不清别表数据）；多区域逐 area 清
        if (clearSource && payload && payload.sheet === sheetId)
          for (const a of payload.areas) tr.clearRange(a.range)
        // cut 负载一次性（移动语义）；copy 负载保留到下次 copy/cut，可反复粘贴
        // （与 Excel 一致：每次粘贴按各自目标偏移公式引用）
        if (payload?.cut) payload = null
        if (tr.steps.length) {
          tr.setSelection(rangeSelection(
            clampRange(target, sheet.rowCount, sheet.colCount),
            { row: target.sr, col: target.sc },
          ))
          view.dispatch(tr)
        }
        return true
      },
    },
  })
}
