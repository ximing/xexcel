// 剪贴板插件。职责：
// - copy/cut（handleCopy）：选区 TSV（evaluator.displayText）写 text/plain，同时在闭包记
//   ClipboardPayload（源区域 raw 网格 + TSV 指纹）；cut 标记移动语义
// - paste（handlePaste）：剪贴板文本与负载 TSV 一致 → 用 raw 网格落格
//   （copy 按目标偏移公式引用；cut 不偏移=移动语义，不相交则清源）；
//   不一致（外部内容）→ 原 TSV 行为：文本原样落格、空串清格、越界裁剪
// 全部经 dispatch transaction，不直接改 doc。
import { CellRange, rangeCellCount, rangesEqual, rangesIntersect } from '../core/addr'
import { Cell, SheetId } from '../core/model'
import { EditorViewLike, Plugin } from '../core/plugin'
import { selectionRange } from '../core/selection'
import { evaluatorFor } from '../formula/engine'
import { shiftFormula } from '../formula/transform'

export interface ClipboardPayload {
  sheet: SheetId
  range: CellRange
  tsv: string // copy 时写入系统剪贴板的文本（指纹：paste 时据此判定内部粘贴）
  raws: (string | null)[][] // 源区域 raw 网格（null=空格）
  cut: boolean
}

export function planPaste(
  payload: ClipboardPayload | null,
  text: string,
  target: CellRange,
  bounds: { rowCount: number; colCount: number },
): { entries: { row: number; col: number; cell: Cell | null }[]; clearSource: boolean } {
  const er = Math.min(target.er, bounds.rowCount - 1)
  const ec = Math.min(target.ec, bounds.colCount - 1)
  const entries: { row: number; col: number; cell: Cell | null }[] = []

  // 内部粘贴：文本与负载指纹一致 → 用 raw 网格（保留公式）
  if (payload && text === payload.tsv) {
    const h = payload.raws.length
    const w = payload.raws[0]?.length ?? 1
    for (let r = target.sr; r <= er; r++) {
      for (let c = target.sc; c <= ec; c++) {
        const src = payload.raws[(r - target.sr) % h][(c - target.sc) % w]
        if (src === null || src === '') {
          entries.push({ row: r, col: c, cell: null })
          continue
        }
        // copy 偏移公式引用；cut 移动语义不偏移（与 Excel 一致）
        const raw =
          !payload.cut && src.startsWith('=')
            ? shiftFormula(src, r - payload.range.sr, c - payload.range.sc)
            : src
        entries.push({ row: r, col: c, cell: { raw } })
      }
    }
    const clearSource =
      payload.cut && !rangesEqual(payload.range, target) && !rangesIntersect(payload.range, target)
    return { entries, clearSource }
  }

  // 外部 TSV：文本原样落格（含 '=' 开头按新公式处理），空串清格，平铺到 target
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop()
  const grid = lines.map((l) => l.split('\t'))
  for (let r = target.sr; r <= er; r++) {
    const rowArr = grid[(r - target.sr) % grid.length]
    for (let c = target.sc; c <= ec; c++) {
      const t = rowArr[(c - target.sc) % rowArr.length]
      entries.push({ row: r, col: c, cell: t === '' ? null : { raw: t } })
    }
  }
  return { entries, clearSource: false }
}

export function clipboard(): Plugin {
  let payload: ClipboardPayload | null = null

  return new Plugin({
    props: {
      handleCopy(view: EditorViewLike, cut: boolean, event: ClipboardEvent): boolean {
        const state = view.state
        const r = selectionRange(state.selection)
        const ev = evaluatorFor(state.doc)
        const sheetId = state.doc.active
        const sheet = state.activeSheet
        const lines: string[] = []
        const raws: (string | null)[][] = []
        for (let row = r.sr; row <= r.er; row++) {
          const cells: string[] = []
          const rawRow: (string | null)[] = []
          for (let col = r.sc; col <= r.ec; col++) {
            cells.push(ev.displayText(sheetId, row, col))
            rawRow.push(sheet.getCell(row, col)?.raw ?? null)
          }
          lines.push(cells.join('\t'))
          raws.push(rawRow)
        }
        const tsv = lines.join('\n')
        event.clipboardData?.setData('text/plain', tsv)
        payload = { sheet: sheetId, range: r, tsv, raws, cut }
        return true // EditorView 侧 preventDefault
      },
      handlePaste(view: EditorViewLike, text: string): boolean {
        if (text === '') return true // 空剪贴板：吞掉事件即可
        const state = view.state
        const sheet = state.activeSheet
        const sheetId = state.doc.active
        const lines = text.replace(/\r\n?/g, '\n').split('\n')
        if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop()
        const grid = lines.map((l) => l.split('\t'))
        const oneCell = grid.length === 1 && grid[0].length === 1
        const selRange = selectionRange(state.selection)
        const start = state.selection.focus
        // 目标区域：单格剪贴板 + 多格选区 → 平铺整个选区；否则从 focus 按剪贴板尺寸展开
        const target: CellRange =
          oneCell && rangeCellCount(selRange) > 1
            ? selRange
            : {
                sr: start.row,
                sc: start.col,
                er: start.row + grid.length - 1,
                ec: start.col + grid[0].length - 1,
              }
        const { entries, clearSource } = planPaste(payload, text, target, {
          rowCount: sheet.rowCount,
          colCount: sheet.colCount,
        })
        const tr = state.tr
        if (entries.length) tr.setCells(sheetId, entries)
        // cut 移动语义：仅当源在本表才清源（跨表粘贴不清别表数据）
        if (clearSource && payload && payload.sheet === sheetId) tr.clearRange(payload.range)
        payload = null
        if (tr.steps.length) {
          tr.setSelection({
            anchor: { row: target.sr, col: target.sc },
            focus: {
              row: Math.min(target.er, sheet.rowCount - 1),
              col: Math.min(target.ec, sheet.colCount - 1),
            },
          })
          view.dispatch(tr)
        }
        return true
      },
    },
  })
}
