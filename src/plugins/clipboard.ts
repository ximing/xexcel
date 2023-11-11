// 剪贴板插件。职责：
// - copy/cut（handleCopy）：选区逐行 TSV（evaluator.displayText，\t 列 \n 行）写 text/plain；
//   cut 在闭包记 cutRange（移动语义），copy 清除之
// - paste（handlePaste）：解析 TSV 二维数组，从 focus 起批量 setCell（越界行列丢弃）；
//   选区 >1 格且剪贴板仅 1 格 → 平铺整个选区；cutRange 存在且与目标不同区域 → 清源区域
// 全部经 dispatch transaction，不直接改 doc。
import { CellRange, rangeCellCount, rangesEqual } from '../core/addr'
import { Cell } from '../core/model'
import { EditorViewLike, Plugin } from '../core/plugin'
import { selectionRange } from '../core/selection'
import { evaluatorFor } from '../formula/engine'

export function clipboard(): Plugin {
  let cutRange: CellRange | null = null

  return new Plugin({
    props: {
      handleCopy(view: EditorViewLike, cut: boolean, event: ClipboardEvent): boolean {
        const state = view.state
        const r = selectionRange(state.selection)
        const ev = evaluatorFor(state.doc)
        const sheetId = state.doc.active
        const lines: string[] = []
        for (let row = r.sr; row <= r.er; row++) {
          const cells: string[] = []
          for (let col = r.sc; col <= r.ec; col++) cells.push(ev.displayText(sheetId, row, col))
          lines.push(cells.join('\t'))
        }
        event.clipboardData?.setData('text/plain', lines.join('\n'))
        cutRange = cut ? r : null
        return true // EditorView 侧 preventDefault
      },
      handlePaste(view: EditorViewLike, text: string): boolean {
        if (text === '') return true // 空剪贴板：吞掉事件即可
        const state = view.state
        const sheet = state.activeSheet
        const sheetId = state.doc.active
        const lines = text.replace(/\r\n?/g, '\n').split('\n')
        if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop() // 去掉尾部换行
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
        const entries: { row: number; col: number; cell: Cell | null }[] = []
        for (let r = target.sr; r <= target.er && r < sheet.rowCount; r++) {
          const rowArr = grid[(r - target.sr) % grid.length]
          for (let c = target.sc; c <= target.ec && c < sheet.colCount; c++) {
            const t = rowArr[(c - target.sc) % rowArr.length]
            entries.push({ row: r, col: c, cell: t === '' ? null : { raw: t } })
          }
        }
        const tr = state.tr
        if (entries.length) tr.setCells(sheetId, entries)
        // cut 移动语义：粘贴目标与源不同区域 → 清源区域
        if (cutRange && !rangesEqual(cutRange, target)) tr.clearRange(cutRange)
        cutRange = null
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
