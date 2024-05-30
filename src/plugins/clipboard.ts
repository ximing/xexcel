// 剪贴板插件。职责：
// - copy/cut（handleCopy）：选区 TSV（evaluator.displayText）写 text/plain，同时在闭包记
//   ClipboardPayload（源区域 raw 网格 + TSV 指纹）；cut 标记移动语义
// - paste（handlePaste）：剪贴板文本与负载 TSV 一致 → 用 raw 网格落格
//   （copy 按目标偏移公式引用；cut 不偏移=移动语义，不相交则清源）；
//   不一致（外部内容）→ 原 TSV 行为：文本原样落格、空串清格、越界裁剪
// 全部经 dispatch transaction，不直接改 doc。
import { CellRange, clampRange, rangeCellCount, rangesEqual, rangesIntersect } from '../core/addr'
import { Cell, SheetId } from '../core/model'
import { EditorViewLike, Plugin } from '../core/plugin'
import { rangeSelection, selectionRange } from '../core/selection'
import { evaluatorFor } from '../formula/engine'
import { normalizedCell } from '../formula/input'
import { shiftFormula } from '../formula/transform'

export interface ClipboardPayload {
  sheet: SheetId
  range: CellRange
  tsv: string // copy 时写入系统剪贴板的文本（指纹：paste 时据此判定内部粘贴）
  raws: (string | null)[][] // 源区域 raw 网格（null=空格）
  cut: boolean
}

// 注意：text 允许传入未规范化文本（CRLF round-trip），函数内先统一成 \n 再比对指纹
export function planPaste(
  payload: ClipboardPayload | null,
  text: string,
  target: CellRange,
  bounds: { rowCount: number; colCount: number },
): { entries: { row: number; col: number; cell: Cell | null }[]; clearSource: boolean } {
  // 防御性规范化：系统剪贴板可能把 \n 变成 \r\n，不先归一会让指纹比对假阴性
  const normText = text.replace(/\r\n?/g, '\n')
  const er = Math.min(target.er, bounds.rowCount - 1)
  const ec = Math.min(target.ec, bounds.colCount - 1)
  const entries: { row: number; col: number; cell: Cell | null }[] = []

  // 内部粘贴：文本与负载指纹一致 → 用 raw 网格（保留公式）
  if (payload && normText === payload.tsv) {
    const h = payload.raws.length
    const w = payload.raws[0]?.length ?? 1
    for (let r = target.sr; r <= er; r++) {
      const i = (r - target.sr) % h
      for (let c = target.sc; c <= ec; c++) {
        const j = (c - target.sc) % w
        const src = payload.raws[i][j]
        if (src === null || src === '') {
          entries.push({ row: r, col: c, cell: null })
          continue
        }
        // copy 偏移公式引用：delta 按 tile 起点算（r-i / c-j）。
        // 源格在源区域内的相对偏移 (i,j) 已含在 raw 文本里，若按 (r-sr, c-sc)
        // 直接偏移会把源内偏移重复计入（多格 copy 越往下偏得越多）。
        // 平铺时每个 tile 独立按自身起点偏移（与 Excel 一致）；cut 移动语义不偏移
        const raw =
          !payload.cut && src.startsWith('=')
            ? shiftFormula(src, r - i - payload.range.sr, c - j - payload.range.sc)
            : src
        entries.push({ row: r, col: c, cell: { raw } })
      }
    }
    const clearSource =
      payload.cut && !rangesEqual(payload.range, target) && !rangesIntersect(payload.range, target)
    return { entries, clearSource }
  }

  // 外部 TSV：文本原样落格（含 '=' 开头按新公式处理），空串清格，平铺到 target
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
  return { entries, clearSource: false }
}

export function clipboard(): Plugin {
  let payload: ClipboardPayload | null = null

  return new Plugin({
    props: {
      handleCopy(view: EditorViewLike, cut: boolean, event: ClipboardEvent): boolean {
        const state = view.state
        const ev = evaluatorFor(state.doc)
        const sheetId = state.doc.active
        const sheet = state.activeSheet
        const ranges = state.selection.ranges
        // 每个 area 的行；多区域间用空行分隔（仅在 area 之间，不在末尾）
        const areas: string[][] = []
        const raws: (string | null)[][] = []
        for (const r of ranges) {
          const areaLines: string[] = []
          for (let row = r.sr; row <= r.er; row++) {
            const cells: string[] = []
            const rawRow: (string | null)[] = []
            for (let col = r.sc; col <= r.ec; col++) {
              cells.push(ev.displayText(sheetId, row, col))
              rawRow.push(sheet.getCell(row, col)?.raw ?? null)
            }
            areaLines.push(cells.join('\t'))
            raws.push(rawRow)
          }
          areas.push(areaLines)
        }
        const tsv = areas.map(a => a.join('\n')).join('\n\n')
        // 单区域 payload 取活动区（ranges[last]）；多区域 T2 改 areas 结构
        const r = ranges[ranges.length - 1]
        event.clipboardData?.setData('text/plain', tsv)
        payload = { sheet: sheetId, range: r, tsv, raws, cut }
        return true // EditorView 侧 preventDefault
      },
      handlePaste(view: EditorViewLike, text: string): boolean {
        if (text === '') return true // 空剪贴板：吞掉事件即可
        const state = view.state
        const sheet = state.activeSheet
        const sheetId = state.doc.active
        // 规范化一次：指纹比对（planPaste 内）与网格解析统一用 \n 文本，
        // 避免系统剪贴板 CRLF round-trip 后误判为外部内容
        const normText = text.replace(/\r\n?/g, '\n')
        const lines = normText.split('\n')
        if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop()
        const grid = lines.map((l) => l.split('\t'))
        const oneCell = grid.length === 1 && grid[0].length === 1
        const selRange = selectionRange(state.selection)
        const start = state.selection.activeCell
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
        const { entries, clearSource } = planPaste(payload, normText, target, {
          rowCount: sheet.rowCount,
          colCount: sheet.colCount,
        })
        const tr = state.tr
        if (entries.length) tr.setCells(sheetId, entries)
        // cut 移动语义：仅当源在本表才清源（跨表粘贴不清别表数据）
        if (clearSource && payload && payload.sheet === sheetId) tr.clearRange(payload.range)
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
