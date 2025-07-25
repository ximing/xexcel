// 工具栏：Lucide 图标钮 + 四个下拉归并（数字格式/行列/冻结/排序）+ 边框面板。
// 一切写操作回走 view.dispatch / applyStylePatch；alert/confirm 已由 showNotice/askConfirm 替换。
import { useRef, useState } from 'react'
import {
  AlignCenter,
  AlignCenterVertical,
  AlignEndVertical,
  AlignLeft,
  AlignRight,
  AlignStartVertical,
  ArrowUpDown,
  Baseline,
  Bold,
  Filter,
  Grid2x2,
  Hash,
  Italic,
  Merge,
  PaintBucket,
  Paintbrush,
  Redo2,
  Search,
  Sheet,
  ShieldCheck,
  Snowflake,
  Split,
  Strikethrough,
  SwatchBook,
  Underline,
  Undo2,
  WrapText,
} from 'lucide-react'
import { applyStylePatch, mergeSelection, unmergeSelection } from '@gmi/excel-core'
import { computeBorderStyles, BorderPreset } from '@gmi/excel-core'
import { selectionRange } from '@gmi/excel-core'
import { redo, redoDepth, undo, undoDepth } from '@gmi/excel-core'
import type { BorderLineStyle, CellStyle } from '@gmi/excel-core'
import type { SheetState } from '@gmi/excel-core'
import type { EditorView } from '@gmi/excel-view'
import { THEME } from '@gmi/excel-view'
import { findBarKey, formatPainterKey, FormatPainterState } from '@gmi/excel-view'
import { useSheetState } from './bridge'
import { adjustDecimals } from '@gmi/excel-core'
import { evaluatorFor } from '@gmi/excel-core'
import { computeSortEntries, sortBlockedByMerges } from '@gmi/excel-core'
import { showNotice } from './notice'
import { SortDialog } from './SortDialog'
import { CondFormatDialog, canCondFormat } from './CondFormatDialog'
import { ValidationDialog } from './ValidationDialog'
import { canValidation } from './validationRules'
import { FileMenu } from './FileMenu'
import { IconButton } from './ui/IconButton'
import { Separator } from './ui/Separator'
import { Select } from './ui/Select'
import { Dropdown } from './ui/Dropdown'
import { askConfirm } from './ui/confirmStore'
import { buildFreezeItems, buildNumberFormatItems, buildRowColItems, buildSortItems } from './toolbarMenus'

interface Props {
  view: EditorView
}

const FONT_SIZES = [10, 11, 12, 13, 14, 16, 18, 24]
const FONT_FAMILIES = [
  { value: '', label: '默认' },
  { value: 'SimSun, serif', label: '宋体' },
  { value: 'SimHei, sans-serif', label: '黑体' },
  { value: 'monospace', label: '等宽' },
]
const BORDER_LINE_OPTIONS = [
  { value: 'thin', label: '细线' },
  { value: 'medium', label: '中线' },
  { value: 'thick', label: '粗线' },
  { value: 'dashed', label: '虚线' },
  { value: 'dotted', label: '点线' },
  { value: 'double', label: '双线' },
  { value: 'hair', label: '极细' },
  { value: 'mediumDashed', label: '中虚线' },
]

// 边框预设示意小格：outer 为外框类，inner 为是否画内十字线
const BORDER_PRESETS: { p: BorderPreset; label: string; outer: string; inner: boolean }[] = [
  { p: 'none', label: '无框', outer: 'border border-transparent', inner: false },
  { p: 'all', label: '全框', outer: 'border border-ink-2', inner: true },
  { p: 'outer', label: '外框', outer: 'border border-ink-2', inner: false },
  { p: 'inner', label: '内框', outer: 'border border-transparent', inner: true },
  { p: 'top', label: '上框线', outer: 'border-t-2 border-ink-2', inner: false },
  { p: 'bottom', label: '下框线', outer: 'border-b-2 border-ink-2', inner: false },
  { p: 'left', label: '左框线', outer: 'border-l-2 border-ink-2', inner: false },
  { p: 'right', label: '右框线', outer: 'border-r-2 border-ink-2', inner: false },
]

export function Toolbar({ view }: Props) {
  const state = useSheetState(view)
  const { row, col } = state.selection.activeCell
  const active: CellStyle = state.activeSheet.getCell(row, col)?.style ?? {}
  const fp = state.getField(formatPainterKey) as FormatPainterState | null | undefined
  const [showSort, setShowSort] = useState(false)
  const [showCF, setShowCF] = useState(false)
  const [showValidation, setShowValidation] = useState(false)
  // 边框面板应用预设后 remount Dropdown 以关闭面板
  const [borderKey, setBorderKey] = useState(0)
  const [borderLine, setBorderLine] = useState<BorderLineStyle>('thin')
  const [borderColor, setBorderColor] = useState<string>(THEME.ink)
  const colorRef = useRef<HTMLInputElement>(null)
  const bgRef = useRef<HTMLInputElement>(null)
  const textColor = active.color ?? THEME.ink
  const bgColor = active.bg ?? THEME.surface

  const applyBorder = (preset: BorderPreset): void => {
    const sheet = view.state.activeSheet
    const edge = preset === 'none' ? null : { style: borderLine, color: borderColor }
    const entries: ReturnType<typeof computeBorderStyles>[] = []
    for (const r of view.state.selection.ranges) entries.push(computeBorderStyles(sheet, r, preset, edge))
    view.dispatch(view.state.tr.setCellStyles(entries.flat()))
    setBorderKey((k) => k + 1)
    view.focus()
  }

  const quickSort = (asc: boolean): void => {
    // 多区域选区拒绝（按钮已禁用，此处兜底）
    const m = sortRejection(view.state)
    if (m) { showNotice(m); return }
    const r = selectionRange(view.state.selection)
    if (r.sr === r.er) return
    const sheet = view.state.activeSheet
    if (sortBlockedByMerges(sheet, r)) {
      showNotice('排序区域包含合并单元格，无法排序')
      return
    }
    const entries = computeSortEntries(sheet, view.state.doc.active, evaluatorFor(view.state.doc), r, [{ col: r.sc, asc }], false)
    view.dispatch(view.state.tr.setCells(view.state.doc.active, entries))
    view.focus()
  }

  const patch = (p: Partial<CellStyle>): void => {
    applyStylePatch(p)(view.state, (tr) => view.dispatch(tr))
  }
  const patchFocus = (p: Partial<CellStyle>): void => {
    patch(p)
    view.focus()
  }

  const doMerge = async (): Promise<void> => {
    const ranges = view.state.selection.ranges
    if (ranges.every(r => r.sr === r.er && r.sc === r.ec)) return
    let nonEmpty = 0
    const sheet = view.state.activeSheet
    for (const r of ranges) {
      sheet.forEachInRange(r, (cell) => {
        if (cell && cell.raw !== '') nonEmpty++
      })
    }
    if (nonEmpty > 1 && !(await askConfirm({
      title: '合并单元格',
      body: '合并仅保留左上角的值，其余内容将被清除。',
      confirmLabel: '合并',
      danger: true,
    }))) return
    mergeSelection(view.state, (tr) => view.dispatch(tr))
    view.focus()
  }

  const numFmtHandlers = {
    thousands: () => patchFocus({ numFmt: '#,##0.00' }),
    percent: () => patchFocus({ numFmt: '0%' }),
    currency: () => patchFocus({ numFmt: '¥#,##0.00' }),
    decInc: () => patchFocus({ numFmt: adjustDecimals(active.numFmt, 1) }),
    decDec: () => patchFocus({ numFmt: adjustDecimals(active.numFmt, -1) }),
  }

  const rowColHandlers = {
    insertRow: () => {
      const r = selectionRange(view.state.selection)
      view.dispatch(view.state.tr.structure('row', r.sr, r.er - r.sr + 1, 'insert'))
      view.focus()
    },
    deleteRow: () => {
      const r = selectionRange(view.state.selection)
      view.dispatch(view.state.tr.structure('row', r.sr, r.er - r.sr + 1, 'delete'))
      view.focus()
    },
    insertCol: () => {
      const r = selectionRange(view.state.selection)
      view.dispatch(view.state.tr.structure('col', r.sc, r.ec - r.sc + 1, 'insert'))
      view.focus()
    },
    deleteCol: () => {
      const r = selectionRange(view.state.selection)
      view.dispatch(view.state.tr.structure('col', r.sc, r.ec - r.sc + 1, 'delete'))
      view.focus()
    },
    hideRow: () => {
      const r = selectionRange(view.state.selection)
      const indices: number[] = []
      for (let i = r.sr; i <= r.er; i++) indices.push(i)
      view.dispatch(view.state.tr.setHidden('row', indices, true))
      view.focus()
    },
    hideCol: () => {
      const r = selectionRange(view.state.selection)
      const indices: number[] = []
      for (let i = r.sc; i <= r.ec; i++) indices.push(i)
      view.dispatch(view.state.tr.setHidden('col', indices, true))
      view.focus()
    },
    unhide: () => {
      const r = selectionRange(view.state.selection)
      const sheet = view.state.activeSheet
      const rows = sheet.hiddenRows.filter((i) => i >= r.sr && i <= r.er)
      const cols = sheet.hiddenCols.filter((i) => i >= r.sc && i <= r.ec)
      const tr = view.state.tr
      if (rows.length) tr.setHidden('row', rows, false)
      if (cols.length) tr.setHidden('col', cols, false)
      view.dispatch(tr)
      view.focus()
    },
    resetSize: () => {
      const r = selectionRange(view.state.selection)
      const sheet = view.state.activeSheet
      const tr = view.state.tr
      if (r.sc === 0 && r.ec === sheet.colCount - 1) {
        for (let i = r.sr; i <= r.er; i++) tr.resize('row', i, null)
      } else if (r.sr === 0 && r.er === sheet.rowCount - 1) {
        for (let i = r.sc; i <= r.ec; i++) tr.resize('col', i, null)
      } else {
        return
      }
      view.dispatch(tr)
      view.focus()
    },
  }

  const freezeHandlers = {
    freezeRow: () => {
      view.dispatch(view.state.tr.setFrozen(state.activeSheet.frozenRows === 1 && state.activeSheet.frozenCols === 0 ? 0 : 1, state.activeSheet.frozenCols))
      view.focus()
    },
    freezeCol: () => {
      view.dispatch(view.state.tr.setFrozen(state.activeSheet.frozenRows, state.activeSheet.frozenCols === 1 && state.activeSheet.frozenRows === 0 ? 0 : 1))
      view.focus()
    },
    freezeTo: () => {
      const { row: ar, col: ac } = view.state.selection.activeCell
      view.dispatch(view.state.tr.setFrozen(ar, ac))
      view.focus()
    },
    unfreeze: () => {
      view.dispatch(view.state.tr.setFrozen(0, 0))
      view.focus()
    },
  }

  const sortHandlers = {
    asc: () => quickSort(true),
    desc: () => quickSort(false),
    custom: () => setShowSort(true),
  }

  return (
    <div className="flex h-10 items-center gap-0.5 border-b border-line bg-surface px-2">
      <FileMenu view={view} />
      <Separator />
      <IconButton
        icon={Undo2}
        tip="撤销"
        kbd="Ctrl+Z"
        disabled={undoDepth(state) === 0}
        onClick={() => {
          undo(view.state, (tr) => view.dispatch(tr))
          view.focus()
        }}
      />
      <IconButton
        icon={Redo2}
        tip="重做"
        kbd="Ctrl+Y"
        disabled={redoDepth(state) === 0}
        onClick={() => {
          redo(view.state, (tr) => view.dispatch(tr))
          view.focus()
        }}
      />
      <Separator />
      <IconButton
        icon={Paintbrush}
        tip="格式刷（单击刷一次，双击锁定连刷，Esc 解除）"
        active={!!fp}
        onClick={(e) => {
          if (e.detail === 2) {
            const src = { ...(view.state.activeSheet.getCell(row, col)?.style ?? {}) }
            view.dispatch(view.state.tr.setMeta(formatPainterKey, { style: src, locked: true }).setMeta('addToHistory', false))
          } else if (fp) {
            view.dispatch(view.state.tr.setMeta(formatPainterKey, null).setMeta('addToHistory', false))
          } else {
            const src = { ...(view.state.activeSheet.getCell(row, col)?.style ?? {}) }
            view.dispatch(view.state.tr.setMeta(formatPainterKey, { style: src, locked: false }).setMeta('addToHistory', false))
          }
          view.focus()
        }}
      />
      <Separator />
      <IconButton icon={Bold} tip="加粗" active={!!active.bold} onClick={() => patchFocus(active.bold ? { bold: undefined } : { bold: true })} />
      <IconButton icon={Italic} tip="斜体" active={!!active.italic} onClick={() => patchFocus(active.italic ? { italic: undefined } : { italic: true })} />
      <IconButton icon={Underline} tip="下划线" active={!!active.underline} onClick={() => patchFocus(active.underline ? { underline: undefined } : { underline: true })} />
      <IconButton icon={Strikethrough} tip="删除线" active={!!active.strikethrough} onClick={() => patchFocus(active.strikethrough ? { strikethrough: undefined } : { strikethrough: true })} />
      <Select
        tip="字号"
        value={String(active.fontSize ?? 13)}
        options={FONT_SIZES.map((s) => ({ value: String(s), label: String(s) }))}
        onChange={(v) => patchFocus({ fontSize: Number(v) })}
      />
      <Select
        tip="字体"
        value={active.fontFamily ?? ''}
        options={FONT_FAMILIES}
        onChange={(v) => patchFocus({ fontFamily: v || undefined })}
      />
      <span className="relative inline-flex">
        <IconButton icon={Baseline} tip="文字颜色" onClick={() => colorRef.current?.click()} />
        <span className="pointer-events-none absolute inset-x-1.5 bottom-1 h-0.5 rounded-full" style={{ background: textColor }} />
        <input
          ref={colorRef}
          type="color"
          aria-label="文字颜色"
          className="sr-only"
          value={textColor}
          onChange={(e) => patchFocus({ color: e.target.value })}
        />
      </span>
      <span className="relative inline-flex">
        <IconButton icon={PaintBucket} tip="背景颜色" onClick={() => bgRef.current?.click()} />
        <span className="pointer-events-none absolute inset-x-1.5 bottom-1 h-0.5 rounded-full" style={{ background: bgColor }} />
        <input
          ref={bgRef}
          type="color"
          aria-label="背景颜色"
          className="sr-only"
          value={bgColor}
          onChange={(e) => patchFocus({ bg: e.target.value })}
        />
      </span>
      <Separator />
      {([['left', AlignLeft, '左对齐'], ['center', AlignCenter, '居中'], ['right', AlignRight, '右对齐']] as const).map(([a, icon, tip]) => (
        <IconButton key={a} icon={icon} tip={tip} active={(active.align ?? 'left') === a} onClick={() => patchFocus({ align: a })} />
      ))}
      <IconButton icon={WrapText} tip="自动换行" active={!!active.wrap} onClick={() => patchFocus({ wrap: active.wrap ? undefined : true })} />
      <Separator />
      {([['top', AlignStartVertical, '顶端对齐'], ['middle', AlignCenterVertical, '垂直居中'], ['bottom', AlignEndVertical, '底端对齐']] as const).map(([v, icon, tip]) => (
        <IconButton key={v} icon={icon} tip={tip} active={(active.vAlign ?? 'bottom') === v} onClick={() => patchFocus({ vAlign: v })} />
      ))}
      <Separator />
      <Dropdown
        trigger={(open, toggle) => <IconButton icon={Hash} tip="数字格式" active={open} onClick={toggle} />}
        entries={buildNumberFormatItems(active.numFmt, numFmtHandlers)}
      />
      <Separator />
      <IconButton icon={Merge} tip="合并单元格" onClick={() => void doMerge()} />
      <IconButton
        icon={Split}
        tip="拆分单元格"
        onClick={() => {
          unmergeSelection(view.state, (tr) => view.dispatch(tr))
          view.focus()
        }}
      />
      <Dropdown key={borderKey} trigger={(open, toggle) => <IconButton icon={Grid2x2} tip="边框" active={open} onClick={toggle} />}>
        <div className="flex w-56 flex-col gap-2 rounded-md border border-line-strong bg-surface p-2 shadow-2">
          <div className="grid grid-cols-4 gap-1">
            {BORDER_PRESETS.map(({ p, label, outer, inner }) => (
              <button
                key={p}
                type="button"
                aria-label={label}
                title={label}
                className="flex h-8 w-full items-center justify-center rounded-sm hover:bg-hover"
                onClick={() => applyBorder(p)}
              >
                <span className={`h-7 w-7 rounded-sm ${outer}`}>
                  {inner && (
                    <span className="grid h-full w-full grid-cols-2 grid-rows-2">
                      <span className="border-r border-b border-ink-2" />
                      <span className="border-b border-ink-2" />
                      <span className="border-r border-ink-2" />
                      <span />
                    </span>
                  )}
                </span>
              </button>
            ))}
          </div>
          <div className="flex items-center gap-2">
            <Select tip="线型" value={borderLine} options={BORDER_LINE_OPTIONS} onChange={(v) => setBorderLine(v as BorderLineStyle)} />
            <input
              type="color"
              aria-label="边框颜色"
              className="h-7 w-9 cursor-pointer rounded-md border border-line bg-surface p-0.5"
              value={borderColor}
              onChange={(e) => setBorderColor(e.target.value)}
            />
          </div>
        </div>
      </Dropdown>
      <Separator />
      <Dropdown
        trigger={(open, toggle) => <IconButton icon={Sheet} tip="行列操作" active={open} onClick={toggle} />}
        entries={buildRowColItems({ fullRow: isFullRowSel(state), fullCol: isFullColSel(state), canUnhide: hasHiddenInSel(state), canReset: canResetSize(state) }, rowColHandlers)}
      />
      <Dropdown
        trigger={(open, toggle) => <IconButton icon={Snowflake} tip="冻结" active={open} onClick={toggle} />}
        entries={buildFreezeItems({ rows: state.activeSheet.frozenRows, cols: state.activeSheet.frozenCols }, freezeHandlers)}
      />
      <Separator />
      <Dropdown
        trigger={(open, toggle) => <IconButton icon={ArrowUpDown} tip="排序" active={open} onClick={toggle} />}
        entries={buildSortItems(canSort(state), sortHandlers)}
      />
      {showSort && state.selection.ranges.length === 1 && <SortDialog view={view} range={selectionRange(view.state.selection)} onClose={() => setShowSort(false)} />}
      <IconButton
        icon={Filter}
        tip="自动筛选（对选区启用/清除全表筛选）"
        active={!!state.activeSheet.filter}
        disabled={!canFilter(state)}
        onClick={() => {
          const fm = filterRejection(view.state)
          if (fm) { showNotice(fm); return }
          const sheet = view.state.activeSheet
          if (sheet.filter) {
            view.dispatch(view.state.tr.setFilter(undefined))
          } else {
            const r = selectionRange(view.state.selection)
            const range = r.sr === r.er && r.sc === r.ec ? sheet.usedRange() : r
            if (range.er <= range.sr) {
              showNotice('筛选区域至少需要两行（表头 + 数据）')
              return
            }
            view.dispatch(view.state.tr.setFilter({ range, criteria: {} }))
          }
          view.focus()
        }}
      />
      <IconButton
        icon={Search}
        tip="查找/替换"
        kbd="Ctrl+F"
        onClick={() => {
          view.dispatch(view.state.tr.setMeta(findBarKey, true).setMeta('addToHistory', false))
        }}
      />
      <Separator />
      <IconButton icon={SwatchBook} tip="条件格式" disabled={!canCondFormat(state)} onClick={() => setShowCF(true)} />
      {showCF && <CondFormatDialog view={view} onClose={() => setShowCF(false)} />}
      <IconButton icon={ShieldCheck} tip="数据验证" disabled={!canValidation(state)} onClick={() => setShowValidation(true)} />
      {showValidation && <ValidationDialog view={view} onClose={() => setShowValidation(false)} />}
    </div>
  )
}

function isFullRowSel(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  return r.sc === 0 && r.ec === state.activeSheet.colCount - 1 && !(r.sr === 0 && r.er === state.activeSheet.rowCount - 1 && r.sc === 0 && r.ec === state.activeSheet.colCount - 1)
}

function isFullColSel(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  return r.sr === 0 && r.er === state.activeSheet.rowCount - 1 && !(r.sr === 0 && r.er === state.activeSheet.rowCount - 1 && r.sc === 0 && r.ec === state.activeSheet.colCount - 1)
}

// 重置准入：选区覆盖全部行或全部列（含全表选区）——与 resetSize 处理函数两分支准入一致
function canResetSize(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  const sheet = state.activeSheet
  return (r.sc === 0 && r.ec === sheet.colCount - 1) || (r.sr === 0 && r.er === sheet.rowCount - 1)
}

function hasHiddenInSel(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  const sheet = state.activeSheet
  return (
    sheet.hiddenRows.some((i) => i >= r.sr && i <= r.er) ||
    sheet.hiddenCols.some((i) => i >= r.sc && i <= r.ec)
  )
}

// 排序可用：多行（er>sr）且单区域；多区域禁用按钮
export function canSort(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  return r.er > r.sr && state.selection.ranges.length === 1
}

// 多区域选区下筛选按钮禁用（单区域零回归：原按钮恒启用，单区域仍启用）
export function canFilter(state: SheetState): boolean {
  return state.selection.ranges.length === 1
}

// 多区域触发排序/筛选的拒绝消息；单区域放行返回 null。按钮 disabled 与触发拒绝共用判定。
export function sortRejection(state: SheetState): string | null {
  return state.selection.ranges.length > 1 ? '排序仅支持单区域选择' : null
}

export function filterRejection(state: SheetState): string | null {
  return state.selection.ranges.length > 1 ? '筛选仅支持单区域选择' : null
}
