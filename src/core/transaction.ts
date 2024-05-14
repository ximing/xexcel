import { CellRange, normalizeRange } from './addr'
import { Cell, CellStyle, CondFormatRule, FilterState, SheetConfig, SheetData, SheetId, Workbook } from './model'
import { Selection } from './selection'
import { PatchStyleStep, ResizeStep, RestoreStyleStep, SetCellsStep, SetCondFormatsStep, SetFilterStep, SetFreezeStep, SetHiddenStep, SetMergesStep, Step, StructureStep } from './steps'
import { InsertSheetStep, MoveSheetStep, RemoveSheetStep, RenameSheetStep, SetActiveSheetStep } from './steps'
import type { PluginKey } from './plugin'
import type { SheetState } from './state'

export class Transaction {
  readonly before: SheetState
  readonly steps: Step[] = []
  readonly docs: Workbook[] = [] // docs[i] = 第 i 步之后的 doc
  selection: Selection | null = null
  scrolledIntoView = false
  private readonly meta: Map<any, unknown> = new Map()

  constructor(before: SheetState) {
    this.before = before
  }

  get doc(): Workbook {
    return this.docs.length ? this.docs[this.docs.length - 1] : this.before.doc
  }

  // 内部 API：立即试应用并记录 doc 链（history 构造反向事务也用它）
  _pushStep(step: Step): this {
    const r = step.apply(this.doc)
    if (!r.ok) throw new Error('step failed: ' + r.failed)
    this.steps.push(step)
    this.docs.push(r.doc!)
    return this
  }

  private get activeSheetId(): SheetId {
    return this.before.doc.active
  }

  setCell(row: number, col: number, raw: string, style?: CellStyle): this {
    const cell: Cell | null = raw === '' && style == null ? null : style ? { raw, style } : { raw }
    return this._pushStep(new SetCellsStep(this.activeSheetId, [{ row, col, cell }]))
  }

  setCells(sheet: SheetId, entries: ReadonlyArray<{ row: number; col: number; cell: Cell | null }>): this {
    return this._pushStep(new SetCellsStep(sheet, entries))
  }

  patchStyle(range: CellRange, patch: Partial<CellStyle>): this {
    return this._pushStep(new PatchStyleStep(this.activeSheetId, range, patch))
  }

  // 逐格整体替换 style（null=删 style 键）；边框预设与格式刷共用（RestoreStyleStep 即此语义）
  setCellStyles(entries: ReadonlyArray<{ row: number; col: number; style: CellStyle | null }>): this {
    return this._pushStep(new RestoreStyleStep(this.activeSheetId, entries))
  }

  resize(axis: 'row' | 'col', index: number, size: number | null): this {
    return this._pushStep(new ResizeStep(this.activeSheetId, axis, index, size))
  }

  clearRange(range: CellRange): this {
    const r = normalizeRange(range)
    const sheet = this.activeSheetId
    const data = this.doc.sheet(sheet)
    const entries: { row: number; col: number; cell: null }[] = []
    for (let row = r.sr; row <= r.er; row++) {
      for (let col = r.sc; col <= r.ec; col++) {
        if (data.getCell(row, col)) entries.push({ row, col, cell: null })
      }
    }
    if (entries.length === 0) return this
    return this._pushStep(new SetCellsStep(sheet, entries))
  }

  // 新建空表并设为 active（config 为新表行列数，通常取当前表尺寸）
  insertSheet(id: SheetId, name: string, config: SheetConfig): this {
    return this._pushStep(new InsertSheetStep(id, name, SheetData.create(config), null, id))
  }

  removeSheet(id: SheetId): this {
    return this._pushStep(new RemoveSheetStep(id))
  }

  renameSheet(id: SheetId, name: string): this {
    return this._pushStep(new RenameSheetStep(id, name))
  }

  // 移动表位置（标签拖动排序/右键左移右移）
  moveSheet(id: SheetId, toIndex: number): this {
    return this._pushStep(new MoveSheetStep(id, toIndex))
  }

  // 调用侧配 setMeta('addToHistory', false)；并自行 setSelection 防止选区越界
  setActiveSheet(id: SheetId): this {
    return this._pushStep(new SetActiveSheetStep(id))
  }

  setMerges(merges: CellRange[]): this {
    return this._pushStep(new SetMergesStep(this.activeSheetId, merges))
  }

  setFrozen(rows: number, cols: number): this {
    return this._pushStep(new SetFreezeStep(this.activeSheetId, rows, cols))
  }

  setHidden(axis: 'row' | 'col', indices: number[], hidden: boolean): this {
    return this._pushStep(new SetHiddenStep(this.activeSheetId, axis, indices, hidden))
  }

  // 设置/清除自动筛选（undefined = 清除）
  setFilter(filter: FilterState | undefined): this {
    return this._pushStep(new SetFilterStep(this.activeSheetId, filter))
  }

  // 整体替换条件格式规则
  setCondFormats(rules: CondFormatRule[]): this {
    return this._pushStep(new SetCondFormatsStep(this.activeSheetId, rules))
  }

  // 插入/删除当前表的行列（公式级联由 StructureStep 内处理）
  structure(axis: 'row' | 'col', index: number, count: number, mode: 'insert' | 'delete'): this {
    return this._pushStep(new StructureStep({ sheet: this.activeSheetId, axis, index, count, mode }, null))
  }

  setSelection(sel: Selection): this {
    this.selection = sel
    return this
  }

  scrollIntoView(): this {
    this.scrolledIntoView = true
    return this
  }

  setMeta(key: string | PluginKey, value: unknown): this {
    this.meta.set(key, value)
    return this
  }

  getMeta(key: string | PluginKey): unknown {
    return this.meta.get(key)
  }
}
