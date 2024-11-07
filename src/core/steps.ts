import { CellRange, normalizeRange, toA1 } from './addr'
import { Cell, CellStyle, CondFormatRule, FilterState, SheetData, SheetId, ValidationRule, Workbook } from './model'

export interface StepResult { ok: boolean; doc?: Workbook; failed?: string }

export abstract class Step {
  abstract apply(doc: Workbook): StepResult
  abstract invert(beforeDoc: Workbook): Step // beforeDoc=本步应用前文档
  abstract toJSON(): unknown
}

export class SetCellsStep extends Step {
  constructor(
    readonly sheet: SheetId,
    readonly entries: ReadonlyArray<{ row: number; col: number; cell: Cell | null }>,
  ) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    // 先全部校验，再统一应用（避免部分生效）
    for (const e of this.entries) {
      if (e.row < 0 || e.row >= data.rowCount || e.col < 0 || e.col >= data.colCount) {
        return { ok: false, failed: `cell out of bounds: ${toA1(e.row, e.col)}` }
      }
    }
    for (const e of this.entries) data = data.setCell(e.row, e.col, e.cell) // 按顺序折叠
    return { ok: true, doc: doc.setSheet(this.sheet, data) }
  }

  invert(beforeDoc: Workbook): Step {
    const data = beforeDoc.sheet(this.sheet)
    return new SetCellsStep(
      this.sheet,
      this.entries.map((e) => ({ row: e.row, col: e.col, cell: data.getCell(e.row, e.col) ?? null })),
    )
  }

  toJSON(): unknown {
    return { type: 'setCells', sheet: this.sheet, entries: this.entries }
  }
}

// PatchStyleStep 的逆：逐格恢复旧 style（null=删除 style 键）
export class RestoreStyleStep extends Step {
  constructor(
    readonly sheet: SheetId,
    readonly entries: ReadonlyArray<{ row: number; col: number; style: CellStyle | null }>,
  ) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    for (const e of this.entries) {
      const cell = data.getCell(e.row, e.col)
      if (!cell) {
        // 格不存在但需恢复非空样式 → 重建纯样式格（如撤销「清空样式导致删格」）
        if (e.style !== null) data = data.setCell(e.row, e.col, { raw: '', style: e.style })
        continue
      }
      if (e.style === null) {
        const next: Cell = { raw: cell.raw }
        data = data.setCell(e.row, e.col, next.raw === '' ? null : next)
      } else {
        data = data.setCell(e.row, e.col, { raw: cell.raw, style: e.style })
      }
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data) }
  }

  invert(beforeDoc: Workbook): Step {
    const data = beforeDoc.sheet(this.sheet)
    return new RestoreStyleStep(
      this.sheet,
      this.entries.map((e) => ({ row: e.row, col: e.col, style: data.getCell(e.row, e.col)?.style ?? null })),
    )
  }

  toJSON(): unknown {
    return { type: 'restoreStyle', sheet: this.sheet, entries: this.entries }
  }
}

export class PatchStyleStep extends Step {
  constructor(
    readonly sheet: SheetId,
    readonly range: CellRange,
    readonly patch: Partial<CellStyle>,
  ) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    const r = normalizeRange(this.range)
    for (let row = r.sr; row <= r.er; row++) {
      for (let col = r.sc; col <= r.ec; col++) {
        if (row < 0 || row >= data.rowCount || col < 0 || col >= data.colCount) {
          return { ok: false, failed: `range out of bounds: ${toA1(row, col)}` }
        }
        const cell = data.getCell(row, col)
        const style: CellStyle = { ...(cell?.style ?? {}) }
        for (const [k, v] of Object.entries(this.patch) as [keyof CellStyle, CellStyle[keyof CellStyle]][]) {
          if (v === undefined) delete style[k]; else (style as Record<string, unknown>)[k] = v
        }
        const raw = cell?.raw ?? ''
        if (Object.keys(style).length === 0 && raw === '') {
          data = data.setCell(row, col, null) // 样式清空且 raw 空 → 删格
        } else {
          data = data.setCell(row, col, Object.keys(style).length === 0 ? { raw } : { raw, style })
        }
      }
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data) }
  }

  invert(beforeDoc: Workbook): Step {
    const data = beforeDoc.sheet(this.sheet)
    const r = normalizeRange(this.range)
    const entries: { row: number; col: number; style: CellStyle | null }[] = []
    for (let row = r.sr; row <= r.er; row++) {
      for (let col = r.sc; col <= r.ec; col++) {
        entries.push({ row, col, style: data.getCell(row, col)?.style ?? null })
      }
    }
    return new RestoreStyleStep(this.sheet, entries)
  }

  toJSON(): unknown {
    return { type: 'patchStyle', sheet: this.sheet, range: this.range, patch: this.patch }
  }
}

export class ResizeStep extends Step {
  constructor(
    readonly sheet: SheetId,
    readonly axis: 'row' | 'col',
    readonly index: number,
    readonly size: number | null, // null=恢复默认
  ) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    const limit = this.axis === 'row' ? data.rowCount : data.colCount
    if (this.index < 0 || this.index >= limit) {
      return { ok: false, failed: `${this.axis} index out of bounds: ${this.index}` }
    }
    data = this.axis === 'row'
      ? data.setRowHeight(this.index, this.size)
      : data.setColWidth(this.index, this.size)
    return { ok: true, doc: doc.setSheet(this.sheet, data) }
  }

  invert(beforeDoc: Workbook): Step {
    const data = beforeDoc.sheet(this.sheet)
    // 旧值：未自定义则记 null（恢复默认）
    const old = this.axis === 'row'
      ? data.customRowHeights.get(this.index) ?? null
      : data.customColWidths.get(this.index) ?? null
    return new ResizeStep(this.sheet, this.axis, this.index, old)
  }

  toJSON(): unknown {
    return { type: 'resize', sheet: this.sheet, axis: this.axis, index: this.index, size: this.size }
  }
}

// 插入新表。data 为新表内容（新建=空表；undo 删表=快照）；
// active 非 null 时插入完成后设为活动表（恢复别表 active 也走此字段）。
export class InsertSheetStep extends Step {
  constructor(
    readonly sheet: SheetId,
    readonly name: string,
    readonly data: SheetData,
    readonly index: number | null, // null = 末尾
    readonly active: SheetId | null,
  ) {
    super()
  }

  apply(doc: Workbook): StepResult {
    if (doc.sheets.has(this.sheet)) return { ok: false, failed: `sheet already exists: ${this.sheet}` }
    const idx = this.index ?? doc.order.length
    if (idx < 0 || idx > doc.order.length) return { ok: false, failed: `index out of bounds: ${this.index}` }
    if (this.active !== null && this.active !== this.sheet && !doc.sheets.has(this.active)) {
      return { ok: false, failed: `active sheet not found: ${this.active}` }
    }
    let d = doc.addSheet(this.sheet, this.data, idx, this.name)
    if (this.active !== null) d = d.setActive(this.active)
    return { ok: true, doc: d }
  }

  invert(beforeDoc: Workbook): Step {
    return new RemoveSheetStep(this.sheet, beforeDoc.active)
  }

  toJSON(): unknown {
    return {
      type: 'insertSheet',
      sheet: this.sheet,
      name: this.name,
      data: this.data.toJSON(),
      index: this.index,
      active: this.active,
    }
  }
}

// 删除表。restoreActive：undo「插入表」时恢复之前的 active（可为 null=不动）。
export class RemoveSheetStep extends Step {
  constructor(readonly sheet: SheetId, readonly restoreActive: SheetId | null = null) {
    super()
  }

  apply(doc: Workbook): StepResult {
    if (!doc.sheets.has(this.sheet)) return { ok: false, failed: `sheet not found: ${this.sheet}` }
    if (doc.order.length <= 1) return { ok: false, failed: 'cannot remove the last sheet' }
    let d = doc.removeSheet(this.sheet)
    if (this.restoreActive !== null && d.sheets.has(this.restoreActive)) {
      d = d.setActive(this.restoreActive)
    }
    return { ok: true, doc: d }
  }

  invert(beforeDoc: Workbook): Step {
    // 快照整张表 + 名称 + 原位置 + 原 active，undo 完整恢复
    return new InsertSheetStep(
      this.sheet,
      beforeDoc.names.get(this.sheet) ?? this.sheet,
      beforeDoc.sheet(this.sheet),
      beforeDoc.order.indexOf(this.sheet),
      beforeDoc.active,
    )
  }

  toJSON(): unknown {
    return { type: 'removeSheet', sheet: this.sheet, restoreActive: this.restoreActive }
  }
}

export class RenameSheetStep extends Step {
  constructor(readonly sheet: SheetId, readonly name: string) {
    super()
  }

  apply(doc: Workbook): StepResult {
    if (!doc.sheets.has(this.sheet)) return { ok: false, failed: `sheet not found: ${this.sheet}` }
    const trimmed = this.name.trim()
    if (trimmed === '') return { ok: false, failed: 'empty sheet name' }
    const lower = trimmed.toLowerCase()
    for (const [id, n] of doc.names) {
      if (id !== this.sheet && n.toLowerCase() === lower) {
        return { ok: false, failed: `duplicate sheet name: ${trimmed}` }
      }
    }
    return { ok: true, doc: doc.renameSheet(this.sheet, trimmed) }
  }

  invert(beforeDoc: Workbook): Step {
    return new RenameSheetStep(this.sheet, beforeDoc.names.get(this.sheet) ?? this.sheet)
  }

  toJSON(): unknown {
    return { type: 'renameSheet', sheet: this.sheet, name: this.name }
  }
}

// 移动工作表位置（标签拖动排序/右键左移右移）
export class MoveSheetStep extends Step {
  constructor(readonly sheet: SheetId, readonly toIndex: number) {
    super()
  }
  apply(doc: Workbook): StepResult {
    if (!doc.sheets.has(this.sheet)) return { ok: false, failed: `sheet not found: ${this.sheet}` }
    if (this.toIndex < 0 || this.toIndex >= doc.order.length) {
      return { ok: false, failed: `index out of bounds: ${this.toIndex}` }
    }
    return { ok: true, doc: doc.moveSheet(this.sheet, this.toIndex) }
  }
  invert(beforeDoc: Workbook): Step {
    return new MoveSheetStep(this.sheet, beforeDoc.order.indexOf(this.sheet))
  }
  toJSON(): unknown {
    return { type: 'moveSheet', sheet: this.sheet, toIndex: this.toIndex }
  }
}

// 切换活动表。调用侧应配 tr.setMeta('addToHistory', false) 不入 undo 栈。
export class SetActiveSheetStep extends Step {
  constructor(readonly sheet: SheetId) {
    super()
  }

  apply(doc: Workbook): StepResult {
    if (!doc.sheets.has(this.sheet)) return { ok: false, failed: `sheet not found: ${this.sheet}` }
    return { ok: true, doc: doc.setActive(this.sheet) }
  }

  invert(beforeDoc: Workbook): Step {
    return new SetActiveSheetStep(beforeDoc.active)
  }

  toJSON(): unknown {
    return { type: 'setActiveSheet', sheet: this.sheet }
  }
}

// 整体替换合并区数组（合并/拆分/结构操作共用）
export class SetMergesStep extends Step {
  constructor(readonly sheet: SheetId, readonly merges: CellRange[]) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data.setMerges(this.merges)) }
  }

  invert(beforeDoc: Workbook): Step {
    return new SetMergesStep(this.sheet, [...beforeDoc.sheet(this.sheet).merges])
  }

  toJSON(): unknown {
    return { type: 'setMerges', sheet: this.sheet, merges: this.merges }
  }
}

// 公式级联注入点：formula/transform 模块加载时注册（保持 core → formula 无依赖）。
// 未注册时 StructureStep 只做结构部分（单测 core 层时可不注入）。
export type StructureCascade = (raw: string, spec: StructureSpecName, hostSheet: string) => string
export interface StructureSpecName {
  sheet: string // 表名
  axis: 'row' | 'col'
  index: number
  count: number
  mode: 'insert' | 'delete'
}
let cascadeFn: StructureCascade | null = null
export function registerStructureCascade(fn: StructureCascade): void {
  cascadeFn = fn
}

export interface StructureStepSpec {
  sheet: SheetId
  axis: 'row' | 'col'
  index: number
  count: number
  mode: 'insert' | 'delete'
}

export interface StructureRestoreEntry {
  sheet: SheetId
  row: number
  col: number
  cell: Cell | null
}

// 逆操作实例的恢复负载：被改公式的原文 + 删除区内的自定义行高/列宽 +
// delete 前目标表的完整 merges 与隐藏行列数组（裁剪后的幸存者无法逐一识别，
// 整体恢复才能保证 undo 恒等；隐藏标记物理丢失，与 merges 同款 wholesale 语义）。
export interface StructureRestore {
  cells: StructureRestoreEntry[]
  sizes: [number, number][] // axis 维度：删除区内 index → 自定义 size
  merges: CellRange[] // delete 模式：目标表完整 merges 原文；insert 模式：空（remap 自身可逆）
  hiddenRows: number[] // delete 模式：目标表完整 hiddenRows 原文；insert 模式：空
  hiddenCols: number[] // 同上
  filter: FilterState | undefined // delete 模式：目标表完整 filter 原文（SheetData 不可变，浅引用即可）
  condFormats: CondFormatRule[] // delete 模式：目标表完整 condFormats 原文；insert 模式：空
  validations: ValidationRule[] // delete 模式：目标表完整 validations 原文；insert 模式：空
}

// 插入/删除行列：物理重索引 + 全簿公式级联（经注入的 cascade）。
// restore 非 null = 逆操作实例：apply = 反向结构 + 公式原文/行高列宽/合并区恢复。
export class StructureStep extends Step {
  constructor(readonly spec: StructureStepSpec, readonly restore: StructureRestore | null = null) {
    super()
  }

  apply(doc: Workbook): StepResult {
    const spec = this.spec
    if (spec.count < 1 || spec.index < 0) return { ok: false, failed: 'bad index/count' }
    let data: SheetData
    try {
      data = doc.sheet(spec.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${spec.sheet}` }
    }
    const limit = spec.axis === 'row' ? data.rowCount : data.colCount
    const mode = this.restore ? (spec.mode === 'insert' ? 'delete' : 'insert') : spec.mode
    // 校验按实际执行方向（逆操作时 mode 已翻转）
    if (mode === 'insert') {
      if (spec.index > limit) return { ok: false, failed: `index out of bounds: ${spec.index}` }
    } else if (spec.index + spec.count > limit) {
      return { ok: false, failed: `delete range out of bounds: ${spec.index}+${spec.count}` }
    }
    const shifted =
      spec.axis === 'row'
        ? mode === 'insert'
          ? data.insertRows(spec.index, spec.count)
          : data.deleteRows(spec.index, spec.count)
        : mode === 'insert'
          ? data.insertCols(spec.index, spec.count)
          : data.deleteCols(spec.index, spec.count)
    let out = doc.setSheet(spec.sheet, shifted)
    if (this.restore) {
      // 恢复负载校验：每个目标格须落在目标表当前边界内（含表存在性），防脏负载越界写
      for (const e of this.restore.cells) {
        const target = out.sheets.get(e.sheet)
        if (!target || e.row < 0 || e.col < 0 || e.row >= target.rowCount || e.col >= target.colCount) {
          return { ok: false, failed: 'restore cell out of bounds' }
        }
      }
      // 逆操作：恢复公式原文与删除区内容
      for (const e of this.restore.cells) {
        out = out.setSheet(e.sheet, out.sheet(e.sheet).setCell(e.row, e.col, e.cell))
      }
      // 恢复删除区内的自定义行高/列宽
      let d = out.sheet(spec.sheet)
      for (const [i, size] of this.restore.sizes) {
        d = spec.axis === 'row' ? d.setRowHeight(i, size) : d.setColWidth(i, size)
      }
      // delete 的 undo：整体恢复 merges 与隐藏行列（裁剪幸存者可能与原文不等，见 StructureRestore 注释）
      if (spec.mode === 'delete') {
        d = d.setMerges(this.restore.merges)
        d = d.withHidden(this.restore.hiddenRows, this.restore.hiddenCols)
        d = d.setFilter(this.restore.filter)
        d = d.setCondFormats(this.restore.condFormats ?? []) // 旧历史 JSON 无此字段
        d = d.setValidations(this.restore.validations ?? []) // 旧历史 JSON 无此字段
      }
      out = out.setSheet(spec.sheet, d)
      return { ok: true, doc: out }
    }
    // 正向：公式级联（未注入 cascade 时跳过）
    if (cascadeFn) {
      const nameSpec: StructureSpecName = {
        sheet: doc.names.get(spec.sheet) ?? spec.sheet,
        axis: spec.axis,
        index: spec.index,
        count: spec.count,
        mode: spec.mode,
      }
      for (const [id, sheetData] of out.sheets) {
        const host = out.names.get(id) ?? id
        const changes: { row: number; col: number; cell: Cell | null }[] = []
        collectFormulaChanges(sheetData, (row, col, cell) => {
          const next = cascadeFn!(cell.raw, nameSpec, host)
          if (next !== cell.raw) changes.push({ row, col, cell: { ...cell, raw: next } })
        })
        if (changes.length) {
          let d = out.sheet(id)
          for (const c of changes) d = d.setCell(c.row, c.col, c.cell)
          out = out.setSheet(id, d)
        }
      }
    }
    return { ok: true, doc: out }
  }

  invert(beforeDoc: Workbook): Step {
    if (this.restore) {
      // 逆操作实例的 invert = 正向实例（cascade 确定性重放）
      return new StructureStep(this.spec, null)
    }
    // 扫描 beforeDoc 全部公式格：级联后有变化 → 记录原文恢复项
    const cells: StructureRestoreEntry[] = []
    const seen = new Set<string>()
    let sizes: [number, number][] = []
    let merges: CellRange[] = []
    let hiddenRows: number[] = []
    let hiddenCols: number[] = []
    let filter: FilterState | undefined
    let condFormats: CondFormatRule[] = []
    let validations: ValidationRule[] = []
    // delete 模式：删除区内的格/行高列宽/隐藏标记物理丢失，原文全部入恢复项（级联只覆盖公式文本）；
    // merges 与隐藏数组记录目标表完整原文（undo 整体恢复）
    if (this.spec.mode === 'delete') {
      const data = beforeDoc.sheet(this.spec.sheet)
      const cross = this.spec.axis === 'row' ? data.colCount : data.rowCount
      const customSizes = this.spec.axis === 'row' ? data.customRowHeights : data.customColWidths
      for (let i = this.spec.index; i < this.spec.index + this.spec.count; i++) {
        const size = customSizes.get(i)
        if (size !== undefined) sizes.push([i, size])
        for (let j = 0; j < cross; j++) {
          const row = this.spec.axis === 'row' ? i : j
          const col = this.spec.axis === 'row' ? j : i
          const cell = data.getCell(row, col)
          if (cell) {
            cells.push({ sheet: this.spec.sheet, row, col, cell })
            seen.add(`${this.spec.sheet}:${row}:${col}`)
          }
        }
      }
      merges = [...data.merges]
      hiddenRows = [...data.hiddenRows]
      hiddenCols = [...data.hiddenCols]
      filter = data.filter
      condFormats = [...data.condFormats]
      validations = [...data.validations]
    }
    if (cascadeFn) {
      const nameSpec: StructureSpecName = {
        sheet: beforeDoc.names.get(this.spec.sheet) ?? this.spec.sheet,
        axis: this.spec.axis,
        index: this.spec.index,
        count: this.spec.count,
        mode: this.spec.mode,
      }
      for (const [id, sheetData] of beforeDoc.sheets) {
        const host = beforeDoc.names.get(id) ?? id
        collectFormulaChanges(sheetData, (row, col, cell) => {
          if (seen.has(`${id}:${row}:${col}`)) return
          if (cascadeFn!(cell.raw, nameSpec, host) !== cell.raw) {
            cells.push({ sheet: id, row, col, cell })
          }
        })
      }
    }
    return new StructureStep(this.spec, { cells, sizes, merges, hiddenRows, hiddenCols, filter, condFormats, validations })
  }

  toJSON(): unknown {
    return { type: 'structure', spec: this.spec, restore: this.restore }
  }
}

// 遍历表中所有公式格（raw 以 = 开头）
function collectFormulaChanges(data: SheetData, cb: (row: number, col: number, cell: Cell) => void): void {
  const r = data.usedRange()
  for (let row = 0; row <= r.er; row++) {
    for (let col = 0; col <= r.ec; col++) {
      const cell = data.getCell(row, col)
      if (cell && cell.raw.startsWith('=')) cb(row, col, cell)
    }
  }
}

// 冻结行列（行/列数，0=不冻结）
export class SetFreezeStep extends Step {
  constructor(readonly sheet: SheetId, readonly rows: number, readonly cols: number) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    if (this.rows < 0 || this.rows >= data.rowCount || this.cols < 0 || this.cols >= data.colCount) {
      return { ok: false, failed: `freeze out of bounds: ${this.rows},${this.cols}` }
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data.setFrozen(this.rows, this.cols)) }
  }

  invert(beforeDoc: Workbook): Step {
    const d = beforeDoc.sheet(this.sheet)
    return new SetFreezeStep(this.sheet, d.frozenRows, d.frozenCols)
  }

  toJSON(): unknown {
    return { type: 'setFreeze', sheet: this.sheet, rows: this.rows, cols: this.cols }
  }
}

// 手动隐藏行列。restore 非 null = 逆操作实例：indices 先全部取消隐藏，再精确恢复 restore 子集，
// 保证 undo 不改动 indices 范围外的状态、且混合前置状态恒等。
export class SetHiddenStep extends Step {
  constructor(
    readonly sheet: SheetId,
    readonly axis: 'row' | 'col',
    readonly indices: number[],
    readonly hidden: boolean,
    readonly restore: number[] | null = null,
  ) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    const limit = this.axis === 'row' ? data.rowCount : data.colCount
    for (const i of this.indices) {
      if (i < 0 || i >= limit) return { ok: false, failed: `${this.axis} index out of bounds: ${i}` }
    }
    if (this.restore) {
      data = data.setHidden(this.axis, this.indices, false)
      data = data.setHidden(this.axis, this.restore, true)
    } else {
      data = data.setHidden(this.axis, this.indices, this.hidden)
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data) }
  }

  invert(beforeDoc: Workbook): Step {
    if (this.restore) return new SetHiddenStep(this.sheet, this.axis, this.indices, this.hidden)
    const data = beforeDoc.sheet(this.sheet)
    const prior = this.axis === 'row' ? data.hiddenRows : data.hiddenCols
    const priorSet = new Set(prior)
    return new SetHiddenStep(
      this.sheet,
      this.axis,
      this.indices,
      this.hidden,
      this.indices.filter((i) => priorSet.has(i)),
    )
  }

  toJSON(): unknown {
    return { type: 'setHidden', sheet: this.sheet, axis: this.axis, indices: this.indices, hidden: this.hidden, restore: this.restore }
  }
}

// 设置/清除自动筛选（filter undefined = 清除）
export class SetFilterStep extends Step {
  constructor(readonly sheet: SheetId, readonly filter: FilterState | undefined) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data.setFilter(this.filter)) }
  }

  invert(beforeDoc: Workbook): Step {
    return new SetFilterStep(this.sheet, beforeDoc.sheet(this.sheet).filter)
  }

  toJSON(): unknown {
    return { type: 'setFilter', sheet: this.sheet, filter: this.filter }
  }
}

// 整体替换条件格式规则（invert=旧值快照）
export class SetCondFormatsStep extends Step {
  constructor(readonly sheet: SheetId, readonly rules: CondFormatRule[]) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data.setCondFormats(this.rules)) }
  }

  invert(beforeDoc: Workbook): Step {
    return new SetCondFormatsStep(this.sheet, [...beforeDoc.sheet(this.sheet).condFormats])
  }

  toJSON(): unknown {
    return { type: 'setCondFormats', sheet: this.sheet, rules: this.rules }
  }
}

// 整体替换数据验证规则（invert=旧值快照）
export class SetValidationsStep extends Step {
  constructor(readonly sheet: SheetId, readonly rules: ValidationRule[]) {
    super()
  }

  apply(doc: Workbook): StepResult {
    let data: SheetData
    try {
      data = doc.sheet(this.sheet)
    } catch {
      return { ok: false, failed: `sheet not found: ${this.sheet}` }
    }
    return { ok: true, doc: doc.setSheet(this.sheet, data.setValidations(this.rules)) }
  }

  invert(beforeDoc: Workbook): Step {
    return new SetValidationsStep(this.sheet, [...beforeDoc.sheet(this.sheet).validations])
  }

  toJSON(): unknown {
    return { type: 'setValidations', sheet: this.sheet, rules: this.rules }
  }
}

export function stepFromJSON(json: any): Step {
  switch (json?.type) {
    case 'setCells':
      return new SetCellsStep(json.sheet, json.entries)
    case 'patchStyle':
      return new PatchStyleStep(json.sheet, json.range, json.patch)
    case 'resize':
      return new ResizeStep(json.sheet, json.axis, json.index, json.size)
    case 'restoreStyle':
      return new RestoreStyleStep(json.sheet, json.entries)
    case 'insertSheet':
      return new InsertSheetStep(json.sheet, json.name, SheetData.fromJSON(json.data), json.index, json.active)
    case 'removeSheet':
      return new RemoveSheetStep(json.sheet, json.restoreActive ?? null)
    case 'renameSheet':
      return new RenameSheetStep(json.sheet, json.name)
    case 'moveSheet':
      return new MoveSheetStep(json.sheet, json.toIndex)
    case 'setActiveSheet':
      return new SetActiveSheetStep(json.sheet)
    case 'setMerges':
      return new SetMergesStep(json.sheet, json.merges)
    case 'structure':
      return new StructureStep(json.spec, json.restore ?? null)
    case 'setFreeze':
      return new SetFreezeStep(json.sheet, json.rows, json.cols)
    case 'setHidden':
      return new SetHiddenStep(json.sheet, json.axis, json.indices, json.hidden, json.restore ?? null)
    case 'setFilter':
      return new SetFilterStep(json.sheet, json.filter)
    case 'setCondFormats':
      return new SetCondFormatsStep(json.sheet, json.rules)
    case 'setValidations':
      return new SetValidationsStep(json.sheet, json.rules)
    default:
      throw new Error(`unknown step type: ${json?.type}`)
  }
}
