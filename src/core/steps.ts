import { CellRange, normalizeRange, toA1 } from './addr'
import { Cell, CellStyle, SheetData, SheetId, Workbook } from './model'

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
    default:
      throw new Error(`unknown step type: ${json?.type}`)
  }
}
