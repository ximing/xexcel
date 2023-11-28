import { CellRange } from './addr'

export const ROW_HEADER_WIDTH = 48
export const COL_HEADER_HEIGHT = 26
export const DEFAULT_ROW_HEIGHT = 24
export const DEFAULT_COL_WIDTH = 96

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  color?: string
  bg?: string
  align?: 'left' | 'center' | 'right'
}
export interface Cell { raw: string; style?: CellStyle }
export interface SheetConfig { rowCount: number; colCount: number }

interface SheetParts {
  rowCount: number
  colCount: number
  cells: Map<number, Map<number, Cell>>
  rowHeights: Map<number, number>
  colWidths: Map<number, number>
  merges: readonly CellRange[]
}

export class SheetData {
  readonly rowCount: number
  readonly colCount: number
  readonly merges: readonly CellRange[]
  private readonly _cells: Map<number, Map<number, Cell>>
  private readonly _rowHeights: Map<number, number>
  private readonly _colWidths: Map<number, number>

  private constructor(parts: SheetParts) {
    this.rowCount = parts.rowCount
    this.colCount = parts.colCount
    this.merges = parts.merges
    this._cells = parts.cells
    this._rowHeights = parts.rowHeights
    this._colWidths = parts.colWidths
  }

  private get _parts(): SheetParts {
    return {
      rowCount: this.rowCount,
      colCount: this.colCount,
      cells: this._cells,
      rowHeights: this._rowHeights,
      colWidths: this._colWidths,
      merges: this.merges,
    }
  }

  private static fromParts(parts: SheetParts): SheetData {
    return new SheetData(parts)
  }

  static create(config: SheetConfig): SheetData {
    return new SheetData({
      rowCount: config.rowCount,
      colCount: config.colCount,
      cells: new Map(),
      rowHeights: new Map(),
      colWidths: new Map(),
      merges: [],
    })
  }

  getCell(row: number, col: number): Cell | undefined {
    return this._cells.get(row)?.get(col)
  }

  rowHeight(row: number): number {
    return this._rowHeights.get(row) ?? DEFAULT_ROW_HEIGHT
  }

  colWidth(col: number): number {
    return this._colWidths.get(col) ?? DEFAULT_COL_WIDTH
  }

  get customRowHeights(): ReadonlyMap<number, number> {
    return this._rowHeights
  }

  get customColWidths(): ReadonlyMap<number, number> {
    return this._colWidths
  }

  // 不可变更新：浅拷贝受影响行 Map + 新外层 Map（行级结构共享）
  setCell(row: number, col: number, cell: Cell | null): SheetData {
    // raw==='' 且无 style → 归一为 null
    const normalized = cell && cell.raw === '' && !cell.style ? null : cell
    const cells = new Map(this._cells)
    const rowMap = new Map(cells.get(row) ?? [])
    if (normalized) rowMap.set(col, normalized); else rowMap.delete(col)
    if (rowMap.size) cells.set(row, rowMap); else cells.delete(row)
    return SheetData.fromParts({ ...this._parts, cells })
  }

  setRowHeight(row: number, h: number | null): SheetData {
    const rowHeights = new Map(this._rowHeights)
    if (h === null) rowHeights.delete(row); else rowHeights.set(row, h)
    return SheetData.fromParts({ ...this._parts, rowHeights })
  }

  setColWidth(col: number, w: number | null): SheetData {
    const colWidths = new Map(this._colWidths)
    if (w === null) colWidths.delete(col); else colWidths.set(col, w)
    return SheetData.fromParts({ ...this._parts, colWidths })
  }

  usedRange(): CellRange {
    let sr = Infinity, sc = Infinity, er = -1, ec = -1
    for (const [row, rowMap] of this._cells) {
      if (rowMap.size === 0) continue
      if (row < sr) sr = row
      if (row > er) er = row
      for (const col of rowMap.keys()) {
        if (col < sc) sc = col
        if (col > ec) ec = col
      }
    }
    if (er < 0) return { sr: 0, sc: 0, er: 0, ec: 0 }
    return { sr, sc, er, ec }
  }

  forEachInRange(r: CellRange, cb: (cell: Cell | undefined, row: number, col: number) => void): void {
    for (let row = r.sr; row <= r.er; row++) {
      for (let col = r.sc; col <= r.ec; col++) {
        cb(this.getCell(row, col), row, col)
      }
    }
  }

  toJSON(): unknown {
    const cells: Record<string, Record<string, Cell>> = {}
    for (const [row, rowMap] of this._cells) {
      const cols: Record<string, Cell> = {}
      for (const [col, cell] of rowMap) cols[col] = cell
      cells[row] = cols
    }
    return {
      rowCount: this.rowCount,
      colCount: this.colCount,
      cells,
      rowHeights: [...this._rowHeights.entries()],
      colWidths: [...this._colWidths.entries()],
      merges: this.merges,
    }
  }

  static fromJSON(json: unknown): SheetData {
    const j = json as {
      rowCount: number
      colCount: number
      cells?: Record<string, Record<string, Cell>>
      rowHeights?: [number, number][]
      colWidths?: [number, number][]
      merges?: CellRange[]
    }
    const cells = new Map<number, Map<number, Cell>>()
    for (const [row, cols] of Object.entries(j.cells ?? {})) {
      const rowMap = new Map<number, Cell>()
      for (const [col, cell] of Object.entries(cols)) rowMap.set(Number(col), cell)
      cells.set(Number(row), rowMap)
    }
    return new SheetData({
      rowCount: j.rowCount,
      colCount: j.colCount,
      cells,
      rowHeights: new Map(j.rowHeights ?? []),
      colWidths: new Map(j.colWidths ?? []),
      merges: j.merges ?? [],
    })
  }
}

export type SheetId = string

export class Workbook {
  readonly sheets: ReadonlyMap<SheetId, SheetData>
  readonly order: readonly SheetId[]
  readonly active: SheetId
  readonly names: ReadonlyMap<SheetId, string>

  private constructor(
    sheets: ReadonlyMap<SheetId, SheetData>,
    order: readonly SheetId[],
    active: SheetId,
    names: ReadonlyMap<SheetId, string>,
  ) {
    this.sheets = sheets
    this.order = order
    this.active = active
    this.names = names
  }

  static create(config: SheetConfig): Workbook {
    return new Workbook(
      new Map([['s1', SheetData.create(config)]]),
      ['s1'],
      's1',
      new Map([['s1', 'Sheet1']]),
    )
  }

  sheet(id: SheetId): SheetData {
    const s = this.sheets.get(id)
    if (!s) throw new Error(`sheet not found: ${id}`)
    return s
  }

  get activeSheet(): SheetData {
    return this.sheet(this.active)
  }

  setSheet(id: SheetId, data: SheetData): Workbook {
    const sheets = new Map(this.sheets)
    sheets.set(id, data)
    const order = sheets.has(id) && !this.sheets.has(id) ? [...this.order, id] : this.order
    return new Workbook(sheets, order, this.active, this.names)
  }

  setActive(id: SheetId): Workbook {
    if (!this.sheets.has(id)) throw new Error(`sheet not found: ${id}`)
    return new Workbook(this.sheets, this.order, id, this.names)
  }

  addSheet(id: SheetId, data: SheetData, index?: number, name?: string): Workbook {
    if (this.sheets.has(id)) throw new Error(`sheet already exists: ${id}`)
    const sheets = new Map(this.sheets)
    sheets.set(id, data)
    const order = [...this.order]
    order.splice(index ?? order.length, 0, id)
    const names = new Map(this.names)
    names.set(id, name ?? `Sheet${sheets.size}`)
    return new Workbook(sheets, order, this.active, names)
  }

  removeSheet(id: SheetId): Workbook {
    if (this.order.length <= 1) throw new Error('cannot remove the last sheet')
    const idx = this.order.indexOf(id)
    if (idx < 0) throw new Error(`sheet not found: ${id}`)
    const sheets = new Map(this.sheets)
    sheets.delete(id)
    const order = this.order.filter((s) => s !== id)
    const names = new Map(this.names)
    names.delete(id)
    // 删除 active 时 active 移到相邻（优先后一个，其次前一个）
    const active = this.active === id ? order[Math.min(idx, order.length - 1)] : this.active
    return new Workbook(sheets, order, active, names)
  }

  renameSheet(id: SheetId, name: string): Workbook {
    if (!this.sheets.has(id)) throw new Error(`sheet not found: ${id}`)
    const names = new Map(this.names)
    names.set(id, name)
    return new Workbook(this.sheets, this.order, this.active, names)
  }

  toJSON(): unknown {
    const sheets: Record<string, unknown> = {}
    for (const [id, data] of this.sheets) sheets[id] = data.toJSON()
    return {
      order: [...this.order],
      active: this.active,
      names: [...this.names.entries()],
      sheets,
    }
  }

  static fromJSON(json: unknown): Workbook {
    const j = json as {
      order: SheetId[]
      active: SheetId
      names?: [SheetId, string][]
      sheets: Record<string, unknown>
    }
    const sheets = new Map<SheetId, SheetData>()
    for (const [id, data] of Object.entries(j.sheets)) sheets.set(id, SheetData.fromJSON(data))
    const names = new Map<SheetId, string>(j.names ?? [])
    for (const id of j.order) if (!names.has(id)) names.set(id, id)
    return new Workbook(sheets, j.order, j.active, names)
  }
}

// 新表 id：取现有 's<N>' 形式 id 的最大 N+1
export function nextSheetId(wb: Workbook): SheetId {
  let max = 0
  for (const id of wb.sheets.keys()) {
    const m = /^s(\d+)$/.exec(id)
    if (m) max = Math.max(max, parseInt(m[1], 10))
  }
  return `s${max + 1}`
}

// 新表名：'Sheet<i>'，从表数+1 起跳过已占用名（不区分大小写）
export function nextSheetName(wb: Workbook): string {
  const taken = new Set([...wb.names.values()].map((n) => n.toLowerCase()))
  let i = wb.order.length + 1
  while (taken.has(`sheet${i}`)) i++
  return `Sheet${i}`
}
