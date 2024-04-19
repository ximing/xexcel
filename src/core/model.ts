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
  numFmt?: string // ECMA-376 格式串（numfmt）；只影响显示，不改 raw
  fontFamily?: string // CSS font-family 串
  fontSize?: number // px，缺省 13
  underline?: boolean
  strikethrough?: boolean
}
export interface Cell { raw: string; style?: CellStyle }
export interface SheetConfig { rowCount: number; colCount: number }

// ---- 自动筛选 ----
export type FilterOp =
  | 'contains' | 'notContains' | 'startsWith' | 'endsWith' // 文本
  | 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'between' // 数值（eq/neq 两域通用）
export interface FilterValuesCriteria {
  type: 'values'
  excluded: string[] // 被排除的显示文本（存排除集：新值默认可见）
}
export interface FilterConditionCriteria {
  type: 'condition'
  field: 'text' | 'num'
  op: FilterOp
  v1: string
  v2?: string // between 上界
}
export type FilterCriteria = FilterValuesCriteria | FilterConditionCriteria
export interface FilterState {
  range: CellRange // 含表头行（sr 行）
  criteria: Record<number, FilterCriteria> // key = 绝对列号
}

interface SheetParts {
  rowCount: number
  colCount: number
  cells: Map<number, Map<number, Cell>>
  rowHeights: Map<number, number>
  colWidths: Map<number, number>
  merges: readonly CellRange[]
  frozenRows: number
  frozenCols: number
  hiddenRows: number[]
  hiddenCols: number[]
  filter?: FilterState
}

export class SheetData {
  readonly rowCount: number
  readonly colCount: number
  readonly merges: readonly CellRange[]
  readonly frozenRows: number
  readonly frozenCols: number
  readonly hiddenRows: number[]
  readonly hiddenCols: number[]
  readonly filter?: FilterState
  private readonly _hiddenRowSet: Set<number>
  private readonly _hiddenColSet: Set<number>
  private readonly _cells: Map<number, Map<number, Cell>>
  private readonly _rowHeights: Map<number, number>
  private readonly _colWidths: Map<number, number>

  private constructor(parts: SheetParts) {
    this.rowCount = parts.rowCount
    this.colCount = parts.colCount
    this.merges = parts.merges
    this.frozenRows = parts.frozenRows
    this.frozenCols = parts.frozenCols
    this.hiddenRows = parts.hiddenRows
    this.hiddenCols = parts.hiddenCols
    this.filter = parts.filter
    this._hiddenRowSet = new Set(parts.hiddenRows)
    this._hiddenColSet = new Set(parts.hiddenCols)
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
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
      hiddenRows: this.hiddenRows,
      hiddenCols: this.hiddenCols,
      filter: this.filter,
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
      frozenRows: 0,
      frozenCols: 0,
      hiddenRows: [],
      hiddenCols: [],
    })
  }

  getCell(row: number, col: number): Cell | undefined {
    return this._cells.get(row)?.get(col)
  }

  rowHeight(row: number): number {
    return this._hiddenRowSet.has(row) ? 0 : this._rowHeights.get(row) ?? DEFAULT_ROW_HEIGHT
  }

  colWidth(col: number): number {
    return this._hiddenColSet.has(col) ? 0 : this._colWidths.get(col) ?? DEFAULT_COL_WIDTH
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

  setMerges(merges: readonly CellRange[]): SheetData {
    return SheetData.fromParts({ ...this._parts, merges: [...merges] })
  }

  // 整体替换隐藏行列数组（结构操作 undo 的 wholesale 恢复用；Set 由构造器从数组重建）
  withHidden(hiddenRows: number[], hiddenCols: number[]): SheetData {
    return SheetData.fromParts({ ...this._parts, hiddenRows, hiddenCols })
  }

  // 命中合并区（含锚点）→ 返回该区域；未命中 → null
  mergeAt(row: number, col: number): CellRange | null {
    for (const m of this.merges) {
      if (row >= m.sr && row <= m.er && col >= m.sc && col <= m.ec) return m
    }
    return null
  }

  setFrozen(rows: number, cols: number): SheetData {
    return SheetData.fromParts({ ...this._parts, frozenRows: rows, frozenCols: cols })
  }

  // 手动隐藏（有序去重）；筛选隐藏不入模型（由 filter 状态实时推导）
  setHidden(axis: 'row' | 'col', indices: number[], hidden: boolean): SheetData {
    const cur = new Set(axis === 'row' ? this.hiddenRows : this.hiddenCols)
    for (const i of indices) {
      if (hidden) cur.add(i); else cur.delete(i)
    }
    const sorted = [...cur].sort((a, b) => a - b)
    return SheetData.fromParts(
      axis === 'row' ? { ...this._parts, hiddenRows: sorted } : { ...this._parts, hiddenCols: sorted },
    )
  }

  // 自动筛选设置/清除（undefined = 清除）
  setFilter(filter: FilterState | undefined): SheetData {
    return SheetData.fromParts({ ...this._parts, filter })
  }

  insertRows(index: number, count: number): SheetData {
    return this.remap('row', index, count, 'insert')
  }

  deleteRows(index: number, count: number): SheetData {
    return this.remap('row', index, count, 'delete')
  }

  insertCols(index: number, count: number): SheetData {
    return this.remap('col', index, count, 'insert')
  }

  deleteCols(index: number, count: number): SheetData {
    return this.remap('col', index, count, 'delete')
  }

  // 物理重索引：cells/宽高/merges 按轴平移；delete 时删除区移除、其后前移。
  private remap(axis: 'row' | 'col', index: number, count: number, mode: 'insert' | 'delete'): SheetData {
    const mapIdx = (x: number): number => {
      if (mode === 'insert') return x >= index ? x + count : x
      if (x >= index && x < index + count) return -1 // 删除区
      return x >= index + count ? x - count : x
    }
    const cells = new Map<number, Map<number, Cell>>()
    for (const [row, rowMap] of this._cells) {
      const nr = axis === 'row' ? mapIdx(row) : row
      if (nr < 0) continue
      const nm = new Map<number, Cell>()
      for (const [col, cell] of rowMap) {
        const nc = axis === 'col' ? mapIdx(col) : col
        if (nc < 0) continue
        nm.set(nc, cell)
      }
      if (nm.size) cells.set(nr, nm)
    }
    const remapSizes = (src: Map<number, number>): Map<number, number> => {
      const out = new Map<number, number>()
      for (const [k, v] of src) {
        const nk = mapIdx(k)
        if (nk >= 0) out.set(nk, v)
      }
      return out
    }
    const merges: CellRange[] = []
    for (const m of this.merges) {
      if (mode === 'insert') {
        merges.push(
          axis === 'row'
            ? { sr: mapIdx(m.sr), sc: m.sc, er: mapIdx(m.er), ec: m.ec }
            : { sr: m.sr, sc: mapIdx(m.sc), er: m.er, ec: mapIdx(m.ec) },
        )
        continue
      }
      // delete：裁剪；完全在删除区内 → 丢弃
      if (axis === 'row') {
        const a = mapIdx(m.sr)
        const b = mapIdx(m.er)
        if (a < 0 && b < 0) continue
        const sr = a < 0 ? index : a
        const er = b < 0 ? index - 1 : b
        if (sr <= er) merges.push({ sr, sc: m.sc, er, ec: m.ec })
      } else {
        const a = mapIdx(m.sc)
        const b = mapIdx(m.ec)
        if (a < 0 && b < 0) continue
        const sc = a < 0 ? index : a
        const ec = b < 0 ? index - 1 : b
        if (sc <= ec) merges.push({ sr: m.sr, sc, er: m.er, ec })
      }
    }
    const delta = mode === 'insert' ? count : -count
    const remapIndices = (src: number[]): number[] => {
      const out = new Set<number>()
      for (const i of src) {
        const ni = mapIdx(i)
        if (ni >= 0) out.add(ni)
      }
      return [...out].sort((a, b) => a - b)
    }
    // 筛选重映射：行轴平移/裁剪 range（表头行被删则整体移除）；列轴另重映射 criteria 键
    const remapFilter = (f: FilterState | undefined): FilterState | undefined => {
      if (!f) return f
      if (axis === 'row') {
        const sr = mapIdx(f.range.sr)
        if (sr < 0) return undefined // 表头行被删 → 筛选整体移除
        const er0 = mapIdx(f.range.er)
        const er = er0 < 0 ? index - 1 : er0 // 删除区下缘裁剪
        if (er < sr) return undefined // 数据区删空
        return { range: { sr, sc: f.range.sc, er, ec: f.range.ec }, criteria: f.criteria }
      }
      const sc = mapIdx(f.range.sc)
      const ec0 = mapIdx(f.range.ec)
      const ec = ec0 < 0 ? index - 1 : ec0
      if (sc < 0 || ec < sc) return undefined
      const criteria: Record<number, FilterCriteria> = {}
      for (const [k, c] of Object.entries(f.criteria)) {
        const nk = mapIdx(Number(k))
        if (nk >= 0) criteria[nk] = c
      }
      return { range: { sr: f.range.sr, sc, er: f.range.er, ec }, criteria }
    }
    return SheetData.fromParts({
      rowCount: this.rowCount + (axis === 'row' ? delta : 0),
      colCount: this.colCount + (axis === 'col' ? delta : 0),
      cells,
      rowHeights: axis === 'row' ? remapSizes(this._rowHeights) : new Map(this._rowHeights),
      colWidths: axis === 'col' ? remapSizes(this._colWidths) : new Map(this._colWidths),
      merges,
      hiddenRows: axis === 'row' ? remapIndices(this.hiddenRows) : [...this.hiddenRows],
      hiddenCols: axis === 'col' ? remapIndices(this.hiddenCols) : [...this.hiddenCols],
      filter: remapFilter(this.filter),
      // 冻结设置：delete 时裁掉落在删除区内的冻结行/列（冻结边界随内容走），insert 与另一轴不动
      frozenRows:
        axis === 'row' && mode === 'delete'
          ? Math.max(0, this.frozenRows - Math.max(0, Math.min(this.frozenRows, index + count) - index))
          : this.frozenRows,
      frozenCols:
        axis === 'col' && mode === 'delete'
          ? Math.max(0, this.frozenCols - Math.max(0, Math.min(this.frozenCols, index + count) - index))
          : this.frozenCols,
    })
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
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
      hiddenRows: this.hiddenRows,
      hiddenCols: this.hiddenCols,
      filter: this.filter,
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
      frozenRows?: number
      frozenCols?: number
      hiddenRows?: number[]
      hiddenCols?: number[]
      filter?: FilterState
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
      frozenRows: j.frozenRows ?? 0,
      frozenCols: j.frozenCols ?? 0,
      hiddenRows: j.hiddenRows ?? [],
      hiddenCols: j.hiddenCols ?? [],
      filter: j.filter,
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
