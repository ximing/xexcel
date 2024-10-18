// src/core/io/xlsx.ts
// Workbook ↔ xlsx 映射（exceljs 装配层）。纯映射无 DOM；exceljs 运行时经 registerExcelJS 注入。
import type ExcelJS from 'exceljs'
import { CellRange, clampRange, parseRange, toA1 } from '../addr'
import {
  Cell, CondFormatRule, FilterOp, FilterState, SheetData, Workbook,
} from '../model'
import { DAY_MS, EPOCH } from '../../formula/date'
import {
  XDiffStyle, cfStyleFromExcel, cfStyleToExcel,
  styleFromExcel, styleToExcelAlignment, styleToExcelBorders, styleToExcelFill, styleToExcelFont,
} from './xlsx-style'

// 导入尺寸下限（值同 src/react/csvImport.ts；core 不能 import react，本地重复）
export const IMPORT_MIN_ROWS = 100
export const IMPORT_MIN_COLS = 26
// values 筛选包含集去重上限：超出则该列退化为无 criteria 的 autoFilter
export const FILTER_VALUES_MAX = 1000
// 导入遍历上限（防整列样式文件撑爆）— Task 4 导入侧使用
export const MAX_IMPORT_ROWS = 20000
export const MAX_IMPORT_COLS = 500
// 行高 pt↔px、列宽 字符↔px（Calibri 11 近似，取整误差容忍）
const PT_PER_PX = 0.75
const COL_CHAR_PX = 7
const COL_PAD_PX = 5

// exceljs 运行时注入点：浏览器经 vendor script + app 层 register；测试直接 register。
// （vite pre-bundle 的 exceljs 在浏览器 load 挂死，故不走模块 import）
let exceljsRef: typeof ExcelJS | null = null
export function registerExcelJS(x: typeof ExcelJS): void {
  exceljsRef = x
}
function requireExcelJS(): typeof ExcelJS {
  if (!exceljsRef) throw new Error('exceljs 未注册（app 启动或测试须先 registerExcelJS）')
  return exceljsRef
}

// ---- 值转换 ----

export function cellValueToExcel(cell: Cell): string | number | boolean | { formula: string } {
  const raw = cell.raw
  if (raw.startsWith('=')) return { formula: raw.slice(1) }
  if (raw === 'TRUE') return true
  if (raw === 'FALSE') return false
  const t = raw.trim()
  if (t !== '') {
    const n = Number(t)
    if (Number.isFinite(n)) return n
  }
  return raw
}

// ---- 筛选导出 ----

const NUM_OP_MAP: Partial<Record<FilterOp, string>> = {
  eq: 'equal',
  neq: 'notEqual',
  gt: 'greaterThan',
  gte: 'greaterThanOrEqual',
  lt: 'lessThan',
  lte: 'lessThanOrEqual',
}

export interface XCustomFilter {
  operator: string
  value: string | number
}
export interface XAutoFilterColumn {
  column: number
  filters?: string[]
  customFilters?: XCustomFilter[]
}
export interface XAutoFilter {
  from: string
  to: string
  filters?: XAutoFilterColumn[]
}

const numOrStr = (v: string): string | number => {
  const n = Number(v)
  return Number.isFinite(n) && v.trim() !== '' ? n : v
}

export function buildAutoFilter(filter: FilterState, sheet: SheetData): XAutoFilter {
  const af: XAutoFilter = {
    from: toA1(filter.range.sr, filter.range.sc),
    to: toA1(filter.range.er, filter.range.ec),
  }
  const cols: XAutoFilterColumn[] = []
  for (const [key, crit] of Object.entries(filter.criteria)) {
    const column = Number(key) - filter.range.sc
    // criteria 键在 range 之外（负值或超出列宽）→ 跳过该列，防负 column 写进 autoFilter
    if (column < 0 || column > filter.range.ec - filter.range.sc) continue
    if (crit.type === 'values') {
      const excluded = new Set(crit.excluded)
      const seen = new Set<string>()
      const included: string[] = []
      let overflow = false
      for (let row = filter.range.sr + 1; row <= filter.range.er; row++) {
        const raw = sheet.getCell(row, Number(key))?.raw ?? ''
        if (raw === '') continue // 空白单元格不进值集合（无 (Blanks) 语义）
        if (seen.has(raw)) continue
        seen.add(raw)
        if (seen.size > FILTER_VALUES_MAX) {
          overflow = true
          break
        }
        if (!excluded.has(raw)) included.push(raw)
      }
      if (overflow) {
        console.warn('xlsx 导出：筛选去重值超过上限，该列省略 criteria', Number(key))
        continue
      }
      cols.push({ column, filters: included })
      continue
    }
    if (crit.field !== 'num') {
      console.warn('xlsx 导出：文本条件筛选不支持，该列省略 criteria', crit.op)
      continue
    }
    if (crit.op === 'between') {
      cols.push({
        column,
        customFilters: [
          { operator: 'greaterThanOrEqual', value: numOrStr(crit.v1) },
          { operator: 'lessThanOrEqual', value: numOrStr(crit.v2 ?? '') },
        ],
      })
      continue
    }
    const operator = NUM_OP_MAP[crit.op]
    if (!operator) {
      console.warn('xlsx 导出：不支持的筛选操作符，该列省略 criteria', crit.op)
      continue
    }
    cols.push({ column, customFilters: [{ operator, value: numOrStr(crit.v1) }] })
  }
  if (cols.length) af.filters = cols
  return af
}

// ---- 条件格式导出 ----

const CF_OP_MAP: Partial<Record<FilterOp, string>> = {
  eq: 'equal',
  neq: 'notEqual',
  gt: 'greaterThan',
  gte: 'greaterThanOrEqual',
  lt: 'lessThan',
  lte: 'lessThanOrEqual',
  between: 'between',
}

export interface XCfRule {
  type: string
  operator?: string
  text?: string
  formulae: string[]
  style?: XDiffStyle
  priority: number
}
export interface XCfBlock {
  ref: string
  rules: XCfRule[]
}

// 单规则单 block（我方规则各自带 range；priority = 数组序 +1）
export function cfRuleToExcel(rule: CondFormatRule, priority: number): XCfBlock | null {
  const ref = `${toA1(rule.range.sr, rule.range.sc)}:${toA1(rule.range.er, rule.range.ec)}`
  const style = cfStyleToExcel(rule.style)
  if (rule.type === 'value') {
    const operator = CF_OP_MAP[rule.op]
    if (!operator) {
      console.warn('xlsx 导出：不支持的 CF 操作符，规则已跳过', rule.op)
      return null
    }
    const formulae = rule.op === 'between' ? [rule.v1, rule.v2 ?? ''] : [rule.v1]
    return { ref, rules: [{ type: 'cellIs', operator, formulae, style, priority }] }
  }
  if (rule.type === 'textContains') {
    const anchor = toA1(rule.range.sr, rule.range.sc)
    const escaped = rule.text.replace(/"/g, '""')
    return {
      ref,
      rules: [{
        type: 'containsText',
        operator: 'containsText',
        text: rule.text,
        formulae: [`NOT(ISERROR(SEARCH("${escaped}",${anchor})))`],
        style,
        priority,
      }],
    }
  }
  // duplicate：exceljs 无 duplicateValues 支持（写入即丢），跳过
  console.warn('xlsx 导出：exceljs 不支持 duplicateValues 条件格式，规则已跳过')
  return null
}

// ---- 装配 ----

function sheetToExcelWS(ws: ExcelJS.Worksheet, sheet: SheetData): void {
  sheet.forEachInRange(sheet.usedRange(), (cell, row, col) => {
    if (!cell) return
    const xc = ws.getCell(row + 1, col + 1)
    xc.value = cellValueToExcel(cell) as ExcelJS.CellValue
    const s = cell.style
    if (!s) return
    const font = styleToExcelFont(s)
    // name 键常驻（值可为 undefined）：exceljs 原样回存对象，缺省字体名交由 Excel 端兜底
    if (font) xc.font = { name: undefined, ...font } as ExcelJS.Font
    const fill = styleToExcelFill(s)
    if (fill) xc.fill = fill as ExcelJS.Fill
    const alignment = styleToExcelAlignment(s)
    if (alignment) xc.alignment = alignment as ExcelJS.Alignment
    const border = styleToExcelBorders(s)
    if (border) xc.border = border as Partial<ExcelJS.Borders>
    if (s.numFmt) xc.numFmt = s.numFmt
  })
  for (const m of sheet.merges) ws.mergeCells(m.sr + 1, m.sc + 1, m.er + 1, m.ec + 1)
  if (sheet.frozenRows || sheet.frozenCols) {
    ws.views = [{ state: 'frozen', xSplit: sheet.frozenCols, ySplit: sheet.frozenRows }]
  }
  // 隐藏行须在有内容后设置（exceljs 不落盘无内容行，空行 hidden 丢失属已知降级）
  for (const i of sheet.hiddenRows) ws.getRow(i + 1).hidden = true
  for (const i of sheet.hiddenCols) ws.getColumn(i + 1).hidden = true
  for (const [i, h] of sheet.customRowHeights) {
    ws.getRow(i + 1).height = Math.round(h * PT_PER_PX * 100) / 100
  }
  for (const [i, w] of sheet.customColWidths) {
    ws.getColumn(i + 1).width = Math.round(((w - COL_PAD_PX) / COL_CHAR_PX) * 100) / 100
  }
  if (sheet.filter) ws.autoFilter = buildAutoFilter(sheet.filter, sheet) as ExcelJS.AutoFilter
  sheet.condFormats.forEach((rule, i) => {
    const block = cfRuleToExcel(rule, i + 1)
    if (block) ws.addConditionalFormatting(block as never)
  })
}

export function workbookToExcelJS(wb: Workbook): ExcelJS.Workbook {
  const ewb = new (requireExcelJS().Workbook)()
  for (const id of wb.order) {
    const ws = ewb.addWorksheet(wb.names.get(id) ?? id)
    sheetToExcelWS(ws, wb.sheet(id))
  }
  ewb.views = [{ activeTab: wb.order.indexOf(wb.active) } as ExcelJS.WorkbookView]
  return ewb
}

// ---- 导入 ----

// exceljs 单元格值的结构镜像（只覆盖我们关心的形状）
type XCellValue =
  | null
  | undefined
  | string
  | number
  | boolean
  | Date
  | { formula?: string; sharedFormula?: string; result?: unknown }
  | { richText: { text: string }[] }
  | { error: string }
  | { text: string; hyperlink?: string }

// 值 → raw 文本；返回 null = 无值（样式可能仍在，由调用侧决定存否）
function excelValueToRaw(value: XCellValue): string | null {
  if (value === null || value === undefined) return null
  if (value instanceof Date) return String((value.getTime() - EPOCH) / DAY_MS)
  if (typeof value === 'string') return value
  if (typeof value === 'number') return String(value)
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if ('richText' in value) return value.richText.map((t) => t.text).join('')
  if ('error' in value) return String(value.error)
  if ('formula' in value && value.formula) return '=' + value.formula
  if ('sharedFormula' in value) {
    // 共享公式从格无公式文本：退化为缓存值
    const res = value.result
    if (typeof res === 'number') return String(res)
    if (typeof res === 'string') return res
    console.warn('xlsx 导入：共享公式从格无缓存值，已跳过')
    return null
  }
  if ('text' in value) return value.text
  return null
}

const CF_OP_REVERSE: Record<string, FilterOp> = {
  equal: 'eq',
  notEqual: 'neq',
  greaterThan: 'gt',
  greaterThanOrEqual: 'gte',
  lessThan: 'lt',
  lessThanOrEqual: 'lte',
  between: 'between',
}

interface XCfRuleIn {
  type: string
  operator?: string
  formulae?: string[]
  text?: string
  style?: XDiffStyle
}

// containsText 回读无 text 字段，从 SEARCH("...",ANCHOR) 提取（"" 反转义）
const SEARCH_RE = /SEARCH\("((?:[^"]|"")*)"/

function cfRuleFromExcel(rule: XCfRuleIn, id: string, range: CellRange): CondFormatRule | null {
  const style = cfStyleFromExcel(rule.style ?? {})
  if (rule.type === 'cellIs') {
    const op = rule.operator ? CF_OP_REVERSE[rule.operator] : undefined
    if (!op) {
      console.warn('xlsx 导入：不支持的 cellIs 操作符，规则已跳过', rule.operator)
      return null
    }
    const out: CondFormatRule = { id, range, type: 'value', op, v1: rule.formulae?.[0] ?? '', style }
    if (op === 'between') out.v2 = rule.formulae?.[1] ?? ''
    return out
  }
  if (rule.type === 'containsText') {
    let text = rule.text
    if (text === undefined) {
      const m = rule.formulae?.[0] ? SEARCH_RE.exec(rule.formulae[0]) : null
      if (!m) {
        console.warn('xlsx 导入：containsText 无法提取文本，规则已跳过')
        return null
      }
      text = m[1].replace(/""/g, '"')
    }
    return { id, range, type: 'textContains', text, style }
  }
  console.warn('xlsx 导入：不支持的条件格式类型，规则已跳过', rule.type)
  return null
}

function sheetFromExcelWS(ws: ExcelJS.Worksheet): SheetData {
  // 内容边界（含 styled 空格；防整列样式文件，截断警告）
  let maxRow = 0
  let maxCol = 0
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > MAX_IMPORT_ROWS) {
      console.warn('xlsx 导入：行数超过上限，已截断', MAX_IMPORT_ROWS)
      return
    }
    maxRow = Math.max(maxRow, rowNumber)
    row.eachCell({ includeEmpty: true }, (_c, colNumber) => {
      if (colNumber <= MAX_IMPORT_COLS) maxCol = Math.max(maxCol, colNumber)
    })
  })
  const rowCount = Math.max(maxRow, IMPORT_MIN_ROWS)
  const colCount = Math.max(maxCol, IMPORT_MIN_COLS)
  let sheet = SheetData.create({ rowCount, colCount })

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > rowCount) return
    row.eachCell({ includeEmpty: true }, (xc, colNumber) => {
      if (colNumber > colCount) return
      const raw = excelValueToRaw(xc.value as XCellValue)
      let style = styleFromExcel({
        font: xc.font as never,
        fill: xc.fill as never,
        alignment: xc.alignment as never,
        border: xc.border as never,
      })
      // numFmt：'General' 视为无；Date 且无 numFmt 时补默认日期格式
      let numFmt = xc.numFmt && xc.numFmt !== 'General' ? xc.numFmt : undefined
      if (xc.value instanceof Date && !numFmt) numFmt = 'yyyy/m/d'
      if (numFmt) style = { ...(style ?? {}), numFmt }
      if (raw === null && !style) return
      sheet = sheet.setCell(rowNumber - 1, colNumber - 1, {
        raw: raw ?? '',
        ...(style ? { style } : {}),
      })
    })
  })

  const clamp = (r: CellRange): CellRange => clampRange(r, rowCount, colCount)
  const merges: CellRange[] = []
  for (const ref of (ws.model.merges as string[] | undefined) ?? []) {
    const r = parseRange(ref)
    if (r) merges.push(clamp(r))
  }
  if (merges.length) sheet = sheet.setMerges(merges)

  const view = (ws.views ?? []).find((v) => v.state === 'frozen') as
    | { xSplit?: number; ySplit?: number }
    | undefined
  if (view && ((view.ySplit ?? 0) > 0 || (view.xSplit ?? 0) > 0)) {
    sheet = sheet.setFrozen(
      Math.min(view.ySplit ?? 0, rowCount - 1),
      Math.min(view.xSplit ?? 0, colCount - 1),
    )
  }

  const hiddenRows: number[] = []
  const hiddenCols: number[] = []
  // 行高/隐藏从 model.rows 读：eachRow(includeEmpty:false) 会跳过只有 height/hidden 的无值行
  const modelRows = (ws.model as { rows?: { number: number; height?: number; hidden?: boolean }[] }).rows ?? []
  for (const r of modelRows) {
    if (r.number > rowCount) continue
    if (r.height) sheet = sheet.setRowHeight(r.number - 1, Math.round(r.height / PT_PER_PX))
    if (r.hidden) hiddenRows.push(r.number - 1)
  }
  // 列宽/隐藏从 model.cols 读：columnCount 只反映内容边界，纯样式列在其外；
  // 无显式宽度的隐藏列带默认 width（isCustomWidth=false），不得当作自定义列宽
  const modelCols = (ws.model as {
    cols?: { min: number; max: number; width?: number; isCustomWidth?: boolean; hidden?: boolean }[]
  }).cols ?? []
  for (const c of modelCols) {
    for (let n = c.min; n <= Math.min(c.max, colCount); n++) {
      if (c.hidden) hiddenCols.push(n - 1)
      if (c.isCustomWidth && c.width) {
        sheet = sheet.setColWidth(n - 1, Math.round(c.width * COL_CHAR_PX + COL_PAD_PX))
      }
    }
  }
  if (hiddenRows.length || hiddenCols.length) sheet = sheet.withHidden(hiddenRows, hiddenCols)

  // 筛选：exceljs 回读只有 ref 字符串 → range-only（criteria 不可得）
  if (ws.autoFilter) {
    const ref = typeof ws.autoFilter === 'string' ? ws.autoFilter : `${ws.autoFilter.from}:${ws.autoFilter.to}`
    const range = parseRange(ref)
    if (range) sheet = sheet.setFilter({ range: clamp(range), criteria: {} })
  }

  const rules: CondFormatRule[] = []
  let cfSeq = 1
  // exceljs 类型未声明 conditionalFormattings，整体强转（同 smoke 测试写法）
  const cfBlocks = (ws.model as never as { conditionalFormattings?: { ref: string; rules: XCfRuleIn[] }[] })
    .conditionalFormattings ?? []
  for (const block of cfBlocks) {
    for (const refPart of String(block.ref).split(/\s+/)) {
      const range = parseRange(refPart)
      if (!range) continue
      for (const rule of block.rules) {
        const mapped = cfRuleFromExcel(rule, `cf${cfSeq}`, clamp(range))
        if (mapped) {
          rules.push(mapped)
          cfSeq++
        }
      }
    }
  }
  if (rules.length) sheet = sheet.setCondFormats(rules)

  return sheet
}

// exceljs workbook → Workbook：sheet id 顺序 s1..sN，active=第一张
export function excelJSToWorkbook(ewb: ExcelJS.Workbook): Workbook {
  if (ewb.worksheets.length === 0) throw new Error('xlsx 中没有工作表')
  let wb = Workbook.create({ rowCount: 1, colCount: 1 })
  ewb.worksheets.forEach((ws, i) => {
    const data = sheetFromExcelWS(ws)
    if (i === 0) {
      wb = wb.setSheet('s1', data)
      wb = wb.renameSheet('s1', ws.name)
    } else {
      wb = wb.addSheet(`s${i + 1}`, data, undefined, ws.name)
    }
  })
  return wb
}

// 二进制 → Workbook 一步入口（FileMenu 用）
export async function parseXlsx(data: Uint8Array): Promise<Workbook> {
  const ewb = new (requireExcelJS().Workbook)()
  await ewb.xlsx.load(data as never)
  return excelJSToWorkbook(ewb)
}
