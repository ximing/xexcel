// src/core/io/xlsx.ts
// Workbook ↔ xlsx 映射（exceljs 装配层）。纯映射无 DOM：vitest 走 node 入口，Vite 经 browser 字段走 dist bundle。
import ExcelJS from 'exceljs'
import { toA1 } from '../addr'
import {
  Cell, CondFormatRule, FilterOp, FilterState, SheetData, Workbook,
} from '../model'
import {
  XDiffStyle, cfStyleToExcel,
  styleToExcelAlignment, styleToExcelBorders, styleToExcelFill, styleToExcelFont,
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
  const ewb = new ExcelJS.Workbook()
  for (const id of wb.order) {
    const ws = ewb.addWorksheet(wb.names.get(id) ?? id)
    sheetToExcelWS(ws, wb.sheet(id))
  }
  ewb.views = [{ activeTab: wb.order.indexOf(wb.active) } as ExcelJS.WorkbookView]
  return ewb
}
