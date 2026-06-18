// 演示工作簿：三张表。销售表 A1 仍为「产品」，兼容既有 e2e。
import {
  SheetData,
  Workbook,
  type CellStyle,
  type CondFormatRule,
  type SheetState,
  type ValidationRule,
} from '@xexcel/core'
import { createStateFromWorkbook } from '@xexcel/react'

const HEADER: CellStyle = { bold: true, bg: '#e8f0fe', align: 'center', vAlign: 'middle' }
const TITLE: CellStyle = { bold: true, fontSize: 18, color: '#1e3a5f' }
const MUTED: CellStyle = { color: '#64748b', wrap: true }
const TOTAL: CellStyle = { bold: true, bg: '#f1f5f9' }
const MONEY: CellStyle = { numFmt: '#,##0.00', align: 'right' }
const QTY: CellStyle = { numFmt: '#,##0', align: 'right' }
const PCT: CellStyle = { numFmt: '0.0%', align: 'right' }
const THIN = { style: 'thin' as const, color: '#cbd5e1' }

const SALES: [string, number, number][] = [
  ['笔记本电脑', 5999, 12],
  ['显示器', 1299, 8],
  ['平板', 2499, 5],
  ['键盘', 199, 30],
  ['耳机', 299, 25],
  ['音箱', 399, 10],
  ['鼠标', 79, 45],
  ['充电器', 49, 15],
  ['数据线', 19, 60],
  ['手机壳', 29, 20],
]

function set(sheet: SheetData, row: number, col: number, raw: string, style?: CellStyle): SheetData {
  return sheet.setCell(row, col, style ? { raw, style } : { raw })
}

function buildSales(): SheetData {
  let s = SheetData.create({ rowCount: 1000, colCount: 26 })
  const headers = ['产品', '单价', '数量', '金额', '占比', '状态']
  headers.forEach((h, col) => {
    s = set(s, 0, col, h, HEADER)
  })

  SALES.forEach(([name, price, qty], i) => {
    const r = i + 1
    const a1 = r + 1
    s = set(s, r, 0, name)
    s = set(s, r, 1, String(price), MONEY)
    s = set(s, r, 2, String(qty), QTY)
    s = set(s, r, 3, `=B${a1}*C${a1}`, { ...MONEY, bold: true })
    s = set(s, r, 4, `=IF($D$12=0,"",D${a1}/$D$12)`, PCT)
    s = set(s, r, 5, `=IF(D${a1}>=1000,"合格","未达标")`, { align: 'center' })
  })

  const totalRow = 11
  s = set(s, totalRow, 0, '合计', TOTAL)
  s = set(s, totalRow, 1, '', TOTAL)
  s = set(s, totalRow, 2, '=SUM(C2:C11)', { ...QTY, ...TOTAL })
  s = set(s, totalRow, 3, '=SUM(D2:D11)', { ...MONEY, ...TOTAL })
  s = set(s, totalRow, 4, '=SUM(E2:E11)', { ...PCT, ...TOTAL })
  s = set(s, totalRow, 5, '', TOTAL)

  s = s.setColWidth(0, 120)
  s = s.setColWidth(1, 100)
  s = s.setColWidth(2, 80)
  s = s.setColWidth(3, 110)
  s = s.setColWidth(4, 80)
  s = s.setColWidth(5, 90)
  s = s.setFrozen(1, 0)
  s = s.setFilter({ range: { sr: 0, sc: 0, er: 10, ec: 5 }, criteria: {} })

  const cf: CondFormatRule[] = [
    {
      id: 'cf-amt',
      range: { sr: 1, sc: 3, er: 10, ec: 3 },
      type: 'value',
      op: 'gte',
      v1: '1000',
      style: { bold: true, bg: '#dcfce7' },
    },
    {
      id: 'cf-ok',
      range: { sr: 1, sc: 5, er: 10, ec: 5 },
      type: 'textContains',
      text: '合格',
      style: { color: '#166534', bold: true },
    },
    {
      id: 'cf-bad',
      range: { sr: 1, sc: 5, er: 10, ec: 5 },
      type: 'textContains',
      text: '未达标',
      style: { color: '#b91c1c' },
    },
  ]
  s = s.setCondFormats(cf)
  return s
}

function buildStyles(): SheetData {
  let s = SheetData.create({ rowCount: 1000, colCount: 26 })
  s = set(s, 0, 0, '样式橱窗', TITLE)
  s = s.setMerges([{ sr: 0, sc: 0, er: 0, ec: 5 }])
  s = s.setRowHeight(0, 36)
  s = set(s, 1, 0, '字体、对齐、边框、合并、换行都走同一套 CellStyle。下面每格是一种样式。', MUTED)
  s = s.setMerges([...s.merges, { sr: 1, sc: 0, er: 1, ec: 5 }])
  s = s.setRowHeight(1, 40)

  const samples: Array<[number, number, string, CellStyle]> = [
    [3, 0, '粗体', { bold: true, bg: '#e8f0fe', align: 'center' }],
    [3, 1, '斜体', { italic: true, align: 'center' }],
    [3, 2, '下划线', { underline: true, align: 'center' }],
    [3, 3, '删除线', { strikethrough: true, align: 'center' }],
    [3, 4, '红字', { color: '#b91c1c', align: 'center' }],
    [3, 5, '底色', { bg: '#fef08a', align: 'center' }],
    [5, 0, '左对齐', { align: 'left', bg: '#f8fafc' }],
    [5, 1, '居中', { align: 'center', bg: '#f8fafc' }],
    [5, 2, '右对齐', { align: 'right', bg: '#f8fafc' }],
    [5, 3, '12px', { fontSize: 12, align: 'center' }],
    [5, 4, '16px', { fontSize: 16, align: 'center' }],
    [5, 5, '20px', { fontSize: 20, align: 'center' }],
    [7, 0, '细边框', { align: 'center', border: { top: THIN, right: THIN, bottom: THIN, left: THIN } }],
    [7, 1, '中边框', {
      align: 'center',
      border: {
        top: { style: 'medium', color: '#334155' },
        right: { style: 'medium', color: '#334155' },
        bottom: { style: 'medium', color: '#334155' },
        left: { style: 'medium', color: '#334155' },
      },
    }],
    [7, 2, '虚线', {
      align: 'center',
      border: {
        top: { style: 'dashed', color: '#0369a1' },
        right: { style: 'dashed', color: '#0369a1' },
        bottom: { style: 'dashed', color: '#0369a1' },
        left: { style: 'dashed', color: '#0369a1' },
      },
    }],
    [7, 3, '双线', {
      align: 'center',
      border: {
        bottom: { style: 'double', color: '#0f172a' },
      },
    }],
  ]
  for (const [r, c, raw, style] of samples) s = set(s, r, c, raw, style)

  s = set(s, 9, 0, '合并单元格', { ...HEADER, bg: '#dbeafe' })
  s = s.setMerges([...s.merges, { sr: 9, sc: 0, er: 10, ec: 2 }])
  s = set(s, 9, 3, '这一段会自动换行：行高由内容撑开，也可以在工具栏里改成手动行高。', {
    wrap: true,
    vAlign: 'top',
    bg: '#f8fafc',
  })
  s = s.setMerges([...s.merges, { sr: 9, sc: 3, er: 10, ec: 5 }])

  s = s.setColWidth(0, 110)
  s = s.setColWidth(1, 110)
  s = s.setColWidth(2, 110)
  s = s.setColWidth(3, 120)
  s = s.setColWidth(4, 100)
  s = s.setColWidth(5, 100)
  s = s.setFrozen(1, 0)
  return s
}

function buildTry(): SheetData {
  let s = SheetData.create({ rowCount: 1000, colCount: 26 })
  s = set(s, 0, 0, '试用这张表', TITLE)
  s = s.setMerges([{ sr: 0, sc: 0, er: 0, ec: 4 }])
  s = s.setRowHeight(0, 36)
  s = set(
    s,
    1,
    0,
    '从左上角「文件」打开你自己的 xlsx / CSV，或直接在下面改公式。清除浏览器存档可回到这份演示。',
    MUTED,
  )
  s = s.setMerges([...s.merges, { sr: 1, sc: 0, er: 1, ec: 4 }])
  s = s.setRowHeight(1, 40)

  s = set(s, 3, 0, '试一试', HEADER)
  s = set(s, 3, 1, '公式', HEADER)
  s = set(s, 3, 2, '说明', HEADER)

  const rows: [string, string, string][] = [
    ['10', '=A5*2', '算术'],
    ['20', '=SUM(A5:A7)', '求和'],
    ['30', '=AVERAGE(A5:A7)', '平均'],
    ['', '=IF(A5>=10,"ok","")', '条件'],
    ['', '=ROUND(C9,1)', '四舍五入（C9 是平均）'],
  ]
  rows.forEach(([a, b, c], i) => {
    const r = 4 + i
    if (a) s = set(s, r, 0, a, { align: 'right' })
    s = set(s, r, 1, b)
    s = set(s, r, 2, c, MUTED)
  })

  s = set(s, 10, 0, '区域', HEADER)
  s = set(s, 11, 0, '华东')
  s = set(s, 12, 0, '华北')
  s = set(s, 13, 0, '华南')
  s = set(s, 10, 1, '点 A12 看序列验证', MUTED)

  const validations: ValidationRule[] = [
    { id: 'dv-region', range: { sr: 11, sc: 0, er: 13, ec: 0 }, type: 'list', items: ['华东', '华北', '华南', '西南'] },
  ]
  s = s.setValidations(validations)

  s = s.setColWidth(0, 100)
  s = s.setColWidth(1, 200)
  s = s.setColWidth(2, 220)
  return s
}

export function createDemoState(): SheetState {
  let wb = Workbook.create({ rowCount: 1000, colCount: 26 })
  wb = wb.setSheet(wb.active, buildSales()).renameSheet(wb.active, '销售')
  wb = wb.addSheet('s2', buildStyles(), undefined, '样式')
  wb = wb.addSheet('s3', buildTry(), undefined, '试用')
  return createStateFromWorkbook(wb)
}
