// RFC 4180 子集：解析宽松（坏行尽力保留），导出严格转义。纯函数，零 DOM。

import type { SheetData } from '../model'

export function csvToGrid(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let inQuotes = false
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0 // 去 BOM

  const pushField = (): void => {
    row.push(field)
    field = ''
  }
  const pushRow = (): void => {
    pushField()
    rows.push(row)
    row = []
  }

  for (; i < text.length; i++) {
    const ch = text[i]
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i++
        } else {
          inQuotes = false
        }
      } else {
        field += ch
      }
      continue
    }
    // 引号仅在字段开头生效；字段中间的引号按字面字符（宽松）
    if (ch === '"' && field === '') {
      inQuotes = true
      continue
    }
    if (ch === ',') {
      pushField()
      continue
    }
    if (ch === '\r') {
      if (text[i + 1] === '\n') i++
      pushRow()
      continue
    }
    if (ch === '\n') {
      pushRow()
      continue
    }
    field += ch
  }
  pushField()
  // 尾部换行后只剩空残余 → 不补幽灵行；纯空文本 → []
  if (row.length > 1 || row[0] !== '') rows.push(row)
  return rows
}

// 含 , " \r \n 的字段按 RFC 4180 加引号、内嵌 " 翻倍
function escapeField(v: string): string {
  return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v
}

export function sheetToCSV(sheet: SheetData, opts?: { bom?: boolean }): string {
  const r = sheet.usedRange()
  if (r.sr === 0 && r.sc === 0 && r.er === 0 && r.ec === 0 && !sheet.getCell(0, 0)) return ''
  const lines: string[] = []
  for (let row = 0; row <= r.er; row++) {
    const fields: string[] = []
    for (let col = 0; col <= r.ec; col++) {
      fields.push(escapeField(sheet.getCell(row, col)?.raw ?? ''))
    }
    lines.push(fields.join(','))
  }
  const body = lines.join('\r\n') + '\r\n'
  return (opts?.bom === false ? '' : '﻿') + body
}
