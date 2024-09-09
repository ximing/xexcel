// RFC 4180 子集：解析宽松（坏行尽力保留），导出严格转义。纯函数，零 DOM。

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
