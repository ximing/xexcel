// src/core/sheetname.ts
// Excel 工作表名规则统一层：非法字符集、长度上限、校验与净化。
// rename UI 用 isValidSheetName 拒绝非法输入；xlsx 导入/导出用 sanitizeSheetName 防 exceljs throw。

// Excel 非法字符集：* ? : \ / [ ]
export const INVALID_SHEET_NAME_CHARS = /[*?:\\/\[\]]/
export const SHEET_NAME_MAX_LEN = 31

// rename 入口校验：非空（trim 后）、无非法字符、≤31
export function isValidSheetName(name: string): boolean {
  return (
    name.trim() !== '' &&
    name.length <= SHEET_NAME_MAX_LEN &&
    !INVALID_SHEET_NAME_CHARS.test(name)
  )
}

// 不区分大小写去重：冲突时追加 ` (2)` 序号；超长时先截 stem 保证整体 ≤31
export function dedupeSheetName(existing: Iterable<string>, base: string): string {
  const taken = new Set([...existing].map((n) => n.toLowerCase()))
  if (!taken.has(base.toLowerCase())) return base
  for (let i = 2; ; i++) {
    const suffix = ` (${i})`
    const stem =
      base.length + suffix.length > SHEET_NAME_MAX_LEN
        ? base.slice(0, SHEET_NAME_MAX_LEN - suffix.length)
        : base
    const cand = `${stem}${suffix}`
    if (!taken.has(cand.toLowerCase())) return cand
  }
}

// 净化：非法字符 → `_`，截断 31，trim 后空回退 'Sheet'，再与 existing 去重
export function sanitizeSheetName(name: string, existing: Iterable<string>): string {
  let cleaned = name.replace(new RegExp(INVALID_SHEET_NAME_CHARS.source, 'g'), '_')
  if (cleaned.length > SHEET_NAME_MAX_LEN) cleaned = cleaned.slice(0, SHEET_NAME_MAX_LEN)
  cleaned = cleaned.trim()
  if (cleaned === '') cleaned = 'Sheet'
  return dedupeSheetName(existing, cleaned)
}
