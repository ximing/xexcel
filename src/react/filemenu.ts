// src/react/filemenu.ts
// 文件菜单纯逻辑（组件外可单测，同 menu.ts 模式）。
export type FileMenuId = 'openCsv' | 'openXlsx' | 'exportCsv' | 'exportXlsx' | 'clearStorage'

export interface FileMenuItem {
  id: FileMenuId
  label: string
  danger: boolean
}

export function fileMenuItems(): FileMenuItem[] {
  return [
    { id: 'openCsv', label: '打开 CSV…', danger: false },
    { id: 'openXlsx', label: '打开 xlsx…', danger: false },
    { id: 'exportCsv', label: '导出 CSV', danger: false },
    { id: 'exportXlsx', label: '导出 xlsx', danger: false },
    { id: 'clearStorage', label: '清除浏览器存档', danger: true },
  ]
}

// 文件名去扩展名作为新 sheet 名（空回退 CSV）
export function csvBaseName(fileName: string): string {
  const stem = fileName.replace(/\.[^.]+$/, '').trim()
  return stem || 'CSV'
}

export const CSV_MAX_BYTES = 5 * 1024 * 1024
export const CSV_MAX_ROWS = 100_000
export const XLSX_MAX_BYTES = 10 * 1024 * 1024

export function isGridEmpty(grid: string[][]): boolean {
  return grid.every((r) => r.every((c) => c === ''))
}
