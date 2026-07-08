// 文件菜单纯逻辑（原 filemenu.ts；与 FileMenu.tsx 同目录仅大小写冲突，tsup dts 无法发射，故改名）
// 文件菜单纯逻辑（组件外可单测，同 menu.ts 模式）。
import type { Locale } from './i18n'
import { t } from './i18n'

export type FileMenuId = 'openCsv' | 'openXlsx' | 'exportCsv' | 'exportXlsx' | 'clearStorage'

export interface FileMenuItem {
  id: FileMenuId
  label: string
  danger: boolean
}

export function fileMenuItems(locale: Locale = 'zh'): FileMenuItem[] {
  return [
    { id: 'openCsv', label: t(locale, 'file.openCsv'), danger: false },
    { id: 'openXlsx', label: t(locale, 'file.openXlsx'), danger: false },
    { id: 'exportCsv', label: t(locale, 'file.exportCsv'), danger: false },
    { id: 'exportXlsx', label: t(locale, 'file.exportXlsx'), danger: false },
    { id: 'clearStorage', label: t(locale, 'file.clearStorage'), danger: true },
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
