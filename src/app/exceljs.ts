// 浏览器 exceljs 桥：vendor script 全局 → core DI。app 层允许碰 window。
import { registerExcelJS } from '../core/io/xlsx'

export function setupExcelJS(): void {
  const w = window as unknown as { ExcelJS?: Parameters<typeof registerExcelJS>[0] }
  if (w.ExcelJS) registerExcelJS(w.ExcelJS)
}
