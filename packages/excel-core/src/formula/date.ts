// 日期序列值工具（Excel 1900 系统）：serial 1 = 1900-01-01，基准 1899-12-30。
export const DAY_MS = 86400000
export const EPOCH = Date.UTC(1899, 11, 30)

// 严格校验（录入归一用）：非法日期（如 2/30）返回 null
export function dateSerial(y: number, m: number, d: number): number | null {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null
  const dt = new Date(Date.UTC(y, m - 1, d))
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null
  return Math.round((dt.getTime() - EPOCH) / DAY_MS)
}

// 宽松转换（Excel DATE 函数语义：月/日可溢出进位，如 DATE(2026,13,1) → 2027-01-01）
export function dateSerialLenient(y: number, m: number, d: number): number {
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS)
}

export function serialToDate(serial: number): { y: number; m: number; d: number } {
  const dt = new Date(EPOCH + Math.floor(serial) * DAY_MS)
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() }
}

export function todaySerial(now: Date = new Date()): number {
  return Math.round((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - EPOCH) / DAY_MS)
}

export function nowSerial(now: Date = new Date()): number {
  const base = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())
  const ms = (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) * 1000
  return (base - EPOCH + ms) / DAY_MS
}
