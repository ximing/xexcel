// 录入归一：把用户输入的日期/时间文本转为 Excel 序列值（1900 系统）+ 对应 numFmt。
// 纯函数，零 DOM 依赖。接入点：编辑框提交、公式栏提交、外部 TSV 粘贴。
import { Cell } from '../core/model'
import { dateSerial } from './date'

// 仅为 =TODAY()/=NOW() 的公式录入自动附日期格式（否则显示 serial 数字）
const TODAY_RE = /^=\s*TODAY\s*\(\s*\)\s*$/i
const NOW_RE = /^=\s*NOW\s*\(\s*\)\s*$/i

const FULL_DATE_RE = /^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$/
const SHORT_DATE_RE = /^(\d{1,2})[/-](\d{1,2})$/
const TIME_RE = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$/

export function normalizeInput(raw: string, now?: Date): { raw: string; numFmt?: string } {
  const t = raw.trim()
  if (t === '' || t.startsWith('=')) return { raw }
  let m = FULL_DATE_RE.exec(t)
  if (m) {
    const s = dateSerial(Number(m[1]), Number(m[2]), Number(m[3]))
    if (s !== null) return { raw: String(s), numFmt: 'yyyy/m/d' }
    return { raw }
  }
  m = SHORT_DATE_RE.exec(t)
  if (m) {
    const y = (now ?? new Date()).getFullYear()
    const s = dateSerial(y, Number(m[1]), Number(m[2]))
    if (s !== null) return { raw: String(s), numFmt: 'm/d' }
    return { raw }
  }
  m = TIME_RE.exec(t)
  if (m) {
    let h = Number(m[1])
    const min = Number(m[2])
    const sec = m[3] ? Number(m[3]) : 0
    const ampm = m[4]?.toUpperCase()
    if (min > 59 || sec > 59) return { raw }
    if (ampm) {
      if (h < 1 || h > 12) return { raw }
      if (ampm === 'PM' && h !== 12) h += 12
      if (ampm === 'AM' && h === 12) h = 0
    } else if (h > 23) {
      return { raw }
    }
    const serial = (h * 3600 + min * 60 + sec) / 86400
    return { raw: String(serial), numFmt: 'h:mm' }
  }
  return { raw }
}

// 提交路径统一入口：返回要写入的 Cell（保留原样式；新识别日期且无 numFmt 时合并 numFmt）
export function normalizedCell(text: string, existing: Cell | undefined, now?: Date): Cell {
  const n = normalizeInput(text, now)
  const style = existing?.style ? { ...existing.style } : undefined
  if (!style?.numFmt) {
    if (n.numFmt) return { raw: n.raw, style: { ...style, numFmt: n.numFmt } }
    if (TODAY_RE.test(text.trim())) return { raw: n.raw, style: { ...style, numFmt: 'yyyy/m/d' } }
    if (NOW_RE.test(text.trim())) return { raw: n.raw, style: { ...style, numFmt: 'yyyy/m/d h:mm' } }
  }
  return style ? { raw: n.raw, style } : { raw: n.raw }
}
