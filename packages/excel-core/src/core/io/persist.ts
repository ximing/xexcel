// Workbook 持久化信封：version 字段为后续格式演进留迁移钩子。纯函数，零 DOM。
import { Workbook } from '../model'

export const PERSIST_VERSION = 1

export class PersistError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'PersistError'
  }
}

interface Envelope {
  version: number
  savedAt: string
  workbook: unknown
}

export function serializeWorkbook(wb: Workbook, savedAt: string = new Date().toISOString()): string {
  const env: Envelope = { version: PERSIST_VERSION, savedAt, workbook: wb.toJSON() }
  return JSON.stringify(env)
}

export function deserializeWorkbook(json: string): Workbook {
  let env: Envelope
  try {
    env = JSON.parse(json) as Envelope
  } catch {
    throw new PersistError('存档不是合法 JSON')
  }
  if (!env || typeof env !== 'object' || env.version !== PERSIST_VERSION) {
    throw new PersistError(`存档版本不符: ${String((env as Envelope | null)?.version)}`)
  }
  try {
    const wb = Workbook.fromJSON(env.workbook)
    // 结构合法但语义空/不自洽的载荷 fromJSON 不抛错，需校验最小不变式
    if (wb.order.length === 0) throw new PersistError('存档无工作表')
    if (!wb.order.includes(wb.active)) throw new PersistError(`活动表不在 order 中: ${wb.active}`)
    for (const id of wb.order) {
      if (!wb.sheets.has(id)) throw new PersistError(`order 中的表缺少数据: ${id}`)
    }
    return wb
  } catch (e) {
    if (e instanceof PersistError) throw e
    throw new PersistError('存档数据损坏: ' + (e as Error).message)
  }
}
