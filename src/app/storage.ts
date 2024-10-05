// localStorage 薄壳：防抖自动保存 + 损坏备份 + 保存状态订阅（StatusBar 用）。
// node（vitest）环境退化为内存 Map：模块 import 期不得触碰 localStorage。
import { deserializeWorkbook, serializeWorkbook } from '../core/io/persist'
import { Workbook } from '../core/model'

export interface StorageLike {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
  removeItem(key: string): void
}

export const STORAGE_KEY = 'xexcel.workbook'

export interface SaveStatus {
  readonly error: string | null
  readonly savedAt: string | null
}

export class WorkbookStorage {
  private pending: (() => Workbook) | null = null
  private timer: ReturnType<typeof setTimeout> | null = null
  private status: SaveStatus = { error: null, savedAt: null }
  private suspended = false
  private readonly listeners = new Set<() => void>()

  constructor(
    private readonly storage: StorageLike,
    private readonly debounceMs = 1000,
  ) {}

  // 无存档/损坏/读异常均返回 null；损坏存档备份到 .corrupt 备查后删原键
  load(): Workbook | null {
    let raw: string | null
    try {
      raw = this.storage.getItem(STORAGE_KEY)
    } catch {
      return null
    }
    if (!raw) return null
    try {
      return deserializeWorkbook(raw)
    } catch {
      try {
        this.storage.setItem(STORAGE_KEY + '.corrupt', raw)
        this.storage.removeItem(STORAGE_KEY)
      } catch {
        // 备份失败不阻塞启动
      }
      return null
    }
  }

  saveNow(wb: Workbook): void {
    if (this.suspended) return
    try {
      const savedAt = new Date().toISOString()
      this.storage.setItem(STORAGE_KEY, serializeWorkbook(wb, savedAt))
      this.setStatus({ error: null, savedAt })
    } catch (e) {
      this.setStatus({ ...this.status, error: (e as Error).message || String(e) })
    }
  }

  // getter 延迟到触发时取值：防抖窗口内的连续编辑只序列化最终 state
  schedule(getWorkbook: () => Workbook): void {
    if (this.suspended) return
    this.pending = getWorkbook
    if (this.timer !== null) clearTimeout(this.timer)
    this.timer = setTimeout(() => this.flush(), this.debounceMs)
  }

  flush(): void {
    if (this.timer !== null) {
      clearTimeout(this.timer)
      this.timer = null
    }
    const get = this.pending
    this.pending = null
    if (get) this.saveNow(get())
  }

  clear(): void {
    try {
      this.storage.removeItem(STORAGE_KEY)
    } catch {
      // 忽略
    }
  }

  // 清除存档并挂起自动保存：取消挂起的防抖 timer/pending，本页面生命周期内不再写入
  // （模块级 latch，刷新即复位；防止「清除」被后续 dispatch 的防抖复活）
  suspend(): void {
    if (this.timer !== null) {
      clearTimeout(this.timer)
      this.timer = null
    }
    this.pending = null
    this.clear()
    this.suspended = true
  }

  // 解除挂起（xlsx 导入完成后恢复自动保存；不补保存，等下一次编辑的 schedule 触发）
  resume(): void {
    this.suspended = false
  }

  subscribeStatus = (cb: () => void): (() => void) => {
    this.listeners.add(cb)
    return () => this.listeners.delete(cb)
  }

  getStatus = (): SaveStatus => this.status

  private setStatus(s: SaveStatus): void {
    this.status = s
    for (const cb of this.listeners) cb()
  }
}

function defaultStorageLike(): StorageLike {
  if (typeof localStorage !== 'undefined') return localStorage
  const m = new Map<string, string>()
  return {
    getItem: (k) => m.get(k) ?? null,
    setItem: (k, v) => void m.set(k, v),
    removeItem: (k) => void m.delete(k),
  }
}

export const workbookStorage = new WorkbookStorage(defaultStorageLike())
