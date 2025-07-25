// src/app/notice.ts
// 瞬态提示微型 store：StatusBar 展示，5s 自动消（文件为空、清除存档成功等）。
let current: string | null = null
let timer: ReturnType<typeof setTimeout> | null = null
const listeners = new Set<() => void>()

function set(msg: string | null): void {
  current = msg
  for (const cb of listeners) cb()
}

export function showNotice(msg: string): void {
  if (timer !== null) clearTimeout(timer)
  set(msg)
  timer = setTimeout(() => {
    timer = null
    set(null)
  }, 5000)
}

export function subscribeNotice(cb: () => void): () => void {
  listeners.add(cb)
  return () => listeners.delete(cb)
}

export function getNotice(): string | null {
  return current
}
