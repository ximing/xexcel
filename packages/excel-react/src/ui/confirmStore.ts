// 自绘确认框微型 store（仿 app/notice.ts）：askConfirm 挂起，ConfirmHost 渲染，resolveConfirm 兑现。
// 前问未决时再 ask → 前问按「取消」兑现 false。
export interface ConfirmRequest {
  title: string
  body?: string
  confirmLabel?: string
  danger?: boolean
}

let pending: { req: ConfirmRequest; resolve: (ok: boolean) => void } | null = null
const listeners = new Set<() => void>()

function emit(): void {
  for (const cb of listeners) cb()
}

export function askConfirm(req: ConfirmRequest): Promise<boolean> {
  if (pending) pending.resolve(false)
  return new Promise<boolean>((resolve) => {
    pending = { req, resolve }
    emit()
  })
}

export function resolveConfirm(ok: boolean): void {
  if (!pending) return
  const { resolve } = pending
  pending = null
  resolve(ok)
  emit()
}

export function getConfirm(): ConfirmRequest | null {
  return pending?.req ?? null
}

export function subscribeConfirm(cb: () => void): () => void {
  listeners.add(cb)
  return () => listeners.delete(cb)
}
