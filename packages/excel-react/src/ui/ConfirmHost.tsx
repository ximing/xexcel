// 挂载于 App：有挂起请求时渲染 ConfirmDialog
import { useSyncExternalStore } from 'react'
import { ConfirmDialog } from './ConfirmDialog'
import { getConfirm, resolveConfirm, subscribeConfirm } from './confirmStore'

export function ConfirmHost() {
  const req = useSyncExternalStore(subscribeConfirm, getConfirm)
  if (!req) return null
  return (
    <ConfirmDialog
      title={req.title}
      body={req.body}
      confirmLabel={req.confirmLabel}
      danger={req.danger}
      onConfirm={() => resolveConfirm(true)}
      onCancel={() => resolveConfirm(false)}
    />
  )
}
