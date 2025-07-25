import { Button } from './Button'
import { Dialog } from './Dialog'

interface Props {
  title: string
  body?: string
  confirmLabel?: string
  danger?: boolean
  onConfirm: () => void
  onCancel: () => void
}

export function ConfirmDialog({ title, body, confirmLabel = '确定', danger, onConfirm, onCancel }: Props) {
  return (
    <Dialog
      title={title}
      onClose={onCancel}
      footer={
        <>
          <Button variant="outline" onClick={onCancel}>取消</Button>
          <Button variant={danger ? 'danger' : 'primary'} onClick={onConfirm}>{confirmLabel}</Button>
        </>
      }
    >
      {body && <div className="text-sm text-ink-2">{body}</div>}
    </Dialog>
  )
}
