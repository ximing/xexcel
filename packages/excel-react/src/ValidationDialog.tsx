// 数据验证对话框：规则列表（增删）+ 行内编辑（范围/类型/条件/序列项）。
// 草稿编辑，确定时一次性 tr.setValidations 提交（一次 undo）。结构镜像 CondFormatDialog。
import { useState } from 'react'
import { showNotice } from './notice'
import { parseRangeA1 } from '@xexcel/core'
import { FilterOp, ValidationRule } from '@xexcel/core'
import { selectionRange } from '@xexcel/core'
import type { EditorView } from '@xexcel/view'
import { useSheetState } from './bridge'
import { Button } from './ui/Button'
import { Dialog } from './ui/Dialog'
import { Select } from './ui/Select'
import { TextInput } from './ui/TextInput'
import { Tooltip } from './ui/Tooltip'
import {
  describeRule,
  nextValidationId,
  normalizeRuleRange,
  parseItems,
  rangeText,
  ruleInvalid,
  validationRejection,
  VALIDATION_OPS,
} from './validationRules'

interface Props {
  view: EditorView
  onClose: () => void
}

const RULE_TYPES: { type: ValidationRule['type']; label: string }[] = [
  { type: 'numRange', label: '数值范围' },
  { type: 'textLen', label: '文本长度' },
  { type: 'list', label: '序列' },
]

export function ValidationDialog({ view, onClose }: Props) {
  const state = useSheetState(view)
  // 草稿深拷贝：range/items 不共享引用
  const [draft, setDraft] = useState<ValidationRule[]>(() =>
    state.activeSheet.validations.map((r) =>
      r.type === 'list'
        ? { ...r, range: { ...r.range }, items: [...r.items] }
        : { ...r, range: { ...r.range } },
    ),
  )
  const [texts, setTexts] = useState<string[]>(() =>
    state.activeSheet.validations.map((r) => rangeText(r.range)),
  )
  // 序列项的文本草稿（与 draft 平行；提交时 parseItems 解析回 items）
  const [itemTexts, setItemTexts] = useState<string[]>(() =>
    state.activeSheet.validations.map((r) => (r.type === 'list' ? r.items.join(', ') : '')),
  )

  const update = (i: number, next: ValidationRule): void => {
    setDraft(draft.map((r, j) => (j === i ? next : r)))
  }
  const remove = (i: number): void => {
    setDraft(draft.filter((_, j) => j !== i))
    setTexts(texts.filter((_, j) => j !== i))
    setItemTexts(itemTexts.filter((_, j) => j !== i))
  }
  const addRule = (): void => {
    // 多区域选区拒绝（入口按钮已禁用，此处兜底）
    const m = validationRejection(view.state)
    if (m) { showNotice(m); return }
    const range = selectionRange(view.state.selection)
    setDraft([
      ...draft,
      { id: nextValidationId(draft), range, type: 'numRange', op: 'between', v1: '', v2: '' },
    ])
    setTexts([...texts, rangeText(range)])
    setItemTexts([...itemTexts, ''])
  }
  const setType = (i: number, t: ValidationRule['type']): void => {
    const r = draft[i]
    const base = { id: r.id, range: r.range }
    const next: ValidationRule =
      t === 'list'
        ? { ...base, type: 'list', items: [] }
        : t === 'numRange'
          ? { ...base, type: 'numRange', op: 'between', v1: '', v2: '' }
          : { ...base, type: 'textLen', op: 'between', v1: '', v2: '' }
    update(i, next)
    if (t === 'list') setItemTexts(itemTexts.map((s, j) => (j === i ? '' : s)))
  }

  // 行非法判定：list 以 items 文本草稿实时解析为准
  const invalid = (i: number): boolean => {
    const r = draft[i]
    const eff = r.type === 'list' ? { ...r, items: parseItems(itemTexts[i]) } : r
    return ruleInvalid(eff, texts[i])
  }
  const anyInvalid = draft.some((_, i) => invalid(i))
  const commit = (): void => {
    const rules = draft.map((r, i) => {
      const range = normalizeRuleRange(parseRangeA1(texts[i])!, state.activeSheet)
      return r.type === 'list' ? { ...r, range, items: parseItems(itemTexts[i]) } : { ...r, range }
    })
    view.dispatch(view.state.tr.setValidations(rules))
    onClose()
    view.focus()
  }

  const close = (): void => {
    onClose()
    view.focus()
  }

  return (
    <Dialog
      title="数据验证"
      width={640}
      onClose={close}
      footer={
        <>
          <Button variant="outline" onClick={close}>
            取消
          </Button>
          <Tooltip tip={anyInvalid ? '存在非法规则（标红行）' : ''}>
            <Button variant="primary" disabled={anyInvalid} onClick={commit}>
              确定
            </Button>
          </Tooltip>
        </>
      }
    >
      <div className="flex max-h-[50vh] flex-col gap-1.5 overflow-y-auto">
        {draft.length === 0 && <div className="text-sm text-ink-2">暂无规则</div>}
        {draft.map((rule, i) => (
          <div
            className={
              'flex items-center gap-1 rounded-md border p-1 ' +
              (invalid(i) ? 'border-danger bg-danger-soft' : 'border-transparent')
            }
            key={rule.id}
            // hover 显示完整规则描述（list 以 items 文本草稿实时解析为准）
            title={describeRule(
              rule.type === 'list' ? { ...rule, items: parseItems(itemTexts[i]) } : rule,
            )}
          >
            <TextInput
              width={110}
              title="应用范围（A1 表示法）"
              invalid={parseRangeA1(texts[i]) === null}
              value={texts[i]}
              onChange={(v) => setTexts(texts.map((t, j) => (j === i ? v : t)))}
              onBlur={() => {
                // 失焦校验：合法则归一化显示，非法保持原文标红
                const nr = parseRangeA1(texts[i])
                if (nr) {
                  setTexts(texts.map((t, j) => (j === i ? rangeText(nr) : t)))
                  update(i, { ...rule, range: nr })
                }
              }}
            />
            <Select
              value={rule.type}
              options={RULE_TYPES.map((t) => ({ value: t.type, label: t.label }))}
              onChange={(v) => setType(i, v as ValidationRule['type'])}
            />
            {rule.type !== 'list' && (
              <>
                <Select
                  value={rule.op}
                  options={VALIDATION_OPS.map((o) => ({ value: o.op, label: o.label }))}
                  onChange={(v) => update(i, { ...rule, op: v as FilterOp })}
                />
                <TextInput
                  width={80}
                  placeholder={rule.type === 'numRange' ? '数值' : '长度'}
                  value={rule.v1}
                  onChange={(v) => update(i, { ...rule, v1: v })}
                />
                {rule.op === 'between' && (
                  <TextInput
                    width={80}
                    placeholder="上界"
                    value={rule.v2 ?? ''}
                    onChange={(v) => update(i, { ...rule, v2: v })}
                  />
                )}
              </>
            )}
            {rule.type === 'list' && (
              <TextInput
                width={200}
                placeholder="序列项，逗号分隔（如：是,否）"
                value={itemTexts[i]}
                onChange={(v) => setItemTexts(itemTexts.map((s, j) => (j === i ? v : s)))}
              />
            )}
            <Tooltip tip="删除规则">
              <Button variant="ghost" size="sm" onClick={() => remove(i)}>
                ×
              </Button>
            </Tooltip>
          </div>
        ))}
      </div>
      <div className="mt-2">
        <Button variant="outline" size="sm" onClick={addRule}>
          + 添加规则
        </Button>
      </div>
    </Dialog>
  )
}
