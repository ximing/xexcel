// 数据验证对话框：规则列表（增删）+ 行内编辑（范围/类型/条件/序列项）。
// 草稿编辑，确定时一次性 tr.setValidations 提交（一次 undo）。结构镜像 CondFormatDialog。
import { useState } from 'react'
import { parseRangeA1 } from '../core/addr'
import { FilterOp, ValidationRule } from '../core/model'
import { selectionRange } from '../core/selection'
import type { EditorView } from '../view/editorview'
import { useSheetState } from './bridge'
import {
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
    if (m) { window.alert(m); return }
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

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="dialog cf-dialog"
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            e.stopPropagation()
            onClose()
            view.focus()
          }
        }}
      >
        <div className="dialog-title">数据验证</div>
        <div className="cf-rules">
          {draft.length === 0 && <div className="cf-empty">暂无规则</div>}
          {draft.map((rule, i) => (
            <div className={'cf-row' + (invalid(i) ? ' invalid' : '')} key={rule.id}>
              <input
                className="cf-range"
                title="应用范围（A1 表示法）"
                value={texts[i]}
                onChange={(e) => setTexts(texts.map((t, j) => (j === i ? e.target.value : t)))}
                onBlur={() => {
                  // 失焦校验：合法则归一化显示，非法保持原文标红
                  const nr = parseRangeA1(texts[i])
                  if (nr) {
                    setTexts(texts.map((t, j) => (j === i ? rangeText(nr) : t)))
                    update(i, { ...rule, range: nr })
                  }
                }}
              />
              <select
                className="tool-select"
                value={rule.type}
                onChange={(e) => setType(i, e.target.value as ValidationRule['type'])}
              >
                {RULE_TYPES.map((t) => (
                  <option key={t.type} value={t.type}>
                    {t.label}
                  </option>
                ))}
              </select>
              {rule.type !== 'list' && (
                <>
                  <select
                    className="tool-select"
                    value={rule.op}
                    onChange={(e) => update(i, { ...rule, op: e.target.value as FilterOp })}
                  >
                    {VALIDATION_OPS.map((o) => (
                      <option key={o.op} value={o.op}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                  <input
                    className="cf-value"
                    placeholder={rule.type === 'numRange' ? '数值' : '长度'}
                    value={rule.v1}
                    onChange={(e) => update(i, { ...rule, v1: e.target.value })}
                  />
                  {rule.op === 'between' && (
                    <input
                      className="cf-value"
                      placeholder="上界"
                      value={rule.v2 ?? ''}
                      onChange={(e) => update(i, { ...rule, v2: e.target.value })}
                    />
                  )}
                </>
              )}
              {rule.type === 'list' && (
                <input
                  className="cf-value"
                  placeholder="序列项，逗号分隔（如：是,否）"
                  value={itemTexts[i]}
                  onChange={(e) => setItemTexts(itemTexts.map((s, j) => (j === i ? e.target.value : s)))}
                />
              )}
              <button className="tool-btn" title="删除规则" onClick={() => remove(i)}>
                ×
              </button>
            </div>
          ))}
        </div>
        <div className="dialog-actions">
          <button onClick={addRule}>+ 添加规则</button>
          <span className="cf-actions-gap" />
          <button disabled={anyInvalid} title={anyInvalid ? '存在非法规则（标红行）' : undefined} onClick={commit}>
            确定
          </button>
          <button
            onClick={() => {
              onClose()
              view.focus()
            }}
          >
            取消
          </button>
        </div>
      </div>
    </div>
  )
}
