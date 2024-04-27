// 条件格式对话框：规则列表（增删/上下移调优先级）+ 行内编辑（范围/类型/条件/样式）。
// 草稿编辑，确定时一次性 tr.setCondFormats 提交（一次 undo）。
import { useState } from 'react'
import { parseRangeA1, toA1 } from '../core/addr'
import { CondFormatRule, FilterOp } from '../core/model'
import { selectionRange } from '../core/selection'
import type { EditorView } from '../view/editorview'
import { useSheetState } from './bridge'

interface Props {
  view: EditorView
  onClose: () => void
}

const VALUE_OPS: { op: FilterOp; label: string }[] = [
  { op: 'gt', label: '>' }, { op: 'gte', label: '>=' }, { op: 'lt', label: '<' },
  { op: 'lte', label: '<=' }, { op: 'eq', label: '=' }, { op: 'neq', label: '≠' },
  { op: 'between', label: '介于' },
]

const RULE_TYPES: { type: CondFormatRule['type']; label: string }[] = [
  { type: 'value', label: '单元格值' },
  { type: 'textContains', label: '文本包含' },
  { type: 'duplicate', label: '重复值' },
]

const STYLE_TOGGLES: { key: 'bold' | 'italic' | 'underline' | 'strikethrough'; label: string }[] = [
  { key: 'bold', label: 'B' },
  { key: 'italic', label: 'I' },
  { key: 'underline', label: 'U' },
  { key: 'strikethrough', label: 'S' },
]

function nextId(rules: CondFormatRule[]): string {
  let max = 0
  for (const r of rules) {
    const m = /^cf(\d+)$/.exec(r.id)
    if (m) max = Math.max(max, parseInt(m[1], 10))
  }
  return `cf${max + 1}`
}

function rangeText(r: { sr: number; sc: number; er: number; ec: number }): string {
  return `${toA1(r.sr, r.sc)}:${toA1(r.er, r.ec)}`
}

// 非法判定：范围解析失败；value 规则 v1 空；between 上界空
function ruleInvalid(rule: CondFormatRule, text: string): boolean {
  if (parseRangeA1(text) === null) return true
  if (rule.type === 'value') {
    if (rule.v1.trim() === '') return true
    if (rule.op === 'between' && (rule.v2 ?? '').trim() === '') return true
  }
  return false
}

export function CondFormatDialog({ view, onClose }: Props) {
  const state = useSheetState(view)
  // 草稿深拷贝：range/style 不共享引用
  const [draft, setDraft] = useState<CondFormatRule[]>(() =>
    state.activeSheet.condFormats.map((r) => ({ ...r, range: { ...r.range }, style: { ...r.style } })),
  )
  const [texts, setTexts] = useState<string[]>(() =>
    state.activeSheet.condFormats.map((r) => rangeText(r.range)),
  )

  const update = (i: number, next: CondFormatRule): void => {
    setDraft(draft.map((r, j) => (j === i ? next : r)))
  }
  const move = (i: number, d: -1 | 1): void => {
    const j = i + d
    if (j < 0 || j >= draft.length) return
    const next = [...draft]
    ;[next[i], next[j]] = [next[j], next[i]]
    setDraft(next)
    const nextTexts = [...texts]
    ;[nextTexts[i], nextTexts[j]] = [nextTexts[j], nextTexts[i]]
    setTexts(nextTexts)
  }
  const remove = (i: number): void => {
    setDraft(draft.filter((_, j) => j !== i))
    setTexts(texts.filter((_, j) => j !== i))
  }
  const addRule = (): void => {
    const range = selectionRange(view.state.selection)
    setDraft([
      ...draft,
      {
        id: nextId(draft),
        range,
        type: 'value',
        op: 'gt',
        v1: '',
        style: { bg: '#ffc7ce', color: '#9c0006' },
      },
    ])
    setTexts([...texts, rangeText(range)])
  }
  const setType = (i: number, t: CondFormatRule['type']): void => {
    const r = draft[i]
    const base = { id: r.id, range: r.range, style: r.style }
    const next: CondFormatRule =
      t === 'value'
        ? { ...base, type: 'value', op: 'gt', v1: '' }
        : t === 'textContains'
          ? { ...base, type: 'textContains', text: '' }
          : { ...base, type: 'duplicate' }
    update(i, next)
  }
  const toggleStyle = (i: number, key: 'bold' | 'italic' | 'underline' | 'strikethrough'): void => {
    const r = draft[i]
    update(i, { ...r, style: { ...r.style, [key]: r.style[key] ? undefined : true } })
  }

  const anyInvalid = draft.some((r, i) => ruleInvalid(r, texts[i]))
  const commit = (): void => {
    const rules = draft.map((r, i) => ({ ...r, range: parseRangeA1(texts[i])! }))
    view.dispatch(view.state.tr.setCondFormats(rules))
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
        <div className="dialog-title">条件格式</div>
        <div className="cf-rules">
          {draft.length === 0 && <div className="cf-empty">暂无规则</div>}
          {draft.map((rule, i) => (
            <div className={'cf-row' + (ruleInvalid(rule, texts[i]) ? ' invalid' : '')} key={rule.id}>
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
                onChange={(e) => setType(i, e.target.value as CondFormatRule['type'])}
              >
                {RULE_TYPES.map((t) => (
                  <option key={t.type} value={t.type}>
                    {t.label}
                  </option>
                ))}
              </select>
              {rule.type === 'value' && (
                <>
                  <select
                    className="tool-select"
                    value={rule.op}
                    onChange={(e) => update(i, { ...rule, op: e.target.value as FilterOp })}
                  >
                    {VALUE_OPS.map((o) => (
                      <option key={o.op} value={o.op}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                  <input
                    className="cf-value"
                    placeholder="值"
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
              {rule.type === 'textContains' && (
                <input
                  className="cf-value"
                  placeholder="包含文本"
                  value={rule.text}
                  onChange={(e) => update(i, { ...rule, text: e.target.value })}
                />
              )}
              <input
                className="tool-color"
                type="color"
                title="文字颜色"
                value={rule.style.color ?? '#000000'}
                onChange={(e) => update(i, { ...rule, style: { ...rule.style, color: e.target.value } })}
              />
              <input
                className="tool-color"
                type="color"
                title="背景颜色"
                value={rule.style.bg ?? '#ffffff'}
                onChange={(e) => update(i, { ...rule, style: { ...rule.style, bg: e.target.value } })}
              />
              {STYLE_TOGGLES.map((t) => (
                <button
                  key={t.key}
                  className={'tool-btn cf-toggle' + (rule.style[t.key] ? ' active' : '')}
                  title={{ bold: '加粗', italic: '斜体', underline: '下划线', strikethrough: '删除线' }[t.key]}
                  onClick={() => toggleStyle(i, t.key)}
                >
                  {t.label}
                </button>
              ))}
              <button className="tool-btn" title="上移（提高优先级）" disabled={i === 0} onClick={() => move(i, -1)}>
                ↑
              </button>
              <button
                className="tool-btn"
                title="下移（降低优先级）"
                disabled={i === draft.length - 1}
                onClick={() => move(i, 1)}
              >
                ↓
              </button>
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
