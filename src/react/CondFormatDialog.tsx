// 条件格式对话框：规则列表（增删/上下移调优先级）+ 行内编辑（范围/类型/条件/样式）。
// 草稿编辑，确定时一次性 tr.setCondFormats 提交（一次 undo）。
import { useState } from 'react'
import { showNotice } from '../app/notice'
import { parseRangeA1, toA1 } from '../core/addr'
import { CondFormatRule, FilterOp } from '../core/model'
import { selectionRange } from '../core/selection'
import type { SheetState } from '../core/state'
import type { EditorView } from '../view/editorview'
import { useSheetState } from './bridge'
import { CFToggleKey, toggleCFStyle, DEFAULT_CF_STYLE, COLOR_INPUT_FALLBACK_TEXT, COLOR_INPUT_FALLBACK_BG } from './cfstyle'
import { Button } from './ui/Button'
import { Dialog } from './ui/Dialog'
import { Select } from './ui/Select'
import { TextInput } from './ui/TextInput'
import { Tooltip } from './ui/Tooltip'

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

const STYLE_TOGGLES: { key: CFToggleKey; label: string }[] = [
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
    // 多区域选区拒绝（入口按钮已禁用，此处兜底）
    const m = condFormatRejection(view.state)
    if (m) { showNotice(m); return }
    const range = selectionRange(view.state.selection)
    setDraft([
      ...draft,
      {
        id: nextId(draft),
        range,
        type: 'value',
        op: 'gt',
        v1: '',
        style: { ...DEFAULT_CF_STYLE },
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
  const toggleStyle = (i: number, key: CFToggleKey): void => {
    const r = draft[i]
    update(i, { ...r, style: toggleCFStyle(r.style, key) })
  }

  const anyInvalid = draft.some((r, i) => ruleInvalid(r, texts[i]))
  const commit = (): void => {
    const rules = draft.map((r, i) => ({ ...r, range: parseRangeA1(texts[i])! }))
    view.dispatch(view.state.tr.setCondFormats(rules))
    onClose()
    view.focus()
  }

  const close = (): void => {
    onClose()
    view.focus()
  }

  return (
    <Dialog
      title="条件格式"
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
              (ruleInvalid(rule, texts[i]) ? 'border-danger bg-danger-soft' : 'border-transparent')
            }
            key={rule.id}
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
              onChange={(v) => setType(i, v as CondFormatRule['type'])}
            />
            {rule.type === 'value' && (
              <>
                <Select
                  value={rule.op}
                  options={VALUE_OPS.map((o) => ({ value: o.op, label: o.label }))}
                  onChange={(v) => update(i, { ...rule, op: v as FilterOp })}
                />
                <TextInput
                  width={80}
                  placeholder="值"
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
            {rule.type === 'textContains' && (
              <TextInput
                width={80}
                placeholder="包含文本"
                value={rule.text}
                onChange={(v) => update(i, { ...rule, text: v })}
              />
            )}
            <input
              className="h-7 w-7 shrink-0 cursor-pointer"
              type="color"
              title="文字颜色"
              value={rule.style.color ?? COLOR_INPUT_FALLBACK_TEXT}
              onChange={(e) => update(i, { ...rule, style: { ...rule.style, color: e.target.value } })}
            />
            <input
              className="h-7 w-7 shrink-0 cursor-pointer"
              type="color"
              title="背景颜色"
              value={rule.style.bg ?? COLOR_INPUT_FALLBACK_BG}
              onChange={(e) => update(i, { ...rule, style: { ...rule.style, bg: e.target.value } })}
            />
            {STYLE_TOGGLES.map((t) => (
              <Tooltip
                key={t.key}
                tip={{ bold: '加粗', italic: '斜体', underline: '下划线', strikethrough: '删除线' }[t.key]}
              >
                <Button variant="ghost" size="sm" active={!!rule.style[t.key]} onClick={() => toggleStyle(i, t.key)}>
                  {t.label}
                </Button>
              </Tooltip>
            ))}
            <Tooltip tip="上移（提高优先级）">
              <Button variant="ghost" size="sm" disabled={i === 0} onClick={() => move(i, -1)}>
                ↑
              </Button>
            </Tooltip>
            <Tooltip tip="下移（降低优先级）">
              <Button variant="ghost" size="sm" disabled={i === draft.length - 1} onClick={() => move(i, 1)}>
                ↓
              </Button>
            </Tooltip>
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

// 多区域选区下条件格式入口按钮禁用（单区域零回归）
export function canCondFormat(state: SheetState): boolean {
  return state.selection.ranges.length === 1
}

// 多区域触发添加规则时的拒绝消息；单区域放行返回 null
export function condFormatRejection(state: SheetState): string | null {
  return state.selection.ranges.length > 1 ? '条件格式仅支持单区域选择' : null
}
