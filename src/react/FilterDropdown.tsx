// 筛选下拉浮层：值勾选列表（搜索/全选/清空）+ 条件筛选区。
// 开启态来自 filterDropdownKey state field（点击箭头写入）；确定/清除经 tr.setFilter 入 undo 栈。
import { useEffect, useMemo, useState } from 'react'
import { FilterConditionCriteria, FilterCriteria, FilterOp, FilterState } from '../core/model'
import { evaluatorFor } from '../formula/engine'
import type { EditorView } from '../view/editorview'
import { filterDropdownKey, FilterDropdownOpen } from '../view/types'
import { useSheetState } from './bridge'
import { Button } from './ui/Button'
import { Select } from './ui/Select'
import { TextInput } from './ui/TextInput'

const TEXT_OPS: { op: FilterOp; label: string }[] = [
  { op: 'contains', label: '包含' },
  { op: 'notContains', label: '不包含' },
  { op: 'eq', label: '等于' },
  { op: 'neq', label: '不等于' },
  { op: 'startsWith', label: '开头是' },
  { op: 'endsWith', label: '结尾是' },
]
const NUM_OPS: { op: FilterOp; label: string }[] = [
  { op: 'eq', label: '=' },
  { op: 'neq', label: '≠' },
  { op: 'gt', label: '>' },
  { op: 'gte', label: '≥' },
  { op: 'lt', label: '<' },
  { op: 'lte', label: '≤' },
  { op: 'between', label: '介于' },
]

// 面板最大高度估值（值列表 max-height 180 + 其余区块），供垂直 clamp
const PANEL_MAX_H = 350

interface Draft {
  mode: 'values' | 'condition'
  excluded: Set<string>
  field: 'text' | 'num'
  op: FilterOp
  v1: string
  v2: string
}

export function FilterDropdown({ view }: { view: EditorView }) {
  const state = useSheetState(view)
  const open = state.getField(filterDropdownKey) as FilterDropdownOpen | null
  const filter = state.activeSheet.filter
  const [draft, setDraft] = useState<Draft | null>(null)
  const [search, setSearch] = useState('')

  // 打开时以该列现有 criteria 初始化草稿
  useEffect(() => {
    if (!open || !filter) {
      setDraft(null)
      setSearch('')
      return
    }
    const crit = filter.criteria[open.col]
    if (crit?.type === 'condition') {
      setDraft({ mode: 'condition', excluded: new Set(), field: crit.field, op: crit.op, v1: crit.v1, v2: crit.v2 ?? '' })
    } else {
      setDraft({ mode: 'values', excluded: new Set(crit?.excluded ?? []), field: 'text', op: 'contains', v1: '', v2: '' })
    }
    setSearch('')
    // 「打开目标列」或 filter 本身（undo 筛选等）变化时重建草稿
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open?.col, open !== null, filter])

  const values = useMemo(() => {
    if (!open || !filter) return []
    const ev = evaluatorFor(state.doc)
    const seen = new Set<string>()
    const out: string[] = []
    for (let row = filter.range.sr + 1; row <= filter.range.er; row++) {
      const t = ev.displayText(state.doc.active, row, open.col)
      if (!seen.has(t)) {
        seen.add(t)
        out.push(t)
      }
    }
    return out
  }, [open, filter, state.doc])

  if (!open || !filter || !draft) return null
  if (open.col < filter.range.sc || open.col > filter.range.ec) return null

  const close = (): void => {
    view.dispatch(view.state.tr.setMeta(filterDropdownKey, null).setMeta('addToHistory', false))
  }

  const applyFilter = (criteria: FilterCriteria | undefined): void => {
    const next: FilterState = { range: filter.range, criteria: { ...filter.criteria } }
    if (criteria) next.criteria[open.col] = criteria
    else delete next.criteria[open.col]
    const tr = view.state.tr.setFilter(next)
    tr.setMeta(filterDropdownKey, null)
    view.dispatch(tr)
    view.focus()
  }

  const applyDraft = (): void => {
    if (draft.mode === 'values') {
      applyFilter(draft.excluded.size > 0 ? { type: 'values', excluded: [...draft.excluded] } : undefined)
    } else {
      const cond: FilterConditionCriteria = { type: 'condition', field: draft.field, op: draft.op, v1: draft.v1 }
      if (draft.op === 'between') cond.v2 = draft.v2
      applyFilter(cond)
    }
  }

  const shown = values.filter((v) => v.toLowerCase().includes(search.toLowerCase()))
  const toggle = (v: string): void => {
    const excluded = new Set(draft.excluded)
    if (excluded.has(v)) excluded.delete(v); else excluded.add(v)
    setDraft({ ...draft, excluded })
  }
  const ops = draft.field === 'text' ? TEXT_OPS : NUM_OPS

  return (
    <div className="fixed inset-0 z-100" onClick={close}>
      <div
        className="fixed flex w-[260px] flex-col gap-1.5 rounded-lg border border-line-strong bg-surface p-2.5 shadow-3"
        style={{ left: Math.min(open.x, window.innerWidth - 280), top: Math.min(open.y, window.innerHeight - PANEL_MAX_H) }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex gap-1">
          <Button variant="ghost" size="sm" active={draft.mode === 'values'} onClick={() => setDraft({ ...draft, mode: 'values' })}>
            值列表
          </Button>
          <Button variant="ghost" size="sm" active={draft.mode === 'condition'} onClick={() => setDraft({ ...draft, mode: 'condition' })}>
            条件
          </Button>
        </div>
        {draft.mode === 'values' ? (
          <>
            <TextInput width={238} placeholder="搜索值" value={search} onChange={setSearch} />
            <div className="flex items-center gap-1.5">
              <Button variant="ghost" size="sm" onClick={() => setDraft({ ...draft, excluded: new Set() })}>
                全选
              </Button>
              <Button variant="ghost" size="sm" onClick={() => setDraft({ ...draft, excluded: new Set(values) })}>
                清空
              </Button>
            </div>
            <div className="flex max-h-[180px] flex-col gap-0.5 overflow-y-auto">
              {shown.map((v) => (
                <label key={v} className="flex h-6 items-center gap-1.5 text-sm">
                  <input type="checkbox" checked={!draft.excluded.has(v)} onChange={() => toggle(v)} />
                  <span>{v === '' ? '（空白）' : v}</span>
                </label>
              ))}
              {shown.length === 0 && <div className="flex items-center text-sm text-ink-2">无匹配值</div>}
            </div>
          </>
        ) : (
          <>
            <div className="flex items-center gap-1.5">
              <Select
                value={draft.field}
                options={[
                  { value: 'text', label: '文本' },
                  { value: 'num', label: '数值' },
                ]}
                onChange={(v) => setDraft({ ...draft, field: v as 'text' | 'num', op: v === 'text' ? 'contains' : 'eq' })}
              />
              <Select
                value={draft.op}
                options={ops.map((o) => ({ value: o.op, label: o.label }))}
                onChange={(v) => setDraft({ ...draft, op: v as FilterOp })}
              />
            </div>
            <div className="flex items-center gap-1.5">
              <TextInput
                width={draft.op === 'between' ? 116 : 238}
                placeholder="值"
                value={draft.v1}
                onChange={(v) => setDraft({ ...draft, v1: v })}
              />
              {draft.op === 'between' && (
                <TextInput
                  width={116}
                  placeholder="至"
                  value={draft.v2}
                  onChange={(v) => setDraft({ ...draft, v2: v })}
                />
              )}
            </div>
          </>
        )}
        <div className="mt-1 flex justify-end gap-2">
          <Button variant="outline" onClick={() => applyFilter(undefined)}>
            清除该列
          </Button>
          <Button variant="outline" onClick={close}>
            关闭
          </Button>
          <Button variant="primary" onClick={applyDraft}>
            确定
          </Button>
        </div>
      </div>
    </div>
  )
}
