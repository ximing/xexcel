import { Workbook } from './model'
import type { SheetState } from './state'
import type { Transaction } from './transaction'

let keySeq = 0

export class PluginKey {
  readonly name: string

  constructor(name?: string) {
    // 每个实例保证唯一名（同名前缀不冲突）
    this.name = (name ?? 'plugin') + '$' + ++keySeq
  }

  getState(state: SheetState): any {
    return state.getField(this)
  }
}

export interface StateField<T> {
  init(config: { doc: Workbook }, state: SheetState): T
  apply(tr: Transaction, value: T, oldState: SheetState, newState: SheetState): T
}

// HitResult 最小结构类型在 core 声明，view/types.ts 再扩展，避免循环依赖
export interface HitResult {
  region: 'cell' | 'rowheader' | 'colheader' | 'corner' | 'fillhandle' | 'outside'
  row: number
  col: number
}

// core 不依赖 view 层：EditorView 的最小鸭子类型，view/editorview.ts 实现
export interface EditorViewLike {
  readonly state: SheetState
  dispatch(tr: Transaction): void
  someProp(name: keyof PluginProps, fn: (prop: any) => boolean): boolean
}

export interface PluginProps {
  handleKeyDown?: (view: EditorViewLike, event: KeyboardEvent) => boolean
  handleMouseDown?: (view: EditorViewLike, event: MouseEvent, hit: HitResult) => boolean
  handleMouseMove?: (view: EditorViewLike, event: MouseEvent, hit: HitResult) => boolean
  handleMouseUp?: (view: EditorViewLike, event: MouseEvent, hit: HitResult) => boolean
  handleDoubleClick?: (view: EditorViewLike, event: MouseEvent, hit: HitResult) => boolean
  handlePaste?: (view: EditorViewLike, text: string) => boolean
  handleCopy?: (view: EditorViewLike, cut: boolean, event: ClipboardEvent) => boolean
}

export interface PluginView {
  update?(view: EditorViewLike, prevState: SheetState): void
  destroy?(): void
}

export interface PluginSpec {
  key?: PluginKey
  state?: StateField<any>
  props?: PluginProps
  appendTransaction?: (
    trs: readonly Transaction[],
    oldState: SheetState,
    newState: SheetState,
  ) => Transaction | null | undefined
  view?: (view: EditorViewLike) => PluginView | void
}

export class Plugin {
  readonly key: PluginKey
  readonly spec: PluginSpec

  constructor(spec: PluginSpec) {
    this.spec = spec
    this.key = spec.key ?? new PluginKey()
  }
}
