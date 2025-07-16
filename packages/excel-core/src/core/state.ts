import { SheetData, Workbook } from './model'
import { Plugin, PluginKey } from './plugin'
import { Selection, singleCell } from './selection'
import { Transaction } from './transaction'

const MAX_APPEND_DEPTH = 10

export class SheetState {
  readonly doc: Workbook
  readonly selection: Selection
  readonly plugins: readonly Plugin[]
  private readonly fields: ReadonlyMap<PluginKey, any>

  private constructor(
    doc: Workbook,
    selection: Selection,
    plugins: readonly Plugin[],
    fields: ReadonlyMap<PluginKey, any>,
  ) {
    this.doc = doc
    this.selection = selection
    this.plugins = plugins
    this.fields = fields
  }

  static create(config: { doc: Workbook; selection?: Selection; plugins?: Plugin[] }): SheetState {
    const plugins = config.plugins ?? []
    const selection = config.selection ?? singleCell(0, 0)
    let state = new SheetState(config.doc, selection, plugins, new Map())
    // 按插件序 init state field（后 init 的插件可见前面字段的值）
    for (const p of plugins) {
      if (!p.spec.state) continue
      const fields = new Map(state.fields)
      fields.set(p.key, p.spec.state.init({ doc: config.doc }, state))
      state = new SheetState(config.doc, selection, plugins, fields)
    }
    return state
  }

  get activeSheet(): SheetData {
    return this.doc.activeSheet
  }

  get tr(): Transaction {
    return new Transaction(this)
  }

  getField(key: PluginKey): any {
    return this.fields.get(key)
  }

  apply(tr: Transaction): SheetState {
    // 逐 step 应用到 doc；任一 failed → 抛错
    let doc = this.doc
    for (const step of tr.steps) {
      const r = step.apply(doc)
      if (!r.ok) throw new Error('step failed: ' + r.failed)
      doc = r.doc!
    }
    const selection = tr.selection ?? this.selection
    // 先搭不含 field 更新的骨架，再按插件序跑 field.apply
    const skeleton = new SheetState(doc, selection, this.plugins, this.fields)
    const fields = new Map(this.fields)
    for (const p of this.plugins) {
      if (!p.spec.state) continue
      fields.set(p.key, p.spec.state.apply(tr, fields.get(p.key), this, skeleton))
    }
    return new SheetState(doc, selection, this.plugins, fields)
  }

  // apply 后依次调插件 appendTransaction，返回 tr 则递归应用（≤10 层防死循环）
  applyTransaction(tr: Transaction): { state: SheetState; trs: Transaction[] } {
    const trs: Transaction[] = [tr]
    let state = this.apply(tr)
    for (let depth = 0; depth < MAX_APPEND_DEPTH; depth++) {
      let appended: Transaction | null | undefined
      for (const p of this.plugins) {
        const fn = p.spec.appendTransaction
        if (!fn) continue
        const res = fn(trs, this, state)
        if (res) {
          appended = res
          break
        }
      }
      if (!appended) return { state, trs }
      trs.push(appended)
      state = state.apply(appended)
    }
    return { state, trs }
  }
}
