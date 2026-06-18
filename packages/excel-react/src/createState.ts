// 应用级 state 装配：history + 内建交互插件。demo 数据工厂与 xlsx 导入共用。
import { history, SheetState } from '@xexcel/core'
import type { Workbook } from '@xexcel/core'
import { builtinPlugins } from '@xexcel/view'

export function createStateFromWorkbook(wb: Workbook): SheetState {
  return SheetState.create({ doc: wb, plugins: [history(), ...builtinPlugins()] })
}
