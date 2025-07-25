// 应用级 state 装配：history + 内建交互插件。demo 数据工厂与 xlsx 导入共用。
import { history, SheetState } from '@gmi/excel-core'
import type { Workbook } from '@gmi/excel-core'
import { builtinPlugins } from '@gmi/excel-view'

export function createStateFromWorkbook(wb: Workbook): SheetState {
  return SheetState.create({ doc: wb, plugins: [history(), ...builtinPlugins()] })
}
