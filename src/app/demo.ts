// demo 数据：spec §7。1000×26，表头/公式/合计/IF 列/H2/I2 逐格规格。
import { history } from '@gmi/excel-core'
import { CellStyle, Workbook } from '@gmi/excel-core'
import { SheetState } from '@gmi/excel-core'
import { builtinPlugins } from '../plugins'

const HEADER_STYLE: CellStyle = { bold: true, bg: '#e8f0fe', align: 'center' }

// 产品/单价/数量：金额 D=B*C 有 >1000 与 ≤1000 的混合，便于演示 IF 列
const ROWS: [string, number, number][] = [
  ['笔记本电脑', 5999, 12],
  ['鼠标', 79, 45],
  ['键盘', 199, 30],
  ['显示器', 1299, 8],
  ['耳机', 299, 25],
  ['数据线', 19, 60],
  ['充电器', 49, 15],
  ['手机壳', 29, 20],
  ['平板', 2499, 5],
  ['音箱', 399, 3],
]

export function createDemoState(): SheetState {
  let wb = Workbook.create({ rowCount: 1000, colCount: 26 })
  let sheet = wb.activeSheet
  const set = (row: number, col: number, raw: string, style?: CellStyle): void => {
    sheet = sheet.setCell(row, col, style ? { raw, style } : { raw })
  }

  const headers = ['产品', '单价', '数量', '金额']
  headers.forEach((h, col) => set(0, col, h, HEADER_STYLE))

  ROWS.forEach(([name, price, qty], i) => {
    const r = i + 1 // 数据行 2..11（0-based 1..10）
    set(r, 0, name)
    set(r, 1, String(price))
    set(r, 2, String(qty))
    set(r, 3, `=B${r + 1}*C${r + 1}`)
    set(r, 5, `=IF(D${r + 1}>1000,"达标","未达标")`)
  })

  set(12, 1, '合计', { bold: true })
  set(12, 3, '=SUM(D2:D11)')
  set(1, 7, '=AVERAGE(B2:B11)')
  set(1, 8, '=ROUND(H2,1)')

  wb = wb.setSheet(wb.active, sheet)
  return createStateFromWorkbook(wb)
}

export function createStateFromWorkbook(wb: Workbook): SheetState {
  return SheetState.create({ doc: wb, plugins: [history(), ...builtinPlugins()] })
}
