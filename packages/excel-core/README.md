# @xexcel/core

Headless spreadsheet engine for [xexcel](https://github.com/ximing/xexcel). Immutable workbook, reversible Steps, Transaction + Plugin, formula evaluator, CSV/xlsx I/O. Zero DOM — works in Node.

```ts
import { Workbook, SheetState, history, evaluatorFor } from '@xexcel/core'

const wb = Workbook.create({ rowCount: 100, colCount: 26 })
const state = SheetState.create({ doc: wb, plugins: [history()] })
const next = state.applyTransaction(state.tr.setCell(0, 0, '=1+2')).state
evaluatorFor(next.doc).get(next.doc.active, 0, 0) // 3
```

See the [root README](https://github.com/ximing/xexcel#readme) for the full editor.
