# @xexcel/react

React shell for [xexcel](https://github.com/ximing/xexcel): `<ExcelEditor/>`, toolbar, formula bar, status bar, dialogs.

```tsx
import { Workbook } from '@xexcel/core'
import { ExcelEditor, createStateFromWorkbook } from '@xexcel/react'
import '@xexcel/react/styles.css'

<ExcelEditor state={createStateFromWorkbook(Workbook.create({ rowCount: 1000, colCount: 26 }))} />
```

Peers: `@xexcel/core`, `@xexcel/view`, `react`, `react-dom`.
