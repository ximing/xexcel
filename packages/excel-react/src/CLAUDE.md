# src/react — React 外壳

## 这个目录负责什么

`ExcelEditor`、Toolbar / FormulaBar / StatusBar / SheetTabBar、对话框、`ui/` 组件、`bridge.ts`。

## 放置约束

- 放：受控 UI 和状态订阅。
- 不放：几何和命中（那是 view）；文档数据副本；绕过 dispatch 的写路径。

## 开发偏好

- `bridge.ts` 用 `useSyncExternalStore` 订 EditorView。组件不要自己存一份表格。
- 写操作走 `view.dispatch(tr)`，不要直接调 core 的 Step / model。
- 内部状态只留纯 UI（输入草稿、菜单开闭）。
- 不用 react-konva。
