---
paths:
  - "src/**/*.tsx"
---

# React 组件规则

- 组件保持纯受控：props 进、事件出，内部状态仅限纯 UI 态。
- 文档数据一律经 `bridge.ts` 的 `useSyncExternalStore` 订阅 EditorView 快照；**组件内不得持有文档数据副本**。
- 写操作回走 `view.dispatch(tr)`，不直接调用 core 的 Step/model API。
- 不使用 react-konva；Konva 只允许出现在 `src/view`。
- 命名：组件文件用 PascalCase（`Toolbar.tsx`），非组件模块用 camelCase（`bridge.ts`）。
