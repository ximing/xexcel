# src/core — 内核：纯数据与状态

## 这个目录负责什么

地址工具、不可变模型（SheetData / Workbook）、Step / Transaction / State、插件抽象、历史、通用命令。

## 放置约束

- 放：纯 TypeScript 的数据结构与状态机。
- 不放：DOM / Konva / React，以及渲染、事件、样式。
- 新的写能力先做成 Step，再经 Transaction 暴露，不要另开一条直接改文档的口子。

## 开发偏好

- 零 DOM 依赖。不要 import DOM 类型或 view / react。`tests/core-nodep.test.ts` 盯着这件事。
- `SheetData` / `Workbook` 的更新返回新对象，行级结构共享。
- Step 要实现 `apply` / `invert` / `toJSON`。invert 用的是这一步之前的文档。新 Step 记得补 `stepFromJSON` 和往返测试。
- `appendTransaction` 最多 10 层；history 最多 200 组；`tr.setMeta('addToHistory', false)` 不入栈。
