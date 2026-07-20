# xexcel 项目指令

## 项目概览

- 浏览器端 Excel 风格电子表格。内核仿 ProseMirror：状态（State）+ 事务（Transaction）+ 插件（Plugin），视图只是状态的投影。
- 技术栈：TypeScript(strict) + Vite + React 19 + Konva（命令式，不经 react-konva）+ Vitest。xlsx 用 exceljs；界面 token 用 Tailwind v4，图标 Lucide。
- 设计说明：`docs/architecture.md`，分主题笔记在 `docs/design/`。

## 分层与依赖方向

pnpm workspace：

```
packages/
  excel-core/   @xexcel/core：core + formula + io，零 DOM
  excel-view/   @xexcel/view：Konva 视图 + 内建交互插件
  excel-react/  @xexcel/react：ExcelEditor 和外壳
apps/
  demo/         演示应用（Pages 部署源）
```

依赖单向：excel-react → excel-view → excel-core。包间 import 用包名，不要跨包相对路径。

## 全局规则

1. 命名和模块划分按本仓库现有结构来，不要按别的表格项目套。
2. 所有对文档的修改只经 Transaction/Step；视图、插件、React 组件不直接改 doc。
3. 行列索引一律 0-based；A1 只在 addr/公式层与 UI 显示层转换。
4. 代码标识符英文，注释用简洁中文，只注释非显然之处。
5. commit message：`feat(core): ...` / `feat(view): ...` 这种。
6. 画布逻辑尺寸常量（`ROW_HEADER_WIDTH` / `COL_HEADER_HEIGHT` / `DEFAULT_ROW_HEIGHT` / `DEFAULT_COL_WIDTH`）从 `packages/excel-core/src/core/model.ts` 导出，不要散落魔术数字。

## 开发入口

- 开发：`pnpm dev`
- 测试：`pnpm -r test`
- 类型检查：`pnpm -r typecheck`
- 构建：`pnpm -r build`
- 改完要 typecheck 零错误、test 全绿、build 成功。

## 局部规则

- `packages/excel-core/src/core/CLAUDE.md` — 内核与事务
- `packages/excel-core/src/formula/CLAUDE.md` — 公式
- `packages/excel-view/src/view/CLAUDE.md` — Konva 视图
- `packages/excel-view/src/plugins/CLAUDE.md` — 插件
- `packages/excel-react/src/CLAUDE.md` — React 外壳

横切：

- `.claude/rules/testing.md`
- `.claude/rules/react-components.md`
