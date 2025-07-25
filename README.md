# xexcel

浏览器端 Excel 风格电子表格，架构仿 ProseMirror——一切皆状态（State）+ 事务（Transaction）+ 插件（Plugin），视图只是状态的投影。

## 在线演示

<https://ximing.github.io/xexcel/>

![xexcel 界面截图](docs/images/screenshot.png)

## 特性

- 编辑与公式：单元格编辑 / 公式引擎（lexer / parser / eval，跨表引用）/ 数字格式
- 样式：字体 / 字号 / 粗斜体 / 下划线删除线 / 对齐 / 换行 / 文字与背景色 / 边框（8 预设 8 线型）/ 格式刷
- 结构：多工作表（增删改名拖动排序）/ 行列插删 / 隐藏 / 行高列宽 / 合并单元格 / 冻结
- 数据：排序（快捷 + 自定义多键）/ 自动筛选 / 查找替换 / 条件格式 / 数据验证（数值范围 / 文本长度 / 序列）
- 互操作：CSV 导入导出 / xlsx 导入导出（exceljs，值 / 公式 / 样式 / 合并 / 冻结 / 筛选 / 条件格式 / 验证全映射）/ 浏览器自动保存（localStorage 防抖 + 损坏自愈）
- 交互：多区域选择 / 拖拽移动 / 富剪贴板 / 行列填充柄 / 右键菜单 / 快捷键

## 架构

pnpm workspace monorepo，分层依赖单向 `excel-react → excel-view → excel-core`：

- `packages/excel-core` 纯数据与状态（model / steps / transaction / state / plugin / history / commands）+ 公式引擎 + xlsx/CSV 互操作，零 DOM 依赖
- `packages/excel-view` 命令式 Konva 视图 + 内建交互插件
- `packages/excel-react` React 外壳 + `ui/` 通用组件层（设计系统）
- `apps/demo` 演示应用（Pages 部署源）

所有文档修改只能经 Transaction/Step；视图 / 插件 / React 永不直接改 doc。

## SDK 集成

仓库为 pnpm workspace monorepo，SDK 三包可独立编译（ESM + d.ts），外部系统经 npm pack / file: 引用集成：

- `@gmi/excel-core`：零 DOM 核心（模型/事务/插件框架/公式引擎/xlsx/CSV 互操作），Node 可用
- `@gmi/excel-view`：Konva 画布视图 + 内建交互插件
- `@gmi/excel-react`：React 外壳（`<ExcelEditor/>` + 工具栏/公式栏/状态栏 + ui 组件层）

```tsx
import { Workbook } from '@gmi/excel-core'
import { ExcelEditor, createStateFromWorkbook } from '@gmi/excel-react'
import '@gmi/excel-react/styles.css'

<ExcelEditor state={createStateFromWorkbook(Workbook.create({ rowCount: 1000, colCount: 26 }))} />
```

## 技术栈

TypeScript(strict) / Vite / React 18 / Konva（命令式）/ Tailwind v4（设计 token @theme）/ Lucide 图标 / exceljs / Vitest

## 开发

| 命令 | 说明 |
|------|------|
| `pnpm dev` | 启动 demo Vite 开发服务器 |
| `pnpm -r test` | 运行 vitest 单元测试 |
| `pnpm -r typecheck` | tsc --noEmit 类型检查 |
| `pnpm -r build` | 类型检查 + 生产构建 |
| `node apps/demo/e2e/run.mjs` | 真实浏览器回归（浏览器本机 Chrome，38 场景） |

测试现状：533 单测（72 文件）+ 38 e2e 场景。

## 设计规范

M5 建立：设计 token 单源（`packages/excel-react/src/theme.css` @theme：色板 / 圆角 / 阴影 / 字号）+ `packages/excel-react/src/ui/` 组件层（Icon / Tooltip / IconButton / Button / Menu / Dropdown / Dialog / ConfirmDialog / Select / TextInput），画布取色镜像 `packages/excel-view/src/view/theme.ts`。

## 文档

- 设计规格：`docs/superpowers/specs/`（M1-M5 各里程碑）
- 实施计划：`docs/superpowers/plans/`
- 项目指令：`CLAUDE.md` 及各分层 CLAUDE.md
