# xexcel

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![npm](https://img.shields.io/npm/v/@xexcel/react.svg)](https://www.npmjs.com/package/@xexcel/react)
[![CI](https://github.com/ximing/xexcel/actions/workflows/ci.yml/badge.svg)](https://github.com/ximing/xexcel/actions/workflows/ci.yml)
[![在线演示](https://img.shields.io/badge/demo-GitHub%20Pages-2ea44f)](https://ximing.github.io/xexcel/)

[English](./README.md) | 简体中文

可嵌入的浏览器端 Excel 风格电子表格。架构仿 ProseMirror：**一切皆状态（State）+ 事务（Transaction）+ 插件（Plugin）**，视图只是状态的投影。

[在线演示](https://ximing.github.io/xexcel/) · [npm `@xexcel/react`](https://www.npmjs.com/package/@xexcel/react) · [贡献指南](./CONTRIBUTING.zh-CN.md)

![xexcel 界面截图](docs/images/screenshot.png)

## 安装

```bash
pnpm add @xexcel/react @xexcel/view @xexcel/core
# 或：npm i @xexcel/react @xexcel/view @xexcel/core
```

```tsx
import { Workbook } from '@xexcel/core'
import { ExcelEditor, createStateFromWorkbook } from '@xexcel/react'
import '@xexcel/react/styles.css'

export function App() {
  const state = createStateFromWorkbook(
    Workbook.create({ rowCount: 1000, colCount: 26 }),
  )
  return <ExcelEditor state={state} locale="zh" />
}
```

`@xexcel/core` 零 DOM，可在 Node 里跑公式和 xlsx/CSV。`@xexcel/view` 是 Konva 画布。`@xexcel/react` 是编辑器外壳。

## 特性

- **编辑与公式**：单元格编辑 / 公式引擎（lexer / parser / eval，跨表引用）/ 数字格式
- **样式**：字体 / 字号 / 粗斜体 / 下划线删除线 / 对齐 / 换行 / 文字与背景色 / 边框（8 预设 × 8 线型）/ 格式刷
- **结构**：多工作表（增删改名拖动排序）/ 行列插删 / 隐藏 / 行高列宽 / 合并单元格 / 冻结
- **数据**：排序（快捷 + 自定义多键）/ 自动筛选 / 查找替换 / 条件格式 / 数据验证（数值范围 / 文本长度 / 序列）
- **互操作**：CSV 导入导出 / xlsx 导入导出（exceljs，值 / 公式 / 样式 / 合并 / 冻结 / 筛选 / 条件格式 / 验证全映射）/ 浏览器自动保存（localStorage 防抖 + 损坏自愈）
- **交互**：多区域选择 / 拖拽移动 / 富剪贴板 / 行列填充柄 / 右键菜单 / 快捷键

## 明确不做

xexcel 是**小而可读的可嵌入内核**，不是办公套件。1.0 不包含：

- 实时协作
- 图表、数据透视、迷你图
- VBA / 宏 / 动态数组公式
- 完整 Excel 函数库（目前是常用函数：`SUM` `AVERAGE` `COUNT` `MIN` `MAX` `IF` `SUMIF` `COUNTIF` `AVERAGEIF` `ABS` `ROUND` 等）
- 打印排版
- 移动端优先的触控编辑

需要协作、图表和完整函数库，看 [Univer](https://github.com/dream-num/univer)。需要从 Luckysheet 长出来的开箱 Excel 外观，看 [FortuneSheet](https://github.com/ruilisi/fortune-sheet)。xexcel 面向「要把一张表嵌进自己的产品，并且读得懂、改得动内核」的团队。

## 包结构

pnpm workspace，依赖单向 `react → view → core`。

| 包 | 职责 |
|---|---|
| [`@xexcel/core`](./packages/excel-core) | 不可变模型、Step、事务、插件框架、公式引擎、xlsx/CSV。零 DOM |
| [`@xexcel/view`](./packages/excel-view) | 命令式 Konva 视图 + 内建交互插件 |
| [`@xexcel/react`](./packages/excel-react) | `<ExcelEditor/>` + 工具栏 / 公式栏 / 状态栏 + ui 组件层 |
| `apps/demo` | 演示应用（GitHub Pages 部署源） |

所有文档修改只能经 Transaction/Step；视图、插件、React 组件永不直接改 doc。

## 快捷键

`Mod` 在 macOS 是 ⌘，Windows/Linux 是 Ctrl。复制 / 剪切 / 粘贴走浏览器事件（`Mod+C` / `Mod+X` / `Mod+V`）；填充柄只有鼠标。

| 按键 | 作用 |
|---|---|
| 方向键 | 移动活动单元格 |
| Shift+方向键 | 扩展选区 |
| Tab / Shift+Tab | 右移 / 左移 |
| Enter / Shift+Enter | 下移 / 上移 |
| Delete / Backspace | 清除选区 |
| F2 | 编辑活动单元格 |
| 可打印字符 | 开编辑并替换内容 |
| Mod+A | 全选 |
| Mod+F | 查找 / 替换 |
| Mod+Z | 撤销 |
| Mod+Shift+Z 或 Ctrl+Y | 重做 |
| Esc | 取消格式刷 |

`<ExcelEditor locale="en" />` 把工具栏 / 文件菜单 / 右键菜单 / 状态栏切成英文。默认 `zh`。

## 开发

| 命令 | 说明 |
|---|---|
| `pnpm dev` | 启动 demo Vite 开发服务器 |
| `pnpm -r test` | 运行 vitest 单元测试 |
| `pnpm -r typecheck` | `tsc --noEmit` 类型检查 |
| `pnpm -r build` | 类型检查 + 生产构建 |
| `pnpm test:e2e` | 真实浏览器回归（38 场景） |

测试现状：536 单测（72 文件）+ 38 e2e 场景。

贡献流程见 [CONTRIBUTING.zh-CN.md](./CONTRIBUTING.zh-CN.md)，发版见 [docs/publishing.md](./docs/publishing.md)。

## 许可证

[MIT](./LICENSE)
