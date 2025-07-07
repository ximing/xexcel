# xexcel

浏览器端 Excel 风格电子表格，架构仿 ProseMirror——一切皆状态（State）+ 事务（Transaction）+ 插件（Plugin），视图只是状态的投影。

## 特性

- 编辑与公式：单元格编辑 / 公式引擎（lexer / parser / eval，跨表引用）/ 数字格式
- 样式：字体 / 字号 / 粗斜体 / 下划线删除线 / 对齐 / 换行 / 文字与背景色 / 边框（8 预设 8 线型）/ 格式刷
- 结构：多工作表（增删改名拖动排序）/ 行列插删 / 隐藏 / 行高列宽 / 合并单元格 / 冻结
- 数据：排序（快捷 + 自定义多键）/ 自动筛选 / 查找替换 / 条件格式 / 数据验证（数值范围 / 文本长度 / 序列）
- 互操作：CSV 导入导出 / xlsx 导入导出（exceljs，值 / 公式 / 样式 / 合并 / 冻结 / 筛选 / 条件格式 / 验证全映射）/ 浏览器自动保存（localStorage 防抖 + 损坏自愈）
- 交互：多区域选择 / 拖拽移动 / 富剪贴板 / 行列填充柄 / 右键菜单 / 快捷键

## 架构

分层依赖单向 `react → view → core/formula`：

- `src/core` 纯数据与状态（model / steps / transaction / state / plugin / history / commands），零 DOM 依赖
- `src/formula` 公式引擎，零 DOM 依赖
- `src/view` 命令式 Konva 视图
- `src/plugins` 内建交互插件
- `src/react` React 外壳 + `ui/` 通用组件层（设计系统）
- `src/app` 应用入口

所有文档修改只能经 Transaction/Step；视图 / 插件 / React 永不直接改 doc。

## 技术栈

TypeScript(strict) / Vite / React 18 / Konva（命令式）/ Tailwind v4（设计 token @theme）/ Lucide 图标 / exceljs / Vitest

## 开发

| 命令 | 说明 |
|------|------|
| `npm run dev` | 启动 Vite 开发服务器 |
| `npm test` | 运行 vitest 单元测试 |
| `npm run typecheck` | tsc --noEmit 类型检查 |
| `npm run build` | 类型检查 + 生产构建 |
| `npm run test:e2e` | 真实浏览器回归（浏览器本机 Chrome，38 场景） |

测试现状：533 单测（72 文件）+ 38 e2e 场景。

## 设计规范

M5 建立：设计 token 单源（`src/app/theme.css` @theme：色板 / 圆角 / 阴影 / 字号）+ `src/react/ui/` 组件层（Icon / Tooltip / IconButton / Button / Menu / Dropdown / Dialog / ConfirmDialog / Select / TextInput），画布取色镜像 `src/view/theme.ts`。

## 文档

- 设计规格：`docs/superpowers/specs/`（M1-M5 各里程碑）
- 实施计划：`docs/superpowers/plans/`
- 项目指令：`CLAUDE.md` 及各分层 CLAUDE.md
