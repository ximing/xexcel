# 参与 xexcel

[English](./CONTRIBUTING.md)

感谢你愿意看一眼。xexcel 是一个小的电子表格内核。有用的 PR 通常是：缺的公式、更清楚的报错、文档修正、或范围很窄的交互 bug。

## 环境

需要 Node 22+ 和 [pnpm](https://pnpm.io/) 9。

```bash
git clone https://github.com/ximing/xexcel.git
cd xexcel
pnpm install
pnpm dev          # 演示在 http://localhost:5173
```

每次改动的完成标准：

```bash
pnpm -r typecheck
pnpm -r test
pnpm -r build
```

三项都必须干净。不要引入 DOM 测试环境；单测放在 `packages/*/tests/`，只覆盖 core、formula 和纯函数型 view 工具。

## 目录

```
packages/excel-core    @xexcel/core     模型 / Step / 事务 / 公式 / io
packages/excel-view    @xexcel/view     Konva 画布 + 交互插件
packages/excel-react   @xexcel/react    React 外壳
apps/demo                               在线演示
```

跨包 import 一律用包名（`@xexcel/core`），禁止跨包相对路径。

硬约束（详见 `AGENTS.md` 和各层 `CLAUDE.md`）：

1. 文档修改只能经 Transaction / Step。
2. 行列索引一律 0-based number；A1 只在 addr 与 UI 显示层转换。
3. 标识符用英文，注释用简洁中文，只注释非显然之处。
4. 画布尺寸常量从 `@xexcel/core` 导出，不散落魔术数字。

## Pull request

1. 一个 PR 只做一件事。
2. 在现有 `tests/<模块>-<主题>.test.ts` 旁补测试。
3. 禁止复制其他电子表格项目的源码、命名或文件组织。
4. 按 PR 模板勾完成项。

适合上手的任务标了 [`good first issue`](https://github.com/ximing/xexcel/labels/good%20first%20issue)。

## 报 bug

用 bug 模板，写清操作、预期、实际、浏览器 / Node 版本，能附最小工作簿或公式最好。

安全问题见 [SECURITY.md](./SECURITY.md)，不要开公开 issue。
