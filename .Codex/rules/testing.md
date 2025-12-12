---
paths:
  - "tests/**/*.ts"
  - "**/*.test.ts"
---

# 测试规则

- 测试框架为 Vitest（`npm test` = `vitest run`）；测试文件集中放 `tests/`，命名 `模块-主题.test.ts`（如 `core-addr.test.ts`）。
- 单测只覆盖 `src/core`、`src/formula` 与纯函数型 view 工具（如 geometry）；不引入 DOM 环境测试。
- `tests/core-nodep.test.ts` 是架构守门测试：断言 core/formula 不 import DOM/上层模块，改动依赖结构时必须保持其通过。
- 核心/公式改动要覆盖 `docs/architecture.md` 里那几类：A1/colName 互逆、range 工具、setCell 不可变性、Step apply/invert 往返、tr 多步 doc 链、插件 field init/apply 次序、appendTransaction 展开、undo/redo 往返与 200 组上限、公式优先级/错误传播/`#CYCLE!`/`#REF!` 等。
- 改完：`pnpm -r typecheck` 零错误，`pnpm -r test` 全绿，`pnpm -r build` 成功。
