# src/formula — 公式引擎

## 这个目录负责什么

词法、递归下降语法、求值、带缓存的 `CellEvaluator`。

## 放置约束

- 放：纯计算。
- 不放：DOM / 渲染。可以依赖 core 的只读数据接口，不要 import view。

## 开发偏好

- 零 DOM，同 core。
- 优先级从低到高：比较 → `&` → 加减 → 乘除 → `^`（右结合）→ 一元 / `%` → 原子。改语法要改测试。
- 空单元格按 0；`/0` → `#DIV/0!`；未知函数 → `#NAME?`；环上格子全部 `#CYCLE!`；越界 → `#REF!`。比较序：布尔 > 字符串 > 数字。
- `evaluatorFor` 按 Workbook 做 WeakMap 缓存，调用侧不要再包一层。
- 数字显示走 `toString`，超过 12 个字符再用 `toPrecision(10)` 去尾零。布尔显示 `TRUE` / `FALSE`。
