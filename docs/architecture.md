# xexcel 架构

2023-11 起稿，之后按实现改过几轮。仓库现在的结构和这里不一致的地方，以代码为准。

## 要做成什么样

浏览器里能嵌的 Excel 风格表格，内核跟常见的「React 里一堆格子组件」不是一条路。状态是唯一真相，视图只负责画和收事件。改数据只走事务。

想法来自 ProseMirror：`State` + `Transaction` + `Plugin`。表格不是文档模型，但这套拆法够用。

- 行列索引全是 0-based。A1 只在公式和 UI 上出现。
- core / formula 不碰 DOM。
- 包之间单向依赖：`@xexcel/react` → `@xexcel/view` → `@xexcel/core`。

## 数据流

```
指针/键盘/工具栏
  → EditorView（Konva + 少量 DOM）
  → 插件 props 链（谁返回 true 谁吃掉）
  → 拼 Transaction（Step + 选区 + meta）
  → view.dispatch(tr)
  → state.apply(tr)
      跑 Step 得到新 Workbook
      跑各插件 field.apply
      appendTransaction 最多再展开 10 层
  → 画布重绘；React 外壳走 subscribe
```

插件、视图、React 组件都不能直接改 `doc`。

## 分层

现在是 pnpm workspace，三个包加一个 demo：

| 包 | 干什么 |
|---|---|
| `@xexcel/core` | 地址、不可变表、Step、事务、插件框架、历史、公式、CSV/xlsx |
| `@xexcel/view` | 命令式 Konva 画布，加上选择/键盘/剪贴板/填充柄/拖动等插件 |
| `@xexcel/react` | `<ExcelEditor/>`、工具栏、公式栏、状态栏、对话框、UI 组件 |
| `apps/demo` | 演示页，GitHub Pages 也是它 |

core 里 `tests/core-nodep.test.ts` 盯着 import，别把 DOM 或上层包引进去。

## 模型

一开始就按多表做，没有「先单表再返工」的阶段。

```ts
Workbook {
  sheets: Map<SheetId, SheetData>
  order: SheetId[]
  active: SheetId
  names: Map<SheetId, string>   // 显示名，跟 id 分开
}

SheetData {
  稀疏单元格
  自定义行高列宽
  merges / freeze / hidden / filter / 条件格式 / 数据验证
}
```

`SheetData.setCell` 这类方法都返回新对象，行级结构共享。空 `raw` 且没有样式就当删格。

画布尺寸写在 `packages/excel-core/src/core/model.ts`：

- 行头宽 48、列头高 26
- 默认行高 24、列宽 96

别处不要再写一遍这些数。

## Step 和事务

文档变更必须是 Step。Step 要能 `apply` / `invert` / `toJSON`，invert 拿的是这一步之前的文档。`stepFromJSON` 漏分支会让持久化和历史对不上。

常见几类：写单元格、改样式、改行高列宽、插删行列、合并、隐藏、多表操作、筛选、条件格式、数据验证。

`Transaction` 是 builder，叠一组 Step，还可以带选区和 meta。`SheetState.applyTransaction` 会把插件的 `appendTransaction` 展开，防止死循环有 10 层顶。

## 插件

插件可以挂：

- `state`：跟文档一起走的 field
- `props`：键盘鼠标剪贴板，返回 true 表示拦截
- `appendTransaction`
- `view`：跟 EditorView 生命周期走，不算 React

历史是插件，`done` 最多 200 组。`tr.setMeta('addToHistory', false)` 的事务不进栈。内建插件由 `@xexcel/view` 的 `builtinPlugins()` 组装，history 由调用方插到最前。

## 公式

递归下降。优先级从低到高：比较 → `&` → 加减 → 乘除 → `^`（右结合）→ 一元和 `%` → 原子。

空格子在算术和比较里当 0。`/0` 是 `#DIV/0!`，未知函数 `#NAME?`，环 `#CYCLE!`（环上每个格子都标），越界 `#REF!`。比较序跟 Excel：布尔 > 字符串 > 数字。

`CellEvaluator` 按 Workbook 做 WeakMap 缓存，外面不要再包一层。

跨表是 `Sheet2!A1` 这种。相对引用在填充和粘贴时要偏移，绝对/`$` 不动。

## 视图

`EditorView` 自己建 Konva.Stage，不走 react-konva。三层：网格、单元格、覆盖层（选区、填充柄、冻结线）。只画可见 range，滚动复用节点。

编辑格子用 DOM textarea 盖在画布上，输入法不能走 Konva。

冻结行列的偏移从 geometry 一直传到 hitTest 和绘制。滚动条是自己画的。

React 侧只有 `bridge.ts`：`useSyncExternalStore` 订 EditorView。组件不准自己存一份表格数据。

## 测试

- 单测：vitest，文件在各包 `tests/`，名字像 `core-addr.test.ts`
- core/formula 要覆盖地址互逆、不可变更新、Step 往返、事务链、插件 field 顺序、undo/redo、公式优先级和那几个错误码
- 浏览器脚本在 `apps/demo/e2e/`，不替代单测

改完至少：`pnpm -r typecheck`、`pnpm -r test`、`pnpm -r build`。
