# M2a 浏览器验收（浏览器）

前置：`npm run dev` 启动；dev 下 `window.__xcell` 为 EditorView。
读格显示值：`__xcell.state` 配合 `evaluatorFor` 不便直接用时，可在控制台执行
`__xcell.state.activeSheet.getCell(r,c)?.raw` 读 raw；显示值目测网格即可。
每步后附【预期】。失败即终止并回报。

## 1. 填充公式引用偏移
- 先用 Ctrl+A / Delete 清空 demo 数据，再在 A1 输入 `1`，B1 输入 `=A1*2`，选中 B1 拖填充柄到 B3
- 【预期】B2 raw `=A2*2`、B3 raw `=A3*2`（A2/A3 为空，显示 0——空单元格按 0 求值）；undo 一次全部撤销

## 2. $ 锁定
- C1 输入 `=$A$1*2`，拖到 C3 → C2/C3 raw 均为 `=$A$1*2`
- D1 输入 `=A$1*2`，向右拖到 E1 → E1 raw `=B$1*2`

## 3. 复制/剪切粘贴
- 复制 B1（`=A1*2`，引用左一格）粘贴到 D5
- 【预期】D5 raw `=C5*2`（相对偏移：左一格同行）
- 剪切 B1 粘贴到 E5 → E5 raw `=A1*2`（移动语义不偏移），B1 已清空
- 从外部（文本编辑器）复制 `=A1*2` 文本粘贴到 F1 → raw 原样 `=A1*2`

## 4. 跨表引用
- 点 `+` 新增 Sheet2，A1 输入 `10`；切回 Sheet1，A5 输入 `=Sheet2!A1*2`
- 【预期】A5 显示 20
- Sheet1 B5 输入 `=SUM(Sheet2!A1:A1)` → 显示 10

## 5. 表改名与已知限制
- 双击 Sheet2 标签改名为 `Data` → Sheet1 A5 显示 `#REF!`（已知限制，不改写公式）
- Ctrl+Z 撤销改名 → A5 恢复 20

## 6. 删表与恢复
- 在 Sheet2 表输入些内容，删除该表（× 按钮，确认）
- 【预期】Sheet1 A5 显示 `#REF!`；undo → 表及内容恢复，A5 恢复 20
- 只剩一张表时 × 按钮不出现

## 7. 跨表循环
- Sheet1 A9 输入 `=Sheet2!B9+1`，Sheet2 B9 输入 `=Sheet1!A9+1`
- 【预期】两格均显示 `#CYCLE!`

## 8. 标签栏综合
- 增 3 张表、改名、切换、删除，穿插 undo/redo（Ctrl+Z / Ctrl+Shift+Z）
- 【预期】各操作可撤销重做；切换表不产生 undo 记录；活动表高亮正确
