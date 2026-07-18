# M4c 浏览器验收：数据验证（E4）

可执行形式为 `apps/demo/e2e/suites/m4c.mjs`（`node apps/demo/e2e/run.mjs m4c`，浏览器 daemon 驱动）；本文档为同步骤的叙述版。前置与防干扰约定同 m4b.md「前置（必读）」：含异步步骤前
bringToFront；notice 断言在触发后 5s 内；React 菜单/对话框点击后等 300ms。

页面侧工具（`installPageTools` 注入，每次 freshPage 后重注）：

- `__setRules(rules)`：`__xcell.dispatch(state.tr.setValidations(rules))` 直接注入规则。
- `__notice()`：读 `.status-notice` 文本（无 → null）。
- `__typeCell(r,c,text)`：W.click + 首字符 keydown 开 editbox → 覆盖全文 → Enter；
  返回 `'open'`（被拒，编辑器保持打开）/ `'closed'`（提交成功）。
- `__typeContinue(text)`：被拒后改当前编辑器文本再 Enter。
- `__setFormula(text)`：公式栏 React 受控 input 经 native setter + input 事件填值，Enter 提交。

测试区域统一 rows 19-23（A20 起），避开 demo 数据区（rows 0-12）。

## 1. numRange 阻止+放行（editbox 与 FormulaBar 双路径）

- `__setRules([{A20:A24 numRange between 1..9}])`
- `__typeCell(19,0,'10')` → 【预期】'open'；A20 raw 仍 null；notice 含「介于 1 与 9」
- `__typeContinue('5')` → 【预期】'closed'；A20 raw='5'；选区下移 A21
- FormulaBar：`__setFormula('10')` → 【预期】A21 raw 仍 null + notice 含「介于 1 与 9」；
  `__setFormula('5')` → 【预期】A21 raw='5'

## 2. textLen

- 规则 `B20:B24 textLen lte 3` → `__typeCell(19,1,'abcd')` → 【预期】'open' + raw null +
  notice「文本长度须小于等于 3」→ `__typeContinue('abc')` → 【预期】'closed' + raw='abc'

## 3. list

- 规则 `C20:C24 list [Apple,Banana]` → 'cherry' 拒（notice「输入值须在序列内：Apple, Banana」）
  → 'apple'（小写，大小写不敏感）放行 raw='apple'

## 4. 对话框配置 + 删除 + undo

先 freshPage（清档刷新，保证 undo 栈干净），W.click(19,3) 选中 D20。

- 点工具栏「验证」（title=数据验证）→ 等 300ms → 「+ 添加规则」→ 等 150ms
- 【预期】默认 numRange between：两个 `.cf-value` 输入框 placeholder=['数值','上界']
- native setter 填 v1='1' / v2='9' → 「确定」→ 等 150ms
- 【预期】model validations=[{D20:D20 numRange between v1:'1' v2:'9'}]
- 再开「验证」→ 点行尾「删除规则」(×) → 「确定」→ 【预期】model 空
- Ctrl+Z 一次 → 【预期】规则恢复（length 1）；再一次 → 【预期】回到空

## 5. xlsx 往返

- `__setRules` 两条：`A20:A24 numRange between 1..9` + `C20:C24 list [Apple,Banana]`
- stub `URL.createObjectURL` 捕获 Blob → 文件 → 导出 xlsx → 等 1200ms → 字节转 base64 回 runner
- **runner 侧** `exceljs` load 字节，校验 `ws.dataValidations.model`（runner 相对纯浏览器
  脚本的优势：直接用项目 node_modules 的 exceljs）：
  - 【预期】A20..A24 逐地址 `{type:'decimal', operator:'between', formulae:[1,9]}`（数字回读为 number）
  - 【预期】C20..C24 逐地址 `{type:'list', formulae:['"Apple,Banana"']}`
- 同一 base64 经「文件 → 打开 xlsx」+ feedFile 喂回导入 → 等 800ms
- 【预期】model validations 归一化后两条规则保真（range/type/op/v1/v2/items 全等，id 重编不计）

## 6. 公式/清空跳过

- 规则 `E20:E24 numRange between 1..9`
- `__typeCell(19,4,'=1+1')` → 【预期】'closed'（公式原文跳过校验），raw='=1+1'
- W.click(19,4) + Delete → 【预期】raw=null（清空不触发校验）且 notice 不含校验拒绝
  文案（请输入/文本长度/输入值须）。注：不采用「等 5s TTL 后断言 notice=null」——隐藏
  tab 定时器密集节流使等待不可靠（实测 visibilityState 恒 hidden）。

---

【验收记录】 `node apps/demo/e2e/run.mjs m4c` 六场景全 PASS（见
）。
