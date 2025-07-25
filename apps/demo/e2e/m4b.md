# M4b 浏览器验收（浏览器）：xlsx 导入导出

前置：`pnpm --filter ./apps/demo dev -- --port 5180`（仓库根，http://localhost:5180）；dev 下 `window.__xcell` 为 EditorView。
先注入 m3c.md 的 `window.W` helper。每步后附【预期】，失败即终止并回报。
localStorage 键：`xexcel.workbook`。行列索引 0-based。

## 前置（必读）：bringToFront 与防干扰

- **必须 bringToFront**：Chrome 对隐藏标签页做 JS 定时器密集节流，exceljs 的异步解析在其中会
  挂死（真实用户前台操作不受影响）。任何含 exceljs 异步（导入/导出）的场景执行前，用 CDP
  `Page.bringToFront` 把验收 tab 置前台；步骤间若长时间停顿致 tab 失焦，重发一次。
- **防用户误触**：bringToFront 会让 tab 出现在用户屏幕上，注意防用户误触干扰（实测中 A1 曾被
  用户 IME 误输入）。断言前若发现 raw 被意外改动，先考虑外部干扰再判 bug。

## 0a. 清档刷新（仅首次执行一次）

```js
// 清空旧存档，刷新从默认 demo 表起
localStorage.removeItem('xexcel.workbook'); location.reload()
```
【预期】页面正常打开，A1='产品'（demo 表）

## 0b. 注入 helper/stub（每次刷新页面后都要重注）

先重注 m3c.md 的 `window.W` helper，再执行：

```js
// 1) 捕获 pickFile 动态创建的 input：存入 window.__lastInput，并把 click 置空，
//    防止真实文件对话框弹出（喂文件走 base64 + DataTransfer，不经系统对话框）
// 2) 记录原始引用，结尾恢复
window.__origCreateElement = document.createElement.bind(document)
window.__origConfirm = window.confirm
window.__lastInput = null
document.createElement = (tag, ...rest) => {
  const el = window.__origCreateElement(tag, ...rest)
  if (tag === 'input') { window.__lastInput = el; el.click = () => {} }
  return el
}
window.confirm = () => true // 默认放行；场景 5 自行替换
window.__archiveBefore = localStorage.getItem('xexcel.workbook')
W.read(0, 0)?.raw
```
【预期】'产品'（W helper 与 stub 注入成功，demo 表可读）

注意：注入后勿打开含 input 的 UI（如查找栏），避免 `__lastInput` 被覆盖。

菜单操作约定（React 异步：点「文件」后等 300ms 再点菜单项）：

```js
// 打开菜单：  document.querySelector('button[aria-label="文件"]').click()
// 等 300ms 后：[...document.querySelectorAll('button[role="menuitem"]')].find(b=>b.textContent.includes('xxx')).click()
```

---

## 1. 导出 xlsx

先在页面构造含样式/合并/筛选/CF/公式的区块（放 rows 20-23，避开 demo 数据区）：

```js
const st = __xcell.state
const sid = st.doc.active
let tr = st.tr.setCells(sid, [
  { row: 20, col: 0, cell: { raw: '汇总区', style: { bold: true, bg: '#e8f0fe' } } },
  { row: 21, col: 0, cell: { raw: '品名', style: { bold: true } } },
  { row: 21, col: 1, cell: { raw: '数量', style: { bold: true } } },
  { row: 22, col: 0, cell: { raw: '苹果', style: { color: '#ff0000', border: { bottom: { style: 'thin' } } } } },
  { row: 22, col: 1, cell: { raw: '3', style: { numFmt: '0.00' } } },
  { row: 23, col: 1, cell: { raw: '=B23*2' } },
])
tr = tr.setMerges([{ sr: 20, sc: 0, er: 20, ec: 1 }])                 // A21:B21
tr = tr.setFilter({ range: { sr: 21, sc: 0, er: 23, ec: 1 }, criteria: {} }) // A22:B24
tr = tr.setCondFormats([
  { id: 'cf1', range: { sr: 22, sc: 1, er: 23, ec: 1 }, type: 'value', op: 'gt', v1: '2',
    style: { bold: true, bg: '#ffff00' } },
])
__xcell.dispatch(tr)
__xcell.state.activeSheet.getCell(23, 1)?.raw
```
【预期】'=B23*2'

stub 下载捕获 Blob，点「导出 xlsx」：

```js
window.__dl = null
const oc = URL.createObjectURL.bind(URL)
URL.createObjectURL = (b) => { window.__dl = b; return oc(b) }
;document.querySelector('button[aria-label="文件"]').click()
await new Promise(r=>setTimeout(r,300))
;[...document.querySelectorAll('button[role="menuitem"]')].find(b=>b.textContent==='导出 xlsx').click()
await new Promise(r=>setTimeout(r,800)) // exportXlsx 内部 await writeBuffer，异步
const bytes = new Uint8Array(await window.__dl.arrayBuffer())
URL.createObjectURL = oc
;[window.__dl.type, bytes.length > 2000, bytes[0] === 0x50 && bytes[1] === 0x4B]
```
【预期】['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', true, true]
（zip 魔数 PK；demo 表较大字节数远超 2KB）

取回字节落盘：evaluate 返回 base64（分块拼，避免 fromCharCode 参数上限）：

```js
let bin = ''
for (let i = 0; i < bytes.length; i += 8192) bin += String.fromCharCode(...bytes.subarray(i, i + 8192))
btoa(bin) // evaluate 返回值：base64 串
```

主 Agent 将 base64 串存 `/tmp/m4b-export.b64`，然后在**项目根目录**执行：

```bash
node -e "require('fs').writeFileSync('/tmp/m4b-export.xlsx', Buffer.from(require('fs').readFileSync('/tmp/m4b-export.b64','utf8').trim(),'base64'))"
```

node + exceljs 回读校验（脚本见附录 B，根目录执行 `node /tmp/m4b-verify-export.cjs`）。
【预期】输出 `VERIFY OK`（校验点：sheet 名 Sheet1；A21 值/粗体/底色 FFE8F0FE；merge A21:B21；
A23 红字 FFFF0000 + bottom thin 边框；B23 数值 3 + numFmt '0.00'；B24 公式 B23*2；
autoFilter 'A22:B24'；CF 块 B23:B24 cellIs/greaterThan/formulae ['2'] + dxf 底色 FFFFFF00）

---

## 2. 导入 xlsx

node 侧先生成 `/tmp/m4b-import.xlsx`（两 sheet：「数据」含粗体红字 A1='标题'、B1=3、B2 公式
=B1*2、合并 D1:E2、冻结首行；「空表」无内容）。脚本见附录 A，根目录执行 `node /tmp/m4b-gen-import.cjs`。
【预期】输出 `written /tmp/m4b-import.xlsx`

记录导入前现场：

```js
window.__archiveBefore = localStorage.getItem('xexcel.workbook')
window.__namesBefore = [...__xcell.state.doc.order].map(id => __xcell.state.doc.names.get(id))
window.__lastInput = null
window.__namesBefore
```
【预期】['Sheet1']（demo 单表）

喂文件（base64 + DataTransfer 主路径；CDP setFileInputFiles 对未挂 DOM 的 input 实测返回
'Not allowed'，降级为附录 C 参考）：

1. 主 Agent 读 `/tmp/m4b-import.xlsx` → base64 串。
2. evaluate 点菜单（confirm 已是默认放行的 `() => true`）：

```js
;document.querySelector('button[aria-label="文件"]').click()
await new Promise(r=>setTimeout(r,300))
;[...document.querySelectorAll('button[role="menuitem"]')].find(b=>b.textContent.includes('打开 xlsx')).click()
await new Promise(r=>setTimeout(r,300))
!!window.__lastInput
```
【预期】true（pickFile 的动态 input 已被捕获）

3. evaluate 构造 File 并手动触发 onchange（`<b64>` 替换为第 1 步的 base64 串）：

```js
const bin = atob('<b64>')
const u8 = new Uint8Array(bin.length)
for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
const f = new File([u8], 'm4b-import.xlsx',
  { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
const dt = new DataTransfer(); dt.items.add(f)
const inp = window.__lastInput
inp.files = dt.files
inp.onchange?.()
'fed'
```
【预期】'fed'

4. evaluate 等待导入完成并断言（**notice 断言须在喂文件后 5s 内执行**：多步 evaluate 往返后
   notice 可能已过期，过期则该子项按跳过处理，勿误判失败）：

```js
await new Promise(r=>setTimeout(r,800))
const wb = __xcell.state.doc
;[[...wb.order].map(id => wb.names.get(id)), wb.order.length, wb.names.get(wb.active)]
```
【预期】[['数据','空表'], 2, '数据']（旧 sheet Sheet1 已消失；active=第一张）

```js
const sh = __xcell.state.activeSheet
const a1 = sh.getCell(0, 0)
;[a1?.raw, a1?.style?.bold, a1?.style?.color, sh.getCell(0, 1)?.raw, sh.getCell(1, 1)?.raw]
```
【预期】['标题', true, '#ff0000', '3', '=B1*2']（粗体红字导入进 model；公式复活为 raw）

```js
// 包源码在 dev server root（apps/demo）之外，须走 vite /@fs 绝对路径（同 suites/m4b.mjs 的 ENGINE_URL）
const { evaluatorFor } = await import('/@fs/Users/ximing/project/mygithub/xexcel/packages/excel-core/src/formula/engine.ts')
evaluatorFor(__xcell.state.doc).get(__xcell.state.doc.active, 1, 1)
```
【预期】6（B1=3 → =B1*2 计算值）

```js
;[sh.merges.some(m=>m.sr===0&&m.sc===3&&m.er===1&&m.ec===4), sh.frozenRows,
  __xcell.state.selection.activeCell, document.querySelector('.status-notice')?.textContent]
```
【预期】[true, 1, {row:0,col:0}, '已打开 m4b-import.xlsx']
（合并 D1:E2；冻结首行；选区在 A1；StatusBar 提示，5s TTL 内有效）

---

## 3. 替换语义与自动保存（接场景 2，中间不刷新）

Task 6 R1 后语义：导入成功**立即落档**（saveNow），存档此刻已是新 workbook。

```js
const arch = localStorage.getItem('xexcel.workbook')
;[arch !== window.__archiveBefore, arch.includes('数据'), arch.includes('空表')]
```
【预期】[true, true, true]（导入成功瞬间存档已是新 workbook，非「编辑后才覆盖」）

随便编辑一格 → 防抖 1s 后存档含编辑内容：

```js
W.setCell(3, 0, '验收标记')
await new Promise(r=>setTimeout(r,1500))
localStorage.getItem('xexcel.workbook').includes('验收标记')
```
【预期】true

刷新 → 新 workbook 恢复：

```js
location.reload(); await new Promise(r=>setTimeout(r,800))
const wb = __xcell.state.doc
;[[...wb.order].map(id => wb.names.get(id)), wb.activeSheet.getCell(0,0)?.raw, wb.activeSheet.getCell(3,0)?.raw]
```
【预期】[['数据','空表'], '标题', '验收标记']

**刷新后必须重注 0b（W helper + createElement/confirm stub）**，再继续场景 4。

---

## 4. 损坏文件（接场景 3，已重注 stub）

喂文件走场景 2 的 base64 主路径。损坏内容仅 7 字节（等价 `printf 'garbage'`），base64 直接内联。
先点菜单（同场景 2 第 2 步，确认 `window.__lastInput` 已捕获），然后：

```js
const bin = atob('Z2FyYmFnZQ==') // 'garbage'
const u8 = new Uint8Array(bin.length)
for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
const f = new File([u8], 'm4b-bad.xlsx',
  { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
const dt = new DataTransfer(); dt.items.add(f)
const inp = window.__lastInput
inp.files = dt.files
inp.onchange?.()
'fed'
```
【预期】'fed'

然后（notice 断言须在喂文件后 5s 内执行）：

```js
await new Promise(r=>setTimeout(r,800))
;[document.querySelector('.status-notice')?.textContent,
  __xcell.state.doc.names.get(__xcell.state.doc.active),
  __xcell.state.activeSheet.getCell(0,0)?.raw,
  __xcell.state.activeSheet.getCell(3,0)?.raw]
```
【预期】['文件无法解析', '数据', '标题', '验收标记']（notice 报无法解析；现场分毫未动）

自动保存已恢复（导入失败后 resume 生效）：

```js
W.setCell(4, 0, '恢复标记')
await new Promise(r=>setTimeout(r,1500))
;[localStorage.getItem('xexcel.workbook').includes('恢复标记'),
  document.querySelector('.status-error')?.textContent ?? null]
```
【预期】[true, null]（存档正常更新；无自动保存失败提示）

---

## 5. 护栏（接场景 4）

先等上场景 notice 过期（5s TTL），避免干扰断言：

```js
await new Promise(r=>setTimeout(r,5000))
document.querySelector('.status-notice')?.textContent ?? null
```
【预期】null

第一道 confirm 否决 → 不出现文件选择、内容不变：

```js
window.__confirmCalls = []
window.confirm = (msg) => { window.__confirmCalls.push(msg); return false }
window.__lastInput = null
const docBefore = __xcell.state.doc
;document.querySelector('button[aria-label="文件"]').click()
await new Promise(r=>setTimeout(r,300))
;[...document.querySelectorAll('button[role="menuitem"]')].find(b=>b.textContent.includes('打开 xlsx')).click()
await new Promise(r=>setTimeout(r,300))
;[window.__confirmCalls.length, window.__confirmCalls[0].includes('替换'), window.__lastInput, __xcell.state.doc === docBefore]
```
【预期】[1, true, null, true]（confirm 文案含「替换」；input 未创建；doc 同一引用未被替换）

confirm 放行 + 取消文件选择 → 无动作无报错：

```js
window.confirm = (msg) => { window.__confirmCalls.push(msg); return true }
;document.querySelector('button[aria-label="文件"]').click()
await new Promise(r=>setTimeout(r,300))
;[...document.querySelectorAll('button[role="menuitem"]')].find(b=>b.textContent.includes('打开 xlsx')).click()
await new Promise(r=>setTimeout(r,300))
// 模拟用户在系统对话框点「取消」：pickFile 监听 cancel 事件（Chrome 113+）
window.__lastInput.dispatchEvent(new Event('cancel'))
await new Promise(r=>setTimeout(r,500))
;[window.__confirmCalls.length, __xcell.state.doc === docBefore,
  document.querySelector('.status-notice')?.textContent ?? null,
  document.querySelector('.status-error')?.textContent ?? null]
```
【预期】[2, true, null, null]（confirm 共 2 次；doc 未动；无 notice 无 error）

收尾恢复 stub：

```js
document.createElement = window.__origCreateElement
window.confirm = window.__origConfirm
'ok'
```
【预期】'ok'

---

## 附录 A：生成 /tmp/m4b-import.xlsx（项目根目录执行）

```bash
cat > /tmp/m4b-gen-import.cjs <<'EOF'
const ExcelJS = require('exceljs')
;(async () => {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('数据')
  const a1 = ws.getCell('A1')
  a1.value = '标题'
  a1.font = { bold: true, color: { argb: 'FFFF0000' } }
  ws.getCell('B1').value = 3
  ws.getCell('B2').value = { formula: 'B1*2', result: 6 }
  ws.getCell('D1').value = '合并区'
  ws.mergeCells('D1:E2')
  ws.views = [{ state: 'frozen', ySplit: 1 }]
  wb.addWorksheet('空表')
  await wb.xlsx.writeFile('/tmp/m4b-import.xlsx')
  console.log('written /tmp/m4b-import.xlsx')
})().catch((e) => { console.error(e); process.exit(1) })
EOF
node /tmp/m4b-gen-import.cjs
```

## 附录 B：校验 /tmp/m4b-export.xlsx（项目根目录执行）

```bash
cat > /tmp/m4b-verify-export.cjs <<'EOF'
const ExcelJS = require('exceljs')
const assert = require('assert')
;(async () => {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile('/tmp/m4b-export.xlsx')
  const ws = wb.worksheets[0]
  assert.strictEqual(ws.name, 'Sheet1', 'sheet 名')
  const a21 = ws.getCell('A21')
  assert.strictEqual(a21.value, '汇总区', 'A21 值')
  assert.strictEqual(a21.font.bold, true, 'A21 粗体')
  assert.strictEqual(a21.fill.fgColor.argb, 'FFE8F0FE', 'A21 底色')
  assert.ok(ws.model.merges.includes('A21:B21'), 'merge A21:B21')
  const a23 = ws.getCell('A23')
  assert.strictEqual(a23.value, '苹果', 'A23 值')
  assert.strictEqual(a23.font.color.argb, 'FFFF0000', 'A23 红字')
  assert.strictEqual(a23.border.bottom.style, 'thin', 'A23 下边框')
  const b23 = ws.getCell('B23')
  assert.strictEqual(b23.value, 3, 'B23 数值')
  assert.strictEqual(b23.numFmt, '0.00', 'B23 numFmt')
  assert.strictEqual(ws.getCell('B24').value.formula, 'B23*2', 'B24 公式')
  assert.strictEqual(ws.autoFilter, 'A22:B24', 'autoFilter ref')
  const cf = ws.model.conditionalFormattings
  const block = cf.find((b) => b.ref === 'B23:B24')
  assert.ok(block, 'CF 块 B23:B24')
  assert.strictEqual(block.rules[0].type, 'cellIs', 'CF 类型')
  assert.strictEqual(block.rules[0].operator, 'greaterThan', 'CF 操作符')
  assert.deepStrictEqual(block.rules[0].formulae, ['2'], 'CF formulae')
  assert.strictEqual(block.rules[0].style.fill.bgColor.argb, 'FFFFFF00', 'CF dxf 底色')
  console.log('VERIFY OK')
})().catch((e) => { console.error('VERIFY FAIL:', e.message); process.exit(1) })
EOF
node /tmp/m4b-verify-export.cjs
```

## 附录 C：CDP 喂文件参考路径（实测失败留档）

原设计主路径，实测 `DOM.setFileInputFiles` 对 pickFile 动态创建且未挂 DOM 的 input 返回
`Not allowed`，已降级为参考；场景 2/4 以 base64 + DataTransfer 为主路径。留档步骤：

1. CDP `Page.enable`
2. CDP `Page.setInterceptFileChooserDialog` `{enabled: true}`
3. evaluate 点菜单（文件 → 打开 xlsx…），`window.__lastInput` 捕获 input
4. evaluate `window.__lastInput`（returnByValue: false）拿 objectId
5. CDP `DOM.setFileInputFiles` `{files: ['/tmp/m4b-import.xlsx'], objectId}` —— 此步实测失败
6. CDP `Page.setInterceptFileChooserDialog` `{enabled: false}`

若未来 pickFile 把 input 挂进 DOM（或 Chrome 放开该限制），此路径可复用。
