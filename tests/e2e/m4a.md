# M4a 浏览器验收（浏览器）

前置：`npm run dev`；dev 下 `window.__xcell` 为 EditorView；先注入 m3c.md 的 `window.W` helper。
每步后附【预期】，失败即终止并回报。localStorage 键：`xexcel.workbook`。

## 1. 自动保存 + 刷新恢复

```js
// 清空旧存档，刷新从默认表演示起
localStorage.removeItem('xexcel.workbook'); location.reload()
```
【预期】页面正常打开，显示 demo 表（A1=产品）

```js
// B20 输入新值（走 W.key 流或直接事务皆可，这里用 UI 流）
W.click(19,1); W.key('9'); W.key('9'); W.key('Enter')
__xcell.state.activeSheet.getCell(19,1)?.raw
```
【预期】'99'

```js
await new Promise(r=>setTimeout(r,1500)); localStorage.getItem('xexcel.workbook')?.length>0
```
【预期】true（防抖 1s 内已写入）

```js
location.reload(); await new Promise(r=>setTimeout(r,800))
__xcell.state.activeSheet.getCell(19,1)?.raw
```
【预期】'99'（刷新后恢复）；且 A1 仍是 '产品'（demo 数据也在存档里）

## 2. 清除存档

```js
// Toolbar「文件」→「清除浏览器存档」→ confirm 自动确认
// React 状态更新异步：点按钮后需等菜单渲染；点菜单项后需等 StatusBar 渲染
await (async () => {
  window.confirm = () => true
  ;[...document.querySelectorAll('.tool-btn')].find(b=>b.textContent==='文件').click()
  await new Promise(r=>setTimeout(r,300))
  ;[...document.querySelectorAll('.file-menu-item')].find(b=>b.textContent.includes('清除')).click()
  await new Promise(r=>setTimeout(r,500))
})()
;[localStorage.getItem('xexcel.workbook'), document.querySelector('.status-notice')?.textContent]
```
【预期】[null, '已清除浏览器存档']（存档已删；StatusBar 提示已渲染）

```js
location.reload(); await new Promise(r=>setTimeout(r,800))
__xcell.state.activeSheet.getCell(19,1)?.raw ?? null
```
【预期】null（回到默认 demo 表，B20 无内容）

## 3. CSV 导出

```js
// 构造含公式/逗号/引号/中文/换行的表
const st = __xcell.state
const tr = st.tr.setCells(st.doc.active, [
  { row: 0, col: 5, cell: { raw: '含,逗号' } },
  { row: 1, col: 5, cell: { raw: '含"引号' } },
  { row: 2, col: 5, cell: { raw: '=B2*2' } },
])
__xcell.dispatch(tr)
// 拦截下载
window.__dl = null
const oc = URL.createObjectURL.bind(URL)
URL.createObjectURL = (b) => { window.__dl = b; return oc(b) }
;[...document.querySelectorAll('.tool-btn')].find(b=>b.textContent==='文件').click()
await new Promise(r=>setTimeout(r,300))
;[...document.querySelectorAll('.file-menu-item')].find(b=>b.textContent==='导出 CSV').click()
await new Promise(r=>setTimeout(r,300))
// blob.text() 解码时会剥离 UTF-8 BOM，BOM 须用 arrayBuffer 验字节
const txt = await window.__dl.text()
const buf = new Uint8Array(await window.__dl.arrayBuffer())
URL.createObjectURL = oc
buf[0]===0xEF && buf[1]===0xBB && buf[2]===0xBF && txt.includes('"含,逗号"') && txt.includes('"含""引号') && txt.includes('=B2*2')
```
【预期】true（BOM 字节 EF BB BF 经 arrayBuffer 校验 + 转义 + 公式原文）

## 4. CSV 导入

```js
// stub createElement：把预制 File 喂给 pickFile 的 input
const csvText = '姓名,数量\r\n苹果,3\r\n=B2*2,"含,逗号"\r\n'
const f = new File([csvText], '进货.csv', { type: 'text/csv' })
const oce = document.createElement.bind(document)
document.createElement = (tag, ...rest) => {
  const el = oce(tag, ...rest)
  if (tag === 'input') setTimeout(() => {
    const dt = new DataTransfer(); dt.items.add(f)
    el.files = dt.files
    el.onchange?.()
  }, 0)
  return el
}
;[...document.querySelectorAll('.tool-btn')].find(b=>b.textContent==='文件').click()
await new Promise(r=>setTimeout(r,300))
;[...document.querySelectorAll('.file-menu-item')].find(b=>b.textContent.includes('打开 CSV')).click()
await new Promise(r=>setTimeout(r,500))
document.createElement = oce
const wb = __xcell.state.doc
({ order: [...wb.order], active: wb.active, name: wb.names.get(wb.active) })
```
【预期】active 为新 sheet，name='进货'（若已存在则 '进货 (2)'）

```js
const sh = __xcell.state.activeSheet
;[sh.getCell(0,0)?.raw, sh.getCell(1,1)?.raw, sh.getCell(2,0)?.raw, sh.getCell(2,1)?.raw]
```
【预期】['姓名','3','=B2*2','含,逗号']

```js
// 公式复活：evaluator 有计算值
const { evaluatorFor } = await import('/src/formula/engine.ts')
evaluatorFor(__xcell.state.doc).get(__xcell.state.doc.active, 2, 0)
```
【预期】6（B2=3 → =B2*2）

```js
// undo 一步（点工具栏撤销按钮）：新 sheet 消失
document.querySelector('.tool-btn[title="撤销"]').click()
__xcell.state.doc.order.length
```
【预期】回到导入前的 sheet 数（undo 一步复原）

## 5. 损坏存档恢复

```js
localStorage.setItem('xexcel.workbook', '{broken')
location.reload(); await new Promise(r=>setTimeout(r,800))
;[!!window.__xcell, localStorage.getItem('xexcel.workbook.corrupt'), localStorage.getItem('xexcel.workbook')]
```
【预期】[true, '{broken', null]（正常启动；损坏档备份 .corrupt；原键已删）

```js
__xcell.state.activeSheet.getCell(0,0)?.raw
```
【预期】'产品'（落默认 demo 表）

## 6. 自动保存失败提示（配额模拟）

```js
const ose = Storage.prototype.setItem
Storage.prototype.setItem = () => { throw new Error('QuotaExceeded') }
W.click(0,9); W.key('x'); W.key('Enter')
await new Promise(r=>setTimeout(r,1500))
Storage.prototype.setItem = ose
document.querySelector('.status-error')?.textContent
```
【预期】'自动保存失败'
