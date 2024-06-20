# M3c 浏览器验收（浏览器）

前置：`npm run dev` 启动；dev 下 `window.__xcell` 为 EditorView。0-based 行列。
读 raw：`__xcell.state.activeSheet.getCell(r,c)?.raw`；style：`...?.style`。每步后附【预期】。
失败即终止并回报。注入 helper（`window.W`）封装像素事件 + 读状态：

```js
window.W = (() => {
  const V = () => window.__xcell
  const root = () => document.querySelector('.xcell-root')
  const stage = () => V().stage.content            // app 的 mousedown 目标
  const rc = () => root().getBoundingClientRect()
  const disp = (type, el, x, y, mods) => { const m = mods || {}
    el.dispatchEvent(new MouseEvent(type, { bubbles:true, cancelable:true, view:window,
      clientX:x, clientY:y, ctrlKey:!!m.ctrl, shiftKey:!!m.shift, metaKey:!!m.meta, button:0,
      buttons: type==='mousedown'?1:0 })) }
  // 单元格中心 viewport 坐标（冻结/zoom 感知，走 cellViewportRect）
  const cellXY = (r,c) => { const cv=V().cellViewportRect(r,c); const R=rc()
    return { clientX: R.left+cv.x+cv.w/2, clientY: R.top+cv.y+cv.h/2 } }
  // 选区右边框点（活动区域四边 ±4px 带，避开右下角填充柄）
  const borderXY = (sr,sc,er,ec) => { const cv=V().cellViewportRect(sr,ec); const R=rc()
    return { clientX: R.left+cv.x+cv.w, clientY: R.top+cv.y+cv.h/2 } }
  const click = (r,c,mods) => { const p=cellXY(r,c); const s=stage()
    disp('mousedown',s,p.clientX,p.clientY,mods); disp('mouseup',window,p.clientX,p.clientY,mods) }
  const dragSelect = (r0,c0,r1,c1,mods) => { const s=stage(); const p0=cellXY(r0,c0); const p1=cellXY(r1,c1)
    disp('mousedown',s,p0.clientX,p0.clientY,mods); disp('mousemove',window,p1.clientX,p1.clientY,mods)
    disp('mouseup',window,p1.clientX,p1.clientY,mods) }
  const dragMove = (sr,sc,er,ec,dR,dC) => { const s=stage(); const b=borderXY(sr,sc,er,ec)
    disp('mousedown',s,b.clientX,b.clientY)            // selborder -> dragging=true，src=活动区域
    const start=V().pointerToCell(b.clientX,b.clientY); const tgt=cellXY(start.row+dR,start.col+dC)
    disp('mousemove',window,tgt.clientX,tgt.clientY); disp('mouseup',window,tgt.clientX,tgt.clientY) }
  const key = (k,mods) => { let el=document.activeElement
    if(!el||el.tagName!=='TEXTAREA') el=document.querySelector('textarea.xcell-input-proxy')
    el.dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,cancelable:true,key:k,code:k,
      ctrlKey:!!(mods&&mods.ctrl),shiftKey:!!(mods&&mods.shift)})) }
  // 剪贴板：mock clipboardData 让内部指纹 tsv 在 copy->paste 间往返
  W_copy = function(){ const V0=V(); V0.focus(); const store={}
    const ev={clipboardData:{setData:(t,d)=>store[t]=d,getData:(t)=>store[t]||''},preventDefault:()=>{}}
    V0.onCopy(ev); window.W.__clipText=store['text/plain']||''; window.W.__clipHtml=store['text/html']||''
    return window.W.__clipText.length }
  W.paste = function(text){ const V0=V(); V0.focus(); const t=(text!=null)?text:(window.W.__clipText||'')
    const ev={clipboardData:{getData:(type)=>type==='text/plain'?t:'',setData:()=>{}},preventDefault:()=>{}}
    V0.onPaste(ev); return t.length }
  const read=(r,c)=>{const cell=V().state.activeSheet.getCell(r,c);return cell?{raw:cell.raw,style:cell.style||null}:null}
  const sel=()=>{const s=V().state.selection;return{ranges:s.ranges,activeCell:s.activeCell}}
  const setCell=(r,c,raw)=>V().dispatch(V().state.tr.setCell(r,c,raw))
  const setStyle=(r,c,st)=>V().dispatch(V().state.tr.setCell(r,c,V().state.activeSheet.getCell(r,c)?.raw??'',st))
  return {click,dragSelect,dragMove,key,read,sel,setCell,setStyle,cellXY,borderXY,V,
    copy:W_copy,paste,render:()=>V().render()}
})()
```

工具栏按钮用 `evaluate` 点 `button[title="加粗"]` 等。React 重渲染有时序：读 DOM `disabled`
属性或操作后须 `await setTimeout(80~150)` 再读。行 51+（视区下方）须先 `V.scrollY=...;V.render()`
滚到可见再操作（合成事件命中需在 canvas 内）。

## F2 多区域选择

### 1. Ctrl+click 多区域 + activeCell=末项
- click(6,0) → click(7,2,{ctrl}) → click(8,4,{ctrl})
- 【预期】sel.ranges = [{A7},{C8},{E8}]，activeCell={8,4}=E8（末加入者）

### 2. Ctrl+click 反选移除
- click(7,2,{ctrl})（C8 在选区内）
- 【预期】ranges=[{A7},{E8}]，activeCell 仍 E8

### 3. Ctrl+click 重叠 LIFO 反选
- click(6,0) → dragSelect(5,0,7,1,{ctrl})（追加 A6:B8，与 A7 重叠）→ sel 应为 [{A7},{A6:B8}]，active=A6:B8
- click(6,0,{ctrl})（A7 同时落在两个 range）→ 【预期】LIFO 移除最后加入者（活动区域 A6:B8），ranges=[{A7}]，activeCell∈ranges[last]
- ⚠ 已知：实现 `toggleRange` 在 clicked∈新末项 range 时保留旧 activeCell，可能落在被移除的 range 内（违反 spec §1.1 不变式）。复现：上一步后 type 一字符 → 落点在 A6:B8 内而非 A7（见 acceptance.md 额外发现 #1）

### 4. Shift+Arrow 生长（不丢起始格）
- click(6,0) → key('ArrowRight',{shift:true}) → key('ArrowRight',{shift:true})
- 【预期】A7:B7 → A7:C7；sr/sc 恒 6/0（起始格不丢），activeCell=焦点端

### 5. 格式作用于全部区域
- click(6,0);click(6,2,{ctrl});click(6,4,{ctrl})（[A7,C7,E7]）→ 工具栏「加粗」(`button[title="加粗"]`)
- 【预期】A7/C7/E7 三格 style.bold=true

### 6. 填充仅活动区域
- setCell(6,4,'1') → [A7,C7,E7]（active=E7）→ dragMove 源=E7，从右下角填充柄往下 2（dragMove 改用 fillhandle：mousedown 活动区域右下角，mousemove 到 E9，mouseup）
- 【预期】E8=1,E9=1（活动区域 E7 填充）；A7/C7 不变

### 7. 排序/筛选多区域禁用
- dragSelect(6,0,8,0);dragSelect(6,2,8,2,{ctrl})（多区域 A7:A9 + C7:C9）→ await 100ms
- 读 `button[title="按选区首列升序"].disabled`、`button[title="自动筛选（对选区启用/清除全表筛选）"].disabled`
- 【预期】两者 disabled=true（单区域基线时 disabled=false）

### 8. 表头并集高亮
- [A7:B7, D9:E9]（rows 6/8，cols 0-1/3-4）→ screenshot
- 【预期】行头 7/9、列头 A/B/D/E 高亮（并集），列头 C 不高亮。layers `colActive/rowActive` 用 `ranges.some` 并集（layers.ts:177-179）

## F3 拖动移动

### 1. 边框拖动 cut 移动（raw+style 搬、公式不 shift）
- setCell(11,1,'5');setCell(11,2,'10');setCell(11,3,'=B11+C11');setStyle(11,1,{bold:true});setStyle(11,3,{bg:'#ffff00'})
- click(11,3) → dragMove(11,3,11,3,2,0)（活动区域 D11，右边框拖到 D13）
- 【预期】D13.raw='=B11+C11'（公式不 shift，cut 语义），D13.style.bg='#ffff00'（style 跟搬），D11=null（清源）

### 2. 相交/merge 拒绝 alert
- 先 stub `window.alert=(m)=>{window.__lastAlert=m}`（防阻塞）
- dragSelect(11,1,12,2) → dragMove(11,1,12,2,1,0)（目标 B12:C13 与源 B11:C12 相交）
- 【预期】__lastAlert='目标区域与源相交、落在合并区或超出表格边界，无法移动'，B11 不变
- merge 路径：setMerges 追加 E12:F12，setCell(14,3,'moveme')，click(14,3)→dragMove(14,3,14,3,-2,1)（落 E12）
- 【预期】同 alert，D14='moveme' 不变

### 3. 越界 clamp+reject（不扩表、不丢数据）
- rowCount=1000/colCount=26。V.scrollY=996*24;V.render()（滚到底）
- setCell(998,0,'rowOob1');setCell(999,0,'rowOob2');dragSelect(998,0,999,0)
- dragMove(998,0,999,0,1,0)（目标行 1000 越界 → clamp 缩为 1 行 < 源 2 行 → reject）
- 【预期】__lastAlert 命中，rowCount 仍 1000（不扩表），A998/A999 不变（不丢数据）

### 4. undo
- setCell(15,1,'undoSrc');setStyle(15,1,{bold:true,italic:true});click(15,1);dragMove(15,1,15,1,2,0)
- 【预期】B17='undoSrc'+bold+italic，B15=null → key('z',{ctrl:true}) → B15 还原='undoSrc'+bold+italic，B17=null（一次 undo）

## F4 富剪贴板

### 1. 内部 copy/paste 保留样式+公式
- setCell(16,1,'42');setCell(16,2,'10');setCell(16,3,'=B16+C16');setStyle(16,3,{bold:true})
- click(16,3);copy() → click(18,3);paste()
- 【预期】D18.raw='=B18+C18'（公式保留为公式，copy 偏移 +2 行），D18.style.bold=true（style 保留），D16 不变

### 2. 多区域 copy 块间空行
- setCell(20,1,'first');setStyle(20,1,{bold:true,color:'#ff0000'});setCell(20,3,'second');setStyle(20,3,{italic:true})
- click(20,1);click(20,3,{ctrl})（[B20,D20]）→ copy()
- 【预期】__clipText='first\n\nsecond'（块间空行 \n\n，Excel 多区域语义）

### 3. 出站带样式（软验收、人工确认）
- 读 __clipHtml：每 area 一段 `<table>`，单元格 `style="..."` 内联 CSS（bold→font-weight:bold、color、italic→font-style:italic）
- 【预期】hasTable/hasBold/hasRed/hasItalic=true；外部应用粘入渲染由人工确认

### 4. 从外部粘入样式丢失（兜底 cell:{raw}）
- click(22,0);paste('alpha\tbeta')（不匹配内部指纹 → TSV 兜底）
- 【预期】A22.raw='alpha'/B22.raw='beta'，style=null（丢失）
- click(23,0);paste('=1+2') → 【预期】A23.raw='=1+2'（normalizedCell 保公式），style=null

### 5. undo paste
- click(16,5);setCell(16,5,'tag');setStyle(16,5,{bold:true,bg:'#00ff00'});click(16,5);copy()
- click(26,5);paste() → 【预期】F26='tag'+bold+green → key('z',{ctrl:true}) → F26=null（raw+style 一次 undo 清），F16 不变

## F5 公式编辑增强

### 1. 编辑公式画布对被引区域彩色框（每域一色）
- click(16,1) → kimi-webbridge `fill` `input.formula-input` 值 `=B16+C16`
- V.render() → 扫 overlay layer 的 Konva.Rect：stroke ∈ REF_PALETTE 且 dash=[3,2]
- 【预期】B16 处 stroke='#1a73e8'(palette[0])，C16 处 stroke='#ea4335'(palette[1])；每域一色循环

### 2. 函数名补全下拉（= 后前缀匹配 ↑↓/Tab/Enter/Esc）
- `fill` `input.formula-input`='=SU' → await 150ms（等 React flush）
- 【预期】`.autocomplete` 下拉，候选=[SUM,SUMIF]（前缀匹配）
- key2('ArrowDown')×2（→SUMIF）→ key2('Tab') → 【预期】input='=SUMIF('，下拉关
- 重置='=SU' → key2('Escape') → 【预期】下拉关，input 仍='=SU'（不补全）
- 重置='=SU' → await → key2('Enter') → 【预期】input='=SUM('（selIndex -1 接受 index 0）
- ='=CO' → ArrowDown→ArrowUp（clamp 0）→ Enter → 【预期】'=COUNT('

---
【验收记录】 kimi-webbridge F2-F5 全 PASS（见 docs/squad/m3c/acceptance.md）：
F2 多区域（Ctrl 加/反选/LIFO/Shift 生长/格式全域/填充活动区/排序筛选禁用/表头并集）✓
F3 拖动（cut raw+style/公式不 shift/相交拒绝/merge 拒绝/越界 clamp+reject 不扩表不丢数据/undo）✓
F4 富剪贴板（内部样式+公式/多区域空行/出站 HTML 内联 CSS/外部兜底 cell:{raw}/undo）✓
F5 公式编辑（被引区域彩色框每域一色/函数名补全 ↑↓Tab Enter Esc）✓
额外发现：①F2.3 toggleRange 移除活动区域后 activeCell 落在被移除 range（§1.1 不变式违反，可致误编辑）②=B16+C16 对文本数字格返回 #VALUE! 而 =B16*C16 强转（+/· 强转不一致，M3a 范畴）
注意：①行 51+ 须先 scrollY 滚到可见再合成事件；②fill 后须 await 150ms 让 React flush 补全态；③mock clipboardData 让内部指纹 tsv 往返 copy->paste；④工具栏 disabled 须 await React 重渲染后读。
