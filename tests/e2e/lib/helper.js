// 页面注入源码：W helper（原样搬自 tests/e2e/m3c.md）+ m4b.md 的 createElement/confirm stub。
// 经 evaluate 注入；每次刷新页面后须重注。
export const HELPER_SOURCE = `window.W = (() => {
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
  // （m3c.md 原文 W_copy/W.paste 依赖注入时全局已有 W，首注必 ReferenceError；改为局部 const）
  const W_copy = function(){ const V0=V(); V0.focus(); const store={}
    const ev={clipboardData:{setData:(t,d)=>store[t]=d,getData:(t)=>store[t]||''},preventDefault:()=>{}}
    V0.onCopy(ev); window.W.__clipText=store['text/plain']||''; window.W.__clipHtml=store['text/html']||''
    return window.W.__clipText.length }
  const paste = function(text){ const V0=V(); V0.focus(); const t=(text!=null)?text:(window.W.__clipText||'')
    const ev={clipboardData:{getData:(type)=>type==='text/plain'?t:'',setData:()=>{}},preventDefault:()=>{}}
    V0.onPaste(ev); return t.length }
  const read=(r,c)=>{const cell=V().state.activeSheet.getCell(r,c);return cell?{raw:cell.raw,style:cell.style||null}:null}
  const sel=()=>{const s=V().state.selection;return{ranges:s.ranges,activeCell:s.activeCell}}
  const setCell=(r,c,raw)=>V().dispatch(V().state.tr.setCell(r,c,raw))
  const setStyle=(r,c,st)=>V().dispatch(V().state.tr.setCell(r,c,V().state.activeSheet.getCell(r,c)?.raw??'',st))
  return {click,dragSelect,dragMove,key,read,sel,setCell,setStyle,cellXY,borderXY,V,
    copy:W_copy,paste,render:()=>V().render()}
})()
window.__origCreateElement = document.createElement.bind(document)
window.__origConfirm = window.confirm
window.__lastInput = null
document.createElement = (tag, ...rest) => {
  const el = window.__origCreateElement(tag, ...rest)
  if (tag === 'input') { window.__lastInput = el; el.click = () => {} }
  return el
}
window.confirm = () => true
'helper-ready'`
