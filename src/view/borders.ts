// 边框渲染裁决：每条网格线只画一次，共享边取线宽权重高者，同重取左/上格。
// 不 import konva（node 可测）；绘制描述由 layers 消费。
import { BorderEdge, BorderLineStyle, SheetData } from '../core/model'

// 线宽权重：hair 0.5 / thin 1 / medium 2 / thick 3；装饰线型按对应线宽档
export function edgeWeight(e: BorderEdge): number {
  switch (e.style) {
    case 'hair': return 0.5
    case 'thin': case 'dotted': case 'dashed': return 1
    case 'medium': case 'mediumDashed': case 'double': return 2
    case 'thick': return 3
  }
}

export function edgeWidth(style: BorderLineStyle): number {
  switch (style) {
    case 'hair': case 'thin': case 'dotted': case 'dashed': return 1
    case 'medium': case 'mediumDashed': return 2
    case 'double': return 3 // 两条 1px + 1px 间距
    case 'thick': return 3
  }
}

export function edgeDash(style: BorderLineStyle): number[] | undefined {
  switch (style) {
    case 'dashed': return [6, 3]
    case 'dotted': return [2, 2]
    case 'mediumDashed': return [8, 4]
    default: return undefined
  }
}

// col c 左缘竖线：c-1 的 right vs c 的 left，同重取左格
export function resolveVEdge(sheet: SheetData, r: number, c: number): BorderEdge | undefined {
  const left = c > 0 ? sheet.getCell(r, c - 1)?.style?.border?.right : undefined
  const right = c < sheet.colCount ? sheet.getCell(r, c)?.style?.border?.left : undefined
  if (!left) return right
  if (!right) return left
  return edgeWeight(left) >= edgeWeight(right) ? left : right
}

// row r 上缘横线：r-1 的 bottom vs r 的 top，同重取上格
export function resolveHEdge(sheet: SheetData, r: number, c: number): BorderEdge | undefined {
  const up = r > 0 ? sheet.getCell(r - 1, c)?.style?.border?.bottom : undefined
  const down = r < sheet.rowCount ? sheet.getCell(r, c)?.style?.border?.top : undefined
  if (!up) return down
  if (!down) return up
  return edgeWeight(up) >= edgeWeight(down) ? up : down
}
