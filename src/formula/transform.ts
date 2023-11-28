// 公式引用平移：fill/copy 时把公式中的相对引用按 (dRow, dCol) 偏移。
// $ 维度锁定不动；偏移后越界（row/col < 0）→ 该引用替换为 #REF! 错误节点。
import { AST, RefTarget, parseFormula } from './parser'
import { serialize } from './serialize'

// null = 越界
function shiftRef(r: RefTarget, dRow: number, dCol: number): RefTarget | null {
  const row = r.rowAbs ? r.row : r.row + dRow
  const col = r.colAbs ? r.col : r.col + dCol
  if (row < 0 || col < 0) return null
  return { ...r, row, col }
}

const REF_ERR: AST = { type: 'err', error: '#REF!' }

export function shiftRefs(node: AST, dRow: number, dCol: number): AST {
  switch (node.type) {
    case 'ref': {
      const r = shiftRef(node.ref, dRow, dCol)
      return r === null ? REF_ERR : { type: 'ref', ref: r }
    }
    case 'range': {
      const a = shiftRef(node.a, dRow, dCol)
      const b = shiftRef(node.b, dRow, dCol)
      if (a === null || b === null) return REF_ERR
      return { type: 'range', a, b }
    }
    case 'call':
      return { type: 'call', name: node.name, args: node.args.map((x) => shiftRefs(x, dRow, dCol)) }
    case 'unary':
      return { type: 'unary', op: node.op, expr: shiftRefs(node.expr, dRow, dCol) }
    case 'binary':
      return {
        type: 'binary',
        op: node.op,
        left: shiftRefs(node.left, dRow, dCol),
        right: shiftRefs(node.right, dRow, dCol),
      }
    case 'percent':
      return { type: 'percent', expr: shiftRefs(node.expr, dRow, dCol) }
    case 'paren':
      return { type: 'paren', expr: shiftRefs(node.expr, dRow, dCol) }
    default:
      return node // num/str/bool/err 叶子不动
  }
}

// cell raw 文本级入口：公式 → 偏移后新文本；非公式/解析失败 → 原文返回
export function shiftFormula(raw: string, dRow: number, dCol: number): string {
  if (!raw.startsWith('=')) return raw
  try {
    return '=' + serialize(shiftRefs(parseFormula(raw.slice(1)), dRow, dCol))
  } catch {
    return raw
  }
}
