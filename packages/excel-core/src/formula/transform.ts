// 公式引用平移：fill/copy 时把公式中的相对引用按 (dRow, dCol) 偏移。
// $ 维度锁定不动；偏移后越界（row/col < 0）→ 该引用替换为 #REF! 错误节点。
import { AST, RefTarget, parseFormula } from './parser'
import { serialize } from './serialize'
import { registerStructureCascade } from '../core/steps'

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

// 结构操作（插入/删除行列）的引用级联。sheet 为被改表表名（大小写不敏感）；
// hostSheet 为公式所在表表名（无表名引用按 host 判定）。
export interface StructureSpec {
  sheet: string
  axis: 'row' | 'col'
  index: number
  count: number
  mode: 'insert' | 'delete'
}

function adjustRefForStructure(r: RefTarget, spec: StructureSpec, hostSheet: string): RefTarget | null {
  const name = (r.sheet ?? hostSheet).toLowerCase()
  if (name !== spec.sheet.toLowerCase()) return r
  const dim = spec.axis === 'row' ? r.row : r.col
  let next: number
  if (spec.mode === 'insert') {
    next = dim >= spec.index ? dim + spec.count : dim
  } else {
    if (dim >= spec.index && dim < spec.index + spec.count) return null // 落在删除区
    next = dim >= spec.index + spec.count ? dim - spec.count : dim
  }
  return spec.axis === 'row' ? { ...r, row: next } : { ...r, col: next }
}

// range 端点落入删除区 → 该 range 节点局部替换为 #REF!（与 M2a shiftRefs 同款语义，
// 不外溢到整个公式：'=SUM(A3:A6)+B1' → '=SUM(#REF!)+B1'）。
export function adjustForStructure(node: AST, spec: StructureSpec, hostSheet: string): AST {
  switch (node.type) {
    case 'ref': {
      const r = adjustRefForStructure(node.ref, spec, hostSheet)
      return r === null ? REF_ERR : { type: 'ref', ref: r }
    }
    case 'range': {
      const a = adjustRefForStructure(node.a, spec, hostSheet)
      const b = adjustRefForStructure(node.b, spec, hostSheet)
      if (a === null || b === null) return REF_ERR
      return { type: 'range', a, b }
    }
    case 'call':
      return { type: 'call', name: node.name, args: node.args.map((x) => adjustForStructure(x, spec, hostSheet)) }
    case 'unary':
      return { type: 'unary', op: node.op, expr: adjustForStructure(node.expr, spec, hostSheet) }
    case 'binary':
      return {
        type: 'binary',
        op: node.op,
        left: adjustForStructure(node.left, spec, hostSheet),
        right: adjustForStructure(node.right, spec, hostSheet),
      }
    case 'percent':
      return { type: 'percent', expr: adjustForStructure(node.expr, spec, hostSheet) }
    case 'paren':
      return { type: 'paren', expr: adjustForStructure(node.expr, spec, hostSheet) }
    default:
      return node
  }
}

// raw 文本级入口：公式 → 级联后新文本；非公式/解析失败 → 原文返回
export function adjustFormulaForStructure(raw: string, spec: StructureSpec, hostSheet: string): string {
  if (!raw.startsWith('=')) return raw
  try {
    return '=' + serialize(adjustForStructure(parseFormula(raw.slice(1)), spec, hostSheet))
  } catch {
    return raw
  }
}

// 注入 core 的 StructureStep（保持 core → formula 无依赖）
registerStructureCascade((raw, spec, host) => adjustFormulaForStructure(raw, spec, host))
