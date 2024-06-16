// F5 公式文本纯函数：被引区域抽取 + 函数名补全候选。零 DOM；view/react 只读引用。
// 跨表引用（Sheet2!A1）不返回（画布高亮仅当前表，跨表为后续项）。
import { CellRange, normalizeRange } from '../core/addr'
import { AST, parseFormula } from './parser'

// 遍历 AST 收集当前表（无 sheet 前缀）的 range/ref 节点 → CellRange[]。
// num/str/bool/err 无子树，switch 默认分支跳过。
function walk(node: AST, out: CellRange[]): void {
  switch (node.type) {
    case 'range':
      if (!node.a.sheet && !node.b.sheet)
        out.push(normalizeRange({ sr: node.a.row, sc: node.a.col, er: node.b.row, ec: node.b.col }))
      break
    case 'ref':
      if (!node.ref.sheet)
        out.push(normalizeRange({ sr: node.ref.row, sc: node.ref.col, er: node.ref.row, ec: node.ref.col }))
      break
    case 'call': for (const a of node.args) walk(a, out); break
    case 'binary': walk(node.left, out); walk(node.right, out); break
    case 'unary': case 'percent': case 'paren': walk(node.expr, out); break
  }
}

// 从公式文本抽取当前表（无表名前缀）的被引区域。非公式 / 语法错 → 空（不抛）。
export function extractCurrentSheetRanges(src: string): CellRange[] {
  if (!src.startsWith('=')) return []
  try {
    const ast = parseFormula(src.slice(1))
    const out: CellRange[] = []
    walk(ast, out)
    return out
  } catch {
    return []
  }
}

// 函数名补全候选：取最后一个 '=' 之后的末尾 [A-Za-z]+ 标识符 token，
// 转大写前缀匹配函数名表（函数名约定大写），截断到 8 个。
// 非公式 / 无末尾标识符 → 空。
export function completionCandidates(text: string, names: string[]): string[] {
  const eq = text.lastIndexOf('=')
  if (eq < 0) return []
  const m = /[A-Za-z]+$/.exec(text.slice(eq + 1))
  if (!m) return []
  const prefix = m[0].toUpperCase()
  return names.filter((n) => n.startsWith(prefix)).slice(0, 8)
}
