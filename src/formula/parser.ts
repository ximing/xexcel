// 公式语法分析：递归下降，产出 AST。
// 优先级（低 → 高）：比较 = <> < <= > >= → 拼接 & → 加减 → 乘除 → 幂 ^（右结合）
//   → 一元 -x / +x / 后缀 x% → 原子。
// 注意：一元负号绑定优先于 ^，即 =-1^2 解析为 (-1)^2 = 1。
//   这是 Excel 的实际语义（Excel 中 -1^2 = 1），有意与数学惯例 -(1^2) 不同，勿改。
// ref/range 节点携带 RefTarget：$ 标志（rowAbs/colAbs）与可选表名（sheet）。
// range 的 a/b 保留书写顺序不归一；归一发生在 eval。
import { parseColName } from '../core/addr'
import { LexError, Token, tokenize } from './lexer'

export interface RefTarget {
  sheet?: string // 表名（原文大小写；eval 时不区分大小写解析为 SheetId）
  row: number
  col: number
  rowAbs: boolean
  colAbs: boolean
}

export type AST =
  | { type: 'num'; value: number }
  | { type: 'str'; value: string }
  | { type: 'bool'; value: boolean }
  | { type: 'err'; error: string } // '#REF!'（shiftRefs 越界/用户原文）或 '#NAME?'（未知裸名）
  | { type: 'ref'; ref: RefTarget }
  | { type: 'range'; a: RefTarget; b: RefTarget }
  | { type: 'call'; name: string; args: AST[] }
  | { type: 'unary'; op: '+' | '-'; expr: AST }
  | { type: 'binary'; op: string; left: AST; right: AST }
  | { type: 'percent'; expr: AST }
  | { type: 'paren'; expr: AST }

export class ParseError extends Error {}

export function parseFormula(src: string): AST {
  const tokens = tokenize(src)
  const p = new Parser(tokens)
  const ast = p.parseCompare()
  if (p.peek() !== null) throw new ParseError(`unexpected token: ${p.peek()!.value}`)
  return ast
}

// '$A$1' / 'A1' 等 cellref token 原文 → RefTarget
function parseCellRefText(text: string, sheet?: string): RefTarget {
  const m = /^(\$?)([A-Za-z]+)(\$?)([0-9]+)$/.exec(text)
  if (!m) throw new ParseError(`bad cell reference: ${text}`)
  const col = parseColName(m[2])
  if (col < 0) throw new ParseError(`bad cell reference: ${text}`)
  const ref: RefTarget = {
    row: parseInt(m[4], 10) - 1,
    col,
    rowAbs: m[3] === '$',
    colAbs: m[1] === '$',
  }
  if (sheet !== undefined) ref.sheet = sheet
  return ref
}

class Parser {
  private pos = 0
  constructor(private tokens: Token[]) {}

  peek(): Token | null {
    return this.tokens[this.pos] ?? null
  }

  private next(): Token {
    const t = this.tokens[this.pos]
    if (!t) throw new ParseError('unexpected end of formula')
    this.pos++
    return t
  }

  private eatOp(...ops: string[]): string | null {
    const t = this.peek()
    if (t && t.type === 'op' && ops.includes(t.value)) {
      this.pos++
      return t.value
    }
    return null
  }

  private peekBang(): boolean {
    const t = this.peek()
    return t !== null && t.type === 'op' && t.value === '!'
  }

  // 比较（最低优先级，左结合）
  parseCompare(): AST {
    let left = this.parseConcat()
    for (;;) {
      const op = this.eatOp('=', '<>', '<=', '>=', '<', '>')
      if (!op) return left
      left = { type: 'binary', op, left, right: this.parseConcat() }
    }
  }

  private parseConcat(): AST {
    let left = this.parseAddSub()
    for (;;) {
      const op = this.eatOp('&')
      if (!op) return left
      left = { type: 'binary', op, left, right: this.parseAddSub() }
    }
  }

  private parseAddSub(): AST {
    let left = this.parseMulDiv()
    for (;;) {
      const op = this.eatOp('+', '-')
      if (!op) return left
      left = { type: 'binary', op, left, right: this.parseMulDiv() }
    }
  }

  private parseMulDiv(): AST {
    let left = this.parsePower()
    for (;;) {
      const op = this.eatOp('*', '/')
      if (!op) return left
      left = { type: 'binary', op, left, right: this.parsePower() }
    }
  }

  // 幂：右结合。底数走一元层 → 一元负号绑定比 ^ 紧（Excel 语义：-1^2 = (-1)^2 = 1）
  private parsePower(): AST {
    const base = this.parseUnary()
    if (this.eatOp('^')) {
      return { type: 'binary', op: '^', left: base, right: this.parsePower() }
    }
    return base
  }

  private parseUnary(): AST {
    const op = this.eatOp('-', '+')
    if (op) return { type: 'unary', op: op as '+' | '-', expr: this.parseUnary() }
    return this.parsePostfix()
  }

  // 后缀百分号：x% → x/100，可叠放
  private parsePostfix(): AST {
    let node = this.parseAtom()
    while (this.eatOp('%')) node = { type: 'percent', expr: node }
    return node
  }

  // 表名! 前缀已消费前的入口：name 为裸名或引号名文本
  private parseSheetRef(name: string): AST {
    this.next() // consume '!'
    const first = this.next()
    if (first.type !== 'cellref') throw new ParseError(`expected cell reference after '!'`)
    const a = parseCellRefText(first.value, name)
    if (this.peek()?.type !== 'colon') return { type: 'ref', ref: a }
    this.next() // consume ':'
    // 第二端允许重复表名前缀（同表），异表 → 语法错误
    let end = this.next()
    if ((end.type === 'ident' || end.type === 'sheetname') && this.peekBang()) {
      if (end.value.toLowerCase() !== name.toLowerCase()) {
        throw new ParseError('range endpoints on different sheets')
      }
      this.next() // consume '!'
      end = this.next()
    }
    if (end.type !== 'cellref') throw new ParseError(`expected cell reference after ':'`)
    return { type: 'range', a, b: parseCellRefText(end.value, name) }
  }

  private parseAtom(): AST {
    const t = this.next()
    switch (t.type) {
      case 'num':
        return { type: 'num', value: Number(t.value) }
      case 'str':
        return { type: 'str', value: t.value }
      case 'errlit':
        return { type: 'err', error: t.value }
      case 'sheetname': {
        if (this.peekBang()) return this.parseSheetRef(t.value)
        throw new ParseError(`unexpected quoted name: ${t.value}`)
      }
      case 'ident': {
        if (this.peekBang()) return this.parseSheetRef(t.value)
        const upper = t.value.toUpperCase()
        if (this.peek()?.type === 'lparen') {
          this.next() // consume '('
          const args: AST[] = []
          if (this.peek()?.type !== 'rparen') {
            for (;;) {
              args.push(this.parseCompare())
              const sep = this.next()
              if (sep.type === 'rparen') break
              if (sep.type !== 'comma') throw new ParseError(`expected ',' or ')', got ${sep.value}`)
            }
          } else {
            this.next() // consume ')'
          }
          return { type: 'call', name: upper, args }
        }
        if (upper === 'TRUE') return { type: 'bool', value: true }
        if (upper === 'FALSE') return { type: 'bool', value: false }
        // 未知裸名（如 NOPE）→ 错误节点而非语法错误：求值期产 #NAME?，可被 IFERROR 兜住
        return { type: 'err', error: '#NAME?' }
      }
      case 'cellref': {
        const a = parseCellRefText(t.value)
        if (this.peek()?.type === 'colon') {
          this.next() // consume ':'
          const end = this.next()
          if (end.type !== 'cellref') throw new ParseError(`expected cell reference after ':'`)
          return { type: 'range', a, b: parseCellRefText(end.value) }
        }
        return { type: 'ref', ref: a }
      }
      case 'lparen': {
        const inner = this.parseCompare()
        const close = this.next()
        if (close.type !== 'rparen') throw new ParseError(`expected ')', got ${close.value}`)
        return { type: 'paren', expr: inner }
      }
      default:
        throw new ParseError(`unexpected token: ${t.value}`)
    }
  }
}

export { LexError }
