// 公式语法分析：递归下降，产出 AST。
// 优先级（低 → 高）：比较 = <> < <= > >= → 拼接 & → 加减 → 乘除 → 幂 ^（右结合）
//   → 一元 -x / +x / 后缀 x% → 原子。
// 注意：一元负号绑定优先于 ^，即 =-1^2 解析为 (-1)^2 = 1。
//   这是 Excel 的实际语义（Excel 中 -1^2 = 1），有意与数学惯例 -(1^2) 不同，勿改。
import { CellAddr, CellRange, fromA1, normalizeRange } from '../core/addr'
import { LexError, Token, tokenize } from './lexer'

export type AST =
  | { type: 'num'; value: number }
  | { type: 'str'; value: string }
  | { type: 'bool'; value: boolean }
  | { type: 'ref'; addr: CellAddr }
  | { type: 'range'; range: CellRange }
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

  private parseAtom(): AST {
    const t = this.next()
    switch (t.type) {
      case 'num':
        return { type: 'num', value: Number(t.value) }
      case 'str':
        return { type: 'str', value: t.value }
      case 'ident': {
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
        throw new ParseError(`unknown name: ${t.value}`)
      }
      case 'cellref': {
        const a = fromA1(t.value)
        if (!a) throw new ParseError(`bad cell reference: ${t.value}`)
        if (this.peek()?.type === 'colon') {
          this.next() // consume ':'
          const end = this.next()
          if (end.type !== 'cellref') throw new ParseError(`expected cell reference after ':'`)
          const b = fromA1(end.value)
          if (!b) throw new ParseError(`bad cell reference: ${end.value}`)
          return {
            type: 'range',
            range: normalizeRange({ sr: a.row, sc: a.col, er: b.row, ec: b.col }),
          }
        }
        return { type: 'ref', addr: a }
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
