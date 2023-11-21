// 公式词法分析：把公式体（不含前导 '='）切成 token 流。
// 纯函数，零 DOM 依赖。
import { fromA1 } from '../core/addr'

export type TokenType =
  | 'num'
  | 'str'
  | 'ident'
  | 'sheetname' // '...' 引号表名（'' 转义）
  | 'cellref' // 可带 $ 前缀，value 保留原文（如 $A$1）
  | 'errlit' // #REF! 字面量
  | 'op'
  | 'lparen'
  | 'rparen'
  | 'comma'
  | 'colon'

export interface Token {
  type: TokenType
  value: string // num/str 为解析后的文本；cellref 为原文（含 $）；sheetname 为去引号文本；op 为运算符
}

export class LexError extends Error {}

const NUM_RE = /^(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/
const CELLREF_RE = /^\$?[A-Za-z]+\$?[0-9]+/

export function tokenize(src: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  while (i < src.length) {
    const ch = src[i]
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++
      continue
    }
    // 数字（含科学计数法）
    if ((ch >= '0' && ch <= '9') || (ch === '.' && src[i + 1] >= '0' && src[i + 1] <= '9')) {
      const m = NUM_RE.exec(src.slice(i))
      if (!m) throw new LexError(`bad number at ${i}`)
      tokens.push({ type: 'num', value: m[0] })
      i += m[0].length
      continue
    }
    // 字符串："..."，`""` 为转义的双引号
    if (ch === '"') {
      let j = i + 1
      let out = ''
      for (;;) {
        if (j >= src.length) throw new LexError('unterminated string')
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            out += '"'
            j += 2
          } else {
            j++
            break
          }
        } else {
          out += src[j]
          j++
        }
      }
      tokens.push({ type: 'str', value: out })
      i = j
      continue
    }
    // 引号表名：'...'，`''` 为转义的单引号
    if (ch === "'") {
      let j = i + 1
      let out = ''
      for (;;) {
        if (j >= src.length) throw new LexError('unterminated sheet name')
        if (src[j] === "'") {
          if (src[j + 1] === "'") {
            out += "'"
            j += 2
          } else {
            j++
            break
          }
        } else {
          out += src[j]
          j++
        }
      }
      tokens.push({ type: 'sheetname', value: out })
      i = j
      continue
    }
    // #REF! 字面量
    if (ch === '#') {
      if (src.startsWith('#REF!', i)) {
        tokens.push({ type: 'errlit', value: '#REF!' })
        i += 5
        continue
      }
      throw new LexError(`unexpected character: #`)
    }
    // cellref（可带 $）或 ident
    if (ch === '$' || (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')) {
      const m = CELLREF_RE.exec(src.slice(i))
      if (m) {
        const text = m[0]
        // 后随 '!' 的是表名前缀而非单元格引用（如 Sheet2!A1；fromA1 不限列名长度，
        // 否则 Sheet2 会被误吞为 cellref）→ 按 ident 交给 parser 的表名分支
        if (src[i + text.length] === '!') {
          tokens.push({ type: 'ident', value: text })
          i += text.length
          continue
        }
        if (fromA1(text.replace(/\$/g, ''))) {
          tokens.push({ type: 'cellref', value: text })
          i += text.length
          continue
        }
        throw new LexError(`bad cell reference: ${text}`)
      }
      if (ch === '$') throw new LexError('unexpected character: $')
      let j = i
      while (j < src.length && /[A-Za-z]/.test(src[j])) j++
      tokens.push({ type: 'ident', value: src.slice(i, j) })
      i = j
      continue
    }
    if (ch === '(') {
      tokens.push({ type: 'lparen', value: '(' })
      i++
      continue
    }
    if (ch === ')') {
      tokens.push({ type: 'rparen', value: ')' })
      i++
      continue
    }
    if (ch === ',') {
      tokens.push({ type: 'comma', value: ',' })
      i++
      continue
    }
    if (ch === ':') {
      tokens.push({ type: 'colon', value: ':' })
      i++
      continue
    }
    // 双字符运算符优先
    const two = src.slice(i, i + 2)
    if (two === '<>' || two === '<=' || two === '>=') {
      tokens.push({ type: 'op', value: two })
      i += 2
      continue
    }
    if ('+-*/^&%=<>!'.includes(ch)) {
      tokens.push({ type: 'op', value: ch })
      i++
      continue
    }
    throw new LexError(`unexpected character: ${ch}`)
  }
  return tokens
}
