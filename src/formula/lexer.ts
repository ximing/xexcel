// 公式词法分析：把公式体（不含前导 '='）切成 token 流。
// 纯函数，零 DOM 依赖。
import { fromA1 } from '../core/addr'

export type TokenType =
  | 'num'
  | 'str'
  | 'ident'
  | 'cellref'
  | 'op'
  | 'lparen'
  | 'rparen'
  | 'comma'
  | 'colon'

export interface Token {
  type: TokenType
  value: string // num/str 为解析后的文本；cellref 为 A1 原文；op 为运算符
}

export class LexError extends Error {}

const NUM_RE = /^(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/

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
    // 字母开头：cellref（字母+数字且是合法 A1）或 ident（纯字母：TRUE/FALSE/函数名）
    if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')) {
      let j = i
      while (j < src.length && /[A-Za-z]/.test(src[j])) j++
      const letters = src.slice(i, j)
      let k = j
      while (k < src.length && /[0-9]/.test(src[k])) k++
      if (k > j) {
        const candidate = src.slice(i, k)
        if (fromA1(candidate)) {
          tokens.push({ type: 'cellref', value: candidate })
          i = k
          continue
        }
        throw new LexError(`bad cell reference: ${candidate}`)
      }
      tokens.push({ type: 'ident', value: letters })
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
    if ('+-*/^&%=<>'.includes(ch)) {
      tokens.push({ type: 'op', value: ch })
      i++
      continue
    }
    throw new LexError(`unexpected character: ${ch}`)
  }
  return tokens
}
