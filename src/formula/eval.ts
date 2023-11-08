// 公式求值：对 AST 递归求值，产出 FormulaValue。
// 错误沿调用链传播（第一个错误胜出）；区域只允许出现在函数参数位置。
import { SheetId } from '../core/model'
import { AST } from './parser'

export type FormulaValue = number | string | boolean | FormulaError
export interface FormulaError {
  error: string // '#REF!' '#DIV/0!' '#NAME?' '#VALUE!' '#CYCLE!'
}

export function isError(v: unknown): v is FormulaError {
  return typeof v === 'object' && v !== null && 'error' in v
}

export interface EvalCtx {
  sheet: SheetId // 当前公式所在表（M1 引用均为本表）
  get(sheet: SheetId, row: number, col: number): FormulaValue
}

// 内部哨兵：标记「空单元格引用」（仅 ref 求值路径产生）。
// 空单元格在算术/比较中按 0、在拼接中按 ''、在聚合函数中被忽略、显示为 ''。
// 与公式里显式写出的空串字面量 "" 区分："" 始终是字符串语义。
// 该哨兵不越过公式边界——engine 在缓存/返回前 unwrap 为 ''。
interface Blank {
  readonly blank: true
}
const BLANK: Blank = { blank: true }
type V = FormulaValue | Blank

export function isBlank(v: unknown): v is Blank {
  return v === BLANK
}

const err = (error: string): FormulaError => ({ error })

export function evalNode(node: AST, ctx: EvalCtx): V {
  switch (node.type) {
    case 'num':
    case 'str':
    case 'bool':
      return node.value
    case 'ref': {
      // 空单元格（''）标记为 BLANK：比较/算术按 0，拼接按 ''（仅 ref 路径）
      const v = ctx.get(ctx.sheet, node.addr.row, node.addr.col)
      return v === '' ? BLANK : v
    }
    case 'range':
      // 区域只在函数参数位置合法（evalCall 特判）；其他位置 → #VALUE!
      return err('#VALUE!')
    case 'paren':
      return evalNode(node.expr, ctx)
    case 'unary': {
      const v = evalNode(node.expr, ctx)
      if (isError(v)) return v
      const n = toNum(v)
      if (isError(n)) return n
      return node.op === '-' ? -n : n
    }
    case 'percent': {
      const v = evalNode(node.expr, ctx)
      if (isError(v)) return v
      const n = toNum(v)
      if (isError(n)) return n
      return n / 100
    }
    case 'binary':
      return evalBinary(node.op, node.left, node.right, ctx)
    case 'call':
      return evalCall(node.name, node.args, ctx)
  }
}

function evalBinary(op: string, l: AST, r: AST, ctx: EvalCtx): V {
  const a = evalNode(l, ctx)
  if (isError(a)) return a
  const b = evalNode(r, ctx)
  if (isError(b)) return b
  if (op === '&') return toText(a) + toText(b)
  if (op === '+' || op === '-' || op === '*' || op === '/' || op === '^') {
    const x = toNum(a)
    if (isError(x)) return x
    const y = toNum(b)
    if (isError(y)) return y
    switch (op) {
      case '+':
        return x + y
      case '-':
        return x - y
      case '*':
        return x * y
      case '/':
        return y === 0 ? err('#DIV/0!') : x / y
      case '^':
        return Math.pow(x, y)
    }
  }
  // 比较
  const c = compareValues(a, b)
  switch (op) {
    case '=':
      return c === 0
    case '<>':
      return c !== 0
    case '<':
      return c < 0
    case '<=':
      return c <= 0
    case '>':
      return c > 0
    case '>=':
      return c >= 0
  }
  return err('#VALUE!')
}

// 算术 coercion：数字原样；布尔 TRUE→1/FALSE→0；空单元格（BLANK）与空串 → 0；
// 纯数字字符串 → 数字（Excel 同款隐式转换）；其余字符串 → #VALUE!
function toNum(v: V): number | FormulaError {
  if (isError(v)) return v
  if (isBlank(v)) return 0
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  const t = v.trim()
  if (t === '') return 0
  const n = Number(t)
  return Number.isNaN(n) ? err('#VALUE!') : n
}

function toText(v: V): string {
  if (isBlank(v)) return '' // 空单元格拼接为空串
  if (typeof v === 'string') return v
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'number') return formatNumber(v)
  return v.error
}

function toBool(v: V): boolean | FormulaError {
  if (isError(v)) return v
  if (isBlank(v)) return false // 空单元格按 0 → false
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  const t = v.trim().toUpperCase()
  if (t === 'TRUE') return true
  if (t === 'FALSE') return false
  return err('#VALUE!')
}

// Excel 比较序：数字 < 字符串 < 布尔；字符串不区分大小写。
// 空单元格（BLANK，仅 ref 路径产生）在比较中按 0；空串字面量 "" 仍是字符串。
function compareValues(a: V, b: V): number {
  const x: FormulaValue = isBlank(a) ? 0 : a
  const y: FormulaValue = isBlank(b) ? 0 : b
  const rank = (v: FormulaValue) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)
  const ra = rank(x)
  const rb = rank(y)
  if (ra !== rb) return ra < rb ? -1 : 1
  if (typeof x === 'number' && typeof y === 'number') return x < y ? -1 : x > y ? 1 : 0
  if (typeof x === 'string' && typeof y === 'string') {
    const lx = x.toLowerCase()
    const ly = y.toLowerCase()
    return lx < ly ? -1 : lx > ly ? 1 : 0
  }
  const bx = x as boolean
  const by = y as boolean
  return bx === by ? 0 : bx ? 1 : -1
}

// ---- 函数 ----

type Fn = (args: AST[], ctx: EvalCtx) => V

// 聚合类（SUM/AVERAGE/COUNT/MAX/MIN）：参数可为值或区域，忽略字符串/空白/布尔
function aggregateArgs(args: AST[], ctx: EvalCtx): number[] | FormulaError {
  const nums: number[] = []
  const push = (v: V): FormulaError | null => {
    if (isError(v)) return v
    if (typeof v === 'number') nums.push(v)
    return null
  }
  for (const arg of args) {
    if (arg.type === 'range') {
      const r = arg.range
      for (let row = r.sr; row <= r.er; row++) {
        for (let col = r.sc; col <= r.ec; col++) {
          const e = push(ctx.get(ctx.sheet, row, col))
          if (e) return e
        }
      }
    } else {
      const e = push(evalNode(arg, ctx))
      if (e) return e
    }
  }
  return nums
}

const AGGREGATES: Record<string, (nums: number[]) => FormulaValue> = {
  SUM: (ns) => ns.reduce((a, b) => a + b, 0),
  AVERAGE: (ns) => (ns.length === 0 ? err('#DIV/0!') : ns.reduce((a, b) => a + b, 0) / ns.length),
  COUNT: (ns) => ns.length,
  MAX: (ns) => (ns.length === 0 ? 0 : Math.max(...ns)),
  MIN: (ns) => (ns.length === 0 ? 0 : Math.min(...ns)),
}

const FUNCTIONS: Record<string, Fn> = {
  ABS: (args, ctx) => {
    if (args.length !== 1) return err('#VALUE!')
    const v = evalNode(args[0], ctx)
    if (isError(v)) return v
    const n = toNum(v)
    return isError(n) ? n : Math.abs(n)
  },
  ROUND: (args, ctx) => {
    if (args.length < 1 || args.length > 2) return err('#VALUE!')
    const v = evalNode(args[0], ctx)
    if (isError(v)) return v
    const n = toNum(v)
    if (isError(n)) return n
    let digits = 0
    if (args.length === 2) {
      const d = evalNode(args[1], ctx)
      if (isError(d)) return d
      const dn = toNum(d)
      if (isError(dn)) return dn
      digits = Math.trunc(dn)
    }
    // Excel ROUND： halves 远离 0（ROUND(2.5)=3, ROUND(-2.5)=-3）
    const f = Math.pow(10, digits)
    return (Math.sign(n) * Math.round(Math.abs(n) * f)) / f
  },
  IF: (args, ctx) => {
    if (args.length !== 3) return err('#VALUE!')
    const c = evalNode(args[0], ctx)
    if (isError(c)) return c
    const cond = toBool(c)
    if (isError(cond)) return cond
    return evalNode(cond ? args[1] : args[2], ctx)
  },
}

function evalCall(name: string, args: AST[], ctx: EvalCtx): V {
  const agg = AGGREGATES[name]
  if (agg) {
    const nums = aggregateArgs(args, ctx)
    return isError(nums) ? nums : agg(nums)
  }
  const fn = FUNCTIONS[name]
  if (!fn) return err('#NAME?')
  return fn(args, ctx)
}

// 数字显示：String(n) 自然行为；长度 >12 → toPrecision(10) 去尾零
export function formatNumber(n: number): string {
  const s = String(n)
  return s.length > 12 ? Number(n.toPrecision(10)).toString() : s
}
