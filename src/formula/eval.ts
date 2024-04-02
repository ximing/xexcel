// 公式求值：对 AST 递归求值，产出 FormulaValue。
// 错误沿调用链传播（第一个错误胜出）；区域在标量位置按隐式交集求值。
import { normalizeRange } from '../core/addr'
import { SheetId } from '../core/model'
import { matchCriteria } from './criteria'
import { dateSerialLenient, nowSerial, serialToDate, todaySerial } from './date'
import { AST } from './parser'

export type FormulaValue = number | string | boolean | FormulaError
export interface FormulaError {
  error: string // '#REF!' '#DIV/0!' '#NAME?' '#VALUE!' '#CYCLE!'
}

export function isError(v: unknown): v is FormulaError {
  return typeof v === 'object' && v !== null && 'error' in v
}

export interface EvalCtx {
  sheet: SheetId // 当前公式所在表
  row: number // 公式所在行（隐式交集用）
  col: number // 公式所在列
  get(sheet: SheetId, row: number, col: number): FormulaValue
  resolveSheet(name: string): SheetId | null // 表名（不区分大小写）→ SheetId；未知 → null
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
    case 'err':
      return err(node.error)
    case 'ref': {
      // 空单元格（''）标记为 BLANK：比较/算术按 0，拼接按 ''（仅 ref 路径）
      const sid = node.ref.sheet !== undefined ? ctx.resolveSheet(node.ref.sheet) : ctx.sheet
      if (sid === null) return err('#REF!')
      const v = ctx.get(sid, node.ref.row, node.ref.col)
      return v === '' ? BLANK : v
    }
    case 'range': {
      // 隐式交集：单列区域按公式行、单行区域按公式列、二维区域需行列均在区域内；
      // 无交集 → #VALUE!（聚合函数参数不走这里，由 evalCall 特判）
      const sid = node.a.sheet !== undefined ? ctx.resolveSheet(node.a.sheet) : ctx.sheet
      if (sid === null) return err('#REF!')
      const r = normalizeRange({ sr: node.a.row, sc: node.a.col, er: node.b.row, ec: node.b.col })
      let hit: { row: number; col: number } | null = null
      if (r.sc === r.ec) {
        if (ctx.row >= r.sr && ctx.row <= r.er) hit = { row: ctx.row, col: r.sc }
      } else if (r.sr === r.er) {
        if (ctx.col >= r.sc && ctx.col <= r.ec) hit = { row: r.sr, col: ctx.col }
      } else if (ctx.row >= r.sr && ctx.row <= r.er && ctx.col >= r.sc && ctx.col <= r.ec) {
        hit = { row: ctx.row, col: ctx.col }
      }
      if (!hit) return err('#VALUE!')
      const v = ctx.get(sid, hit.row, hit.col)
      return v === '' ? BLANK : v
    }
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
      const sid = arg.a.sheet !== undefined ? ctx.resolveSheet(arg.a.sheet) : ctx.sheet
      if (sid === null) return err('#REF!')
      const r = normalizeRange({ sr: arg.a.row, sc: arg.a.col, er: arg.b.row, ec: arg.b.col })
      for (let row = r.sr; row <= r.er; row++) {
        for (let col = r.sc; col <= r.ec; col++) {
          const e = push(ctx.get(sid, row, col))
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

// 逐值迭代（函数参数可为值或区域）；cb 返回 FormulaError 即中断传播
function eachValue(args: AST[], ctx: EvalCtx, cb: (v: V) => FormulaError | null | void): FormulaError | null {
  for (const arg of args) {
    if (arg.type === 'range') {
      const sid = arg.a.sheet !== undefined ? ctx.resolveSheet(arg.a.sheet) : ctx.sheet
      if (sid === null) return err('#REF!')
      const r = normalizeRange({ sr: arg.a.row, sc: arg.a.col, er: arg.b.row, ec: arg.b.col })
      for (let row = r.sr; row <= r.er; row++) {
        for (let col = r.sc; col <= r.ec; col++) {
          const cv = ctx.get(sid, row, col)
          // 区域中的空单元格映射为 BLANK，与 ref 求值路径同语义（AND/OR 据此跳过）
          const e = cb(cv === '' ? BLANK : cv)
          if (e) return e
        }
      }
    } else {
      const e = cb(evalNode(arg, ctx))
      if (e) return e
    }
  }
  return null
}

// 单参数文本函数辅助
function oneTextArg(nameArgs: AST[], ctx: EvalCtx): string | FormulaError {
  if (nameArgs.length !== 1) return err('#VALUE!')
  const v = evalNode(nameArgs[0], ctx)
  if (isError(v)) return v
  return toText(v)
}

// 条件聚合的区域参数：ref 视为 1×1；其余 → #VALUE!
function critRangeOf(arg: AST, ctx: EvalCtx): { sid: SheetId; sr: number; sc: number; er: number; ec: number } | FormulaError {
  if (arg.type === 'ref') {
    const sid = arg.ref.sheet !== undefined ? ctx.resolveSheet(arg.ref.sheet) : ctx.sheet
    if (sid === null) return err('#REF!')
    return { sid, sr: arg.ref.row, sc: arg.ref.col, er: arg.ref.row, ec: arg.ref.col }
  }
  if (arg.type === 'range') {
    const sid = arg.a.sheet !== undefined ? ctx.resolveSheet(arg.a.sheet) : ctx.sheet
    if (sid === null) return err('#REF!')
    const r = normalizeRange({ sr: arg.a.row, sc: arg.a.col, er: arg.b.row, ec: arg.b.col })
    return { sid, ...r }
  }
  return err('#VALUE!')
}

// 条件聚合公共实现：mode='sum'|'count'|'average'
function condAggregate(mode: 'sum' | 'count' | 'average', args: AST[], ctx: EvalCtx): V {
  if (args.length < 2 || args.length > 3) return err('#VALUE!')
  if (mode === 'count' && args.length !== 2) return err('#VALUE!')
  const target = critRangeOf(args[0], ctx)
  if (isError(target)) return target
  const crit = evalNode(args[1], ctx)
  if (isError(crit)) return crit
  const critVal: number | string | boolean = isBlank(crit) ? '' : crit
  let sumZone = target
  if (args.length === 3) {
    const sz = critRangeOf(args[2], ctx)
    if (isError(sz)) return sz
    sumZone = sz
  }
  let sum = 0
  let count = 0
  for (let row = target.sr; row <= target.er; row++) {
    for (let col = target.sc; col <= target.ec; col++) {
      if (!matchCriteria(critVal, ctx.get(target.sid, row, col))) continue
      count++
      if (mode === 'count') continue
      // 求和域按条件域尺寸锚定其左上角（Excel 语义）；非数值按 0
      const v = ctx.get(sumZone.sid, sumZone.sr + (row - target.sr), sumZone.sc + (col - target.sc))
      if (isError(v)) return v
      if (typeof v === 'number') sum += v
    }
  }
  if (mode === 'count') return count
  if (mode === 'sum') return sum
  return count === 0 ? err('#DIV/0!') : sum / count
}

function leftRight(side: 'left' | 'right', args: AST[], ctx: EvalCtx): V {
  if (args.length < 1 || args.length > 2) return err('#VALUE!')
  const t = oneTextArg([args[0]], ctx)
  if (isError(t)) return t
  let n = 1
  if (args.length === 2) {
    const nn = numArg(args[1], ctx)
    if (isError(nn)) return nn
    n = Math.floor(nn)
    if (n < 0) return err('#VALUE!')
  }
  return side === 'left' ? t.slice(0, n) : n === 0 ? '' : t.slice(-n)
}

function numArg(arg: AST, ctx: EvalCtx): number | FormulaError {
  const v = evalNode(arg, ctx)
  if (isError(v)) return v
  return toNum(v)
}

function datePart(part: 'y' | 'm' | 'd', args: AST[], ctx: EvalCtx): V {
  if (args.length !== 1) return err('#VALUE!')
  const n = numArg(args[0], ctx)
  if (isError(n)) return n
  if (n < 0) return err('#VALUE!') // Excel 为 #NUM!，本引擎无该错误码
  return serialToDate(n)[part]
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
  // ---- 逻辑 ----
  AND: (args, ctx) => {
    let seen = false
    let result = true
    const e = eachValue(args, ctx, (v) => {
      if (isBlank(v)) return // 区域中的空格跳过
      const b = toBool(v)
      if (isError(b)) return b
      seen = true
      if (!b) result = false
    })
    if (e) return e
    return seen ? result : err('#VALUE!')
  },
  OR: (args, ctx) => {
    let seen = false
    let result = false
    const e = eachValue(args, ctx, (v) => {
      if (isBlank(v)) return
      const b = toBool(v)
      if (isError(b)) return b
      seen = true
      if (b) result = true
    })
    if (e) return e
    return seen ? result : err('#VALUE!')
  },
  NOT: (args, ctx) => {
    if (args.length !== 1) return err('#VALUE!')
    const v = evalNode(args[0], ctx)
    if (isError(v)) return v
    const b = toBool(v)
    return isError(b) ? b : !b
  },
  IFERROR: (args, ctx) => {
    if (args.length !== 2) return err('#VALUE!')
    const v = evalNode(args[0], ctx)
    return isError(v) ? evalNode(args[1], ctx) : v
  },
  // ---- 文本 ----
  LEN: (args, ctx) => {
    const t = oneTextArg(args, ctx)
    return isError(t) ? t : t.length
  },
  LEFT: (args, ctx) => leftRight('left', args, ctx),
  RIGHT: (args, ctx) => leftRight('right', args, ctx),
  MID: (args, ctx) => {
    if (args.length !== 3) return err('#VALUE!')
    const t = oneTextArg([args[0]], ctx)
    if (isError(t)) return t
    const s = numArg(args[1], ctx)
    if (isError(s)) return s
    const n = numArg(args[2], ctx)
    if (isError(n)) return n
    if (s < 1 || n < 0) return err('#VALUE!')
    return t.slice(Math.floor(s) - 1, Math.floor(s) - 1 + Math.floor(n))
  },
  UPPER: (args, ctx) => {
    const t = oneTextArg(args, ctx)
    return isError(t) ? t : t.toUpperCase()
  },
  LOWER: (args, ctx) => {
    const t = oneTextArg(args, ctx)
    return isError(t) ? t : t.toLowerCase()
  },
  TRIM: (args, ctx) => {
    const t = oneTextArg(args, ctx)
    return isError(t) ? t : t.trim().replace(/ +/g, ' ') // Excel TRIM 只处理 ASCII 空格
  },
  CONCAT: (args, ctx) => {
    let out = ''
    const e = eachValue(args, ctx, (v) => {
      if (isError(v)) return v
      out += toText(v)
    })
    return e ?? out
  },
  // ---- 条件聚合 ----
  SUMIF: (args, ctx) => condAggregate('sum', args, ctx),
  COUNTIF: (args, ctx) => condAggregate('count', args, ctx),
  AVERAGEIF: (args, ctx) => condAggregate('average', args, ctx),
  // ---- 日期 ----
  TODAY: (args) => (args.length !== 0 ? err('#VALUE!') : todaySerial()),
  NOW: (args) => (args.length !== 0 ? err('#VALUE!') : nowSerial()),
  YEAR: (args, ctx) => datePart('y', args, ctx),
  MONTH: (args, ctx) => datePart('m', args, ctx),
  DAY: (args, ctx) => datePart('d', args, ctx),
  DATE: (args, ctx) => {
    if (args.length !== 3) return err('#VALUE!')
    const y = numArg(args[0], ctx)
    if (isError(y)) return y
    const m = numArg(args[1], ctx)
    if (isError(m)) return m
    const d = numArg(args[2], ctx)
    if (isError(d)) return d
    return dateSerialLenient(Math.trunc(y), Math.trunc(m), Math.trunc(d))
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
