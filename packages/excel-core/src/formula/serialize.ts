// 公式序列化：AST → 公式体文本（不含前导 '='）。
// 与 parser 互逆：parse(serialize(ast)) 与 ast 结构等价（round-trip 由单测保证）。
import { colName } from '../core/addr'
import { AST, RefTarget } from './parser'

// 裸表名安全字符：字母/下划线开头，后接字母数字下划线点；否则单引号包裹（' 转义为 ''）
export function serializeSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}

function serializeRef(r: RefTarget, omitSheet: boolean): string {
  const sheet = !omitSheet && r.sheet !== undefined ? serializeSheetName(r.sheet) + '!' : ''
  return `${sheet}${r.colAbs ? '$' : ''}${colName(r.col)}${r.rowAbs ? '$' : ''}${r.row + 1}`
}

export function serialize(node: AST): string {
  switch (node.type) {
    case 'num':
      return String(node.value)
    case 'str':
      return `"${node.value.replace(/"/g, '""')}"`
    case 'bool':
      return node.value ? 'TRUE' : 'FALSE'
    case 'err':
      return node.error
    case 'ref':
      return serializeRef(node.ref, false)
    case 'range':
      // 同表区域表名只写一次（Sheet!A1:B2）
      return `${serializeRef(node.a, false)}:${serializeRef(node.b, node.a.sheet === node.b.sheet)}`
    case 'call':
      return `${node.name}(${node.args.map(serialize).join(',')})`
    case 'unary':
      return node.op + serialize(node.expr)
    case 'binary':
      return `${serialize(node.left)}${node.op}${serialize(node.right)}`
    case 'percent':
      return serialize(node.expr) + '%'
    case 'paren':
      return `(${serialize(node.expr)})`
  }
}
