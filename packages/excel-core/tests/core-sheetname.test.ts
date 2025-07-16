// tests/core-sheetname.test.ts
import { describe, expect, it } from 'vitest'
import {
  dedupeSheetName,
  INVALID_SHEET_NAME_CHARS,
  isValidSheetName,
  sanitizeSheetName,
  SHEET_NAME_MAX_LEN,
} from '../src/core/sheetname'

describe('INVALID_SHEET_NAME_CHARS', () => {
  it('逐个命中 Excel 非法字符 * ? : \\ / [ ]', () => {
    for (const ch of ['*', '?', ':', '\\', '/', '[', ']']) {
      expect(INVALID_SHEET_NAME_CHARS.test(ch)).toBe(true)
    }
    expect(INVALID_SHEET_NAME_CHARS.test('中文名')).toBe(false)
    expect(INVALID_SHEET_NAME_CHARS.test('Sheet 1')).toBe(false)
  })
})

describe('isValidSheetName', () => {
  it('普通名与中文名合法', () => {
    expect(isValidSheetName('Sheet1')).toBe(true)
    expect(isValidSheetName('数据表')).toBe(true)
    expect(isValidSheetName('Q1 报表 (final)')).toBe(true)
  })

  it('空串与纯空白非法', () => {
    expect(isValidSheetName('')).toBe(false)
    expect(isValidSheetName('   ')).toBe(false)
  })

  it('含任一非法字符即非法', () => {
    for (const ch of ['*', '?', ':', '\\', '/', '[', ']']) {
      expect(isValidSheetName(`a${ch}b`)).toBe(false)
    }
  })

  it('边界：31 合法，32 非法', () => {
    expect(isValidSheetName('a'.repeat(SHEET_NAME_MAX_LEN))).toBe(true)
    expect(isValidSheetName('a'.repeat(SHEET_NAME_MAX_LEN + 1))).toBe(false)
  })
})

describe('dedupeSheetName', () => {
  it('无冲突直接用；大小写不敏感冲突追加序号', () => {
    expect(dedupeSheetName(['Sheet1'], 'data')).toBe('data')
    expect(dedupeSheetName(['Data', 'data (2)'], 'DATA')).toBe('DATA (3)')
  })

  it('stem 超长时先截断再加序号，整体 ≤31', () => {
    const base = 'a'.repeat(SHEET_NAME_MAX_LEN)
    const cand = dedupeSheetName([base], base)
    expect(cand.length).toBeLessThanOrEqual(SHEET_NAME_MAX_LEN)
    expect(cand.endsWith(' (2)')).toBe(true)
  })
})

describe('sanitizeSheetName', () => {
  it('非法字符逐个替换为 _', () => {
    expect(sanitizeSheetName('a*b?c:d\\e/f[g]h', [])).toBe('a_b_c_d_e_f_g_h')
  })

  it('超长截断到 31', () => {
    const long = 'x'.repeat(40)
    expect(sanitizeSheetName(long, [])).toBe('x'.repeat(SHEET_NAME_MAX_LEN))
  })

  it('trim 后空回退 Sheet', () => {
    expect(sanitizeSheetName('', [])).toBe('Sheet')
    expect(sanitizeSheetName('   ', [])).toBe('Sheet')
    expect(sanitizeSheetName('///', [])).toBe('___')
  })

  it('中文名原样保留', () => {
    expect(sanitizeSheetName('财务报表', [])).toBe('财务报表')
  })

  it('与 existing 不区分大小写去重', () => {
    expect(sanitizeSheetName('Data', ['data'], )).toBe('Data (2)')
    expect(sanitizeSheetName('a/b', ['a_b'])).toBe('a_b (2)')
  })

  it('边界：31 原样；32 截断', () => {
    const n31 = '中'.repeat(SHEET_NAME_MAX_LEN)
    expect(sanitizeSheetName(n31, [])).toBe(n31)
    expect(sanitizeSheetName(n31 + '中', [])).toBe(n31)
  })
})
