import { describe, it, expect } from 'vitest'
import { normalizeInput, normalizedCell } from '../src/formula/input'

// serial 基准：2024-01-01 = 45292；2024-07-31 = 45504（2024 为闰年）
const NOW = new Date(2024, 6, 15) // 2024-07-15 本地时间

describe('normalizeInput', () => {
  it('yyyy/m/d、yyyy-m-d、yyyy.m.d', () => {
    expect(normalizeInput('2024/1/1')).toEqual({ raw: '45292', numFmt: 'yyyy/m/d' })
    expect(normalizeInput('2024-01-01')).toEqual({ raw: '45292', numFmt: 'yyyy/m/d' })
    expect(normalizeInput('2024.1.1')).toEqual({ raw: '45292', numFmt: 'yyyy/m/d' })
    expect(normalizeInput('2024/12/31')).toEqual({ raw: '45657', numFmt: 'yyyy/m/d' }) // 45292+365（闰年）
  })
  it('m/d 与 m-d（无年 → 当年）', () => {
    expect(normalizeInput('7/31', NOW)).toEqual({ raw: '45504', numFmt: 'm/d' })
    expect(normalizeInput('1-1', NOW)).toEqual({ raw: '45292', numFmt: 'm/d' })
  })
  it('h:mm 与 h:mm:ss 与 AM/PM', () => {
    // 13:00 = 13/24 ≈ 0.5416666667
    expect(normalizeInput('13:00')).toEqual({ raw: String(13 / 24), numFmt: 'h:mm' })
    expect(normalizeInput('1:30 PM')).toEqual({ raw: String(13.5 / 24), numFmt: 'h:mm' })
    expect(normalizeInput('0:00')).toEqual({ raw: '0', numFmt: 'h:mm' })
  })
  it('非日期原样返回', () => {
    expect(normalizeInput('hello')).toEqual({ raw: 'hello' })
    expect(normalizeInput('=A1+1')).toEqual({ raw: '=A1+1' })
    expect(normalizeInput('123')).toEqual({ raw: '123' })
    expect(normalizeInput('')).toEqual({ raw: '' })
    expect(normalizeInput('50%')).toEqual({ raw: '50%' })
  })
  it('非法日期不识别', () => {
    expect(normalizeInput('2024/2/30')).toEqual({ raw: '2024/2/30' })
    expect(normalizeInput('2024/13/1')).toEqual({ raw: '2024/13/1' })
    expect(normalizeInput('25:00')).toEqual({ raw: '25:00' })
  })
})

describe('normalizedCell', () => {
  it('日期且原格无 numFmt → 合并 numFmt', () => {
    expect(normalizedCell('2024/1/1', { raw: '', style: { bold: true } })).toEqual({
      raw: '45292',
      style: { bold: true, numFmt: 'yyyy/m/d' },
    })
  })
  it('日期但原格已有 numFmt → 只转 serial，保留原格式', () => {
    expect(normalizedCell('2024/1/1', { raw: 'x', style: { numFmt: '0%' } })).toEqual({
      raw: '45292',
      style: { numFmt: '0%' },
    })
  })
  it('非日期 → 保留原样式（修复编辑丢样式）', () => {
    expect(normalizedCell('abc', { raw: '1', style: { bold: true } })).toEqual({
      raw: 'abc',
      style: { bold: true },
    })
    expect(normalizedCell('abc', undefined)).toEqual({ raw: 'abc' })
  })
})
