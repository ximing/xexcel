// tests/core-io-xlsx-style.test.ts
import { describe, expect, it } from 'vitest'
import { CFStyle, CellStyle } from '../src/core/model'
import {
  argbToCss,
  cfStyleFromExcel,
  cfStyleToExcel,
  cssToArgb,
  styleFromExcel,
  styleToExcelAlignment,
  styleToExcelBorders,
  styleToExcelFill,
  styleToExcelFont,
} from '../src/core/io/xlsx-style'

describe('cssToArgb / argbToCss', () => {
  it('#rrggbb → FFRRGGBB（大写）', () => {
    expect(cssToArgb('#e8f0fe')).toBe('FFE8F0FE')
  })
  it('#rgb 展开', () => {
    expect(cssToArgb('#abc')).toBe('FFAABBCC')
  })
  it('非 hex 返回 null', () => {
    expect(cssToArgb('red')).toBeNull()
    expect(cssToArgb('rgb(1,2,3)')).toBeNull()
  })
  it('argb → #rrggbb 小写；非法 null', () => {
    expect(argbToCss('FFE8F0FE')).toBe('#e8f0fe')
    expect(argbToCss('FFF')).toBeNull()
  })
})

describe('styleToExcel*', () => {
  it('font 全字段', () => {
    expect(
      styleToExcelFont({
        bold: true, italic: true, underline: true, strikethrough: true,
        color: '#ff0000', fontFamily: 'Arial', fontSize: 14,
      }),
    ).toEqual({
      bold: true, italic: true, underline: true, strike: true,
      color: { argb: 'FFFF0000' }, name: 'Arial', size: 14,
    })
  })
  it('无 font 字段返回 undefined；非 hex 颜色省略 color', () => {
    expect(styleToExcelFont({ bg: '#fff' })).toBeUndefined()
    expect(styleToExcelFont({ color: 'red' })).toEqual({})
  })
  it('fill / alignment / borders', () => {
    expect(styleToExcelFill({ bg: '#00ff00' })).toEqual({
      type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' },
    })
    expect(styleToExcelFill({})).toBeUndefined()
    expect(styleToExcelAlignment({ align: 'center', vAlign: 'middle', wrap: true })).toEqual({
      horizontal: 'center', vertical: 'middle', wrapText: true,
    })
    expect(styleToExcelAlignment({})).toBeUndefined()
    expect(
      styleToExcelBorders({ border: { top: { style: 'thin' }, bottom: { style: 'double', color: '#0000ff' } } }),
    ).toEqual({
      top: { style: 'thin' },
      bottom: { style: 'double', color: { argb: 'FF0000FF' } },
    })
    expect(styleToExcelBorders({})).toBeUndefined()
  })
})

describe('styleFromExcel', () => {
  it('全字段往返', () => {
    const s: CellStyle = {
      bold: true, color: '#ff0000', bg: '#e8f0fe', align: 'right', vAlign: 'top',
      wrap: true, fontFamily: 'Arial', fontSize: 12, underline: true, strikethrough: true,
      border: { left: { style: 'dashed', color: '#00ff00' } },
    }
    const back = styleFromExcel({
      font: styleToExcelFont(s),
      fill: styleToExcelFill(s),
      alignment: styleToExcelAlignment(s),
      border: styleToExcelBorders(s),
    })
    expect(back).toEqual(s)
  })
  it('空输入 undefined；theme 色（无 argb）忽略', () => {
    expect(styleFromExcel({})).toBeUndefined()
    expect(styleFromExcel({ font: { color: { theme: 1 } as never } })).toBeUndefined()
  })
  it('未知 vertical/border style 忽略', () => {
    expect(styleFromExcel({ alignment: { vertical: 'justify' as never } })).toBeUndefined()
    expect(styleFromExcel({ border: { top: { style: 'slantDashDot' } } })).toBeUndefined()
  })
})

describe('CF differential style（fill 用 bgColor）', () => {
  it('互转往返', () => {
    const s: CFStyle = { color: '#ff0000', bg: '#ffff00', bold: true, underline: true }
    expect(cfStyleFromExcel(cfStyleToExcel(s)!)).toEqual(s)
    expect(cfStyleToExcel(s)!.fill).toEqual({
      type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' },
    })
  })
  it('空 style → undefined；空 dxf → {}', () => {
    expect(cfStyleToExcel({})).toBeUndefined()
    expect(cfStyleFromExcel({})).toEqual({})
  })
})
