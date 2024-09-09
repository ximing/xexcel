import { describe, expect, it } from 'vitest'
import { SheetData } from '../src/core/model'
import { csvToGrid, sheetToCSV } from '../src/core/io/csv'

describe('csvToGrid', () => {
  it('普通网格', () => {
    expect(csvToGrid('a,b,c\n1,2,3')).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
    ])
  })

  it('引号包裹内嵌逗号/换行，双引号转义', () => {
    expect(csvToGrid('"a,b","x\ny","他""好"')).toEqual([['a,b', 'x\ny', '他"好']])
  })

  it('CRLF 与 LF 混用', () => {
    expect(csvToGrid('a,b\r\nc,d\ne,f')).toEqual([['a', 'b'], ['c', 'd'], ['e', 'f']])
  })

  it('尾部换行不产生幽灵行；中间空行保留', () => {
    expect(csvToGrid('a\n\nb\n')).toEqual([['a'], [''], ['b']])
  })

  it('去开头 BOM', () => {
    expect(csvToGrid('﻿a,b')).toEqual([['a', 'b']])
  })

  it('引号不配对宽松解析（余下入字段）', () => {
    expect(csvToGrid('a,"b\nc')).toEqual([['a', 'b\nc']])
  })

  it('字段中间的引号按字面字符（宽松）', () => {
    expect(csvToGrid('ab"c,d')).toEqual([['ab"c', 'd']])
  })

  it('ragged 行保留原长', () => {
    expect(csvToGrid('a,b,c\n1')).toEqual([['a', 'b', 'c'], ['1']])
  })

  it('空文本返回 []', () => {
    expect(csvToGrid('')).toEqual([])
  })
})

describe('sheetToCSV', () => {
  const mkSheet = (rows: string[][]): SheetData => {
    let s = SheetData.create({ rowCount: Math.max(rows.length, 5), colCount: 5 })
    rows.forEach((r, ri) => r.forEach((v, ci) => {
      if (v !== '') s = s.setCell(ri, ci, { raw: v })
    }))
    return s
  }

  it('特殊字符加引号转义；默认带 BOM；CRLF 行尾', () => {
    const csv = sheetToCSV(mkSheet([['a,b', 'x"y', '行1\n行2']]))
    expect(csv).toBe('﻿"a,b","x""y","行1\n行2"\r\n')
  })

  it('公式导出 raw 原文', () => {
    const csv = sheetToCSV(mkSheet([['=A1+B1', '3']]), { bom: false })
    expect(csv).toBe('=A1+B1,3\r\n')
  })

  it('空表返回空串', () => {
    expect(sheetToCSV(SheetData.create({ rowCount: 10, colCount: 5 }))).toBe('')
  })

  it('只导到 usedRange，尾部空行空列不输出', () => {
    const csv = sheetToCSV(mkSheet([['a', '', ''], ['b', '', '']]), { bom: false })
    expect(csv).toBe('a\r\nb\r\n')
  })

  it('往返：grid→sheet→csv→grid 恒等', () => {
    const grid = [['姓名', '数量'], ['苹果,梨', '3'], ['=B2*2', '']]
    const back = csvToGrid(sheetToCSV(mkSheet(grid), { bom: false }))
    expect(back).toEqual([
      ['姓名', '数量'],
      ['苹果,梨', '3'],
      ['=B2*2', ''],
    ])
  })
})
