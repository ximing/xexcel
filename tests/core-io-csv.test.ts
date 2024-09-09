import { describe, expect, it } from 'vitest'
import { csvToGrid } from '../src/core/io/csv'

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
