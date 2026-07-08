// tests/react-filemenu.test.ts
import { describe, expect, it } from 'vitest'
import { csvBaseName, fileMenuItems, isGridEmpty } from '../src/fileMenuCore'

describe('fileMenuItems', () => {
  it('五项：打开 CSV / 打开 xlsx / 导出 CSV / 导出 xlsx / 清除存档（danger）', () => {
    expect(fileMenuItems()).toEqual([
      { id: 'openCsv', label: '打开 CSV…', danger: false },
      { id: 'openXlsx', label: '打开 xlsx…', danger: false },
      { id: 'exportCsv', label: '导出 CSV', danger: false },
      { id: 'exportXlsx', label: '导出 xlsx', danger: false },
      { id: 'clearStorage', label: '清除浏览器存档', danger: true },
    ])
  })
  it('locale=en 切英文', () => {
    expect(fileMenuItems('en').map((i) => i.label)).toEqual([
      'Open CSV…',
      'Open xlsx…',
      'Export CSV',
      'Export xlsx',
      'Clear browser archive',
    ])
  })
})

describe('csvBaseName', () => {
  it('去扩展名；空名回退 CSV', () => {
    expect(csvBaseName('销售数据.csv')).toBe('销售数据')
    expect(csvBaseName('a.b.csv')).toBe('a.b')
    expect(csvBaseName('.csv')).toBe('CSV')
  })
})

describe('isGridEmpty', () => {
  it('空网格与全空行均为空', () => {
    expect(isGridEmpty([])).toBe(true)
    expect(isGridEmpty([[''], ['', '']])).toBe(true)
    expect(isGridEmpty([['', 'x']])).toBe(false)
  })
})
