import { describe, expect, it } from 'vitest'
import { Workbook } from '../src/core/model'
import {
  deserializeWorkbook,
  PersistError,
  PERSIST_VERSION,
  serializeWorkbook,
} from '../src/core/io/persist'

describe('core/io/persist', () => {
  it('serialize/deserialize 往返（含样式/合并/条件格式）', () => {
    let wb = Workbook.create({ rowCount: 10, colCount: 5 })
    let sheet = wb.activeSheet
    sheet = sheet.setCell(0, 0, { raw: '=B1*2', style: { bold: true, numFmt: '0.00' } })
    sheet = sheet.setCell(0, 1, { raw: '21' })
    sheet = sheet.setMerges([{ sr: 2, sc: 0, er: 3, ec: 1 }])
    sheet = sheet.setCondFormats([
      { id: 'cf1', range: { sr: 0, sc: 0, er: 9, ec: 4 }, type: 'value', op: 'gt', v1: '10',
        style: { bg: '#ff0000' } },
    ])
    wb = wb.setSheet(wb.active, sheet)

    const json = serializeWorkbook(wb, '2026-08-02T00:00:00.000Z')
    expect(JSON.parse(json).version).toBe(PERSIST_VERSION)
    expect(JSON.parse(json).savedAt).toBe('2026-08-02T00:00:00.000Z')

    const back = deserializeWorkbook(json)
    expect(back.toJSON()).toEqual(wb.toJSON())
  })

  it('version 不符抛 PersistError', () => {
    const bad = JSON.stringify({ version: 999, savedAt: 'x', workbook: {} })
    expect(() => deserializeWorkbook(bad)).toThrow(PersistError)
  })

  it('非法 JSON 抛 PersistError', () => {
    expect(() => deserializeWorkbook('{oops')).toThrow(PersistError)
  })

  it('workbook 载荷损坏抛 PersistError', () => {
    const bad = JSON.stringify({ version: PERSIST_VERSION, savedAt: 'x', workbook: { nope: 1 } })
    expect(() => deserializeWorkbook(bad)).toThrow(PersistError)
  })

  it('语义损坏信封（order 空/active 缺席）抛 PersistError', () => {
    const bad = JSON.stringify({
      version: PERSIST_VERSION,
      savedAt: '2026-08-02T00:00:00.000Z',
      workbook: { order: [], active: 's1', names: [], sheets: {} },
    })
    expect(() => deserializeWorkbook(bad)).toThrow(PersistError)
  })

  it('active 不在 order 中抛 PersistError', () => {
    const good = Workbook.create({ rowCount: 3, colCount: 3 })
    const j = good.toJSON() as { order: string[]; active: string; sheets: Record<string, unknown> }
    const bad = JSON.stringify({
      version: PERSIST_VERSION,
      savedAt: 'x',
      workbook: { ...j, active: 's999' },
    })
    expect(() => deserializeWorkbook(bad)).toThrow(PersistError)
  })
})
