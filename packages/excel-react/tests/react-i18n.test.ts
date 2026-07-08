import { describe, expect, it } from 'vitest'
import { t } from '../src/i18n'

describe('t', () => {
  it('默认中文；en 切英文', () => {
    expect(t('zh', 'status.ready')).toBe('就绪')
    expect(t('en', 'status.ready')).toBe('Ready')
    expect(t('en', 'file.clearConfirm')).toBe('Clear')
  })
  it('替换占位符', () => {
    expect(t('en', 'file.sizeWarn', { size: 12 })).toContain('12')
  })
})
