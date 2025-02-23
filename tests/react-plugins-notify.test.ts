import { describe, expect, it } from 'vitest'
import { pluginNotice, registerPluginNotice } from '../src/plugins/notify'

describe('plugins/notify', () => {
  it('注入后转发', () => {
    const got: string[] = []
    registerPluginNotice((m) => got.push(m))
    pluginNotice('测试')
    expect(got).toEqual(['测试'])
  })
})
