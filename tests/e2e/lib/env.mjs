import { spawn } from 'node:child_process'
import { bringToFront, cmd, evaluateJS } from './bridge.mjs'
import { HELPER_SOURCE } from './helper.js'

const APP = process.env.E2E_APP || 'http://localhost:5180'
let devProc = null

export async function ensureApp() {
  // daemon 可达性
  try { await cmd('list_tabs') } catch { throw new Error('kimi-webbridge daemon 不可达（~/.kimi-webbridge/bin/kimi-webbridge start）') }
  // dev server：已起则复用，否则 spawn（strictPort 防撞端口）
  try { await fetch(APP, { signal: AbortSignal.timeout(1500) }) } catch {
    devProc = spawn('npm', ['run', 'dev', '--', '--port', '5180', '--strictPort'], {
      cwd: new URL('../../..', import.meta.url).pathname, stdio: 'ignore',
    })
    for (let i = 0; i < 60; i++) {
      try { await fetch(APP); break } catch { await new Promise(r => setTimeout(r, 500)) }
      if (i === 59) throw new Error('dev server 启动超时')
    }
  }
  await cmd('navigate', { url: APP, newTab: true, group_title: 'xexcel e2e 回归' })
  await new Promise(r => setTimeout(r, 2500))
  await bringToFront()
}

// 清档刷新 + 注入 helper/stub（每个 suite 开头一次；suite 内刷新页面后须重注）
// 注意：App 注册了 beforeunload → workbookStorage.flush()，removeItem 后 reload 会把
// 防抖未落盘的旧现场重新写回存档（M4c 实测场景 3 规则泄漏进场景 4）。beforeunload 在
// window 自身派发时 at-target 按注册序触发，后注册挡不住先注册的 App 监听；改为临时
// 置空 Storage.prototype.setItem 吞掉 unload 写盘（reload 后新 realm 自动还原）。
export async function freshPage() {
  await bringToFront()
  await evaluateJS(`
    Storage.prototype.setItem = function () {}
    localStorage.removeItem('xexcel.workbook')
    location.reload()
    'ok'`)
  await new Promise(r => setTimeout(r, 2500))
  await bringToFront()
  return evaluateJS(`return (()=>{ ${HELPER_SOURCE} })()`)
}

// 喂文件主路径（M4b 实测）：base64 → DataTransfer → 手动 onchange
export async function feedFile(b64, name) {
  return evaluateJS(`
    const bin = atob(${JSON.stringify(b64)})
    const u8 = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i)
    const f = new File([u8], ${JSON.stringify(name)}, { type: 'application/octet-stream' })
    const dt = new DataTransfer(); dt.items.add(f)
    if (!window.__lastInput) throw new Error('未捕获 pickFile input（先点菜单）')
    window.__lastInput.files = dt.files
    window.__lastInput.onchange?.()
    return 'fed'`)
}

export function cleanup() { devProc?.kill() }
