// kimi-webbridge daemon HTTP 封装：单 session 驱动浏览器。
const DAEMON = 'http://127.0.0.1:10086/command'
const SESSION = process.env.E2E_SESSION || 'xexcel-e2e'

export async function cmd(action, args = {}) {
  const res = await fetch(DAEMON, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action, args, session: SESSION }),
  })
  const j = await res.json()
  if (!j.ok) throw new Error(`${action} 失败: ${j.error?.message ?? JSON.stringify(j.error)}`)
  return j.data
}

// evaluate：自动包 async IIFE；返回值须可 JSON 序列化
export async function evaluateJS(code) {
  const data = await cmd('evaluate', { code: `(async()=>{${code}})()` })
  return data.value
}

export const bringToFront = () => cmd('cdp', { method: 'Page.bringToFront', params: {} })
