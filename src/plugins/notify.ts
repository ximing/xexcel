// plugins 层不能 import app（方向倒置），通知经注入（仿 core/validation.ts registerValidationNotice）
let fn: (msg: string) => void = (m) => console.warn(m)

export function registerPluginNotice(f: (msg: string) => void): void {
  fn = f
}

export function pluginNotice(msg: string): void {
  fn(msg)
}
