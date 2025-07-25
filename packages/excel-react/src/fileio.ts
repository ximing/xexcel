// src/app/fileio.ts
// 浏览器文件读写薄壳（DOM 专属，无单测；逻辑均在 core/react 纯函数侧）。
export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.addEventListener('cancel', () => resolve(null)) // 取消选择（Chrome 113+）
    input.click()
  })
}

export function readFileText(file: File): Promise<string> {
  return file.text()
}

export function readFileArrayBuffer(file: File): Promise<ArrayBuffer> {
  return file.arrayBuffer()
}

export function downloadBlob(name: string, blob: Blob): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
