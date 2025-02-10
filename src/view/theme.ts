// 画布取色：与 src/app/theme.css @theme 同值（Konva 读不了 CSS 变量）。
// 镜像一致性由 tests/app-theme-mirror.test.ts 守护。
export const THEME = {
  primary: '#1a73e8',
  primarySoft: '#e8f0fe',
  ink: '#202124',
  ink2: '#5f6368',
  ink3: '#9aa0a6',
  line: '#e0e0e0',
  lineStrong: '#d9dce1',
  surface: '#ffffff',
  surface2: '#f8f9fa',
  hover: '#f1f3f4',
  scrollbar: '#c1c7cd',
} as const
