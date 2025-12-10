# 界面

2025-02

功能面已经够用，界面还是各写各的颜色和 `alert`/`confirm`。这一轮只动表现，不改事务和存档。

## token

颜色、圆角、阴影、z-index 只从一份 `@theme` 出（`packages/excel-react/src/theme.css`）。画布在 `view/theme.ts` 镜像同一组色值，不要在 layers 里再写 `#1a73e8`。

色板还是原来那套蓝白，没有暗色模式。

Tailwind 的 utility 尽量关在 `src/ui/` 里。外面的业务组件用 variant / size / disabled。布局用的 flex/gap 可以松一点。

## 组件

`ui/` 里有按钮、图标钮、菜单、下拉、对话框、确认框、输入、分隔线、Tooltip。Tooltip 延迟大约 500ms，快捷键可以写在旁边，不用原生 `title`。

Lucide 做图标。图标钮必须带 `aria-label`，e2e 靠无障碍树点。

原来的 `alert` 进 notice；`confirm` 进 `ConfirmDialog` + `confirmStore`。插件里的 `window.alert` 也改成注入的 `pluginNotice`。

## 画布

绘制结构和命中逻辑不动，只把硬编码颜色换成 theme 常量。
