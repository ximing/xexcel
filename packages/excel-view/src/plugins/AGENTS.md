# src/plugins — 内建交互插件

## 这个目录负责什么

selection、keymap、clipboard、fillhandle、dragmove，以及 uistate 里那些临时 UI 字段。`builtinPlugins()` 负责组装。history 不在这个数组里，由调用方插到最前。

## 放置约束

- 放：PluginSpec（props、field、appendTransaction、plugin view）。
- 不放：绘制细节、数据层逻辑。插件之间不要互相 import 内部实现，读状态用 PluginKey。

## 开发偏好

- 改文档必须 `view.dispatch(tr)`。props 返回 true 表示拦截。
- `builtinPlugins()` 顺序：selection → fillhandle → clipboard → keymap → dragmove → uistate 系列。新插件插入时注意拦截优先级。
- 跨插件通信用 `tr.setMeta` + PluginKey。
- 改选区的操作带 `scrollIntoView()`。
- 填充预览、菜单开闭这类临时状态放插件 field，不进 doc，不进 history。
