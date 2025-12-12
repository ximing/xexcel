# src/view — 命令式 Konva 视图

## 这个目录负责什么

`EditorView`、geometry、layers、editbox、types。

## 放置约束

- 放：渲染、事件、坐标、编辑覆盖层。
- 不放：React；数据模型和公式（那是 core / formula）。

## 开发偏好

- 不 import React。对外只暴露 `subscribe(listener)` 给外壳用。
- 不直接改 doc。事件 → `hitTest` → 插件 `someProp` → 默认行为 → `dispatch(tr)` → rAF 合帧重绘。
- 只实例化可见 range 里的节点，滚动复用。
- 冻结偏移要传到 hitTest 和绘制。
- 单元格编辑器用 DOM textarea，不要用 Konva 文本。
