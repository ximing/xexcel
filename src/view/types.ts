// 视图层公共类型。本文件不得 import konva（需在 node 环境单测）。
import { PluginKey } from '../core/plugin'
import type { HitResult as CoreHitResult } from '../core/plugin'

export interface Rect {
  x: number
  y: number
  w: number
  h: number
}

// region 全集在 core/plugin.ts 声明（含 colborder/rowborder 调宽边界），此处仅取别名
export type HitRegion = CoreHitResult['region']

export interface HitResult {
  region: HitRegion
  row: number
  col: number
}

// 填充手柄预览区域的 plugin state field（fillhandle 插件以其为 key；layers 读取并画 1px 虚线框）
export const fillPreviewKey = new PluginKey('fillPreview')

// F3 拖动目标预览：dragmove 插件以本 key 写 CellRange|null；layers 读非空画虚线框
export const dragPreviewKey = new PluginKey('dragPreview')

// 行列调宽拖拽参考线：selection 插件写入，layers 读取画竖/横虚线（内容坐标像素）
export interface ResizeGuide {
  axis: 'row' | 'col'
  pos: number // axis='col' → 竖线 x；axis='row' → 横线 y
}
export const resizeGuideKey = new PluginKey('resizeGuide')

// 筛选下拉面板开启态：点击箭头时的列号与 client 坐标（React 浮层定位）；null=关闭
export interface FilterDropdownOpen {
  col: number
  x: number
  y: number
}
export const filterDropdownKey = new PluginKey('filterDropdown')

// 查找栏开关（true=显示）
export const findBarKey = new PluginKey('findBar')

// 格式刷状态：源格 style 快照 + 是否锁定连刷；null=未激活
export interface FormatPainterState {
  style: import('../core/model').CellStyle
  locked: boolean
}
export const formatPainterKey = new PluginKey('formatPainter')

// 右键菜单开启态（React 浮层定位）；null=关闭
export interface ContextMenuOpen {
  kind: 'cell' | 'rowheader' | 'colheader' | 'tab'
  x: number // client 坐标
  y: number
  row: number
  col: number
  sheet?: import('../core/model').SheetId // kind='tab'
}
export const contextMenuKey = new PluginKey('contextMenu')

// 标签栏重命名请求（ContextMenu → SheetTabBar 进入改名输入态）
export const tabRenameKey = new PluginKey('tabRename')

// 每表缩放档位（{ sheetId: zoom }，zoom=1 为 100%）；非文档态，不入 undo、不持久化
export const zoomKey = new PluginKey('zoom')

// F5 引用高亮：编辑公式时被引区域在画布画彩色虚线框；值 = 当前编辑文本（以 = 开头）| null。
// 非文档态（视图态），setMeta(...).setMeta('addToHistory', false) 不入 undo。
export const refHighlightKey = new PluginKey('refHighlight')

// 引用高亮调色板：按出现顺序循环取色（layers 按 ranges 下标 i % length）
export const REF_PALETTE = ['#1a73e8', '#ea4335', '#fbbc04', '#34a853', '#9334e6', '#ff6d00', '#00897b', '#7cb342']
