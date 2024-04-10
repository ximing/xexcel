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
