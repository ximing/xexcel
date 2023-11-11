// 视图层公共类型。本文件不得 import konva（需在 node 环境单测）。
import { PluginKey } from '../core/plugin'
import type { HitResult as CoreHitResult } from '../core/plugin'

export interface Rect {
  x: number
  y: number
  w: number
  h: number
}

// 在 core HitResult 基础上预留行列调宽的边界区域（Task 6 使用）
export type HitRegion = CoreHitResult['region'] | 'colborder' | 'rowborder'

export interface HitResult {
  region: HitRegion
  row: number
  col: number
}

// 填充手柄预览区域的 plugin state field（Task 6 fillhandle 写入；layers 读取并画 1px 虚线框）
export const fillPreviewKey = new PluginKey('fillPreview')
