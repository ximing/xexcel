# 存档、CSV、xlsx、数据验证

2024 下半年

## 持久化信封

`serializeWorkbook` / `parseWorkbook`。JSON 外面包一层 version。解析失败丢明确的 `PersistError`，不要 silent。

浏览器自动保存走 localStorage，防抖。坏档备份一份再拒绝，避免把好档盖掉。xlsx 导入期间先 pause 自动保存，成功后再写；失败要把当前文档落回去，编辑到一半的东西别丢。

## CSV

导出 RFC 4180，公式出原文（带 `=`），UTF-8 BOM，范围用 `usedRange`。导入宽松一点，进新 sheet，单元格走 `normalizedCell`，整次导入一个事务。

## xlsx

用 exceljs 4.4.0，不用 SheetJS。导入导出按整本替换。

能对上的：值、公式、字体/填充/对齐/边框、合并、冻结、隐藏行列、筛选、条件格式、数据验证。

对不上就降级，不要假装成功：

- 重复的条件格式规则不映射
- autoFilter 导入只保留 range
- 隐藏空行一类 exceljs 回读的怪癖按实测来
- 超出现在行列上限的，警告一次（不要每个格子打一遍）
- 非法 sheet 名导入时净化、去重

样式颜色要在主题色、RGB、ARGB 之间换算，单测里有往返。

浏览器里 exceljs 曾经用 script 标签 + DI 绕 pre-bundle 挂死的问题，后来确认是后台标签页的计时器被节流，前台静态 import 是正常的，那次 DI 回退了。

## 数据验证

模型：`ValidationRule` + `SetValidationsStep`。输入走 `validateInput`，拒绝就不提交，notice 提示。公式原文跳过校验。清空格子也不拦。

对话框按整表规则列表改，不是只编当前选区那一条。between 缺第二个值、范围倒过来、超出表界，都当非法。
