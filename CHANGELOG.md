# Changelog

`@xexcel/core`、`@xexcel/view`、`@xexcel/react` 共用版本号。

## 1.0.0 — 2026-07-18

第一个标成稳定的版本。包名 `@xexcel/*`，MIT。

内核还是同一套：不可变 Workbook、可逆 Step、事务、插件、公式引擎。画布是 Konva，外壳是 React。

这一版能用的东西：

- 编辑、数字格式、跨表公式
- 字体、颜色、边框、换行、合并、冻结、隐藏、行高列宽
- 多工作表（增、改名、拖动排序）
- 排序、筛选、查找替换、条件格式、数据验证
- CSV / xlsx 导入导出（值、公式、样式、合并、冻结、筛选、条件格式、验证）
- 浏览器 localStorage 自动保存，坏档能恢复
- `<ExcelEditor locale="en" />`（默认中文）
- 千分位、日期、会计数字格式

内部包名从 `@gmi/excel-*` 改过，那个 scope 没有发过 npm。demo 的存档 key 是 `xexcel.workbook.v2`，老标签页如果还挂着旧档，用「文件 → 清除浏览器存档」。
