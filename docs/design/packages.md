# 拆包

2025-07

单仓库 npm 包撑不住对外集成，拆成三个包。行为不改，只搬文件和改 import。

```
packages/excel-core    @xexcel/core     model + formula + io，零 DOM
packages/excel-view    @xexcel/view     Konva + 内建插件
packages/excel-react   @xexcel/react    ExcelEditor 和外壳
apps/demo                               演示和 Pages
```

依赖单向。包与包之间用包名 import，不要相对路径跨包。

每包 tsup 出 ESM 和 `.d.ts`。react / konva 当 peer 或 external。`@xexcel/react` 另外编一份 `styles.css`。

测试跟着代码走：`core-*` / `formula-*` / io 在 core，画布和插件在 view，外壳和 ui 在 react。e2e 放到 `apps/demo/e2e`。

对外名字后来从 `@gmi/excel-*` 改成 `@xexcel/*`，`@gmi` 没在 npm 上发过。
