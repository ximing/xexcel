# Contributing to xexcel

[简体中文](./CONTRIBUTING.zh-CN.md)

Thanks for taking a look. xexcel is a small spreadsheet kernel. Useful PRs are usually: a missing formula, a clearer error, a docs fix, or a tightly scoped interaction bug.

## Setup

Requires Node 22+ and [pnpm](https://pnpm.io/) 9.

```bash
git clone https://github.com/ximing/xexcel.git
cd xexcel
pnpm install
pnpm dev          # demo at http://localhost:5173
```

Definition of done for every change:

```bash
pnpm -r typecheck
pnpm -r test
pnpm -r build
```

All three must be clean. Do not add a DOM test harness; unit tests stay in `packages/*/tests/` and cover core, formula, and pure view helpers.

## Layout

```
packages/excel-core    @xexcel/core     model / steps / tx / formula / io
packages/excel-view    @xexcel/view     Konva canvas + interaction plugins
packages/excel-react   @xexcel/react    React chrome
apps/demo                               live demo
```

Imports across packages use the package name (`@xexcel/core`), never a relative path into another package.

Hard rules (see `AGENTS.md` and the per-directory `CLAUDE.md` files):

1. Never mutate the workbook except through a Transaction / Step.
2. Row and column indexes are 0-based numbers. A1 notation lives in `addr` and the UI only.
3. Identifiers in English; comments in short Chinese, and only when the code is not obvious.
4. Canvas size constants come from `@xexcel/core` (`ROW_HEADER_WIDTH` and friends). No magic numbers.

## Pull requests

1. One concern per PR.
2. Add or extend a test next to the existing `tests/<area>-<topic>.test.ts` files.
3. Do not copy code, names, or file layout from other spreadsheet projects.
4. Fill in the PR template checklist.

Good first issues are tagged [`good first issue`](https://github.com/ximing/xexcel/labels/good%20first%20issue).

## Reporting bugs

Use the bug template. Include:

- what you did
- what you expected
- what happened
- browser / Node version
- a minimal workbook or formula if you can

Security issues: see [SECURITY.md](./SECURITY.md). Do not open a public issue for those.
