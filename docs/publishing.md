# Publishing `@xexcel/*`

Packages share one version. A GitHub Release on tag `vX.Y.Z` runs `.github/workflows/publish.yml` and publishes:

- `@xexcel/core`
- `@xexcel/view`
- `@xexcel/react`

## One-time setup (you must do this)

Two auth options. Prefer **Trusted Publishing**; keep `NPM_TOKEN` only as a fallback.

### A. npm Trusted Publishing (recommended)

No long-lived token. Each package on npmjs.com needs a trusted publisher:

1. Log into [npmjs.com](https://www.npmjs.com/) as a member of the **`xexcel` org** (the machine user `yeanzhi` currently gets 403 on that org — add that account, or publish from an account that already belongs).
2. Create the three empty packages if they do not exist yet, or wait for the first publish.
3. For each of `@xexcel/core`, `@xexcel/view`, `@xexcel/react`:
   - Package → **Settings** → **Trusted Publisher** → GitHub Actions
   - Repository: `ximing/xexcel`
   - Workflow: `publish.yml`
   - Environment: `npm` (must match the workflow `environment:` name)
4. In this GitHub repo: **Settings → Environments → New environment → `npm`**. Optionally require a reviewer.

### B. Classic `NPM_TOKEN` fallback

1. On npm, as an `xexcel` org member, create a granular access token:
   - Read and write to `@xexcel/core`, `@xexcel/view`, `@xexcel/react`
   - Bypass 2FA / automation token as required by current npm policy
2. GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**
   - Name: `NPM_TOKEN`
   - Value: the token

The workflow sends `NODE_AUTH_TOKEN` from that secret when present. If the secret is missing, the job relies on OIDC trusted publishing.

## Cut a release

1. Bump `"version"` in the three package `package.json` files (keep them equal).
2. Add a section to `CHANGELOG.md`.
3. Merge to `main`, then:

```bash
git tag v1.0.0
git push origin v1.0.0
gh release create v1.0.0 --title "v1.0.0" --notes-file CHANGELOG.md
```

The `release: published` event starts the publish job. Watch **Actions → Publish to npm**.

## Dry run locally

```bash
pnpm install --frozen-lockfile
pnpm -r typecheck && pnpm -r test && pnpm -r build
pnpm --filter @xexcel/core pack
pnpm --filter @xexcel/view pack
pnpm --filter @xexcel/react pack
```

Inspect the `.tgz` contents before the first real publish. Delete the tarballs afterwards (they are gitignored).
