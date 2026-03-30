# Build System

This document describes how the `teams.ts` monorepo is structured, built, tested, and published.

## Overview

`teams.ts` is an **npm workspaces** monorepo orchestrated with [Turborepo](https://turbo.build/). It is written entirely in TypeScript and targets Node.js ≥ 20. The repository is divided into three workspace areas:

| Area | Path | Purpose |
|---|---|---|
| Core packages | `packages/*` | SDK libraries published to npm |
| External packages | `external/*` | Optional integrations / plugins |
| Examples | `examples/*` | Reference apps and end-to-end test fixtures |

---

## Package Manager

- **npm** (v10+, see `engines` in `package.json`)
- Dependencies are locked via `package-lock.json`.
- Workspace installs are performed with a single `npm install` at the repository root.

---

## Monorepo Orchestration – Turborepo

[Turborepo](https://turbo.build/) (`turbo` v2) manages task execution across packages. The root `turbo.json` defines the task graph:

```json
{
  "tasks": {
    "clean":    { "dependsOn": ["^clean"],    "cache": false },
    "build":    { "dependsOn": ["^build"],    "cache": false },
    "lint":     { "dependsOn": ["^lint"],     "cache": false },
    "lint:fix": { "dependsOn": ["^lint:fix"], "cache": false },
    "test":     {                             "cache": false },
    "dev":      { "cache": false, "persistent": true },
    "start":    { "cache": false, "persistent": true }
  }
}
```

The `"dependsOn": ["^build"]` convention ensures that a package's dependencies are built **before** the package itself, so the full dependency graph is always up to date.

### Root scripts

```bash
npm run build            # build every workspace
npm run build:packages   # build packages/* only
npm run build:essential  # build packages/* excluding heavy graph packages
npm run build:examples   # build examples/* only
npm run build:external   # build external/* only

npm run clean            # delete all dist directories
npm run dev              # start all example apps in watch mode
npm run lint             # lint all workspaces (100% concurrency)
npm run lint:fix         # auto-fix lint errors
npm run test             # run all tests (100% concurrency)
```

> Any script can also be scoped to a single workspace:
> ```bash
> npm run build --workspace=@microsoft/teams.apps
> ```

---

## Per-Package Build – tsup

Individual packages are built with [tsup](https://tsup.egoist.dev/) (v8), a zero-config TypeScript bundler backed by esbuild. A shared `tsup.config.js` lives in `packages/config` and is referenced by each package.

Output per package (`dist/`):

| File | Format | Description |
|---|---|---|
| `index.js` | CommonJS | Node.js `require()` entry point |
| `index.mjs` | ESM | `import` entry point |
| `index.d.ts` | TypeScript | CJS type declarations |
| `index.d.mts` | TypeScript | ESM type declarations |

Source maps and tree-shaking are enabled for all outputs.

Package `exports` map (dual CJS / ESM):

```json
"exports": {
  ".": {
    "types":   { "require": "./dist/index.d.ts", "import": "./dist/index.d.mts" },
    "require": "./dist/index.js",
    "import":  "./dist/index.mjs"
  }
}
```

---

## TypeScript Configuration

Shared TypeScript configs are provided by `packages/config`:

| File | Module system | `module` | `target` | Notes |
|---|---|---|---|---|
| `tsconfig.node.json` | CJS | `NodeNext` | `ESNext` | Used for `dist/*.js` output |
| `tsconfig.esm.json` | ESM | `ESNext` | `ESNext` | `moduleResolution: bundler` |

All configs enable `strict` mode and emit `.d.ts` declarations.

---

## Testing – Jest

Tests are run with [Jest](https://jestjs.io/) (v29) using the `ts-jest` preset so `.spec.ts` files are executed without a separate compilation step.

A shared `jest.config.js` in `packages/config` is used by every package:

- Test file pattern: `src/**/*.spec.ts`
- Environment: `node`
- Coverage collection: enabled by default
- Clear mocks and verbose output on by default

Run all tests:

```bash
npm run test
```

---

## Linting – ESLint

ESLint (v9) uses the new **flat config** format (`eslint.config.js`) from `packages/config`.

Key plugins / rule sets:

| Plugin | Purpose |
|---|---|
| `@eslint/js` | Core JS rules |
| `typescript-eslint` (v8) | TypeScript-specific rules (strict mode) |
| `eslint-plugin-import` | Import ordering and resolution |
| `@stylistic/eslint-plugin` | Formatting rules (replaces Prettier) |

---

## Versioning

Version numbers are managed by [Nerdbank.GitVersioning](https://github.com/dotnet/Nerdbank.GitVersioning) (`nbgv`). The version is stored in `version.json` and stamped into each package's `package.json` before publishing:

```bash
npm run version:stamp   # write computed version into package.json files
npm run version:get     # print the current NpmPackageVersion
```

---

## CI/CD – GitHub Actions

Two workflows live in `.github/workflows/`:

### `build-test-lint.yml`

Triggered on pull requests, pushes to `main`, and manual dispatch.

```
Checkout → Setup Node.js (24.x / 25.x matrix) → npm install → lint → build → test
```

Permissions are scoped to `read-all` for security.

### `template-sync.yml`

Ensures that changes to templates in `packages/cli` are kept in sync with the corresponding `examples/` directories. Skippable by adding `skip-test-verification` to the PR description.

---

## Packages

### Core (`packages/`)

| Package | Description |
|---|---|
| `@microsoft/teams.apps` | App development framework (main entry point) |
| `@microsoft/teams.ai` | AI integration utilities |
| `@microsoft/teams.api` | API bindings |
| `@microsoft/teams.botbuilder` | Bot Builder SDK adapter |
| `@microsoft/teams.cards` | Adaptive Cards helpers |
| `@microsoft/teams.cli` | Project scaffolding CLI |
| `@microsoft/teams.client` | Browser / client-side SDK |
| `@microsoft/teams.common` | Shared utilities consumed by other packages |
| `@microsoft/teams.config` | *(private)* Shared ESLint / Jest / tsup config |
| `@microsoft/teams.dev` | Development utilities and plugins |
| `@microsoft/teams.devtools` | Web frontend for the DevTools plugin |
| `@microsoft/teams.graph` | Microsoft Graph API client |
| `@microsoft/teams.graph-endpoints` | Generated Graph API endpoint definitions |
| `@microsoft/teams.graph-endpoints-beta` | Beta Graph API endpoint definitions |
| `@microsoft/teams.graph-tools` | CLI for generating Graph endpoints from OpenAPI specs |
| `@microsoft/teams.openai` | OpenAI integration |

### External (`external/`)

| Package | Description |
|---|---|
| `@microsoft/teams.mcp` | Model Context Protocol plugin |
| `@microsoft/teams.mcpclient` | MCP client implementation |
| `@microsoft/teams.a2a` | App-to-App integration |
