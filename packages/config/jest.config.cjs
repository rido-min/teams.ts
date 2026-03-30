const fs = require('fs');
const path = require('path');

// Build moduleNameMapper to resolve workspace packages to source TS files.
// This is needed because the workspace packages produce ESM-only output,
// and Jest (running in CJS mode) cannot require ESM modules directly.
// By mapping to source, ts-jest compiles everything in-process.
function buildModuleNameMapper () {
  const root = path.resolve(__dirname, '../..');
  const map = {};
  for (const parent of ['packages', 'external']) {
    const parentDir = path.join(root, parent);
    if (!fs.existsSync(parentDir)) continue;
    for (const name of fs.readdirSync(parentDir)) {
      const pkgPath = path.join(parentDir, name, 'package.json');
      if (!fs.existsSync(pkgPath)) continue;
      const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
      if (!pkg.name) continue;
      const srcDir = path.join(parentDir, name, 'src');
      if (!fs.existsSync(srcDir)) continue;
      const escaped = pkg.name.replace(/[.]/g, '\\.');
      // Map root import
      map[`^${escaped}$`] = `${srcDir}/index`;
      // Map subpath exports
      if (pkg.exports && typeof pkg.exports === 'object') {
        for (const exp of Object.keys(pkg.exports)) {
          if (exp === '.') continue;
          if (exp.includes('*')) {
            // Wildcard export: ./foo/* → src/foo/$1
            const prefix = exp.replace(/\/\*$/, '').replace(/^\.\//, '');
            map[`^${escaped}/${prefix}/(.+)$`] = `${srcDir}/${prefix}/$1`;
          } else {
            const sub = exp.replace(/^\.\//, '');
            map[`^${escaped}/${sub}$`] = `${srcDir}/${sub}`;
          }
        }
      }
    }
  }
  return map;
}

/** @type {import('jest').Config} */
module.exports = {
  roots: ['<rootDir>/src'],
  transform: {
    '^.+\\.ts?$': 'ts-jest',
  },
  moduleNameMapper: buildModuleNameMapper(),
  collectCoverage: true,
  preset: 'ts-jest',
  coverageDirectory: 'coverage',
  passWithNoTests: true,
  clearMocks: true,
  collectCoverageFrom: [
    '**/*.ts',
    '!**/*.d.ts',
    '!**/dist/**',
    '!**/coverage/**',
    '!**/node_modules/**',
    '!**/testing/**',
    '!**/index.ts',
  ],
  coverageThreshold: {
    // global: {
    //   branches: 90,
    //   functions: 90,
    //   lines: 90,
    //   statements: 90,
    // },
  },
  testEnvironment: 'node',
  silent: true,
  verbose: true,
};
