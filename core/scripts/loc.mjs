// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
//
// Lines-of-code report for the core workspace.
// Usage: node scripts/loc.mjs

import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = fileURLToPath(new URL('..', import.meta.url))

function walk (dir, ext) {
  let files = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name === 'dist') continue
    const full = join(dir, entry.name)
    if (entry.isDirectory()) {
      files = files.concat(walk(full, ext))
    } else if (entry.isFile() && entry.name.endsWith(ext)) {
      files.push(full)
    }
  }
  return files
}

function countLines (files) {
  let total = 0
  for (const f of files) {
    const lines = readFileSync(f, 'utf8').split('\n').length
    total += lines
  }
  return total
}

function categorise (dir) {
  const all = walk(dir, '.ts')
  const tests = all.filter(f => f.endsWith('.spec.ts'))
  const src = all.filter(f => !f.endsWith('.spec.ts'))
  return { src, tests }
}

// --- collect by category ---
const packages = categorise(join(root, 'packages'))
const samplesFiles = walk(join(root, 'samples'), '.ts')

const groups = [
  { label: 'Production', files: packages.src },
  { label: 'Tests', files: packages.tests },
  { label: 'Samples', files: samplesFiles },
]

const total = groups.reduce((n, g) => n + countLines(g.files), 0)

// --- render ---
const col1 = 12
const col2 = 7
const col3 = 7
const divider = '-'.repeat(col1 + col2 + col3 + 6)

console.log('\nLines of code\n' + divider)
console.log(
  'Category'.padEnd(col1) + '  ' +
  'Files'.padStart(col2) + '  ' +
  'Lines'.padStart(col3)
)
console.log(divider)

for (const { label, files } of groups) {
  const lines = countLines(files)
  console.log(
    label.padEnd(col1) + '  ' +
    String(files.length).padStart(col2) + '  ' +
    String(lines).padStart(col3)
  )
}

console.log(divider)
console.log(
  'Total'.padEnd(col1) + '  ' +
  String(groups.reduce((n, g) => n + g.files.length, 0)).padStart(col2) + '  ' +
  String(total).padStart(col3)
)
console.log()
