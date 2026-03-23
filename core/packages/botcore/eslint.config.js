// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import eslint from '@eslint/js';
import tseslint from 'typescript-eslint';
import stylistic from '@stylistic/eslint-plugin';

export default tseslint.config(
  {
    ignores: ['**/dist/**', '**/node_modules/**'],
  },
  {
    extends: [eslint.configs.recommended, ...tseslint.configs.recommended],
    plugins: {
      '@stylistic': stylistic,
    },
    files: ['src/**/*.ts'],
    languageOptions: {
      parser: tseslint.parser,
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // typescript
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-empty-object-type': 'off',
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],

      // stylistic
      '@stylistic/linebreak-style': ['error', 'unix'],
      '@stylistic/quotes': ['error', 'single'],
      '@stylistic/semi': ['error', 'always'],
      '@stylistic/space-before-blocks': 'error',
      '@stylistic/block-spacing': 'error',
      '@stylistic/keyword-spacing': 'error',
    },
  },
  {
    files: ['src/**/*.spec.ts'],
    languageOptions: {
      parserOptions: {
        projectService: false,
      },
    },
  }
);
