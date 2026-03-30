import neostandard from 'neostandard';

export default [
  ...neostandard({ ts: true, semi: true }),
  {
    ignores: [
      '**/dist/**',
      '**/coverage/**',
      '**/node_modules/**',
      'packages/cli/templates/**',
      'packages/xxcards/src/**',
      'packages/graph*/src/**',
    ]
  },
  {
    files: ['**/*.ts'],
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-unused-vars': 'off',
    }
  }
];
