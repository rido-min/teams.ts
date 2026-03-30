import neostandard, { plugins } from 'neostandard';

export default [
  ...neostandard({ ts: true, semi: true }),
  {
    ignores: [
      '**/dist/**',
      '**/coverage/**',
      '**/node_modules/**'
    ]
  },
  {
    files: ['src/**/*.ts'],
    plugins: {
      'typescript-eslint': plugins['typescript-eslint']
    },
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-unused-vars': 'off',
    }
  }
];
