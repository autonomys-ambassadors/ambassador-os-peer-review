import { eslint } from 'eslint';

export default [
  {
    files: ['**/*.js', '**/*.ts'],
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'module',
    },
    plugins: {
      prettier: eslint.plugin.prettier,
    },
    rules: {
      quotes: ['error', 'single'],
      indent: ['error', 2],
      semi: ['error', 'always'],
      'no-trailing-spaces': 'error',
      'eol-last': ['error', 'always'],
      'prettier/prettier': ['error'], // Prettier rules as ESLint errors
    },
  },
];
