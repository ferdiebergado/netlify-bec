module.exports = {
  env: {
    es2021: true,
    node: true,
  },
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
    'plugin:@typescript-eslint/recommended-type-checked',
    'plugin:jest/recommended',
    'prettier',
  ],
  plugins: ['@typescript-eslint', 'import', 'security', 'n', 'promise', 'jest'],
  parser: '@typescript-eslint/parser',
  parserOptions: {
    project: ['./tsconfig.eslint.json'],
    tsconfigRootDir: __dirname,
    ecmaVersion: 'latest',
    sourceType: 'module',
    // typescript-eslint specific options
    warnOnUnsupportedTypeScriptVersion: true,
    EXPERIMENTAL_useProjectService: true,
  },
  ignorePatterns: ['**/*.js', 'node_modules', 'out', '.netlify', 'netlify'],
  overrides: [
    {
      files: ['.eslintrc.{js,cjs}'],
      parserOptions: {
        sourceType: 'script',
      },
    },
    {
      files: ['**/*.test.ts', '**/*.test.tsx'],
      env: {
        jest: true,
      },
    },
    {
      files: ['src/client/**/*.ts'],
      env: {
        browser: true,
      },
    },
  ],
  root: true,
};
