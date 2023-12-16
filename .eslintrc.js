module.exports = {
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
    'plugin:@typescript-eslint/recommended-type-checked',
    'prettier',
  ],
  plugins: ['@typescript-eslint', 'import', 'security', 'n', 'promise'],
  parser: '@typescript-eslint/parser',
  parserOptions: {
    project: true,
    tsconfigRootDir: __dirname,
  },
  root: true,
  ignorePatterns: ['**/*.js', 'node_modules', 'out', '.netlify', 'netlify'],
}
