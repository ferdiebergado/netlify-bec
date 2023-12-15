module.exports = {
  plugins: ['import', 'security'],
  extends: ['airbnb-base', 'airbnb-typescript/base', 'prettier'],
  parserOptions: {
    project: true,
  },
  ignorePatterns: ['**/*.js', 'node_modules', 'out', '.netlify', 'netlify'],
  root: true,
}
