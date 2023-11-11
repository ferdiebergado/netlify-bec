module.exports = {
  plugins: ["import"],
  extends: ["airbnb-typescript/base", "prettier"],
  parserOptions: {
    project: "./tsconfig.eslint.json",
  },
};
