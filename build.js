const esbuild = require('esbuild');
const postcss = require('postcss');
const autoprefixer = require('autoprefixer');
const postcssPresetEnv = require('postcss-preset-env');
const { sassPlugin } = require('esbuild-sass-plugin');

esbuild
  .build({
    entryPoints: ['src/app.ts'],
    bundle: true,
    outdir: 'public',
    entryNames: '[name]',
    minify: true,
    plugins: [
      sassPlugin({
        async transform(source, _resolveDir) {
          const { css } = await postcss([
            autoprefixer,
            postcssPresetEnv({ stage: 0 }),
          ]).process(source, { from: undefined });
          return css;
        },
      }),
    ],
  })
  .catch(() => process.exit(1));
