const esbuild = require('esbuild');
const postcss = require('postcss');
const autoprefixer = require('autoprefixer');
const postcssPresetEnv = require('postcss-preset-env');
const { sassPlugin } = require('esbuild-sass-plugin');

esbuild
  .context({
    entryPoints: ['src/app.ts', 'src/sass/app.scss'],
    bundle: true,
    outdir: 'public',
    entryNames: '[name]',
    sourcemap: true,
    logLevel: 'info',
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
  .then(ctx => {
    ctx.watch();
    console.log('watching...');
  })
  .catch(() => process.exit(1));
