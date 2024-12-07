import { context } from 'esbuild';
import postcss from 'postcss';
import autoprefixer from 'autoprefixer';
import postcssPresetEnv from 'postcss-preset-env';
import { sassPlugin } from 'esbuild-sass-plugin';

context({
  entryPoints: ['src/app.ts'],
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
