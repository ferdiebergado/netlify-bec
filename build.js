import { build } from 'esbuild';
import postcss from 'postcss';
import autoprefixer from 'autoprefixer';
import postcssPresetEnv from 'postcss-preset-env';
import { sassPlugin } from 'esbuild-sass-plugin';

build({
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
}).catch(() => process.exit(1));
