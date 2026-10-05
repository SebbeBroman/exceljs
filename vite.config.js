import {defineConfig} from 'vite-plus';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));

/**
 * Vite+ config for this ESM library.
 *
 * - Source is strict TypeScript under lib/ and excel.ts
 * - `pnpm build` emits JS to `dist/` for Node consumers and publish
 * - Tests: Vitest via `vp test` (loads TS sources directly)
 * - Lint/format: Oxlint + Oxfmt via `vp check` / `vp lint` / `vp fmt`
 */
export default defineConfig({
  staged: {
    // Keep pre-commit scoped to maintained surface
    '{lib,scripts,spec}/**/*.{ts,js,mjs},excel.ts,vite.config.js,package.json,index.d.ts,benchmark.js':
      'vp check --fix',
  },

  lint: {
    plugins: ['oxc', 'typescript', 'unicorn'],
    categories: {
      correctness: 'error',
    },
    env: {
      builtin: true,
      es2022: true,
      node: true,
    },
    ignorePatterns: [
      'node_modules/**',
      'build/**',
      'dist/**',
      'out/**',
      'coverage/**',
      'spec/out/**',
      'spec/manual/**',
      'spec/browser/**',
      'spec/dist/**',
      'spec/typescript/**',
    ],
    rules: {
      // Spreadsheet fixtures use holey arrays intentionally
      'no-sparse-arrays': 'off',
      // Hot read/write paths pre-allocate fixed-length buffers with
      // `new Array(n)` then fill by index — deliberately not `Array.from`,
      // which would allocate an extra array per row.
      'unicorn/no-new-array': 'off',
      // XML_SPECIAL must match C0 control chars to skip encoding them.
      'no-control-regex': 'off',
      // Closures capture `this` via a local alias (once-listeners, zip sinks).
      'typescript/no-this-alias': 'off',
      'no-console': ['error', {allow: ['warn', 'error']}],
      'prefer-const': 'warn',
      'no-var': 'error',
      'no-unused-vars': [
        'error',
        {
          vars: 'all',
          args: 'none',
          ignoreRestSiblings: true,
          caughtErrors: 'none',
        },
      ],
      'no-empty': ['error', {allowEmptyCatch: true}],
      'no-constant-condition': ['error', {checkLoops: false}],
      'vite-plus/prefer-vite-plus-imports': 'error',
    },
    overrides: [
      {
        files: ['spec/**/*.{ts,js,mjs,cjs}'],
        rules: {
          'no-console': 'off',
          // Chai-style property assertions (`expect(x).to.be.ok`) read as
          // unused expressions to the linter but are valid assertions.
          'no-unused-expressions': 'off',
        },
        globals: {
          verquire: 'readonly',
          context: 'readonly',
          before: 'readonly',
          after: 'readonly',
        },
        env: {
          vitest: true,
        },
      },
      {
        files: ['benchmark.js', 'scripts/**/*.{js,mjs,cjs}', '*.config.js'],
        rules: {
          'no-console': 'off',
          'no-unused-vars': 'off',
          // Bench harness builds fixed-size row buffers up front.
          'unicorn/no-new-array': 'off',
        },
      },
    ],
    // Hand-written public index.d.ts; full check via `pnpm typecheck` (tsc --strict)
    options: {
      typeAware: false,
      typeCheck: false,
    },
    jsPlugins: [
      {
        name: 'vite-plus',
        specifier: 'vite-plus/oxlint-plugin',
      },
    ],
  },

  fmt: {
    bracketSpacing: false,
    printWidth: 100,
    trailingComma: 'all',
    arrowParens: 'avoid',
    singleQuote: true,
    sortPackageJson: false,
    ignorePatterns: [
      'node_modules/**',
      'build/**',
      'dist/**',
      'out/**',
      'coverage/**',
      'spec/**',
      'spec/out/**',
      'spec/manual/public/**',
      '.github/**',
      'package-lock.json',
      'pnpm-lock.yaml',
      '*.xlsx',
      '*.min.js',
      'README.md',
      'README_zh.md',
      'TODO.txt',
      'MODEL.md',
      'UPGRADE-4.0.md',
    ],
  },

  test: {
    globals: true,
    environment: 'node',
    setupFiles: [path.join(root, 'spec/config/vitest.setup.ts')],
    include: [
      'spec/unit/**/*.spec.ts',
      'spec/integration/**/*.spec.ts',
      'spec/end-to-end/**/*.spec.ts',
    ],
    exclude: [
      'node_modules/**',
      'spec/browser/**',
      'spec/dist/**',
      'spec/manual/**',
      'spec/typescript/**',
      // Optional network e2e deps (express/got) — skip unless installed
      'spec/end-to-end/express.spec.ts',
    ],
    testTimeout: 30000,
    hookTimeout: 30000,
    // Integration tests share paths under spec/out/ — run files serially
    fileParallelism: false,
    sequence: {
      concurrent: false,
    },
    pool: 'forks',
  },
});
