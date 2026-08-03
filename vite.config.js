import {defineConfig} from 'vite-plus';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));

/**
 * Vite+ config for this ESM library.
 *
 * - Publish is source-only (`excel.js` + `lib/`) — no `vp build` / `vp pack` required
 * - Tests: Vitest via `vp test` (`vite-plus/test`)
 * - Lint/format: Oxlint + Oxfmt via `vp check` / `vp lint` / `vp fmt`
 */
export default defineConfig({
  staged: {
    // Keep pre-commit scoped to maintained surface
    '{lib,scripts}/**/*.{js,mjs},excel.js,vite.config.js,package.json,index.d.ts,benchmark.js':
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
      'scripts/cjs-to-esm.mjs',
      'scripts/migrate-specs-to-vitest.mjs',
      // Legacy ad-hoc scripts (many still CJS)
      'test/**',
    ],
    rules: {
      // Spreadsheet fixtures use holey arrays intentionally
      'no-sparse-arrays': 'off',
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
        files: ['spec/**/*.{js,mjs,cjs}'],
        rules: {
          'no-console': 'off',
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
        },
      },
    ],
    // Pure JS library with a hand-written index.d.ts — skip TS type-aware lint
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
      'test/**',
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
      'ESM.md',
      'MODEL.md',
      'UPGRADE-4.0.md',
      // One-shot migration helpers
      'scripts/cjs-to-esm.mjs',
      'scripts/migrate-specs-to-vitest.mjs',
    ],
  },

  test: {
    globals: true,
    environment: 'node',
    setupFiles: [path.join(root, 'spec/config/vitest.setup.js')],
    include: [
      'spec/unit/**/*.spec.js',
      'spec/integration/**/*.spec.js',
      'spec/end-to-end/**/*.spec.js',
    ],
    exclude: [
      'node_modules/**',
      'spec/browser/**',
      'spec/dist/**',
      'spec/manual/**',
      'spec/typescript/**',
      // Optional network e2e deps (express/got) — skip unless installed
      'spec/end-to-end/express.spec.js',
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
