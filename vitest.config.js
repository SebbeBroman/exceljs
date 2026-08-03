import {defineConfig} from 'vitest/config';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
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
