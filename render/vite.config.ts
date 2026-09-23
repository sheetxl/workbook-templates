import path from 'node:path';

import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// The templates are served as the page's static files, so `?src=finance/loan-calculator.xlsx`
// resolves beside it.
export default defineConfig({
  root: import.meta.dirname,
  publicDir: path.resolve(import.meta.dirname, '../templates'),
  plugins: [react()],
  build: {
    outDir: path.resolve(import.meta.dirname, '../.render'),
    emptyOutDir: true,
    chunkSizeWarningLimit: 20000
  },
  logLevel: 'warn'
});
