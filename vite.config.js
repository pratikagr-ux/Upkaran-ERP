import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
export default defineConfig(({ command }) => ({
  // GitHub Pages serves this app from https://pratikagr.github.io/upkaran-erp/,
  // so production asset URLs need that path prefix. The dev server still runs
  // at the site root.
  base: command === 'build' ? '/upkaran-erp/' : '/',
  plugins: [react()],
  server: { port: 3000, open: true }
}));
