import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
export default defineConfig(({ command }) => ({
  // GitHub Pages serves this app from https://pratikagr-ux.github.io/Upkaran-ERP/,
  // so production asset URLs need that path prefix (case-sensitive). The dev
  // server still runs at the site root.
  base: command === 'build' ? '/Upkaran-ERP/' : '/',
  plugins: [react()],
  server: { port: 3000, open: true }
}));
