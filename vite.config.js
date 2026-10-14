import { defineConfig } from 'vite';

export default defineConfig({
  server: {
    port: 5173,
    host: '127.0.0.1',
    watch: {
      ignored: [
        '**/diagram_assets/**',
        '**/.web_outputs/**',
        '**/.web_uploads/**',
        '**/tests/**',
        '**/*.db',
        '**/*.db-journal',
        '**/*.mmd',
        '**/*.png',
        '**/*.svg',
        '**/*.docx',
        '**/*.pptx',
        '**/*.pdf',
      ],
    },
    proxy: {
      '/api': {
        target: 'http://127.0.0.1:8000',
        changeOrigin: true,
      },
    },
  },
  build: {
    outDir: 'dist',
    sourcemap: true,
  },
});
