import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    host: '127.0.0.1',
    watch: {
      ignored: [
        '**/diagram_assets/**',
        '**/.web_outputs/**',
        '**/.web_uploads/**',
        '**/scratch/**',
        '**/tests/**',
        '**/*.db',
        '**/*.db-journal',
        '**/*.db-wal',
        '**/*.mmd',
        '**/*.png',
        '**/*.svg',
        '**/*.docx',
        '**/*.pptx',
        '**/*.pdf',
        '**/*.xlsx',
        '**/*.tmp',
        '**/~$*',
        '**/.~lock.*',
      ],
    },
    proxy: {
      '/api': {
        target: 'http://127.0.0.1:8000',
        changeOrigin: true,
        configure: (proxy) => {
          proxy.on('error', (err, _req, res) => {
            if (err.code === 'ECONNREFUSED' && res && !res.headersSent) {
              res.writeHead(503, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ error: 'API Bridge connecting', connecting: true }));
            }
          });
        },
      },
    },
  },
  build: {
    outDir: 'dist',
    sourcemap: false,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('node_modules/react')) return 'react-vendor';
          if (id.includes('node_modules/lucide-react')) return 'lucide';
        },
      },
    },
  },
});
