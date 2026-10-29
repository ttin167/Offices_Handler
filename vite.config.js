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
        '**/scratch/**',
        '**/tests/**',
        '**/*.db',
        '**/*.db-journal',
        '**/*.mmd',
        '**/*.png',
        '**/*.svg',
        '**/*.docx',
        '**/*.pptx',
        '**/*.pdf',
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
    sourcemap: true,
  },
});
