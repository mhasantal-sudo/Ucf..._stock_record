import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'node:fs';
import path from 'node:path';
import { defineConfig, type Plugin } from 'vite';
import { VitePWA } from 'vite-plugin-pwa';

function emitSourceSnapshotPlugin(): Plugin {
  return {
    name: 'emit-source-snapshot',
    apply: 'build',
    closeBundle() {
      const rootDir = import.meta.dirname;
      const distDir = path.join(rootDir, 'dist');
      if (!fs.existsSync(distDir)) return;

      const textExts = new Set([
        '.ts', '.tsx', '.js', '.jsx', '.css', '.html', '.json',
        '.svg', '.md', '.example', '.gitignore', '.yml', '.yaml', '.rules',
      ]);
      const files: Array<{ path: string; category: string; size: number; content: string }> = [];

      const walk = (dir: string, prefix: string, category: string) => {
        if (!fs.existsSync(dir)) return;
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
          if (entry.name === 'node_modules' || entry.name === 'dist' || entry.name === '.git') continue;
          const full = path.join(dir, entry.name);
          const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
          if (rel.startsWith('public/assets/aistudio')) continue;
          if (entry.isDirectory()) {
            walk(full, rel, category);
          } else if (entry.isFile()) {
            const ext = path.extname(entry.name).toLowerCase();
            if (!textExts.has(ext) && entry.name !== '.gitignore' && entry.name !== '.env.example') continue;
            const content = fs.readFileSync(full, 'utf-8');
            files.push({
              path: rel,
              category,
              size: Buffer.byteLength(content, 'utf-8'),
              content,
            });
          }
        }
      };

      const rootFiles = [
        'package.json',
        'tsconfig.json',
        'vite.config.ts',
        'index.html',
        'server.ts',
        'metadata.json',
        'firebase-applet-config.json',
        'firestore.rules',
        'firebase-blueprint.json',
        'README.md',
        '.env.example',
        '.gitignore',
      ];
      for (const fileName of rootFiles) {
        const full = path.join(rootDir, fileName);
        if (fs.existsSync(full)) {
          const content = fs.readFileSync(full, 'utf-8');
          files.push({
            path: fileName,
            category: 'Root & Config',
            size: Buffer.byteLength(content, 'utf-8'),
            content,
          });
        }
      }

      walk(path.join(rootDir, 'src'), 'src', 'Frontend (src)');
      walk(path.join(rootDir, 'server'), 'server', 'Server');
      walk(path.join(rootDir, 'public'), 'public', 'Public Assets');

      fs.writeFileSync(
        path.join(distDir, 'source-snapshot.json'),
        JSON.stringify({ ok: true, totalFiles: files.length, files }),
        'utf-8'
      );
    },
  };
}

export default defineConfig(() => {
  return {
    plugins: [
      react(),
      tailwindcss(),
      emitSourceSnapshotPlugin(),
      VitePWA({
        registerType: 'autoUpdate',
        includeAssets: [
          'icon.svg',
          'apple-touch-icon.png',
          'pwa-192x192.png',
          'pwa-512x512.png',
          'pwa-maskable-512x512.png',
          'default-signature.png',
        ],
        manifest: {
          id: '/',
          name: 'খাদ্যশস্য বরাদ্দ ও বিতরণ ব্যবস্থাপনা',
          short_name: 'খাদ্যবরাদ্দ',
          description:
            'উপজেলা খাদ্য নিয়ন্ত্রকের কার্যালয়ের খাদ্যশস্য বরাদ্দ, উপবরাদ্দ, ডিও ইস্যু, চালান যাচাই এবং নিকোশ ফন্টে ১৫টি রেজিস্টার ও রিপোর্ট।',
          theme_color: '#0d6e4e',
          background_color: '#f5f7f6',
          display: 'standalone',
          start_url: '/',
          scope: '/',
          icons: [
            {
              src: '/pwa-192x192.png',
              sizes: '192x192',
              type: 'image/png',
              purpose: 'any',
            },
            {
              src: '/pwa-512x512.png',
              sizes: '512x512',
              type: 'image/png',
              purpose: 'any',
            },
            {
              src: '/pwa-maskable-512x512.png',
              sizes: '512x512',
              type: 'image/png',
              purpose: 'maskable',
            },
          ],
        },
        workbox: {
          maximumFileSizeToCacheInBytes: 5 * 1024 * 1024,
          navigateFallbackDenylist: [/^\/api\//],
          globPatterns: ['**/*.{js,css,html,ico,png,svg,ttf,woff,woff2}'],
          runtimeCaching: [
            {
              urlPattern: /^https:\/\/fonts\.googleapis\.com\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'google-fonts-cache',
                expiration: {
                  maxEntries: 10,
                  maxAgeSeconds: 60 * 60 * 24 * 365,
                },
                cacheableResponse: {
                  statuses: [0, 200],
                },
              },
            },
            {
              urlPattern: /^https:\/\/fonts\.gstatic\.com\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'gstatic-fonts-cache',
                expiration: {
                  maxEntries: 10,
                  maxAgeSeconds: 60 * 60 * 24 * 365,
                },
                cacheableResponse: {
                  statuses: [0, 200],
                },
              },
            },
          ],
        },
        devOptions: {
          enabled: true,
          type: 'module',
        },
      }),
    ],
    resolve: {
      alias: {
        '@': path.resolve(import.meta.dirname, '.'),
      },
    },
    server: {
      hmr: process.env.DISABLE_HMR !== 'true',
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});

