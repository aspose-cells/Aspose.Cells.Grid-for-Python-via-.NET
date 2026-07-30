import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

const backend = 'http://127.0.0.1:2022';

export default defineConfig({
  base: '/demos/vue/',
  plugins: [vue()],
  server: {
    proxy: {
      '/GridJs': { target: backend, changeOrigin: true },
    },
  },
  build: {
    outDir: '../../static/demos/vue',
    emptyOutDir: true,
  },
});
