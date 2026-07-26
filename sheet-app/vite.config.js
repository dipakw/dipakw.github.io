import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
// Served from https://dipakw.github.io/sheet/ — assets must be prefixed with /sheet/.
// https://vitejs.dev/config/
export default defineConfig({
    base: '/sheet/',
    plugins: [react()],
    build: {
        // Emit straight into the Jekyll web root so the site publishes it at /sheet.
        outDir: '../website/public/sheet',
        emptyOutDir: true,
    },
});
