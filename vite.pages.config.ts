import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: "/SY-LAND/",
  plugins: [react()],
  build: {
    outDir: "pages-dist",
    emptyOutDir: true,
    rollupOptions: { input: "github-pages.html" },
  },
});
