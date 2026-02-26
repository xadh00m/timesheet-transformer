import { defineConfig } from "vite";

export default defineConfig({
  base: "./",
  server: {
    port: 4567,
  },
  build: {
    outDir: "dist",
  },
});
