import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  base: "./",
  test: {
    poolOptions: {
      // Node >=22 injects a broken localStorage global that shadows jsdom's;
      // disable it so tests get the real Storage implementation.
      threads: { execArgv: ["--no-experimental-webstorage"] },
      forks: { execArgv: ["--no-experimental-webstorage"] },
    },
  },
  server: { host: "127.0.0.1", port: 5173 },
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes("xlsx")) return "spreadsheet-engine";
          if (id.includes("lucide-react")) return "icons";
          if (id.includes("react")) return "react";
        },
      },
    },
  },
});
