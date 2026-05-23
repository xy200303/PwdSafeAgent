import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["tests/**/*.{test,spec}.{ts,tsx}"],
    exclude: ["node_modules", "out", "data"],
    globals: true,
    restoreMocks: true,
    testTimeout: 15000
  }
});
