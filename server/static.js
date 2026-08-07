import express from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

// Works in both native ESM (dev) and esbuild CJS bundle (production)
const __dirname = typeof __filename !== "undefined"
  ? path.dirname(__filename)
  : path.dirname(fileURLToPath(import.meta.url));

export function serveStatic(app) {
  const distPath = path.resolve(__dirname, "public");
  if (!fs.existsSync(distPath)) {
    throw new Error(
      `Could not find the build directory: ${distPath}, make sure to build the client first`,
    );
  }

  app.use(express.static(distPath));

  app.use("*", (_req, res) => {
    res.sendFile(path.resolve(distPath, "index.html"));
  });
}
