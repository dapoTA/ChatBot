process.on("unhandledRejection", (reason) => {
  console.error("=== UNHANDLED REJECTION ===");
  console.error("raw reason:", reason);
  console.error("type:", reason?.constructor?.name);
  console.error("message:", reason?.message);
  console.error("stack:", reason?.stack);
  console.error("cause:", reason?.cause);
  console.error("error:", reason?.error);
  console.error("target:", reason?.target);
  console.error("stringified:", JSON.stringify(reason, Object.getOwnPropertyNames(reason)));
});

process.on("uncaughtException", (err) => {
  console.error("=== UNCAUGHT EXCEPTION ===");
  console.error("raw error:", err);
  console.error("type:", err?.constructor?.name);
  console.error("message:", err?.message);
  console.error("stack:", err?.stack);
});

process.on("uncaughtExceptionMonitor", (err) => {
  console.error("=== UNCAUGHT EXCEPTION MONITOR ===");
  console.error(err);
});

import "dotenv/config";
import { warnIfKeyMissing } from "./crypto.js";

console.log("ENV deployment:", process.env.AZURE_OPENAI_DEPLOYMENT);
console.log("ENV api version:", process.env.AZURE_OPENAI_API_VERSION);
console.log("ENV endpoint:", process.env.AZURE_OPENAI_ENDPOINT);
warnIfKeyMissing();

import express from "express";
import { registerRoutes } from "./routes.js";
import { serveStatic } from "./static.js";
import { createServer } from "http";

const app = express();
const httpServer = createServer(app);

app.use(
  express.json({
    verify: (req, _res, buf) => {
      req.rawBody = buf;
    },
  }),
);

app.use(express.urlencoded({ extended: false }));

export function log(message, source = "express") {
  const formattedTime = new Date().toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true,
  });

  console.log(`${formattedTime} [${source}] ${message}`);
}

app.use((req, res, next) => {
  try {
  const start = Date.now();
  const path = req.path;
  let capturedJsonResponse = undefined;

  const originalResJson = res.json;
  res.json = function (bodyJson, ...args) {
    capturedJsonResponse = bodyJson;
    return originalResJson.apply(res, [bodyJson, ...args]);
  };

  res.on("finish", () => {
    const duration = Date.now() - start;
    if (path.startsWith("/api")) {
      let logLine = `${req.method} ${path} ${res.statusCode} in ${duration}ms`;
      if (capturedJsonResponse) {
        logLine += ` :: ${JSON.stringify(capturedJsonResponse)}`;
      }

      log(logLine);
    }
  });

  next();
  } catch (err) {
    console.error("❌ Server startup failed:");
    console.error(err);

    // Optional: cleaner exit
    process.exit(1);
  }
});

(async () => {
  try {
    console.log("A: entering startup");

    console.log("B: before registerRoutes");
    await registerRoutes(httpServer, app);
    console.log("C: after registerRoutes");

    if (process.env.NODE_ENV === "production") {
      console.log("D: before serveStatic");
      serveStatic(app);
      console.log("E: after serveStatic");
    } else {
      console.log("D: before import vite");
      const { setupVite } = await import("./vite.js");
      console.log("E: after import vite");

      console.log("F: before setupVite");
      await setupVite(httpServer, app);
      console.log("G: after setupVite");
    }

    // iisnode passes a named pipe path as PORT (e.g. \\.\pipe\...)
    // so we must NOT parseInt — pass the raw value and let Node handle both pipes and numbers
    const port = process.env.PORT || 5000;
    const listenArg = typeof port === "string" && port.startsWith("\\\\.\\")
      ? port
      : { port: Number(port), host: "0.0.0.0" };

      console.log("H: before listen");

      httpServer.listen(
        listenArg,
        () => {
          log(`serving on ${port}`);
        }
      );
    console.log("I: after listen call");
  } catch (err) {
    console.error("=== STARTUP FAILED ===");
    console.error("raw:", err);
    console.error("type:", err?.constructor?.name);
    console.error("message:", err?.message);
    console.error("stack:", err?.stack);
    console.error("cause:", err?.cause);
    process.exit(1);
  }
})();