import esbuild from "esbuild";
import fs from "node:fs";
import path from "node:path";

const watch = process.argv.includes("--watch");
const distDir = path.resolve("dist");
const srcUi = path.resolve("src/ui.html");
const outUi = path.resolve("dist/ui.html");

fs.mkdirSync(distDir, { recursive: true });

function copyUi() {
  fs.copyFileSync(srcUi, outUi);
  console.log("[build] copied ui.html");
}

copyUi();

const ctx = await esbuild.context({
  entryPoints: ["src/code.ts"],
  bundle: true,
  outfile: "dist/code.js",
  format: "iife",
  platform: "browser",
  target: ["es2017"],
  logLevel: "info",
});

if (watch) {
  await ctx.watch();

  fs.watch("src", { recursive: true }, function (eventType, filename) {
    if (filename === "ui.html") {
      copyUi();
    }
  });

  console.log("[build] watching...");
} else {
  await ctx.rebuild();
  await ctx.dispose();
}
