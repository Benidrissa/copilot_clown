/**
 * Build script that:
 * 1. Runs webpack production build
 * 2. Generates manifest.xml with the correct BASE_URL
 *
 * Usage:
 *   node scripts/build.js                              → uses https://localhost:3000 (dev)
 *   node scripts/build.js https://user.github.io/repo  → uses GitHub Pages URL
 *   BASE_URL=https://example.com node scripts/build.js  → uses env var
 */
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const DEFAULT_URL = "https://localhost:3000";
const baseUrl = process.argv[2] || process.env.BASE_URL || DEFAULT_URL;

// Remove trailing slash
const cleanUrl = baseUrl.replace(/\/+$/, "");

console.log(`\n  Building with BASE_URL: ${cleanUrl}\n`);

// 1. Run webpack
console.log("  [1/3] Running webpack production build...");
execSync("npx webpack --mode production", { stdio: "inherit", cwd: path.resolve(__dirname, "..") });

// 2. Read manifest template and replace placeholder
console.log("  [2/3] Generating manifest.xml...");
const manifestTemplate = fs.readFileSync(path.resolve(__dirname, "..", "manifest.xml"), "utf-8");
const manifest = manifestTemplate.replace(/\{\{BASE_URL\}\}/g, cleanUrl);
fs.writeFileSync(path.resolve(__dirname, "..", "dist", "manifest.xml"), manifest);

// 3. Done
console.log("  [3/3] Build complete!");
console.log(`\n  Output: dist/`);
console.log(`  Manifest: dist/manifest.xml (configured for ${cleanUrl})`);
console.log(`\n  To sideload in Excel, use the manifest.xml from the dist/ folder.\n`);
