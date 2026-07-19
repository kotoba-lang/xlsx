/**
 * Post-build script: read Vite output and produce a single self-contained HTML file.
 * Inlines JS and CSS into the HTML so it can be served from a Hono route without static assets.
 */
import { readFileSync, writeFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

const outDir = join(import.meta.dirname, "../../_svelte");
const assetsDir = join(outDir, "assets");

// Find JS and CSS files
const files = readdirSync(assetsDir);
const jsFile = files.find(f => f.endsWith(".js"));
const cssFile = files.find(f => f.endsWith(".css"));

if (!jsFile) { console.error("No JS file found in _svelte/assets/"); process.exit(1); }

const jsContent = readFileSync(join(assetsDir, jsFile), "utf-8");
const cssContent = cssFile ? readFileSync(join(assetsDir, cssFile), "utf-8") : "";

const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>XLSX Editor — xlsx.etzhayyim.com</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:100%;height:100%;overflow:hidden}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#fff;color:#333}
${cssContent}
</style>
</head>
<body>
<div id="app"></div>
<script type="module">${jsContent}<\/script>
</body>
</html>`;

writeFileSync(join(outDir, "inline.html"), html, "utf-8");
// Also overwrite index.html so ASSETS binding serves the full inlined editor at /
writeFileSync(join(outDir, "index.html"), html, "utf-8");

// Patch src/app.ts — replace EDITOR_HTML base64 content
const srcDir = join(import.meta.dirname, "../../src");
const appTs = readFileSync(join(srcDir, "app.ts"), "utf-8");
const b64 = Buffer.from(html, "utf-8").toString("base64");
const marker = 'const EDITOR_HTML_B64 = ""; // AUTO-REPLACED BY BUILD';
const replacement = `const EDITOR_HTML_B64 = "${b64}"; // AUTO-REPLACED BY BUILD`;
if (appTs.includes(marker)) {
  writeFileSync(join(srcDir, "app.ts"), appTs.replace(marker, replacement), "utf-8");
  console.log(`Patched app.ts EDITOR_HTML_B64: ${Math.round(b64.length / 1024)}KB base64`);
} else if (appTs.includes("// AUTO-REPLACED BY BUILD")) {
  const patched = appTs.replace(/const EDITOR_HTML_B64 = "[^]*?"; \/\/ AUTO-REPLACED BY BUILD/, replacement);
  writeFileSync(join(srcDir, "app.ts"), patched, "utf-8");
  console.log(`Re-patched app.ts EDITOR_HTML_B64: ${Math.round(b64.length / 1024)}KB base64`);
} else {
  console.warn("WARN: marker not found in app.ts — EDITOR_HTML not updated");
}
console.log(`inline.html: ${Math.round(html.length / 1024)}KB (JS: ${Math.round(jsContent.length / 1024)}KB, CSS: ${Math.round(cssContent.length / 1024)}KB)`);
