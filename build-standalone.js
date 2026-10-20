import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const cssContent = fs.readFileSync(path.join(__dirname, "dist", "styles.css"), "utf8");
const storesContent = fs.readFileSync(path.join(__dirname, "dist", "stores.js"), "utf8")
  .replace(/export\s+const/g, "const");
const plannerContent = fs.readFileSync(path.join(__dirname, "dist", "route-planner.js"), "utf8")
  .replace(/export\s+function/g, "function")
  .replace(/export\s+const/g, "const");
const appContent = fs.readFileSync(path.join(__dirname, "dist", "app.js"), "utf8")
  .replace(/import\s+[\s\S]*?from\s+["'].*?["'];?/g, "");

const standaloneJs = `
// === stores.js ===
${storesContent}

// === route-planner.js ===
${plannerContent}

// === app.js ===
${appContent}
`;

const htmlTemplate = fs.readFileSync(path.join(__dirname, "dist", "index.html"), "utf8");
const bundledHtml = htmlTemplate
  .replace('<link rel="stylesheet" href="./styles.css" />', `<style>\n${cssContent}\n</style>`)
  .replace('<script type="module" src="./app.js"></script>', `<script>\n${standaloneJs}\n</script>`);

fs.writeFileSync(path.join(__dirname, "index.html"), bundledHtml, "utf8");
fs.writeFileSync(path.join(__dirname, "dist", "standalone.html"), bundledHtml, "utf8");
console.log("成功生成根目錄 index.html 與 dist/standalone.html！");
