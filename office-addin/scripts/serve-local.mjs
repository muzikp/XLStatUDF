import fs from "node:fs";
import fsp from "node:fs/promises";
import https from "node:https";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.dirname(__dirname);
const distDir = path.join(rootDir, "dist");
const certDir = path.join(rootDir, ".certs");
const port = Number(process.env.EVALYTICS_OFFICE_PORT ?? process.env.XLSTAT_OFFICE_PORT ?? 3000);

const pfxPath = path.join(certDir, "localhost-devcert.pfx");
const passphrasePath = path.join(certDir, "localhost-devcert.pass.txt");

if (!fs.existsSync(pfxPath) || !fs.existsSync(passphrasePath)) {
  console.error("Missing local dev certificate. Run the 'office-addin: ensure dev cert' task first.");
  process.exit(1);
}

const pfx = await fsp.readFile(pfxPath);
const passphrase = (await fsp.readFile(passphrasePath, "utf8")).trim();

const contentTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "application/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".map", "application/json; charset=utf-8"],
  [".xml", "application/xml; charset=utf-8"],
  [".png", "image/png"],
  [".svg", "image/svg+xml"],
  [".ico", "image/x-icon"]
]);

const server = https.createServer({ pfx, passphrase }, async (req, res) => {
  try {
    const urlPath = new URL(req.url ?? "/", `https://localhost:${port}`).pathname;
    let relativePath = urlPath === "/" ? "/functions.html" : urlPath;
    if (relativePath.includes("..")) {
      res.writeHead(400);
      res.end("Invalid path");
      return;
    }

    const filePath = path.join(distDir, relativePath);
    const stat = await fsp.stat(filePath);
    if (!stat.isFile()) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "Content-Type": contentTypes.get(ext) ?? "application/octet-stream",
      "Cache-Control": "no-store"
    });
    fs.createReadStream(filePath).pipe(res);
  } catch {
    res.writeHead(404);
    res.end("Not found");
  }
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Evalytics Office add-in served at https://localhost:${port}/functions.html`);
  console.log(`Local manifest: ${path.join(distDir, "manifest.local.xml")}`);
});
