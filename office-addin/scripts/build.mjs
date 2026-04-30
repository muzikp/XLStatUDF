import { mkdir, copyFile, readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import esbuild from "esbuild";

const rootDir = dirname(dirname(fileURLToPath(import.meta.url)));
const distDir = join(rootDir, "dist");
const buildStamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);

await mkdir(distDir, { recursive: true });

await esbuild.build({
  entryPoints: [join(rootDir, "src", "custom-functions", "functions.ts")],
  bundle: true,
  outfile: join(distDir, "functions.js"),
  format: "iife",
  target: "es2020",
  minify: true,
  sourcemap: false,
  define: {
    __EVALYTICS_BUILD_STAMP__: JSON.stringify(buildStamp)
  },
});

const functionsHtml = (await readFile(join(rootDir, "src", "public", "functions.html"), "utf8"))
  .replace("./functions.js", `./functions.js?v=${buildStamp}`);
await writeFile(join(distDir, "functions.html"), functionsHtml, "utf8");
await copyFile(join(rootDir, "src", "public", "functions.json"), join(distDir, "functions.json"));
await copyFile(join(rootDir, "src", "public", "function-docs.json"), join(distDir, "function-docs.json"));
await copyFile(join(rootDir, "src", "public", "commands.html"), join(distDir, "commands.html"));
await copyFile(join(rootDir, "src", "public", "commands.js"), join(distDir, "commands.js"));

const taskpaneCss = await readFile(join(rootDir, "src", "public", "taskpane.css"), "utf8");
const taskpaneHtml = (await readFile(join(rootDir, "src", "public", "taskpane.html"), "utf8"))
  .replaceAll("__BUILD_STAMP__", buildStamp)
  .replace('<link rel="stylesheet" href="./taskpane.css" />', `<style>\n${taskpaneCss}\n</style>`)
  .replace("./taskpane.js", `./taskpane.js?v=${buildStamp}`);
await writeFile(join(distDir, "taskpane.html"), taskpaneHtml, "utf8");
await copyFile(join(rootDir, "src", "public", "taskpane.js"), join(distDir, "taskpane.js"));
await copyFile(join(rootDir, "src", "public", "taskpane.css"), join(distDir, "taskpane.css"));
await copyFile(join(rootDir, "src", "public", "icon-16.png"), join(distDir, "icon-16.png"));
await copyFile(join(rootDir, "src", "public", "icon-32.png"), join(distDir, "icon-32.png"));
await copyFile(join(rootDir, "src", "public", "icon-80.png"), join(distDir, "icon-80.png"));
await copyFile(join(rootDir, "manifest.xml"), join(distDir, "manifest.xml"));
await copyFile(join(rootDir, "manifest.docs.xml"), join(distDir, "manifest.docs.xml"));
