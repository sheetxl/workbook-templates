/**
 * Renders a thumbnail for every template into `.thumbnails/<category>/<name>.webp`, which
 * `config/build.ts` then publishes beside the workbook.
 *
 *   npm run thumbnails [-- <filter>]
 *
 * Each template is opened by `render/` (the pinned @sheetxl/studio-mui) in headless Chromium,
 * photographed once the grid has painted and its fonts have loaded, and scaled down to webp. A
 * hand-made `templates/<category>/<name>.png|.webp` wins over the rendered one, so it is skipped.
 *
 * Needs a Chromium for Playwright: `npx playwright install chromium` (CI adds `--with-deps`).
 */
import fs from "node:fs/promises";
import path from "node:path";

import { build, preview } from "vite";
import { chromium } from "playwright";
import sharp from "sharp";

const SRC = path.resolve("templates");
const OUT = path.resolve(".thumbnails");
const RENDER_CONFIG = path.resolve("render/vite.config.ts");

/** The rendered viewport, in CSS pixels. */
const VIEWPORT = { width: 560, height: 350 };
const SCALE = 2;
/** The published thumbnail, in pixels. Same 8:5 shape as the viewport. */
const THUMBNAIL = { width: 480, height: 300 };
const LOAD_TIMEOUT = 60_000;

const HAND_MADE = [".png", ".webp"];

async function listTemplates(filter?: string): Promise<string[]> {
  const found: string[] = [];
  for (const dir of await fs.readdir(SRC, { withFileTypes: true })) {
    if (!dir.isDirectory() || dir.name.startsWith(".")) continue;
    const names = new Set(await fs.readdir(path.join(SRC, dir.name)));
    for (const name of names) {
      if (!name.toLowerCase().endsWith(".xlsx")) continue;
      const base = name.slice(0, -".xlsx".length);
      if (HAND_MADE.some((ext) => names.has(base + ext))) continue;
      const rel = `${dir.name}/${name}`;
      if (!filter || rel.includes(filter)) found.push(rel);
    }
  }
  return found.sort();
}

async function main(): Promise<void> {
  const templates = await listTemplates(process.argv[2]);
  if (templates.length === 0) {
    console.log("No templates need a rendered thumbnail.");
    return;
  }

  await build({ configFile: RENDER_CONFIG });
  const server = await preview({ configFile: RENDER_CONFIG, preview: { port: 0, host: "127.0.0.1" } });
  const origin = server.resolvedUrls?.local[0];
  if (!origin) throw new Error("The render page did not start.");

  const browser = await chromium.launch();
  const failures: string[] = [];
  try {
    const context = await browser.newContext({ viewport: VIEWPORT, deviceScaleFactor: SCALE });
    for (const rel of templates) {
      const page = await context.newPage();
      const errors: string[] = [];
      page.on("pageerror", (e) => errors.push(e.message));
      try {
        await page.goto(`${origin}?src=${encodeURIComponent(rel)}`);
        await page.waitForFunction(() => window.__thumbnail !== undefined, undefined, { timeout: LOAD_TIMEOUT });
        const outcome = await page.evaluate(() => window.__thumbnail);
        if (outcome !== "ready") throw new Error(outcome?.error ?? "did not load");
        await page.evaluate(async () => {
          await document.fonts.ready;
          await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        });
        const png = await page.screenshot({ type: "png" });
        const webp = await sharp(png).resize(THUMBNAIL.width, THUMBNAIL.height, { fit: "cover", position: "left top" }).webp({ quality: 82 }).toBuffer();
        const out = path.join(OUT, rel.replace(/\.xlsx$/i, ".webp"));
        await fs.mkdir(path.dirname(out), { recursive: true });
        await fs.writeFile(out, webp);
        console.log(`  ok  ${rel}  ${(webp.length / 1024).toFixed(1)} KB`);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        failures.push(`${rel}: ${message}${errors.length ? ` (${errors.join("; ")})` : ""}`);
        console.log(`FAIL  ${rel}  ${message}`);
      } finally {
        await page.close();
      }
    }
  } finally {
    await browser.close();
    await server.close();
  }

  if (failures.length) {
    console.error(`\n${failures.length} thumbnail(s) failed:`);
    for (const f of failures) console.error(`  - ${f}`);
    process.exit(1);
  }
  console.log(`\nRendered ${templates.length} thumbnail(s) into .thumbnails/`);
}

declare global {
  interface Window { __thumbnail?: "ready" | { error: string } }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
