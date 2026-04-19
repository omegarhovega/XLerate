// @ts-check
/**
 * One-shot script that downloads Fluent UI System Icons (MIT, Microsoft)
 * as SVGs and renders PNGs at the three sizes the Office manifest needs
 * (16, 32, 80) into `assets/icons/`. The generated PNGs are checked into
 * the repo; this script is not part of `ci:all`.
 *
 * Run with: npm run icons
 *
 * Fluent icons ship only as SVG (and PDF). SVG is vector, so the "size"
 * in the filename is a design hint — the 20-px variant renders crisply
 * at any output dimension. We use size-20 regular SVGs across the board
 * for visual consistency; a few icons that lack a 20 variant fall back
 * to 24.
 */

import { mkdirSync, writeFileSync, existsSync } from "node:fs";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import sharp from "sharp";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(__dirname, "..");
const OUT_DIR = resolve(REPO_ROOT, "assets/icons");
const OUT_SIZES = [16, 32, 80];

// Button/group -> Fluent icon spec. `dir` is the icon folder name in the
// Fluent repo (URL-encoded at fetch time). `slug` is the snake_case name
// used in the SVG filename.
/** @type {Array<{ key: string; dir: string; slug: string; sourceSize: number }>} */
const ICONS = [
  { key: "trace-precedents", dir: "Arrow Between Down", slug: "arrow_between_down", sourceSize: 20 },
  { key: "trace-dependents", dir: "Arrow Between Up", slug: "arrow_between_up", sourceSize: 20 },
  { key: "switch-sign", dir: "Add Subtract Circle", slug: "add_subtract_circle", sourceSize: 20 },
  { key: "smart-fill", dir: "Arrow Autofit Width", slug: "arrow_autofit_width", sourceSize: 20 },
  { key: "consistency", dir: "Checkmark Square", slug: "checkmark_square", sourceSize: 20 },
  { key: "cycle-number", dir: "Number Symbol", slug: "number_symbol", sourceSize: 20 },
  { key: "cycle-date", dir: "Calendar", slug: "calendar", sourceSize: 20 },
  { key: "cycle-cell", dir: "Table Simple", slug: "table_simple", sourceSize: 20 },
  { key: "cycle-text", dir: "Text Font", slug: "text_font", sourceSize: 20 },
  { key: "auto-color", dir: "Color", slug: "color", sourceSize: 20 },
  { key: "show-taskpane", dir: "Panel Right", slug: "panel_right", sourceSize: 20 },
  { key: "group-formulas", dir: "Math Formula", slug: "math_formula", sourceSize: 20 },
  { key: "group-auditing", dir: "Checkmark Starburst", slug: "checkmark_starburst", sourceSize: 20 },
  { key: "group-formatting", dir: "Paint Brush", slug: "paint_brush", sourceSize: 20 },
  { key: "group-settings", dir: "Settings", slug: "settings", sourceSize: 20 },
];

function fluentSvgUrl(dir, slug, sourceSize) {
  const encodedDir = encodeURIComponent(dir);
  return `https://raw.githubusercontent.com/microsoft/fluentui-system-icons/main/assets/${encodedDir}/SVG/ic_fluent_${slug}_${sourceSize}_regular.svg`;
}

async function fetchText(url) {
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`GET ${url} -> ${res.status}`);
  }
  return res.text();
}

async function main() {
  if (!existsSync(OUT_DIR)) {
    mkdirSync(OUT_DIR, { recursive: true });
  }

  for (const icon of ICONS) {
    const url = fluentSvgUrl(icon.dir, icon.slug, icon.sourceSize);
    process.stdout.write(`  ${icon.key.padEnd(20)} `);
    const svg = await fetchText(url);
    const svgBuffer = Buffer.from(svg, "utf-8");

    for (const size of OUT_SIZES) {
      const outPath = join(OUT_DIR, `${icon.key}-${size}.png`);
      await sharp(svgBuffer, { density: Math.round((size / icon.sourceSize) * 72) })
        .resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
        .png()
        .toFile(outPath);
      process.stdout.write(`${size} `);
    }
    writeFileSync(join(OUT_DIR, `${icon.key}.svg`), svg);
    process.stdout.write("✓\n");
  }

  console.log(`\nDone. ${ICONS.length} icons × ${OUT_SIZES.length} sizes -> ${OUT_DIR}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
