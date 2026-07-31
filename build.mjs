/**
 * Bundle the add-on for clasp.
 *
 * Apps Script has no module system and no npm, so everything — including
 * @jobo-ai/connector-core — is inlined into a single `Code.gs`. IIFE format
 * keeps the module scope private; Code.ts assigns the handlers Apps Script needs
 * onto `globalThis` explicitly.
 *
 * Also emits an ESM copy under dist-test/ so the pure helpers can be unit tested
 * on Node without an Apps Script runtime.
 */
import { build } from "esbuild";
import { copyFileSync, mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const dist = join(root, "dist");
const distTest = join(root, "dist-test");
const connectorUi = join(root, "..", "packages", "connector-ui");

mkdirSync(dist, { recursive: true });
mkdirSync(distTest, { recursive: true });

// ── The bundle clasp pushes ──────────────────────────────────────────────────
await build({
  entryPoints: [join(root, "src", "Code.ts")],
  outfile: join(dist, "Code.js"),
  bundle: true,
  format: "iife",
  // Apps Script's V8 runtime does not accept ES2020+ syntax in every position;
  // es2019 is the safe floor and costs nothing here.
  target: "es2019",
  platform: "neutral",
  legalComments: "none",
});

// Apps Script expects .gs for server code.
renameSync(join(dist, "Code.js"), join(dist, "Code.gs"));

// ── Sidebar: inline the shared UI assets from packages/connector-ui ──────────
// The sidebar HTML carries @inline markers instead of committed copies, so
// tokens/combobox can never drift from their single source. combobox.js is
// authored as an ES module for its tests; stripping `export ` leaves plain
// function declarations a <script> block can use.
{
  const inlines = [
    { marker: "/* @inline:tokens.css */", file: "tokens.css", transform: (s) => s },
    { marker: "/* @inline:combobox.css */", file: "combobox.css", transform: (s) => s },
    { marker: "// @inline:combobox.js", file: "combobox.js", transform: (s) => s.replace(/^export /gm, "") },
  ];

  let sidebar = readFileSync(join(root, "src", "Sidebar.html"), "utf8");
  for (const { marker, file, transform } of inlines) {
    if (!sidebar.includes(marker)) {
      throw new Error(`Sidebar.html is missing the inline marker "${marker}" for ${file}`);
    }
    sidebar = sidebar.replace(marker, transform(readFileSync(join(connectorUi, file), "utf8")));
  }
  writeFileSync(join(dist, "Sidebar.html"), sidebar);
}

copyFileSync(join(root, "appsscript.json"), join(dist, "appsscript.json"));

// ── Testable ESM copy ────────────────────────────────────────────────────────
await build({
  entryPoints: [join(root, "src", "Code.ts")],
  outfile: join(distTest, "Code.js"),
  bundle: true,
  format: "esm",
  target: "es2022",
  platform: "neutral",
  legalComments: "none",
});

await build({
  entryPoints: [join(root, "src", "Code.test.ts")],
  outfile: join(distTest, "Code.test.js"),
  bundle: true,
  format: "esm",
  target: "es2022",
  platform: "node",
  external: ["node:*"],
  legalComments: "none",
});

console.log("built dist/Code.gs, dist/Sidebar.html, dist/appsscript.json");
