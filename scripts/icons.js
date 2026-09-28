#!/usr/bin/env node
/*
 * Renders src/assets/icon.svg to every PNG size the add-in and the store listing use.
 *
 *   npm run icons
 *
 *  16, 32, 80  ribbon button (manifest.xml)
 *  32, 64      add-in icon and high-resolution icon (manifest.xml); 32 is also in the task pane
 *  300         store logo for Microsoft Partner Center
 */
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const { Resvg } = require("@resvg/resvg-js");

const SIZES = [16, 32, 64, 80, 300];
const dir = path.join(__dirname, "..", "src", "assets");
const svg = fs.readFileSync(path.join(dir, "icon.svg"));

for (const size of SIZES) {
  const png = new Resvg(svg, { fitTo: { mode: "width", value: size } }).render().asPng();
  fs.writeFileSync(path.join(dir, `icon-${size}.png`), png);
  console.log(`Wrote src/assets/icon-${size}.png`);
}
