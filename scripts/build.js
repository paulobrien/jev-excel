#!/usr/bin/env node
/*
 * Builds ./dist for the Cloudflare Worker (see wrangler.toml): the add-in's
 * files from ./src, plus manifest.xml, which the Worker rewrites to its own URL.
 */
"use strict";
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");

fs.rmSync(dist, { recursive: true, force: true });
fs.cpSync(path.join(root, "src"), dist, { recursive: true });
fs.copyFileSync(path.join(root, "manifest.xml"), path.join(dist, "manifest.xml"));
// Excel on the web fetches functions.json from another origin.
fs.writeFileSync(path.join(dist, "_headers"), "/*\n  Access-Control-Allow-Origin: *\n");
console.log("Built dist/");
