"use strict";
/* Checks that the files which must agree with each other do. */
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Jev = require("../src/jev-core.js");

const read = (p) => fs.readFileSync(path.join(__dirname, "..", p), "utf8");

test("version numbers match everywhere", () => {
  const version = JSON.parse(read("package.json")).version;
  assert.equal(Jev.VERSION, version, "src/jev-core.js");
  assert.ok(read("manifest.xml").includes(`<Version>${version}.0</Version>`), "manifest.xml");
  assert.ok(read("vba/modJev.bas").includes(`JEV_VERSION As String = "${version}"`), "vba/modJev.bas");
  assert.ok(read("CHANGELOG.md").includes(`## ${version} `), "CHANGELOG.md");
});

test("every function in functions.json is registered in functions.js, and vice versa", () => {
  const declared = JSON.parse(read("src/functions.json")).functions.map((f) => f.id).sort();
  const associated = [...read("src/functions.js").matchAll(/CustomFunctions\.associate\("([A-Z]+)"/g)]
    .map((m) => m[1])
    .sort();
  assert.deepEqual(associated, declared);
});
