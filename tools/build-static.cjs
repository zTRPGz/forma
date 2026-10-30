"use strict";

const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const out = path.join(root, "dist");

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });

for (const entry of ["index.html", "styles.css", ".nojekyll", "js", "vendor"]) {
  fs.cpSync(path.join(root, entry), path.join(out, entry), { recursive: true });
}

console.log("dist собран: " + out);
