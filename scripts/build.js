// Build: combine the shared engine with each consumer "part" into the files
// the manifests actually load. Zero dependencies on purpose — this project has
// no npm packages, and simple concatenation keeps the output ES2016-safe by
// construction (we write the sources in that style).
//
//   src/lib/engine.js + src/lib/sendRuntime.part.js   -> src/sendTestRuntime.js
//   src/lib/engine.js + src/lib/taskpaneCore.part.js  -> src/recipientGuardCore.js
//
// Run: npm run build   (or: node scripts/build.js)

const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const lib = (name) => fs.readFileSync(path.join(root, "src", "lib", name), "utf8");

const banner = (parts) =>
  "// GENERATED FILE — DO NOT EDIT.\n" +
  "// Built by scripts/build.js from: " + parts.join(" + ") + "\n" +
  "// Edit those sources, then run `npm run build`.\n\n";

function emit(outName, partName) {
  const out =
    banner(["src/lib/engine.js", "src/lib/" + partName]) +
    "(function () {\n  \"use strict\";\n\n" +
    indent(lib("engine.js")) + "\n" +
    indent(lib(partName)) +
    "})();\n";
  fs.writeFileSync(path.join(root, "src", outName), out);
  console.log("built src/" + outName);
}

function indent(source) {
  return source
    .split("\n")
    .map((line) => (line.trim() === "" ? "" : "  " + line))
    .join("\n");
}

emit("sendTestRuntime.js", "sendRuntime.part.js");
emit("recipientGuardCore.js", "taskpaneCore.part.js");
