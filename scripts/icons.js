// Regenerate the PNG icons from assets/icon.svg (the source of truth).
//   npm run icons
//
// Sizes: 16/32/80 are the manifest ribbon icons, 64 is IconUrl, 128 is
// HighResolutionIconUrl, 215 is the Entra app-registration logo (Entra requires
// 215x215 PNG, <100KB — it's what shows on the consent screen), 300 is the
// AppSource store logo.
//
// Uses sharp via npx so the repo keeps no image dependency.

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const assets = path.join(root, "assets");
const svg = path.join(assets, "icon.svg");
const tmp = path.join(root, ".tmp-icons");
const SIZES = [16, 32, 64, 80, 128, 215, 300];

fs.rmSync(tmp, { recursive: true, force: true });
fs.mkdirSync(tmp, { recursive: true });

SIZES.forEach((size) => {
  execSync(
    `npx -y sharp-cli -i "${svg}" -o "${tmp}" resize ${size} ${size} --format png`,
    { stdio: "pipe" }
  );
  fs.renameSync(path.join(tmp, "icon.png"), path.join(assets, `icon-${size}.png`));
  console.log("built assets/icon-" + size + ".png");
});

fs.rmSync(tmp, { recursive: true, force: true });
