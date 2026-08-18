/**
 * deploy-site.js — one command to publish the website.
 *
 * Does the whole website publish in a single step:
 *   1. regenerates the 4 translated homepages from English (build-i18n.js)
 *   2. stages every change
 *   3. commits (with your message, or a default)
 *   4. pushes to GitHub — which auto-publishes via GitHub Pages
 *
 * Usage:
 *   npm run deploy:site -- "what you changed"
 *   npm run deploy:site              (uses a default message)
 *
 * After it finishes, GitHub Pages takes ~1–2 minutes to serve the new version.
 * That wait is normal — during it the old page may still show; that is the CDN
 * cache, not a failure. You do NOT need to push again.
 *
 * This is the website equivalent of `npm run deploy:control` for the guide portal.
 */
const { execSync } = require("child_process");

const msg = process.argv.slice(2).join(" ").trim() || "Update site";
const run = (cmd) => execSync(cmd, { stdio: "inherit" });

console.log("1/4  Regenerating translated pages…");
run("node tools/build-i18n.js");

console.log("2/4  Staging changes…");
run("git add -A");

console.log("3/4  Committing…");
try {
  run(`git commit -m "${msg.replace(/"/g, '\\"')}"`);
} catch (e) {
  console.log("Nothing to commit — the site is already up to date.");
  process.exit(0);
}

console.log("4/4  Pushing to GitHub…");
run("git push");

console.log("\n✅ Pushed. GitHub Pages will publish in ~1–2 minutes.");
console.log("   The live site may show the old version for a minute (CDN cache) — that is normal, don't push again.");
