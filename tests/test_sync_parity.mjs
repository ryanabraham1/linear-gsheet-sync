import assert from "node:assert/strict";
import fs from "node:fs/promises";

const localSource = await fs.readFile("Code.local.gs", "utf8");
const publicSource = await fs.readFile("Code.gs", "utf8");
const gitignore = await fs.readFile(".gitignore", "utf8");
const keyPattern = /(LINEAR_API_KEY:\s*')[^']+(')/;

assert.match(localSource, /LINEAR_API_KEY:\s*'lin_api_[^']+'/);
assert.doesNotMatch(publicSource, /lin_api_[A-Za-z0-9]+/);
assert.match(gitignore, /^Code\.local\.gs$/m);

const normalizedLocal = localSource.replace(
  keyPattern,
  "$1PASTE_LINEAR_PERSONAL_API_KEY_HERE$2"
);
assert.equal(
  publicSource,
  normalizedLocal,
  "Code.gs drifted from Code.local.gs. Run: npm run sync-public"
);

console.log("Local and publishable scripts match; only the API key differs.");
