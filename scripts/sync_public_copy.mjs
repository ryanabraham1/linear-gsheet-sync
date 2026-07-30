import fs from "node:fs/promises";

const localPath = "Code.local.gs";
const publicPath = "Code.gs";
const placeholder = "PASTE_LINEAR_PERSONAL_API_KEY_HERE";

const localSource = await fs.readFile(localPath, "utf8");
const keyPattern = /(LINEAR_API_KEY:\s*')[^']+(')/;
if (!keyPattern.test(localSource)) {
  throw new Error(`Could not find LINEAR_API_KEY in ${localPath}.`);
}

const publicSource = localSource.replace(keyPattern, `$1${placeholder}$2`);
await fs.writeFile(publicPath, publicSource);
console.log(`Updated ${publicPath} from ${localPath} with the API key redacted.`);
