#!/usr/bin/env bash
set -euo pipefail

readonly expected_node_version="$(<.node-version)"

if [[ "$(node --version)" != "$expected_node_version" ]]; then
  printf 'Expected Node.js %s; found %s\n' "$expected_node_version" "$(node --version)" >&2
  exit 1
fi

readonly distribution_files=(
  README.md
  index.html
  manifest.xml
  functions.html
  functions.js
  functions.json
)

for file in "${distribution_files[@]}"; do
  test -s "$file"
done

test ! -e functions.js.map
node --check functions.js
node - <<'NODE'
const assert = require("node:assert/strict");
const fs = require("node:fs");

const manifest = fs.readFileSync("manifest.xml", "utf8");
const functionsSource = fs.readFileSync("functions.js", "utf8");
const tokens = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/?[\w:.-]+(?:\s+[^<>]*?)?\/?>/g;
const openTag = /^<([\w:.-]+)(?:\s+[^<>]*?)?\/?>$/;
const closingTag = /^<\/([\w:.-]+)\s*>$/;
const stack = [];
let cursor = 0;

for (const match of manifest.matchAll(tokens)) {
  assert.doesNotMatch(manifest.slice(cursor, match.index), /[<>]/, "manifest.xml contains malformed XML");
  cursor = match.index + match[0].length;

  const token = match[0];
  if (token.startsWith("<?") || token.startsWith("<!")) {
    continue;
  }

  const closing = token.match(closingTag);
  if (closing) {
    assert.equal(stack.pop(), closing[1], `Unexpected closing tag: ${closing[1]}`);
    continue;
  }

  const opening = token.match(openTag);
  assert.ok(opening, `Malformed XML tag: ${token}`);
  if (!token.endsWith("/>")) {
    stack.push(opening[1]);
  }
}

assert.match(manifest.slice(cursor), /^\s*$/, "manifest.xml contains malformed XML");
assert.deepEqual(stack, [], "manifest.xml contains unclosed XML tags");
assert.match(manifest, /<Id>c2a1f3b3-8d4f-4c6a-9b2e-9a4e2b7f1a6d<\/Id>/);
const manifestVersion = manifest.match(/<Version>([^<]+)<\/Version>/)?.[1];
assert.match(manifestVersion ?? "", /^\d+\.\d+\.\d+\.\d+$/);

const expectedUrls = new Set([
  "https://koizumihiroo.github.io/j2e-excel-addin/functions.html",
  "https://koizumihiroo.github.io/j2e-excel-addin/functions.js",
  "https://koizumihiroo.github.io/j2e-excel-addin/functions.json",
  "https://github.com/koizumihiroo/j2e-excel-addin/issues",
]);
const urls = [...manifest.matchAll(/\bDefaultValue=(["'])(.*?)\1/g)].map((match) => match[2])
  .filter((value) => value.startsWith("http://") || value.startsWith("https://"));

for (const url of expectedUrls) {
  assert.ok(urls.includes(url), `Missing URL: ${url}`);
}
assert.ok(urls.every((url) => url.startsWith("https://")), "All manifest URLs must use HTTPS");
assert.ok(!functionsSource.includes("sourceMappingURL"), "functions.js must not include a source map");
NODE
