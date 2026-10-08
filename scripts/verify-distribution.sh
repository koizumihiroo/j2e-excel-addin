#!/usr/bin/env bash
set -euo pipefail

readonly expected_node_version="$(<.node-version)"

if [[ "$(node --version)" != "$expected_node_version" ]]; then
  printf 'Expected Node.js %s; found %s\n' "$expected_node_version" "$(node --version)" >&2
  exit 1
fi

readonly distribution_files=(
  README.md
  public/index.html
  public/manifest.xml
  public/functions.html
  public/functions.js
  public/functions.json
)
readonly root_release_files=(
  index.html
  manifest.xml
  functions.html
  functions.js
  functions.json
  example.xlsx
)

if [[ ! -d public || -L public ]]; then
  printf 'Expected a regular public directory\n' >&2
  exit 1
fi

for file in "${root_release_files[@]}"; do
  if [[ -e "$file" || -L "$file" ]]; then
    printf 'Release files must be stored under public/: %s\n' "$file" >&2
    exit 1
  fi
done

if [[ -e public/example.xlsx || -L public/example.xlsx ]]; then
  printf 'Retired release file must not be present: public/example.xlsx\n' >&2
  exit 1
fi

for file in "${distribution_files[@]}"; do
  if [[ ! -f "$file" || -L "$file" ]]; then
    printf 'Expected a regular distribution file: %s\n' "$file" >&2
    exit 1
  fi
  test -s "$file"
done

test ! -e public/functions.js.map
node --check public/functions.js
node - <<'NODE'
const assert = require("node:assert/strict");
const fs = require("node:fs");

const manifest = fs.readFileSync("public/manifest.xml", "utf8");
const functionsSource = fs.readFileSync("public/functions.js", "utf8");
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
