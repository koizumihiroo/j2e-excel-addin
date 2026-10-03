#!/usr/bin/env bash
set -euo pipefail

site_dir="${1:-_site}"

if [[ -e "$site_dir" ]]; then
  printf 'Refusing to overwrite existing Pages directory: %s\n' "$site_dir" >&2
  exit 1
fi

mkdir -p "$site_dir"
cp README.md index.html manifest.xml functions.html functions.js functions.json "$site_dir/"
