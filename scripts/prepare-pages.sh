#!/usr/bin/env bash
set -euo pipefail

site_dir="${1:-_site}"
marker_name=".j2e-pages-generated"
marker_content="j2e-excel-addin-pages-v1"
distribution_files=(
  README.md
  index.html
  manifest.xml
  functions.html
  functions.js
  functions.json
)
marker_path="$site_dir/$marker_name"

refuse_update() {
  printf 'Refusing to update Pages directory: %s\n' "$1" >&2
  exit 1
}

if [[ -L "$site_dir" ]]; then
  refuse_update "site path is a symlink: $site_dir"
fi

if [[ -e "$site_dir" ]]; then
  if [[ ! -d "$site_dir" ]]; then
    refuse_update "site path is not a directory: $site_dir"
  fi

  shopt -s dotglob nullglob
  existing_entries=("$site_dir"/*)
  shopt -u dotglob nullglob

  for entry in "${existing_entries[@]}"; do
    entry_name="${entry##*/}"
    case "$entry_name" in
      "$marker_name"|README.md|index.html|manifest.xml|functions.html|functions.js|functions.json) ;;
      *) refuse_update "unexpected content: $entry" ;;
    esac
    if [[ ! -f "$entry" || -L "$entry" ]]; then
      refuse_update "expected a regular file: $entry"
    fi
  done

  for file in "${distribution_files[@]}"; do
    if [[ ! -f "$site_dir/$file" || -L "$site_dir/$file" ]]; then
      refuse_update "missing or unsafe distribution file: $site_dir/$file"
    fi
  done

  if [[ -e "$marker_path" && "$(<"$marker_path")" != "$marker_content" ]]; then
    refuse_update "invalid generation marker: $marker_path"
  fi
else
  mkdir -p "$site_dir"
fi

cp "${distribution_files[@]}" "$site_dir/"
printf '%s\n' "$marker_content" > "$marker_path"
