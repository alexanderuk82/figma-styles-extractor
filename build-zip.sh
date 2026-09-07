#!/bin/sh
# Build the distributable zip for the DS Styles Extractor plugin.
#
#   ./build-zip.sh            → ../ds-styles-extractor-v<version>.zip (version read from manifest.json)
#
# The repo only holds placeholders for the shared Figma API token used by the Audit tab.
# This script injects the real token at packaging time from ../.key-figma.md, which must
# live OUTSIDE the repository and never be committed. File format:
#   line 1: figd_...                 (the personal access token)
#   line 2: expires: YYYY-MM-DD      (optional — powers the expiry reminder in the Audit tab)
set -eu
cd "$(dirname "$0")"
VERSION=$(sed -n 's/.*"name": *"DS Styles Extractor V\.\([0-9.]*\)".*/\1/p' manifest.json)
[ -n "$VERSION" ] || { echo "Could not read version from manifest.json"; exit 1; }
KEYFILE="../.key-figma.md"
TOKEN=""; EXPIRES=""
if [ -f "$KEYFILE" ]; then
  TOKEN=$(sed -n '1p' "$KEYFILE" | tr -d '[:space:]')
  EXPIRES=$(sed -n 's/^expires:[[:space:]]*\([0-9-]*\).*/\1/p' "$KEYFILE" | head -1)
else
  echo "WARNING: $KEYFILE not found — the zip will ship WITHOUT an API token (Graveyard check off)."
fi
# dist/ must mirror the root files exactly (with placeholders — never the real token)
cp code.js ui.html manifest.json dist/
TMP=$(mktemp -d)
cp code.js manifest.json "$TMP/"
sed -e "s|__FIGMA_TOKEN__|$TOKEN|g" -e "s|__FIGMA_TOKEN_EXPIRES__|$EXPIRES|g" ui.html > "$TMP/ui.html"
OUT="../ds-styles-extractor-v$VERSION.zip"
rm -f "$OUT"
(cd "$TMP" && zip -q -j "$OLDPWD/$OUT" manifest.json code.js ui.html)
# local/ = the same injected files, for testing in Figma via "Import plugin from manifest"
# (gitignored — it contains the real token)
rm -rf local && mkdir local && cp "$TMP/manifest.json" "$TMP/code.js" "$TMP/ui.html" local/
rm -rf "$TMP"
echo "Local test copy: local/manifest.json"
echo "Built $OUT (v$VERSION, token: $([ -n "$TOKEN" ] && echo injected || echo none), expires: ${EXPIRES:-not set})"
