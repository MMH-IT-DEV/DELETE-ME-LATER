#!/usr/bin/env bash
# =============================================================================
# deploy.sh — Push all GAS projects to Google Apps Script via clasp
#
# ONE-TIME SETUP (do this once, never again):
#   npm install -g @google/clasp
#   clasp login
#   cd src && clasp pull    ← pulls appsscript.json from remote (needed once)
#   cd ../engin-src && fill in scriptId in engin-src/.clasp.json
#   cd ../engin-src && clasp pull   ← pulls engin appsscript.json
#
# USAGE (every time you want to deploy):
#   bash scripts/deploy.sh
# =============================================================================

set -e  # Exit on any error

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"

echo ""
echo "======================================================"
echo "  WASP-Katana Deploy"
echo "======================================================"

# ── 1. Main project: src/ ─────────────────────────────────────────────────────
echo ""
echo "▶  Pushing main project (src/)..."
cd "$REPO_ROOT/src"
clasp push --force
echo "✓  Main project pushed"

# ── 2. Engin project: engin-src/ ─────────────────────────────────────────────
echo ""
echo "▶  Pushing engin project (engin-src/)..."

ENGIN_SRC="$REPO_ROOT/engin-src"
TMP="$REPO_ROOT/.deploy-tmp-engin"

# Clean and recreate temp dir
rm -rf "$TMP"
mkdir -p "$TMP"

# Copy .clasp.json (contains the engin scriptId)
cp "$ENGIN_SRC/.clasp.json" "$TMP/.clasp.json"

# Copy appsscript.json if it exists
if [ -f "$ENGIN_SRC/appsscript.json" ]; then
  cp "$ENGIN_SRC/appsscript.json" "$TMP/appsscript.json"
fi

# Rename .txt → .gs so clasp recognises them as Apps Script files
for f in "$ENGIN_SRC"/*.txt; do
  [ -f "$f" ] || continue
  base=$(basename "$f" .txt)
  cp "$f" "$TMP/${base}.gs"
done

# Push from temp dir, then clean up
cd "$TMP"
clasp push --force
cd "$REPO_ROOT"
rm -rf "$TMP"

echo "✓  Engin project pushed"

echo ""
echo "======================================================"
echo "  ✅  All projects deployed successfully"
echo "======================================================"
echo ""
