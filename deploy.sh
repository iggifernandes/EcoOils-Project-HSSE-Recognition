#!/usr/bin/env bash
# Push Code.gs to Apps Script and update the LIVE web-app deployment (same URL, new version).
# Usage: ./deploy.sh        (needs: npm i -g @google/clasp && clasp login, once)
set -euo pipefail
cd "$(dirname "$0")"
DEPLOYMENT_ID="AKfycbx2wdQ_TGsu6Kfj8xk8Tjz8vuqvzdKTuKKlLZFb214dai5J32Kw781Z1DOZ9Amd3GNV"
VER="$(head -1 Code.gs | sed -E 's#^// *##' | cut -c1-60)"
echo "Pushing: $VER"
clasp push -f
clasp deploy --deploymentId "$DEPLOYMENT_ID" --description "$VER"
echo "Done. Live at https://script.google.com/macros/s/$DEPLOYMENT_ID/exec"
