#!/usr/bin/env bash
# git-push.sh – stage, commit, and push the current branch
set -euo pipefail

# Read the commit message from the arguments or prompt interactively
if [[ $# -gt 0 ]]; then
  MSG="$*"
else
  read -rp "Commit message: " MSG
fi

if [[ -z "$MSG" ]]; then
  echo "ERROR: No commit message was provided."
  exit 1
fi

BRANCH=$(git rev-parse --abbrev-ref HEAD)

git add .
git commit -m "$MSG"
git push origin "$BRANCH"

echo ""
echo "Pushed: $(git rev-parse --short HEAD) to $BRANCH – $MSG"
