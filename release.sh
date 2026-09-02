#!/usr/bin/env bash
# release.sh – create the matching release tag and push it to GitHub
# GitHub Actions builds and publishes the installer and checksum automatically.

set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "Usage: $0 <version> (for example: $0 1.2.2)"
  exit 1
fi

VERSION="${1#v}"
TAG="v${VERSION}"

if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+([+-][0-9A-Za-z.-]+)?$ ]]; then
  echo "ERROR: Invalid version: $VERSION"
  exit 1
fi

PROJECT_VERSION=$(sed -n 's:.*<Version>\([^<]*\)</Version>.*:\1:p' OneNoteExporter.csproj | head -n 1)
if [[ "$VERSION" != "$PROJECT_VERSION" ]]; then
  echo "ERROR: Version $VERSION does not match project version $PROJECT_VERSION."
  exit 1
fi

if [[ -n "$(git status --porcelain)" ]]; then
  echo "ERROR: The working tree contains uncommitted changes."
  exit 1
fi

git fetch origin --tags
if git rev-parse --verify --quiet "refs/tags/$TAG" >/dev/null; then
  echo "ERROR: Tag $TAG already exists."
  exit 1
fi

git tag -a "$TAG" -m "Release $TAG"
git push origin "$TAG"

echo "Tag $TAG was pushed. GitHub Actions will now create the release."
