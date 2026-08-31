#!/usr/bin/env bash
# release.sh – passenden Release-Tag erstellen und zu GitHub pushen
# GitHub Actions baut und veröffentlicht Installer und Prüfsumme automatisch.

set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "Verwendung: $0 <version> (zum Beispiel: $0 1.2.0)"
  exit 1
fi

VERSION="${1#v}"
TAG="v${VERSION}"

if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+([+-][0-9A-Za-z.-]+)?$ ]]; then
  echo "FEHLER: Ungültige Version: $VERSION"
  exit 1
fi

PROJECT_VERSION=$(sed -n 's:.*<Version>\([^<]*\)</Version>.*:\1:p' OneNoteExporter.csproj | head -n 1)
if [[ "$VERSION" != "$PROJECT_VERSION" ]]; then
  echo "FEHLER: Version $VERSION stimmt nicht mit der Projektversion $PROJECT_VERSION überein."
  exit 1
fi

if [[ -n "$(git status --porcelain)" ]]; then
  echo "FEHLER: Das Arbeitsverzeichnis enthält nicht committete Änderungen."
  exit 1
fi

git fetch origin --tags
if git rev-parse --verify --quiet "refs/tags/$TAG" >/dev/null; then
  echo "FEHLER: Tag $TAG existiert bereits."
  exit 1
fi

git tag -a "$TAG" -m "Release $TAG"
git push origin "$TAG"

echo "Tag $TAG wurde gepusht. GitHub Actions erstellt jetzt das Release."
