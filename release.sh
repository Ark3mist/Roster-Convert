#!/usr/bin/env bash
# Bump the app version everywhere in one step, then verify nothing was missed.
#
#   ./release.sh            next version: today's date + next number (2026.11.01-27)
#   ./release.sh 2026.12.01-01   set an exact version
#   ./release.sh --check    only verify every reference matches (exit 1 if not)
#
# The version lives in importers.js (APP_VERSION, the source of truth) and is
# repeated as ?v=… on every asset URL in index.html and style.css so browsers
# fetch fresh files after a release instead of stale cached ones.
# Uses perl (preinstalled on macOS and Linux) so it behaves the same on both.
set -euo pipefail
cd "$(dirname "$0")"

FILES=(index.html style.css importers.js)
VER_RE='[0-9]{4}\.[0-9]{2}\.[0-9]{2}-[0-9]{2,}'

current() { perl -ne "print \$1 if /const APP_VERSION = '($VER_RE)'/" importers.js; }

check() {
  local want="$1" bad=0 found
  found=$(perl -nle "print \"\$ARGV:\$.: \$1\" while /(?:\\?v=|APP_VERSION = ')($VER_RE)/g; close ARGV if eof" "${FILES[@]}")
  [ -n "$found" ] || { echo "No version references found" >&2; return 1; }
  while IFS= read -r line; do
    [ "${line##* }" = "$want" ] || { echo "MISMATCH  $line (expected $want)" >&2; bad=1; }
  done <<< "$found"
  # Local assets referenced without any ?v= would be served stale from cache.
  local bare
  bare=$(perl -nle 'print "$ARGV:$.: $&" while /(?:src|href)="(?!https?:|data:|#)[^"?]+\.(?:js|css)"|url\((?!data:)[^)?]+\.woff2\)/g; close ARGV if eof' index.html style.css)
  [ -z "$bare" ] || { echo "NO ?v=     $bare" >&2; bad=1; }
  [ $bad -eq 0 ] && echo "OK  all $(wc -l <<< "$found" | tr -d ' ') references are $want"
  return $bad
}

OLD=$(current)
[ -n "$OLD" ] || { echo "APP_VERSION not found in importers.js" >&2; exit 1; }

if [ "${1:-}" = "--check" ]; then check "$OLD"; exit; fi

if [ -n "${1:-}" ]; then
  NEW="$1"
  [[ "$NEW" =~ ^$VER_RE$ ]] || { echo "Version must look like 2026.11.01-27" >&2; exit 1; }
else
  TODAY=$(date +%Y.%m.%d)
  if [ "${OLD%-*}" = "$TODAY" ]; then N=$((10#${OLD##*-} + 1)); else N=1; fi
  NEW=$(printf '%s-%02d' "$TODAY" "$N")
fi

# Replace every version reference (not just the old value) so a copy that was
# missed in an earlier release is repaired too.
perl -pi -e "s/(\\?v=|APP_VERSION = ')$VER_RE/\${1}$NEW/g" "${FILES[@]}"
echo "$OLD -> $NEW"
check "$NEW"
