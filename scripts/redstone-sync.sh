#!/usr/bin/env bash
# Redstone fork sync.
#
#   main      exact mirror of elie222/rakazo main. Never commit to it.
#   redstone  what we deploy: upstream plus our patches. Upstream arrives only by merge,
#             through a sync/<ref> pull request. Never rebase or force-push redstone.
#
# Usage: scripts/redstone-sync.sh [upstream-tag-or-branch]
#   With no argument, merges the latest upstream release tag.
# Exit codes: 0 synced or nothing to do, 2 merge conflict (nothing pushed but main).
set -euo pipefail

UPSTREAM_URL=${UPSTREAM_URL:-https://github.com/elie222/rakazo.git}
UPSTREAM_REPO=${UPSTREAM_REPO:-elie222/rakazo}
REPO=${GH_REPO:-redstone-labs-ai/rakazo}
REF=${1:-}

git remote get-url upstream >/dev/null 2>&1 || git remote add upstream "$UPSTREAM_URL"
git fetch --no-tags upstream \
  "+refs/heads/main:refs/remotes/upstream/main" \
  "+refs/tags/*:refs/upstream-tags/*"
git fetch --no-tags origin "+refs/heads/redstone:refs/remotes/origin/redstone"

# Mirror. A non-fast-forward means someone committed to main; stop and look.
git push origin refs/remotes/upstream/main:refs/heads/main

if [ -z "$REF" ]; then
  REF=$(gh release view -R "$UPSTREAM_REPO" --json tagName --jq .tagName)
fi
TARGET="refs/upstream-tags/$REF"
if ! git rev-parse -q --verify "$TARGET^{commit}" >/dev/null; then
  git fetch --no-tags upstream "+refs/heads/$REF:refs/remotes/upstream/$REF"
  TARGET="refs/remotes/upstream/$REF"
fi

if git merge-base --is-ancestor "$TARGET" refs/remotes/origin/redstone; then
  echo "redstone already contains upstream $REF"
  exit 0
fi

BRANCH="sync/$(printf %s "$REF" | tr -c 'A-Za-z0-9._-' '-')"
git switch -C "$BRANCH" refs/remotes/origin/redstone
if ! git merge --no-ff --no-edit -m "Merge upstream $REF into redstone" "$TARGET"; then
  echo "Conflicts merging upstream $REF into redstone:"
  git diff --name-only --diff-filter=U
  git merge --abort
  exit 2
fi

git push --force-with-lease origin "$BRANCH"
if ! gh pr view "$BRANCH" -R "$REPO" >/dev/null 2>&1; then
  gh pr create -R "$REPO" --base redstone --head "$BRANCH" \
    --title "Merge upstream $REF" \
    --body "Merges upstream \`$REF\` into \`redstone\` with a merge commit. Redstone patches stay as they are. Check redstone-ci before merging. Merge with **Create a merge commit**, never squash or rebase."
fi
echo "Opened or updated $BRANCH -> redstone"
