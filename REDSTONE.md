# Redstone fork of Rakazo

This fork runs Bot at bot.managebot.ai. Upstream is [elie222/rakazo](https://github.com/elie222/rakazo). The deploy procedure lives in the Managebot repo, `docs/runbooks/bot-managebot-ai.md`.

## Branches

| Branch | What it is | Rule |
| --- | --- | --- |
| `main` | Exact mirror of upstream `main` | Never commit. Only `scripts/redstone-sync.sh` pushes it. |
| `redstone` | Default branch. What we build and deploy | Upstream arrives only by merge commit. Never rebase, squash or force-push. |
| `feat/*` | One change each, based on upstream `main` | Shaped as an upstream pull request. Merged into `redstone` with a merge commit. |
| `sync/<ref>` | Upstream release merged onto `redstone` | Opened by the sync script. Merge with "Create a merge commit". |

## Keep the patch small

Upstream ships several releases a week. Every line we change in an upstream file is a line that can conflict.

- Prefer new files. Redstone-only files are this one, `scripts/redstone-sync.sh` and `.github/workflows/redstone-*.yml`.
- A feature that upstream could use goes on a `feat/*` branch in upstream style (their lint, tests, docs and CHANGELOG voice) with no Redstone names, hosts or tenant ids. Offer it upstream. When upstream merges it, the next sync makes our copy disappear.
- No Redstone configuration in code. Hosts, tenants and secrets are environment variables set at deploy time.

## Current patches

| Branch | Upstream PR | Notes |
| --- | --- | --- |
| `feat/teams-messaging` | not opened yet | Microsoft Teams chat app on the Chat SDK `@chat-adapter/teams`, with group chats and channels. |
| `feat/group-run-approvals` | not opened yet | A connector call that can change data in a group chat run asks the bot owner every time, even in Never ask and past Always allow. Fork #2. |
| `fix/group-approval-reply` | not opened yet | A group run resumed by an approval card still posts its reply to the group; an edited instruction change stays out. Fork #3. |
| `feat/group-runs-read-only` | not opened yet | `RAKAZO_GROUP_RUNS_READ_ONLY=true` refuses connector writes in group chat runs instead of asking the owner, audited as `group_read_only`. Built-in tools are unchanged. Fork #5. |
| `feat/default-user-presets` | not opened yet | `RAKAZO_DEFAULT_TIMEZONE`, `_LOCALE`, `_BOT_INSTRUCTIONS` and `_ROUTINES` give new accounts a starting setup: first bot instructions and paused routines. Bad values are logged and ignored. Fork #6. |
| `feat/marketplace-hidden-plugins` | not opened yet | `RAKAZO_HIDDEN_MARKETPLACE_PLUGINS` leaves plugin sources out of the marketplace and refuses their preview and install; installed ones stay. Used to hide Cursor's Microsoft plugins, which cannot sign in here, in favor of Managebot's own Outlook, Teams and SharePoint connectors. Fork #8. |
| `feat/server-default-mcp-servers` | not opened yet | `RAKAZO_DEFAULT_MCP_SERVERS` adds MCP servers to each person at sign-in, once per person (a deleted one stays deleted), optionally only for an `audience` of roles; `oidcRoles` match the SSO ID token's `roles` claim. Replaces the proxy-injected browser bootstrap, which uses the same slugs. Fork #9. |
| `ci/expo-check-offline` | Redstone only | `redstone-ci.yml` passes `EXPO_OFFLINE` so the mobile check uses the installed Expo SDK. Fork #4. |

## Sync upstream

```sh
scripts/redstone-sync.sh            # latest upstream release
scripts/redstone-sync.sh v0.9.0     # a given tag
scripts/redstone-sync.sh main       # upstream main
```

It mirrors `main`, merges the upstream ref into a `sync/<ref>` branch from `redstone` and opens a pull request into `redstone`. On a conflict it pushes nothing but `main`, lists the files and exits 2: merge by hand on a `sync/<ref>` branch and open the pull request yourself.

`redstone-sync.yml` runs the same script daily when the `REDSTONE_SYNC_TOKEN` secret exists (contents, pull requests and workflows write on this repository only). `GITHUB_TOKEN` cannot push commits that touch `.github/workflows`, which upstream merges often do.

`gh pr create` in a fork defaults to the parent repository. Always pass `-R redstone-labs-ai/rakazo --base redstone`.

## Workflows

Upstream workflows are disabled in this fork except `publish-server-image.yml`, which `redstone-image.yml` calls. Upstream `ci.yml` deploys upstream production from `main`; `redstone-ci.yml` runs the same lint, typecheck and unit suite for `redstone`.

| Workflow | Trigger | Result |
| --- | --- | --- |
| `redstone-ci.yml` | push and pull request to `redstone` | Lint, typecheck, unit tests |
| `redstone-image.yml` | push to `redstone` | `ghcr.io/redstone-labs-ai/rakazo/{app,updater,computer}:sha-<commit>` |
| `redstone-sync.yml` | daily, manual | Mirror and sync pull request |

Deploys pin image digests, never a moving tag.
