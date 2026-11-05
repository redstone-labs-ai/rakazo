import { parsePluginSource } from "@rakazo/adapters";
import { getLogger } from "@rakazo/logging";

/** The source a plugin installs as: its repository and folder, without a pinned ref. */
function installedSource(source: string): string {
  const ref = parsePluginSource(source);
  return (ref.gitUrl + (ref.subdir ? `#${ref.subdir}` : "")).toLowerCase();
}

/**
 * Plugins a deployment keeps out of the marketplace (RAKAZO_HIDDEN_MARKETPLACE_PLUGINS). Each
 * entry is a plugin source such as `owner/repo#folder` (on GitHub) or a full HTTPS source URL; a
 * trailing `*` matches every source that starts with the rest. Invalid entries are logged and
 * left out, so a typo never stops the server.
 */
export function hiddenPluginsFromEnv(
  value: string | undefined,
  warn: (message: string) => void = (message) => getLogger().warn(message),
): string[] {
  const hidden: string[] = [];
  for (const item of (value ?? "").split(/[\s,]+/).filter(Boolean)) {
    const prefix = item.endsWith("*");
    const body = prefix ? item.slice(0, -1) : item;
    const url = /^https:\/\//i.test(body) ? body : `https://github.com/${body}`;
    try {
      if (body.includes("*")) throw new Error("only a trailing * is supported");
      if (prefix) hidden.push(`${new URL(url).toString().toLowerCase()}*`);
      else hidden.push(installedSource(url));
    } catch (error) {
      warn(
        `RAKAZO_HIDDEN_MARKETPLACE_PLUGINS: ignoring "${item}": ${error instanceof Error ? error.message : "invalid"}`,
      );
    }
  }
  return hidden;
}

/** Whether a plugin source, pinned or not, is one the deployment hides. */
export function isHiddenPlugin(hidden: readonly string[] | undefined, source: string): boolean {
  if (!hidden?.length) return false;
  let installed: string;
  try {
    installed = installedSource(source);
  } catch {
    return false;
  }
  return hidden.some((entry) =>
    entry.endsWith("*") ? installed.startsWith(entry.slice(0, -1)) : installed === entry,
  );
}
