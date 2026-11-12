import { crc32, deflateRawSync, deflateSync } from "node:zlib";

/** Same manifest schema as the shared Teams app (managebot infra/teams-app, v1.2.x). */
export const TEAMS_MANIFEST_SCHEMA_VERSION = "1.29";
export const AGENT_RSC_PERMISSIONS = ["ChatMessage.Read.Chat"] as const;

export interface AgentManifestDeveloper {
  name: string;
  websiteUrl: string;
  privacyUrl: string;
  termsOfUseUrl: string;
}

export interface AgentManifestInput {
  /** The agent's Entra app id: the manifest id, the bot id and the webApplicationInfo id. */
  appId: string;
  name: string;
  description: string;
  /** App package version, e.g. "1.0.0". */
  version: string;
  developer: AgentManifestDeveloper;
  accentColor?: string;
}

export interface AgentTeamsManifest {
  $schema: string;
  manifestVersion: string;
  version: string;
  id: string;
  developer: AgentManifestDeveloper;
  icons: { color: string; outline: string };
  name: { short: string; full: string };
  description: { short: string; full: string };
  accentColor: string;
  bots: Array<{
    botId: string;
    scopes: string[];
    supportsFiles: boolean;
    isNotificationOnly: boolean;
  }>;
  webApplicationInfo: { id: string; resource: string };
  authorization: {
    permissions: { resourceSpecific: Array<{ name: string; type: "Application" }> };
  };
  permissions: string[];
  validDomains: string[];
}

const SHORT_NAME_MAX = 30;
const FULL_NAME_MAX = 100;
const SHORT_DESCRIPTION_MAX = 80;
const FULL_DESCRIPTION_MAX = 4000;

function clip(value: string, max: number): string {
  const trimmed = value.trim().replace(/\s+/g, " ");
  return trimmed.length > max ? `${trimmed.slice(0, max - 1).trimEnd()}…` : trimmed;
}

/**
 * The Teams manifest for one agent's own app. `webApplicationInfo.id` must equal the bot's app
 * id: Teams only lets an app consent to RSC and install itself into a chat when they match.
 */
export function buildAgentTeamsManifest(input: AgentManifestInput): AgentTeamsManifest {
  const name = clip(input.name, SHORT_NAME_MAX) || "Agent";
  const description = input.description.trim() || `${name}, a Rakazo agent in Teams.`;
  return {
    $schema: `https://developer.microsoft.com/json-schemas/teams/v${TEAMS_MANIFEST_SCHEMA_VERSION}/MicrosoftTeams.schema.json`,
    manifestVersion: TEAMS_MANIFEST_SCHEMA_VERSION,
    version: input.version,
    id: input.appId,
    developer: input.developer,
    icons: { color: "color.png", outline: "outline.png" },
    name: { short: name, full: clip(input.name, FULL_NAME_MAX) || name },
    description: {
      short: clip(description, SHORT_DESCRIPTION_MAX),
      full: clip(description, FULL_DESCRIPTION_MAX),
    },
    accentColor: input.accentColor ?? "#ff3443",
    bots: [
      {
        botId: input.appId,
        scopes: ["personal", "team", "groupChat"],
        supportsFiles: false,
        isNotificationOnly: false,
      },
    ],
    webApplicationInfo: { id: input.appId, resource: "https://RscBasedStoreApp" },
    authorization: {
      permissions: {
        resourceSpecific: AGENT_RSC_PERMISSIONS.map((name) => ({
          name,
          type: "Application" as const,
        })),
      },
    },
    permissions: ["identity", "messageTeamMembers"],
    validDomains: [],
  };
}

export interface AgentIcons {
  /** 192x192 full-color PNG. */
  color: Uint8Array;
  /** 32x32 white-on-transparent PNG. */
  outline: Uint8Array;
}

/** The app package Graph `POST /appCatalogs/teamsApps` takes: manifest.json plus both icons. */
export function buildManifestZip(manifest: AgentTeamsManifest, icons: AgentIcons): Buffer {
  return zip([
    { name: "manifest.json", data: Buffer.from(JSON.stringify(manifest, null, 2)) },
    { name: "color.png", data: Buffer.from(icons.color) },
    { name: "outline.png", data: Buffer.from(icons.outline) },
  ]);
}

/** Icons for an agent: its avatar when it has one, else a colored initial. */
export function agentIcons(name: string, options: { avatar?: Uint8Array; color?: string } = {}) {
  const rgb = parseHexColor(options.color) ?? colorForName(name);
  return {
    color:
      options.avatar ?? initialPng(name, 192, { background: rgb, foreground: [255, 255, 255] }),
    outline: initialPng(name, 32, { foreground: [255, 255, 255] }),
  } satisfies AgentIcons;
}

type Rgb = readonly [number, number, number];

function parseHexColor(value: string | undefined): Rgb | undefined {
  const match = value?.trim().match(/^#?([0-9a-f]{6})$/i);
  if (!match?.[1]) return undefined;
  const n = Number.parseInt(match[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

const PALETTE: readonly Rgb[] = [
  [0xe5, 0x48, 0x4d],
  [0x3e, 0x63, 0xdd],
  [0x29, 0xa3, 0x83],
  [0xf7, 0x6b, 0x15],
  [0x8e, 0x4e, 0xc6],
  [0x12, 0xa5, 0x94],
];

function colorForName(name: string): Rgb {
  let hash = 0;
  for (const char of name) hash = (hash * 31 + (char.codePointAt(0) ?? 0)) >>> 0;
  return PALETTE[hash % PALETTE.length] ?? [0x3e, 0x63, 0xdd];
}

/** 5x7 bitmap glyphs, one 5-bit row per entry, for the initial on generated icons. */
const GLYPHS: Record<string, readonly number[]> = {
  A: [14, 17, 17, 31, 17, 17, 17],
  B: [30, 17, 17, 30, 17, 17, 30],
  C: [14, 17, 16, 16, 16, 17, 14],
  D: [30, 17, 17, 17, 17, 17, 30],
  E: [31, 16, 16, 30, 16, 16, 31],
  F: [31, 16, 16, 30, 16, 16, 16],
  G: [14, 17, 16, 23, 17, 17, 15],
  H: [17, 17, 17, 31, 17, 17, 17],
  I: [14, 4, 4, 4, 4, 4, 14],
  J: [7, 2, 2, 2, 2, 18, 12],
  K: [17, 18, 20, 24, 20, 18, 17],
  L: [16, 16, 16, 16, 16, 16, 31],
  M: [17, 27, 21, 21, 17, 17, 17],
  N: [17, 17, 25, 21, 19, 17, 17],
  O: [14, 17, 17, 17, 17, 17, 14],
  P: [30, 17, 17, 30, 16, 16, 16],
  Q: [14, 17, 17, 17, 21, 18, 13],
  R: [30, 17, 17, 30, 20, 18, 17],
  S: [15, 16, 16, 14, 1, 1, 30],
  T: [31, 4, 4, 4, 4, 4, 4],
  U: [17, 17, 17, 17, 17, 17, 14],
  V: [17, 17, 17, 17, 17, 10, 4],
  W: [17, 17, 17, 21, 21, 21, 10],
  X: [17, 17, 10, 4, 10, 17, 17],
  Y: [17, 17, 10, 4, 4, 4, 4],
  Z: [31, 1, 2, 4, 8, 16, 31],
  "0": [14, 17, 19, 21, 25, 17, 14],
  "1": [4, 12, 4, 4, 4, 4, 14],
  "2": [14, 17, 1, 2, 4, 8, 31],
  "3": [31, 2, 4, 2, 1, 17, 14],
  "4": [2, 6, 10, 18, 31, 2, 2],
  "5": [31, 16, 30, 1, 1, 17, 14],
  "6": [6, 8, 16, 30, 17, 17, 14],
  "7": [31, 1, 2, 4, 8, 8, 8],
  "8": [14, 17, 17, 14, 17, 17, 14],
  "9": [14, 17, 17, 15, 1, 2, 12],
  "?": [14, 17, 1, 2, 4, 0, 4],
};

/** A square RGBA PNG with the name's first letter centered. No background means transparent. */
export function initialPng(
  name: string,
  size: number,
  colors: { background?: Rgb; foreground: Rgb },
): Buffer {
  const first = name.trim().charAt(0).toUpperCase();
  const glyph = GLYPHS[first] ?? GLYPHS["?"] ?? [];
  const scale = Math.max(1, Math.floor((size * 0.55) / 7));
  const left = Math.floor((size - 5 * scale) / 2);
  const top = Math.floor((size - 7 * scale) / 2);
  const stride = size * 4 + 1;
  const raw = Buffer.alloc(stride * size);
  for (let y = 0; y < size; y++) {
    raw[y * stride] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const gx = Math.floor((x - left) / scale);
      const gy = Math.floor((y - top) / scale);
      const lit = x >= left && y >= top && gx < 5 && gy < 7 && ((glyph[gy] ?? 0) >> (4 - gx)) & 1;
      const color = lit ? colors.foreground : colors.background;
      const offset = y * stride + 1 + x * 4;
      if (color) {
        raw[offset] = color[0];
        raw[offset + 1] = color[1];
        raw[offset + 2] = color[2];
        raw[offset + 3] = 255;
      }
    }
  }
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header[8] = 8; // bit depth
  header[9] = 6; // RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", header),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

function pngChunk(type: string, data: Buffer): Buffer {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const typed = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(typed));
  return Buffer.concat([length, typed, crc]);
}

/** A minimal deflate zip writer; enough for a three-file Teams app package. */
export function zip(entries: ReadonlyArray<{ name: string; data: Buffer }>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, "utf8");
    const compressed = deflateRawSync(entry.data);
    const crc = crc32(entry.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4); // version needed
    local.writeUInt16LE(0x0800, 6); // utf-8 names
    local.writeUInt16LE(8, 8); // deflate
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(entry.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(entry.data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, name, compressed);
    centrals.push(central, name);
    offset += local.length + name.length + compressed.length;
  }
  const centralSize = centrals.reduce((sum, part) => sum + part.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, ...centrals, end]);
}
