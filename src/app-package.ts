/**
 * Generate a Teams app manifest package (.zip) for sideloading the bot.
 *
 * Creates a minimal manifest.json with the bot registration and two
 * placeholder icon PNGs, then bundles them into a zip file that can
 * be uploaded via Teams → Apps → Manage your apps → Upload a custom app.
 */
import { writeFileSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { deflateSync, crc32 } from "node:zlib";
import type { InstallContext } from "@openacp/plugin-sdk";

/** Minimal solid-color 1x1 PNG (valid but tiny — Teams accepts it). */
function createPlaceholderPng(): Buffer {
  // Minimal valid PNG: 1x1 pixel, RGB, blue (#4F6BED)
  const header = Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, // PNG signature
  ]);
  const ihdr = createChunk("IHDR", Buffer.from([
    0x00, 0x00, 0x00, 0x01, // width: 1
    0x00, 0x00, 0x00, 0x01, // height: 1
    0x08,                   // bit depth: 8
    0x02,                   // color type: RGB
    0x00, 0x00, 0x00,       // compression, filter, interlace
  ]));
  // Raw pixel data: filter byte (0) + RGB
  const raw = Buffer.from([0x00, 0x4f, 0x6b, 0xed]);
  const idat = createChunk("IDAT", deflateSync(raw));
  const iend = createChunk("IEND", Buffer.alloc(0));
  return Buffer.concat([header, ihdr, idat, iend]);
}

function createChunk(type: string, data: Buffer): Buffer {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const typeBytes = Buffer.from(type, "ascii");
  const crcInput = Buffer.concat([typeBytes, data]);
  const crcVal = Buffer.alloc(4);
  crcVal.writeUInt32BE(crc32(crcInput) >>> 0);
  return Buffer.concat([len, typeBytes, data, crcVal]);
}

function createZip(files: { name: string; data: Buffer }[]): Buffer {
  // Minimal ZIP implementation — no compression (STORE method)
  const centralHeaders: Buffer[] = [];
  const localEntries: Buffer[] = [];
  let offset = 0;

  for (const file of files) {
    const nameBytes = Buffer.from(file.name, "utf-8");
    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);  // local file header signature
    localHeader.writeUInt16LE(20, 4);           // version needed
    localHeader.writeUInt16LE(0, 6);            // flags
    localHeader.writeUInt16LE(0, 8);            // compression: STORE
    localHeader.writeUInt16LE(0, 10);           // mod time
    localHeader.writeUInt16LE(0, 12);           // mod date
    // CRC-32
    const crc = crc32(file.data) >>> 0;
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(file.data.length, 18); // compressed size
    localHeader.writeUInt32LE(file.data.length, 22); // uncompressed size
    localHeader.writeUInt16LE(nameBytes.length, 26);  // filename length
    localHeader.writeUInt16LE(0, 28);                  // extra field length

    const localEntry = Buffer.concat([localHeader, nameBytes, file.data]);
    localEntries.push(localEntry);

    // Central directory header
    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);  // central directory signature
    centralHeader.writeUInt16LE(20, 4);           // version made by
    centralHeader.writeUInt16LE(20, 6);           // version needed
    centralHeader.writeUInt16LE(0, 8);            // flags
    centralHeader.writeUInt16LE(0, 10);           // compression
    centralHeader.writeUInt16LE(0, 12);           // mod time
    centralHeader.writeUInt16LE(0, 14);           // mod date
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(file.data.length, 20);
    centralHeader.writeUInt32LE(file.data.length, 24);
    centralHeader.writeUInt16LE(nameBytes.length, 28);
    centralHeader.writeUInt16LE(0, 30);           // extra field length
    centralHeader.writeUInt16LE(0, 32);           // comment length
    centralHeader.writeUInt16LE(0, 34);           // disk number start
    centralHeader.writeUInt16LE(0, 36);           // internal file attributes
    centralHeader.writeUInt32LE(0, 38);           // external file attributes
    centralHeader.writeUInt32LE(offset, 42);      // relative offset

    centralHeaders.push(Buffer.concat([centralHeader, nameBytes]));
    offset += localEntry.length;
  }

  const centralDir = Buffer.concat(centralHeaders);
  const endRecord = Buffer.alloc(22);
  endRecord.writeUInt32LE(0x06054b50, 0);               // end of central directory signature
  endRecord.writeUInt16LE(0, 4);                          // disk number
  endRecord.writeUInt16LE(0, 6);                          // central dir disk
  endRecord.writeUInt16LE(files.length, 8);               // entries on this disk
  endRecord.writeUInt16LE(files.length, 10);              // total entries
  endRecord.writeUInt32LE(centralDir.length, 12);         // central dir size
  endRecord.writeUInt32LE(offset, 16);                    // central dir offset

  return Buffer.concat([...localEntries, centralDir, endRecord]);
}

/**
 * Generate a Teams app package zip and write it to the plugin data directory.
 * Returns the path to the zip file, or null if generation fails.
 */
export async function generateTeamsAppPackage(
  botAppId: string,
  ctx: InstallContext,
): Promise<string | null> {
  const manifest = {
    $schema: "https://developer.microsoft.com/en-us/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
    manifestVersion: "1.17",
    version: "1.0.0",
    id: botAppId,
    developer: {
      name: "OpenACP",
      websiteUrl: "https://openacp.dev",
      privacyUrl: "https://openacp.dev/privacy",
      termsOfUseUrl: "https://openacp.dev/terms",
    },
    name: { short: "OpenACP Bot", full: "OpenACP Teams Bot" },
    description: {
      short: "OpenACP AI assistant for Teams",
      full: "OpenACP AI assistant that integrates with Microsoft Teams for interactive agent sessions.",
    },
    icons: { color: "color.png", outline: "outline.png" },
    accentColor: "#4F6BED",
    bots: [
      {
        botId: botAppId,
        scopes: ["personal", "team", "groupChat"],
        supportsFiles: true,
        isNotificationOnly: false,
        commandLists: [
          {
            scopes: ["personal", "team", "groupChat"],
            commands: [
              { title: "help", description: "Show available commands" },
              { title: "new", description: "Start a new agent session" },
              { title: "agents", description: "List available agents" },
              { title: "status", description: "Show session status" },
              { title: "cancel", description: "Cancel current session" },
              { title: "menu", description: "Show action menu" },
            ],
          },
        ],
      },
    ],
    permissions: ["identity", "messageTeamMembers"],
    validDomains: [],
    webApplicationInfo: {
      id: botAppId,
      resource: `api://botid-${botAppId}`,
    },
    authorization: {
      permissions: {
        resourceSpecific: [
          {
            name: "ChannelMessage.Read.Group",
            type: "Application",
          },
        ],
      },
    },
  };

  const manifestJson = Buffer.from(JSON.stringify(manifest, null, 2), "utf-8");
  const colorPng = createPlaceholderPng();
  const outlinePng = createPlaceholderPng();

  const zip = createZip([
    { name: "manifest.json", data: manifestJson },
    { name: "color.png", data: colorPng },
    { name: "outline.png", data: outlinePng },
  ]);

  if (!ctx.dataDir) {
    return null;
  }

  const outPath = join(ctx.dataDir, "openacp-bot.zip");
  mkdirSync(dirname(outPath), { recursive: true });
  writeFileSync(outPath, zip);
  return outPath;
}
