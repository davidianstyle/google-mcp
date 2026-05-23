import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult, mimeShortcut } from "../../utils/formatting.js";

import { drive_v3 } from "googleapis";
import { mkdir, writeFile, readdir, stat, unlink } from "node:fs/promises";
import { createWriteStream } from "node:fs";
import { pipeline } from "node:stream/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { randomBytes } from "node:crypto";
import type { Readable } from "node:stream";

const DOWNLOAD_CACHE_DIR = join(tmpdir(), "google-mcp");
const DOWNLOAD_CACHE_TTL_MS = 24 * 60 * 60 * 1000;
// Hard cap for inline (returnContent=true) downloads to keep tool responses
// from blowing past MCP/LLM context budgets and to avoid OOM. Callers that
// need larger payloads should use the default disk mode.
const MAX_INLINE_BYTES = 10 * 1024 * 1024;
const FOLDER_MIME = "application/vnd.google-apps.folder";

// Output MIME types we treat as text in inline mode. Everything else is
// returned as base64 to avoid corrupting binary bytes through UTF-8.
// Workspace files can be exported as text (markdown/csv/html) OR as
// binary (PDF/DOCX/XLSX), so the source format doesn't determine
// encoding — the OUTPUT MIME does.
//
// Covers text/* (markdown, plain, csv, html, xml, etc.), structured
// data formats commonly served as text (json, xml, javascript, sql,
// yaml), and any structured suffix (+json, +xml, +yaml). When unsure,
// we fall through to base64 — a base64-wrapped text payload is
// recoverable, but a UTF-8-decoded binary is corrupted.
const TEXT_MIMES = new Set([
  "application/json",
  "application/xml",
  "application/javascript",
  "application/ecmascript",
  "application/sql",
  "application/yaml",
  "application/x-yaml",
  "application/x-sh",
  "application/x-www-form-urlencoded",
]);
function isTextMime(mime: string): boolean {
  if (mime.startsWith("text/")) return true;
  if (TEXT_MIMES.has(mime)) return true;
  // Structured suffixes like application/atom+xml, application/ld+json.
  return /\+(?:json|xml|yaml)$/.test(mime);
}

// Sweeps the cache dir of files older than DOWNLOAD_CACHE_TTL_MS.
// Used both on first download per process (to handle a freshly started
// server inheriting a stale cache dir) and periodically thereafter (to
// keep long-running processes from accumulating downloads).
async function sweepStaleDownloads(): Promise<void> {
  const entries = await readdir(DOWNLOAD_CACHE_DIR);
  const cutoff = Date.now() - DOWNLOAD_CACHE_TTL_MS;
  // Sequential sweep keeps file-descriptor and IO pressure bounded even
  // when the cache dir has accumulated many entries.
  for (const entry of entries) {
    const path = join(DOWNLOAD_CACHE_DIR, entry);
    try {
      const s = await stat(path);
      if (s.isFile() && s.mtimeMs < cutoff) await unlink(path);
    } catch {
      // Entry was removed concurrently or otherwise inaccessible. Ignore.
    }
  }
}

// One-shot initialization of the cache dir: ensures the directory exists
// and runs an initial sweep. Subsequent downloads re-run the sweep at most
// once per DOWNLOAD_CACHE_TTL_MS interval (lastSweepAt), so a long-running
// process doesn't get stuck on the first sweep forever. If init fails, the
// cached promise is cleared so the next call can retry. Per-entry cleanup
// errors are swallowed inside sweepStaleDownloads, but a top-level
// mkdir/readdir failure rejects and resets so callers see the real error
// and the next call gets a fresh attempt.
let cacheInitPromise: Promise<void> | null = null;
let lastSweepAt = 0;
function ensureCacheInitialized(): Promise<void> {
  if (cacheInitPromise) return cacheInitPromise;
  cacheInitPromise = (async () => {
    await mkdir(DOWNLOAD_CACHE_DIR, { recursive: true });
    await sweepStaleDownloads();
    lastSweepAt = Date.now();
  })().catch((err) => {
    cacheInitPromise = null;
    throw err;
  });
  return cacheInitPromise;
}

// Run an additional sweep if more than DOWNLOAD_CACHE_TTL_MS has passed
// since the last one. Fire-and-forget so it never blocks a download.
function maybePeriodicSweep(): void {
  if (Date.now() - lastSweepAt < DOWNLOAD_CACHE_TTL_MS) return;
  lastSweepAt = Date.now(); // optimistic: prevents concurrent re-entry
  sweepStaleDownloads().catch(() => {
    // Best-effort; an error here is non-fatal for downloads.
  });
}

const MIME_TO_EXT: Record<string, string> = {
  "text/markdown": "md",
  "text/plain": "txt",
  "text/csv": "csv",
  "text/html": "html",
  "application/pdf": "pdf",
  "application/json": "json",
  "application/zip": "zip",
  "application/rtf": "rtf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
};

function extensionFor(mime: string | null | undefined): string {
  if (!mime) return "bin";
  return MIME_TO_EXT[mime] || mime.split("/").pop()?.split(".").pop() || "bin";
}

function safeFileName(name: string): string {
  return name.replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 80);
}

function formatFile(f: drive_v3.Schema$File): Record<string, unknown> {
  return {
    id: f.id,
    name: f.name,
    mimeType: f.mimeType,
    size: f.size,
    modifiedTime: f.modifiedTime,
    createdTime: f.createdTime,
    owner: f.owners?.[0]?.displayName,
    url: f.webViewLink,
  };
}

export function registerDriveTools(server: McpServer, ctx: ServiceContext): void {
  const api = () => google.drive({ version: "v3", auth: ctx.auth });

  server.tool("drive_list_files", "List files in Drive with optional filtering", {
    folderId: z.string().optional().describe("Folder ID to list. Use 'root' for top-level."),
    mimeType: z.string().optional().describe("Filter by MIME type. Shortcuts: document, spreadsheet, presentation, folder, pdf, zip"),
    maxResults: z.number().optional().default(20),
    orderBy: z.enum(["name", "modifiedTime", "createdTime", "quotaBytesUsed"]).optional().default("modifiedTime"),
    sortDirection: z.enum(["asc", "desc"]).optional().default("desc"),
    ownedByMe: z.boolean().optional(),
    modifiedAfter: z.string().optional().describe("ISO 8601 date filter"),
  }, async ({ folderId, mimeType, maxResults, orderBy, sortDirection, ownedByMe, modifiedAfter }) => {
    const qParts: string[] = ["trashed = false"];
    if (folderId) qParts.push(`'${folderId}' in parents`);
    if (mimeType) qParts.push(`mimeType = '${mimeShortcut(mimeType)}'`);
    if (ownedByMe) qParts.push("'me' in owners");
    if (modifiedAfter) qParts.push(`modifiedTime > '${modifiedAfter}'`);

    const order = `${orderBy} ${sortDirection === "asc" ? "" : "desc"}`.trim();
    const res = await api().files.list({
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      q: qParts.join(" and "),
      pageSize: maxResults,
      orderBy: order,
      fields: "files(id,name,mimeType,size,modifiedTime,createdTime,owners,webViewLink)",
    });
    return textResult(res.data.files?.map(formatFile) || []);
  });

  server.tool("drive_list_folder_contents", "List files and subfolders in a Drive folder", {
    folderId: z.string().describe("Folder ID. Use 'root' for top-level."),
    includeFiles: z.boolean().optional().default(true),
    includeSubfolders: z.boolean().optional().default(true),
    maxResults: z.number().optional().default(50),
  }, async ({ folderId, includeFiles, includeSubfolders, maxResults }) => {
    const drive = api();
    const folders: unknown[] = [];
    const files: unknown[] = [];

    if (includeSubfolders) {
      const fRes = await drive.files.list({
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        q: `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
        pageSize: maxResults,
        orderBy: "name",
        fields: "files(id,name,modifiedTime)",
      });
      folders.push(...(fRes.data.files || []));
    }

    if (includeFiles) {
      const fRes = await drive.files.list({
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        q: `'${folderId}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false`,
        pageSize: maxResults,
        orderBy: "name",
        fields: "files(id,name,mimeType,modifiedTime)",
      });
      files.push(...(fRes.data.files || []));
    }

    return textResult({ folders, files });
  });

  server.tool("drive_search_files", "Search Drive by name or content", {
    query: z.string().describe("Search term"),
    searchIn: z.enum(["name", "content", "both"]).optional().default("both"),
    folderId: z.string().optional(),
    mimeType: z.string().optional(),
    maxResults: z.number().optional().default(10),
    orderBy: z.enum(["name", "modifiedTime", "createdTime"]).optional().default("modifiedTime"),
    sortDirection: z.enum(["asc", "desc"]).optional().default("desc"),
    modifiedAfter: z.string().optional(),
    pageToken: z.string().optional(),
  }, async ({ query, searchIn, folderId, mimeType, maxResults, orderBy, sortDirection, modifiedAfter, pageToken }) => {
    const qParts: string[] = ["trashed = false"];
    if (searchIn === "name") qParts.push(`name contains '${query}'`);
    else if (searchIn === "content") qParts.push(`fullText contains '${query}'`);
    else qParts.push(`(name contains '${query}' or fullText contains '${query}')`);
    if (folderId) qParts.push(`'${folderId}' in parents`);
    if (mimeType) qParts.push(`mimeType = '${mimeShortcut(mimeType)}'`);
    if (modifiedAfter) qParts.push(`modifiedTime > '${modifiedAfter}'`);

    const res = await api().files.list({
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      q: qParts.join(" and "),
      pageSize: maxResults,
      pageToken,
      orderBy: `${orderBy} ${sortDirection === "asc" ? "" : "desc"}`.trim(),
      fields: "nextPageToken,files(id,name,mimeType,size,modifiedTime,createdTime,owners,webViewLink)",
    });

    return textResult({
      files: res.data.files?.map(formatFile) || [],
      total: res.data.files?.length || 0,
      hasMore: !!res.data.nextPageToken,
      nextPageToken: res.data.nextPageToken,
    });
  });

  server.tool("drive_move_file", "Move a file to a different folder", {
    fileId: z.string(),
    newParentId: z.string().describe("Destination folder ID. Use 'root' for top-level."),
    removeFromAllParents: z.boolean().optional().default(false).describe("Remove the file from all current parents. Always treated as true for Shared Drive items, which can only have one parent."),
  }, async ({ fileId, newParentId, removeFromAllParents }) => {
    const drive = api();
    // Always fetch parents + driveId. Shared Drive items can only have one
    // parent, so adding a new parent without removing the existing one
    // will fail. Force-remove existing parents on Shared Drive items.
    const file = await drive.files.get({
      supportsAllDrives: true,
      fileId,
      fields: "parents,driveId",
    });
    const isSharedDrive = !!file.data.driveId;
    let removeParents: string | undefined;
    if (removeFromAllParents || isSharedDrive) {
      removeParents = file.data.parents?.join(",");
    }
    const res = await drive.files.update({
      supportsAllDrives: true,
      fileId,
      addParents: newParentId,
      removeParents,
      fields: "id,name,parents",
    });
    return textResult({ id: res.data.id, name: res.data.name, message: `Moved to ${newParentId}` });
  });

  server.tool("drive_copy_file", "Copy a file", {
    fileId: z.string(),
    name: z.string().optional().describe("Name for the copy"),
    parentId: z.string().optional().describe("Destination folder ID"),
  }, async ({ fileId, name, parentId }) => {
    const res = await api().files.copy({
      supportsAllDrives: true,
      fileId,
      requestBody: { name, parents: parentId ? [parentId] : undefined },
      fields: "id,name,webViewLink",
    });
    return textResult({ id: res.data.id, name: res.data.name, url: res.data.webViewLink });
  });

  server.tool("drive_rename_file", "Rename a file", {
    fileId: z.string(),
    newName: z.string(),
  }, async ({ fileId, newName }) => {
    const res = await api().files.update({ supportsAllDrives: true, fileId, requestBody: { name: newName }, fields: "id,name" });
    return textResult({ id: res.data.id, name: res.data.name });
  });

  server.tool("drive_delete_file", "Move a file to trash or permanently delete it", {
    fileId: z.string(),
    permanent: z.boolean().optional().default(false),
  }, async ({ fileId, permanent }) => {
    const drive = api();
    if (permanent) {
      await drive.files.delete({ supportsAllDrives: true, fileId });
      return textResult({ success: true, action: "deleted", fileId });
    }
    await drive.files.update({ supportsAllDrives: true, fileId, requestBody: { trashed: true } });
    const file = await drive.files.get({ supportsAllDrives: true, fileId, fields: "id,name" });
    return textResult({ success: true, action: "trashed", fileId, fileName: file.data.name });
  });

  server.tool("drive_download_file", `Download a file from Drive. Default writes to a local temp path and returns { name, mimeType, path, bytes }; caller uses Read/Bash on the path. Pass returnContent=true to skip disk and return the body inline as { name, mimeType, content, bytes, encoding } where encoding is 'utf-8' for text MIME types (text/* and application/json) and 'base64' for everything else — including Workspace exports to binary formats like PDF/DOCX/XLSX. Inline mode is capped at ${MAX_INLINE_BYTES} bytes; larger files must use disk mode. Folders cannot be downloaded; use drive_list_folder_contents to enumerate them. Files older than 24h in the cache dir are cleaned up on first use per process.`, {
    fileId: z.string(),
    mimeType: z.string().optional().describe("Export MIME type for Google Workspace files (e.g., 'text/markdown', 'text/plain', 'application/pdf')"),
    returnContent: z.boolean().optional().default(false).describe("If true, return the file content inline (utf-8 for text MIME, base64 otherwise) instead of writing to disk. Default false: write to disk and return a path."),
  }, async ({ fileId, mimeType, returnContent }) => {
    const drive = api();
    const meta = await drive.files.get({
      supportsAllDrives: true,
      fileId,
      fields: "name,mimeType,size",
    });

    // Folders aren't downloadable — surface a clear error instead of letting
    // the API call below fail opaquely. Callers should use
    // drive_list_folder_contents to enumerate folder children.
    if (meta.data.mimeType === FOLDER_MIME) {
      throw new Error(
        `Cannot download a folder (fileId=${fileId}, name=${meta.data.name}). Use drive_list_folder_contents to list its contents.`
      );
    }

    const isWorkspace = !!meta.data.mimeType?.startsWith("application/vnd.google-apps.");
    const outMime = isWorkspace
      ? mimeType || "text/plain"
      : meta.data.mimeType || "application/octet-stream";

    // Inline mode: return body in the response, never touch disk.
    // Encoding choice keys off the OUTPUT MIME, not the source format —
    // Workspace docs can be exported to binary (PDF/DOCX) which MUST be
    // base64 to preserve bytes.
    if (returnContent) {
      // Pre-flight size guard for native files (meta.size is reliable for
      // them). Workspace exports don't have a meta.size — they're bounded
      // by Google's own export caps (~10MB), and we re-check post-download
      // below as a belt-and-braces guard.
      const reportedSize = meta.data.size ? Number(meta.data.size) : undefined;
      if (!isWorkspace && reportedSize !== undefined && reportedSize > MAX_INLINE_BYTES) {
        throw new Error(
          `File too large for inline mode: ${reportedSize} bytes > ${MAX_INLINE_BYTES} byte cap. Use the default disk mode (omit returnContent) for files this size.`
        );
      }
      const arraybufRes = isWorkspace
        ? await drive.files.export(
            { fileId, mimeType: outMime },
            { responseType: "arraybuffer" }
          )
        : await drive.files.get(
            { supportsAllDrives: true, fileId, alt: "media" },
            { responseType: "arraybuffer" }
          );
      const body = Buffer.from(arraybufRes.data as ArrayBuffer);
      if (body.byteLength > MAX_INLINE_BYTES) {
        throw new Error(
          `Downloaded body too large for inline mode: ${body.byteLength} bytes > ${MAX_INLINE_BYTES} byte cap. Use the default disk mode (omit returnContent).`
        );
      }
      const encoding: "utf-8" | "base64" = isTextMime(outMime) ? "utf-8" : "base64";
      return textResult({
        name: meta.data.name,
        mimeType: outMime,
        content: body.toString(encoding),
        bytes: body.byteLength,
        encoding,
      });
    }

    // Disk mode (default): ensure cache dir is initialized (with stale-file
    // sweep), then write to a temp path and return it. Trigger a periodic
    // re-sweep in the background if it's been more than DOWNLOAD_CACHE_TTL_MS
    // since the last one — handles long-running MCP processes.
    await ensureCacheInitialized();
    maybePeriodicSweep();

    const ext = extensionFor(outMime);
    const safeName = safeFileName(meta.data.name || fileId);
    const suffix = randomBytes(4).toString("hex");
    const path = join(DOWNLOAD_CACHE_DIR, `${safeName}-${Date.now()}-${suffix}.${ext}`);

    let bytes: number;
    try {
      if (isWorkspace) {
        // Workspace exports are bounded by Google (~10MB Doc export cap); buffer is fine.
        const res = await drive.files.export(
          { fileId, mimeType: outMime },
          { responseType: "arraybuffer" }
        );
        const body = Buffer.from(res.data as ArrayBuffer);
        await writeFile(path, body);
        bytes = body.byteLength;
      } else {
        // Native files can be arbitrarily large; stream directly to disk.
        const res = await drive.files.get(
          { supportsAllDrives: true, fileId, alt: "media" },
          { responseType: "stream" }
        );
        await pipeline(res.data as Readable, createWriteStream(path));
        bytes = (await stat(path)).size;
      }
    } catch (err) {
      // Download failed mid-flight (network error, stream interruption,
      // export rejection). Remove any partial file before propagating so
      // callers don't see a path that points at corrupt/incomplete bytes.
      await unlink(path).catch(() => {});
      throw err;
    }

    return textResult({
      name: meta.data.name,
      mimeType: outMime,
      path,
      bytes,
    });
  });

  server.tool("drive_create_folder", "Create a new folder", {
    name: z.string(),
    parentId: z.string().optional().describe("Parent folder ID"),
  }, async ({ name, parentId }) => {
    const res = await api().files.create({
      supportsAllDrives: true,
      requestBody: {
        name,
        mimeType: "application/vnd.google-apps.folder",
        parents: parentId ? [parentId] : undefined,
      },
      fields: "id,name,webViewLink",
    });
    return textResult({ id: res.data.id, name: res.data.name, url: res.data.webViewLink });
  });

  server.tool("drive_get_folder_info", "Get folder metadata and size", {
    folderId: z.string(),
  }, async ({ folderId }) => {
    const drive = api();
    const meta = await drive.files.get({ supportsAllDrives: true, fileId: folderId, fields: "id,name,modifiedTime,createdTime,owners,webViewLink" });
    const children = await drive.files.list({
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      q: `'${folderId}' in parents and trashed = false`,
      fields: "files(id)",
      pageSize: 1000,
    });
    return textResult({ ...meta.data, childCount: children.data.files?.length || 0 });
  });
}
