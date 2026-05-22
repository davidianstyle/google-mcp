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

// One-shot cleanup of stale downloads, run on first use per process.
// Removes files in DOWNLOAD_CACHE_DIR older than DOWNLOAD_CACHE_TTL_MS so
// long-running MCP servers don't accumulate downloads indefinitely.
// Best-effort: any IO error is swallowed since cleanup is non-essential.
let cacheCleanupDone = false;
async function cleanupStaleDownloads(): Promise<void> {
  if (cacheCleanupDone) return;
  cacheCleanupDone = true;
  try {
    const entries = await readdir(DOWNLOAD_CACHE_DIR);
    const cutoff = Date.now() - DOWNLOAD_CACHE_TTL_MS;
    await Promise.all(
      entries.map(async (entry) => {
        const path = join(DOWNLOAD_CACHE_DIR, entry);
        try {
          const s = await stat(path);
          if (s.isFile() && s.mtimeMs < cutoff) await unlink(path);
        } catch {
          // Ignore — entry may have been removed concurrently, or a permission issue.
        }
      })
    );
  } catch {
    // Cache dir doesn't exist yet (first run) — nothing to clean.
  }
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
        q: `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
        pageSize: maxResults,
        orderBy: "name",
        fields: "files(id,name,modifiedTime)",
      });
      folders.push(...(fRes.data.files || []));
    }

    if (includeFiles) {
      const fRes = await drive.files.list({
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
    removeFromAllParents: z.boolean().optional().default(false),
  }, async ({ fileId, newParentId, removeFromAllParents }) => {
    const drive = api();
    let removeParents: string | undefined;
    if (removeFromAllParents) {
      const file = await drive.files.get({ fileId, fields: "parents" });
      removeParents = file.data.parents?.join(",");
    }
    const res = await drive.files.update({
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
    const res = await api().files.update({ fileId, requestBody: { name: newName }, fields: "id,name" });
    return textResult({ id: res.data.id, name: res.data.name });
  });

  server.tool("drive_delete_file", "Move a file to trash or permanently delete it", {
    fileId: z.string(),
    permanent: z.boolean().optional().default(false),
  }, async ({ fileId, permanent }) => {
    const drive = api();
    if (permanent) {
      await drive.files.delete({ fileId });
      return textResult({ success: true, action: "deleted", fileId });
    }
    await drive.files.update({ fileId, requestBody: { trashed: true } });
    const file = await drive.files.get({ fileId, fields: "id,name" });
    return textResult({ success: true, action: "trashed", fileId, fileName: file.data.name });
  });

  server.tool("drive_download_file", "Download a file from Drive. Default writes to a local temp path and returns { name, mimeType, path, bytes }; caller uses Read/Bash on the path. Pass returnContent=true to skip disk and return the body inline as { name, mimeType, content, bytes, encoding } — Workspace exports as utf-8 text, native files as base64. Use inline mode only for small files where disk indirection is wasteful; for binary or large files prefer the default path mode. Files older than 24h in the cache dir are cleaned up on first use per process.", {
    fileId: z.string(),
    mimeType: z.string().optional().describe("Export MIME type for Google Workspace files (e.g., 'text/markdown', 'text/plain', 'application/pdf')"),
    returnContent: z.boolean().optional().default(false).describe("If true, return the file content inline (utf-8 for Workspace, base64 for binary) instead of writing to disk. Default false: write to disk and return a path."),
  }, async ({ fileId, mimeType, returnContent }) => {
    const drive = api();
    const meta = await drive.files.get({ supportsAllDrives: true, fileId, fields: "name,mimeType" });

    const isWorkspace = !!meta.data.mimeType?.startsWith("application/vnd.google-apps.");
    const outMime = isWorkspace
      ? mimeType || "text/plain"
      : meta.data.mimeType || "application/octet-stream";

    // Inline mode: return body in the response, never touch disk.
    if (returnContent) {
      if (isWorkspace) {
        const res = await drive.files.export(
          { fileId, mimeType: outMime },
          { responseType: "arraybuffer" }
        );
        const body = Buffer.from(res.data as ArrayBuffer);
        return textResult({
          name: meta.data.name,
          mimeType: outMime,
          content: body.toString("utf-8"),
          bytes: body.byteLength,
          encoding: "utf-8",
        });
      }
      const res = await drive.files.get(
        { supportsAllDrives: true, fileId, alt: "media" },
        { responseType: "arraybuffer" }
      );
      const body = Buffer.from(res.data as ArrayBuffer);
      return textResult({
        name: meta.data.name,
        mimeType: outMime,
        content: body.toString("base64"),
        bytes: body.byteLength,
        encoding: "base64",
      });
    }

    // Disk mode (default): write to a temp path and return the path.
    await cleanupStaleDownloads();
    await mkdir(DOWNLOAD_CACHE_DIR, { recursive: true });

    const ext = extensionFor(outMime);
    const safeName = safeFileName(meta.data.name || fileId);
    const suffix = randomBytes(4).toString("hex");
    const path = join(DOWNLOAD_CACHE_DIR, `${safeName}-${Date.now()}-${suffix}.${ext}`);

    let bytes: number;
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
    const meta = await drive.files.get({ fileId: folderId, fields: "id,name,modifiedTime,createdTime,owners,webViewLink" });
    const children = await drive.files.list({
      q: `'${folderId}' in parents and trashed = false`,
      fields: "files(id)",
      pageSize: 1000,
    });
    return textResult({ ...meta.data, childCount: children.data.files?.length || 0 });
  });
}
