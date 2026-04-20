import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult, mimeShortcut } from "../../utils/formatting.js";

import { drive_v3 } from "googleapis";

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

  server.tool("drive_download_file", "Download file content", {
    fileId: z.string(),
    mimeType: z.string().optional().describe("Export MIME type for Google Workspace files (e.g., 'text/plain', 'application/pdf')"),
  }, async ({ fileId, mimeType }) => {
    const drive = api();
    const meta = await drive.files.get({ fileId, fields: "name,mimeType" });

    if (meta.data.mimeType?.startsWith("application/vnd.google-apps.")) {
      const exportMime = mimeType || "text/plain";
      const res = await drive.files.export({ fileId, mimeType: exportMime }, { responseType: "text" });
      return textResult({ name: meta.data.name, mimeType: exportMime, content: res.data });
    }

    const res = await drive.files.get({ fileId, alt: "media" }, { responseType: "text" });
    return textResult({ name: meta.data.name, mimeType: meta.data.mimeType, content: res.data });
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
