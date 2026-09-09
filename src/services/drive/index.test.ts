import { describe, expect, it } from "vitest";
import { drive_v3 } from "googleapis";
import { formatFileForList, resolveUploadMimeTypes } from "./index.js";

describe("formatFileForList", () => {
  const file: drive_v3.Schema$File = {
    id: "file1",
    name: "Report.docx",
    mimeType: "application/vnd.google-apps.document",
    size: "12345",
    modifiedTime: "2026-07-01T00:00:00Z",
    createdTime: "2026-01-01T00:00:00Z",
    parents: ["folder1"],
    owners: [{ displayName: "Alice" }],
    webViewLink: "https://docs.google.com/document/d/file1/edit",
  };

  it("prunes to id/name/mimeType/modifiedTime/size/parents", () => {
    expect(formatFileForList(file)).toEqual({
      id: "file1",
      name: "Report.docx",
      mimeType: "application/vnd.google-apps.document",
      modifiedTime: "2026-07-01T00:00:00Z",
      size: "12345",
      parents: ["folder1"],
    });
  });

  it("drops owner, webViewLink, and createdTime", () => {
    const result = formatFileForList(file) as Record<string, unknown>;
    expect("owner" in result).toBe(false);
    expect("owners" in result).toBe(false);
    expect("url" in result).toBe(false);
    expect("webViewLink" in result).toBe(false);
    expect("createdTime" in result).toBe(false);
  });

  it("handles a file with no parents (e.g. a Shared Drive root item)", () => {
    const { parents: _parents, ...noParents } = file;
    expect(formatFileForList(noParents).parents).toBeUndefined();
  });
});

describe("resolveUploadMimeTypes", () => {
  it("guesses media type from the extension with no conversion by default", () => {
    expect(resolveUploadMimeTypes("/tmp/notes.md")).toEqual({
      mediaMimeType: "text/markdown",
      targetMimeType: undefined,
    });
  });

  it("uses an explicit non-google mime_type as the media type", () => {
    expect(resolveUploadMimeTypes("/tmp/data.bin", "application/x-custom")).toEqual({
      mediaMimeType: "application/x-custom",
      targetMimeType: undefined,
    });
  });

  it("maps convert_to to a google-apps target while keeping the real media type", () => {
    expect(resolveUploadMimeTypes("/tmp/report.docx", undefined, "document")).toEqual({
      mediaMimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      targetMimeType: "application/vnd.google-apps.document",
    });
    expect(resolveUploadMimeTypes("/tmp/deck.pptx", undefined, "presentation").targetMimeType)
      .toBe("application/vnd.google-apps.presentation");
    expect(resolveUploadMimeTypes("/tmp/rows.csv", undefined, "spreadsheet")).toEqual({
      mediaMimeType: "text/csv",
      targetMimeType: "application/vnd.google-apps.spreadsheet",
    });
  });

  it("treats a google-apps mime_type as convert_to and guesses the media type", () => {
    expect(resolveUploadMimeTypes("/tmp/notes.md", "application/vnd.google-apps.document")).toEqual({
      mediaMimeType: "text/markdown",
      targetMimeType: "application/vnd.google-apps.document",
    });
  });

  it("lets convert_to win over a conflicting google-apps mime_type", () => {
    expect(resolveUploadMimeTypes("/tmp/x.csv", "application/vnd.google-apps.document", "spreadsheet").targetMimeType)
      .toBe("application/vnd.google-apps.spreadsheet");
  });

  it("falls back to octet-stream for unknown extensions", () => {
    expect(resolveUploadMimeTypes("/tmp/blob.xyz").mediaMimeType).toBe("application/octet-stream");
  });
});
