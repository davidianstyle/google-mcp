export function textResult(data: unknown): { content: Array<{ type: "text"; text: string }> } {
  const text = typeof data === "string" ? data : JSON.stringify(data, null, 2);
  return { content: [{ type: "text" as const, text }] };
}

export function mimeShortcut(input: string): string {
  const shortcuts: Record<string, string> = {
    document: "application/vnd.google-apps.document",
    spreadsheet: "application/vnd.google-apps.spreadsheet",
    presentation: "application/vnd.google-apps.presentation",
    folder: "application/vnd.google-apps.folder",
    form: "application/vnd.google-apps.form",
    pdf: "application/pdf",
    zip: "application/zip",
  };
  return shortcuts[input] || input;
}
