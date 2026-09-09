/**
 * What Drive's markdown -> Google Docs converter does and does not render.
 * Shared by every docs_* tool that accepts markdown so the guidance stays
 * consistent.
 */
export const MARKDOWN_SUPPORT_NOTE =
  "Drive's markdown converter SUPPORTS: headings (#..######), bold, italic, strikethrough, links, nested bullet and numbered lists, and pipe tables. " +
  "It does NOT support fenced code blocks, blockquotes, horizontal rules, images, checkboxes, or footnotes — those come through as plain text, so avoid them or add them afterwards with docs_insert_image / docs_apply_text_style.";

/**
 * Pure helpers for the Drive-native markdown round-trip used by the docs
 * markdown tools. The heavy lifting (markdown <-> native Docs conversion) is
 * done by Drive's own export/import; these helpers only handle the string
 * plumbing around it.
 */

/**
 * Joins an existing markdown export with new markdown to append. Ensures the
 * two blocks are separated by a blank line so the appended content starts a
 * fresh paragraph/heading rather than merging into the last line of the
 * existing document. Drive's markdown export normally ends with a trailing
 * newline; we normalize trailing whitespace before inserting the separator so
 * the result is deterministic regardless of how many trailing newlines the
 * export carried.
 */
export function concatMarkdownForAppend(existingMarkdown: string, newMarkdown: string): string {
  const existing = existingMarkdown.replace(/\s+$/, "");
  if (existing === "") return newMarkdown;
  const addition = newMarkdown.replace(/^\n+/, "");
  return `${existing}\n\n${addition}`;
}
