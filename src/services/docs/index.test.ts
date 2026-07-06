import { describe, expect, it } from "vitest";
import { docs_v1 } from "googleapis";
import { findInsertedTable } from "./index.js";

/**
 * Per the Docs API (Schema$InsertTableRequest.location): "A newline
 * character will be inserted before the inserted table, therefore the
 * table start index will be at the specified location index + 1."
 */
describe("findInsertedTable", () => {
  const tableAt = (startIndex: number, marker: number): docs_v1.Schema$StructuralElement => ({
    startIndex,
    endIndex: startIndex + 10,
    table: { rows: marker, columns: 1, tableRows: [] },
  });

  const paragraphAt = (startIndex: number): docs_v1.Schema$StructuralElement => ({
    startIndex,
    endIndex: startIndex + 1,
    paragraph: { elements: [] },
  });

  it("selects the table at insertionIndex + 1, not a decoy table elsewhere", () => {
    const content: docs_v1.Schema$StructuralElement[] = [
      paragraphAt(1),
      tableAt(5, 111), // decoy: pre-existing table earlier in the doc
      paragraphAt(15),
      tableAt(21, 222), // the just-inserted table (insertion at index 20 -> startIndex 21)
      paragraphAt(31),
    ];
    const found = findInsertedTable(content, 20);
    expect(found?.rows).toBe(222);
  });

  it("does not match a table sitting exactly at the insertion index", () => {
    // A table at startIndex === insertionIndex is NOT the inserted table
    // (that position is impossible for the new table; insertion shifts it +1).
    const content = [tableAt(20, 111)];
    expect(findInsertedTable(content, 20)).toBeUndefined();
  });

  it("returns undefined when no table matches (so callers must error, not report success)", () => {
    const content = [paragraphAt(1), tableAt(5, 111)];
    expect(findInsertedTable(content, 40)).toBeUndefined();
  });

  it("returns undefined for missing content", () => {
    expect(findInsertedTable(undefined, 10)).toBeUndefined();
  });
});
