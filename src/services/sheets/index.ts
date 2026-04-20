import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult } from "../../utils/formatting.js";

export function registerSheetsTools(server: McpServer, ctx: ServiceContext): void {
  const api = () => google.sheets({ version: "v4", auth: ctx.auth });
  const driveApi = () => google.drive({ version: "v3", auth: ctx.auth });

  server.tool("sheets_read", "Read data from a spreadsheet range", {
    spreadsheetId: z.string(),
    range: z.string().describe("A1 notation (e.g., 'Sheet1!A1:C10')"),
    valueRenderOption: z.enum(["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]).optional().default("FORMATTED_VALUE"),
  }, async ({ spreadsheetId, range, valueRenderOption }) => {
    const res = await api().spreadsheets.values.get({ spreadsheetId, range, valueRenderOption });
    return textResult({ range: res.data.range, values: res.data.values });
  });

  server.tool("sheets_write", "Write data to a spreadsheet range", {
    spreadsheetId: z.string(),
    range: z.string(),
    values: z.array(z.array(z.unknown())).describe("2D array of values"),
    valueInputOption: z.enum(["RAW", "USER_ENTERED"]).optional().default("USER_ENTERED"),
  }, async ({ spreadsheetId, range, values, valueInputOption }) => {
    const res = await api().spreadsheets.values.update({
      spreadsheetId, range, valueInputOption,
      requestBody: { values },
    });
    return textResult({ updatedRange: res.data.updatedRange, updatedCells: res.data.updatedCells });
  });

  server.tool("sheets_batch_write", "Write data to multiple ranges at once", {
    spreadsheetId: z.string(),
    data: z.array(z.object({ range: z.string(), values: z.array(z.array(z.unknown())) })),
    valueInputOption: z.enum(["RAW", "USER_ENTERED"]).optional().default("USER_ENTERED"),
  }, async ({ spreadsheetId, data, valueInputOption }) => {
    const res = await api().spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption, data },
    });
    return textResult({ totalUpdatedCells: res.data.totalUpdatedCells, totalUpdatedSheets: res.data.totalUpdatedSheets });
  });

  server.tool("sheets_create", "Create a new spreadsheet", {
    title: z.string(),
    sheetTitles: z.array(z.string()).optional().describe("Names of initial sheets"),
  }, async ({ title, sheetTitles }) => {
    const sheets = sheetTitles?.map((t) => ({ properties: { title: t } }));
    const res = await api().spreadsheets.create({
      requestBody: { properties: { title }, sheets },
    });
    return textResult({
      spreadsheetId: res.data.spreadsheetId,
      url: res.data.spreadsheetUrl,
      sheets: res.data.sheets?.map((s) => ({ sheetId: s.properties?.sheetId, title: s.properties?.title })),
    });
  });

  server.tool("sheets_get_info", "Get spreadsheet metadata", {
    spreadsheetId: z.string(),
  }, async ({ spreadsheetId }) => {
    const res = await api().spreadsheets.get({ spreadsheetId });
    return textResult({
      spreadsheetId: res.data.spreadsheetId,
      title: res.data.properties?.title,
      url: res.data.spreadsheetUrl,
      sheets: res.data.sheets?.map((s) => ({
        sheetId: s.properties?.sheetId,
        title: s.properties?.title,
        rowCount: s.properties?.gridProperties?.rowCount,
        columnCount: s.properties?.gridProperties?.columnCount,
      })),
    });
  });

  server.tool("sheets_list", "List spreadsheets in Drive", {
    maxResults: z.number().optional().default(20),
  }, async ({ maxResults }) => {
    const res = await driveApi().files.list({
      q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
      pageSize: maxResults,
      orderBy: "modifiedTime desc",
      fields: "files(id,name,modifiedTime,webViewLink)",
    });
    return textResult(res.data.files || []);
  });

  server.tool("sheets_add_sheet", "Add a new sheet (tab) to a spreadsheet", {
    spreadsheetId: z.string(),
    title: z.string(),
    rowCount: z.number().optional(),
    columnCount: z.number().optional(),
  }, async ({ spreadsheetId, title, rowCount, columnCount }) => {
    const res = await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title, gridProperties: { rowCount: rowCount || 1000, columnCount: columnCount || 26 } } } }],
      },
    });
    const sheet = res.data.replies?.[0]?.addSheet;
    return textResult({ sheetId: sheet?.properties?.sheetId, title: sheet?.properties?.title });
  });

  server.tool("sheets_delete_sheet", "Delete a sheet from a spreadsheet", {
    spreadsheetId: z.string(),
    sheetId: z.number().describe("Sheet ID (not the sheet name)"),
  }, async ({ spreadsheetId, sheetId }) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ deleteSheet: { sheetId } }] },
    });
    return textResult({ success: true, sheetId });
  });

  server.tool("sheets_rename_sheet", "Rename a sheet", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    newTitle: z.string(),
  }, async ({ spreadsheetId, sheetId, newTitle }) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ updateSheetProperties: { properties: { sheetId, title: newTitle }, fields: "title" } }],
      },
    });
    return textResult({ success: true, sheetId, newTitle });
  });

  server.tool("sheets_duplicate_sheet", "Duplicate a sheet", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    newTitle: z.string().optional(),
    insertIndex: z.number().optional(),
  }, async ({ spreadsheetId, sheetId, newTitle, insertIndex }) => {
    const res = await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ duplicateSheet: { sourceSheetId: sheetId, newSheetName: newTitle, insertSheetIndex: insertIndex } }],
      },
    });
    const dup = res.data.replies?.[0]?.duplicateSheet;
    return textResult({ sheetId: dup?.properties?.sheetId, title: dup?.properties?.title });
  });

  server.tool("sheets_append_rows", "Append rows to the end of a sheet", {
    spreadsheetId: z.string(),
    range: z.string().describe("A1 range to search for a table (e.g., 'Sheet1')"),
    values: z.array(z.array(z.unknown())),
    valueInputOption: z.enum(["RAW", "USER_ENTERED"]).optional().default("USER_ENTERED"),
  }, async ({ spreadsheetId, range, values, valueInputOption }) => {
    const res = await api().spreadsheets.values.append({
      spreadsheetId, range, valueInputOption,
      requestBody: { values },
    });
    return textResult({ updatedRange: res.data.updates?.updatedRange, updatedRows: res.data.updates?.updatedRows });
  });

  server.tool("sheets_clear_range", "Clear all values from a range", {
    spreadsheetId: z.string(),
    range: z.string(),
  }, async ({ spreadsheetId, range }) => {
    const res = await api().spreadsheets.values.clear({ spreadsheetId, range });
    return textResult({ clearedRange: res.data.clearedRange });
  });

  server.tool("sheets_delete_range", "Delete rows or columns from a sheet", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    dimension: z.enum(["ROWS", "COLUMNS"]),
    startIndex: z.number(),
    endIndex: z.number(),
  }, async ({ spreadsheetId, sheetId, dimension, startIndex, endIndex }) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ deleteDimension: { range: { sheetId, dimension, startIndex, endIndex } } }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_read_cell_format", "Read formatting of cells in a range", {
    spreadsheetId: z.string(),
    range: z.string(),
  }, async ({ spreadsheetId, range }) => {
    const res = await api().spreadsheets.get({
      spreadsheetId,
      ranges: [range],
      includeGridData: true,
    });
    const grid = res.data.sheets?.[0]?.data?.[0];
    const formats = grid?.rowData?.map((row) =>
      row.values?.map((cell) => ({
        value: cell.formattedValue,
        format: cell.effectiveFormat,
      }))
    );
    return textResult(formats);
  });

  server.tool("sheets_format_cells", "Apply formatting to a cell range", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    startRowIndex: z.number(),
    endRowIndex: z.number(),
    startColumnIndex: z.number(),
    endColumnIndex: z.number(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    fontSize: z.number().optional(),
    backgroundColor: z.object({ red: z.number(), green: z.number(), blue: z.number() }).optional(),
    horizontalAlignment: z.enum(["LEFT", "CENTER", "RIGHT"]).optional(),
    numberFormat: z.object({ type: z.string(), pattern: z.string() }).optional(),
  }, async (opts) => {
    const format: Record<string, unknown> = {};
    const fields: string[] = [];
    if (opts.bold !== undefined || opts.italic !== undefined || opts.fontSize !== undefined) {
      const tf: Record<string, unknown> = {};
      if (opts.bold !== undefined) tf.bold = opts.bold;
      if (opts.italic !== undefined) tf.italic = opts.italic;
      if (opts.fontSize !== undefined) tf.fontSize = opts.fontSize;
      format.textFormat = tf;
      fields.push("userEnteredFormat.textFormat");
    }
    if (opts.backgroundColor) {
      format.backgroundColor = opts.backgroundColor;
      fields.push("userEnteredFormat.backgroundColor");
    }
    if (opts.horizontalAlignment) {
      format.horizontalAlignment = opts.horizontalAlignment;
      fields.push("userEnteredFormat.horizontalAlignment");
    }
    if (opts.numberFormat) {
      format.numberFormat = opts.numberFormat;
      fields.push("userEnteredFormat.numberFormat");
    }

    await api().spreadsheets.batchUpdate({
      spreadsheetId: opts.spreadsheetId,
      requestBody: {
        requests: [{
          repeatCell: {
            range: { sheetId: opts.sheetId, startRowIndex: opts.startRowIndex, endRowIndex: opts.endRowIndex, startColumnIndex: opts.startColumnIndex, endColumnIndex: opts.endColumnIndex },
            cell: { userEnteredFormat: format },
            fields: fields.join(","),
          },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_add_conditional_formatting", "Add conditional formatting to a range", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    startRowIndex: z.number(),
    endRowIndex: z.number(),
    startColumnIndex: z.number(),
    endColumnIndex: z.number(),
    type: z.enum(["NUMBER_GREATER", "NUMBER_LESS", "TEXT_CONTAINS", "CUSTOM_FORMULA"]),
    value: z.string().describe("Comparison value or formula"),
    backgroundColor: z.object({ red: z.number(), green: z.number(), blue: z.number() }),
  }, async (opts) => {
    const conditionType = opts.type === "CUSTOM_FORMULA" ? "CUSTOM_FORMULA"
      : opts.type === "NUMBER_GREATER" ? "NUMBER_GREATER"
      : opts.type === "NUMBER_LESS" ? "NUMBER_LESS" : "TEXT_CONTAINS";

    await api().spreadsheets.batchUpdate({
      spreadsheetId: opts.spreadsheetId,
      requestBody: {
        requests: [{
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: opts.sheetId, startRowIndex: opts.startRowIndex, endRowIndex: opts.endRowIndex, startColumnIndex: opts.startColumnIndex, endColumnIndex: opts.endColumnIndex }],
              booleanRule: {
                condition: { type: conditionType, values: [{ userEnteredValue: opts.value }] },
                format: { backgroundColor: opts.backgroundColor },
              },
            },
            index: 0,
          },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_set_column_widths", "Set column widths", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    startIndex: z.number(),
    endIndex: z.number(),
    pixelSize: z.number(),
  }, async ({ spreadsheetId, sheetId, startIndex, endIndex, pixelSize }) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          updateDimensionProperties: {
            range: { sheetId, dimension: "COLUMNS", startIndex, endIndex },
            properties: { pixelSize },
            fields: "pixelSize",
          },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_auto_resize_columns", "Auto-resize columns to fit content", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    startIndex: z.number().optional().default(0),
    endIndex: z.number().optional(),
  }, async ({ spreadsheetId, sheetId, startIndex, endIndex }) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          autoResizeDimensions: {
            dimensions: { sheetId, dimension: "COLUMNS", startIndex, endIndex: endIndex || undefined },
          },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_freeze_rows_and_columns", "Freeze rows and/or columns", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    frozenRowCount: z.number().optional(),
    frozenColumnCount: z.number().optional(),
  }, async ({ spreadsheetId, sheetId, frozenRowCount, frozenColumnCount }) => {
    const gridProperties: Record<string, number> = {};
    const fields: string[] = [];
    if (frozenRowCount !== undefined) { gridProperties.frozenRowCount = frozenRowCount; fields.push("gridProperties.frozenRowCount"); }
    if (frozenColumnCount !== undefined) { gridProperties.frozenColumnCount = frozenColumnCount; fields.push("gridProperties.frozenColumnCount"); }

    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ updateSheetProperties: { properties: { sheetId, gridProperties }, fields: fields.join(",") } }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_set_dropdown_validation", "Set dropdown validation on cells", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    startRowIndex: z.number(),
    endRowIndex: z.number(),
    startColumnIndex: z.number(),
    endColumnIndex: z.number(),
    values: z.array(z.string()).describe("Allowed dropdown values"),
    strict: z.boolean().optional().default(true).describe("Reject input not in the list"),
  }, async (opts) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId: opts.spreadsheetId,
      requestBody: {
        requests: [{
          setDataValidation: {
            range: { sheetId: opts.sheetId, startRowIndex: opts.startRowIndex, endRowIndex: opts.endRowIndex, startColumnIndex: opts.startColumnIndex, endColumnIndex: opts.endColumnIndex },
            rule: {
              condition: { type: "ONE_OF_LIST", values: opts.values.map((v) => ({ userEnteredValue: v })) },
              strict: opts.strict,
              showCustomUi: true,
            },
          },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_group_rows", "Group rows (collapse/expand)", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    startIndex: z.number(),
    endIndex: z.number(),
  }, async ({ spreadsheetId, sheetId, startIndex, endIndex }) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          addDimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex, endIndex } },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("sheets_ungroup_all_rows", "Remove all row groupings from a sheet", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
  }, async ({ spreadsheetId, sheetId }) => {
    const info = await api().spreadsheets.get({ spreadsheetId });
    const sheet = info.data.sheets?.find((s) => s.properties?.sheetId === sheetId);
    const rowCount = sheet?.properties?.gridProperties?.rowCount || 1000;

    try {
      await api().spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{
            deleteDimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: rowCount } },
          }],
        },
      });
    } catch {
      // No groups to delete
    }
    return textResult({ success: true });
  });

  server.tool("sheets_insert_chart", "Insert a chart into a sheet", {
    spreadsheetId: z.string(),
    sheetId: z.number(),
    chartType: z.enum(["BAR", "LINE", "PIE", "COLUMN", "AREA", "SCATTER"]),
    title: z.string().optional(),
    dataRange: z.string().describe("A1 notation of the data range for the chart"),
    anchorRowIndex: z.number().optional().default(0),
    anchorColumnIndex: z.number().optional().default(0),
  }, async ({ spreadsheetId, sheetId, chartType, title, dataRange, anchorRowIndex, anchorColumnIndex }) => {
    const res = await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          addChart: {
            chart: {
              spec: {
                title,
                basicChart: {
                  chartType,
                  domains: [{ domain: { sourceRange: { sources: [{ sheetId, startRowIndex: 0, endRowIndex: 100, startColumnIndex: 0, endColumnIndex: 1 }] } } }],
                  series: [{ series: { sourceRange: { sources: [{ sheetId, startRowIndex: 0, endRowIndex: 100, startColumnIndex: 1, endColumnIndex: 2 }] } } }],
                },
              },
              position: { overlayPosition: { anchorCell: { sheetId, rowIndex: anchorRowIndex, columnIndex: anchorColumnIndex } } },
            },
          },
        }],
      },
    });
    const chart = res.data.replies?.[0]?.addChart?.chart;
    return textResult({ chartId: chart?.chartId, title: chart?.spec?.title });
  });

  server.tool("sheets_delete_chart", "Delete a chart from a spreadsheet", {
    spreadsheetId: z.string(),
    chartId: z.number(),
  }, async ({ spreadsheetId, chartId }) => {
    await api().spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ deleteEmbeddedObject: { objectId: chartId } }] },
    });
    return textResult({ success: true, chartId });
  });

  server.tool("sheets_replace_range_with_markdown", "Replace a range with markdown-formatted content (as plain text)", {
    spreadsheetId: z.string(),
    range: z.string(),
    markdown: z.string(),
  }, async ({ spreadsheetId, range, markdown }) => {
    const rows = markdown.split("\n").map((line) => [line]);
    const res = await api().spreadsheets.values.update({
      spreadsheetId, range,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    });
    return textResult({ updatedRange: res.data.updatedRange, updatedCells: res.data.updatedCells });
  });
}
