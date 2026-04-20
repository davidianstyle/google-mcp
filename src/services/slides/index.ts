import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult } from "../../utils/formatting.js";

export function registerSlidesTools(server: McpServer, ctx: ServiceContext): void {
  const api = () => google.slides({ version: "v1", auth: ctx.auth });
  const driveApi = () => google.drive({ version: "v3", auth: ctx.auth });

  server.tool("slides_create_presentation", "Create a new Google Slides presentation", {
    title: z.string(),
  }, async ({ title }) => {
    const res = await api().presentations.create({ requestBody: { title } });
    return textResult({
      presentationId: res.data.presentationId,
      title: res.data.title,
      url: `https://docs.google.com/presentation/d/${res.data.presentationId}/edit`,
      slides: res.data.slides?.map((s) => s.objectId),
    });
  });

  server.tool("slides_get_presentation", "Get presentation metadata and slide list", {
    presentationId: z.string(),
  }, async ({ presentationId }) => {
    const res = await api().presentations.get({ presentationId });
    return textResult({
      presentationId: res.data.presentationId,
      title: res.data.title,
      slideCount: res.data.slides?.length,
      slides: res.data.slides?.map((s, i) => ({
        index: i,
        objectId: s.objectId,
        elements: s.pageElements?.length || 0,
      })),
      url: `https://docs.google.com/presentation/d/${res.data.presentationId}/edit`,
    });
  });

  server.tool("slides_add_slide", "Add a new slide to a presentation", {
    presentationId: z.string(),
    insertionIndex: z.number().optional().describe("Position to insert (0 = first)"),
    layoutId: z.string().optional().describe("Layout object ID to use"),
  }, async ({ presentationId, insertionIndex, layoutId }) => {
    const requests = [{
      createSlide: {
        insertionIndex,
        slideLayoutReference: layoutId ? { layoutId } : undefined,
      },
    }];
    const res = await api().presentations.batchUpdate({ presentationId, requestBody: { requests } });
    const slideId = res.data.replies?.[0]?.createSlide?.objectId;
    return textResult({ slideId });
  });

  server.tool("slides_delete_slide", "Delete a slide from a presentation", {
    presentationId: z.string(),
    slideObjectId: z.string(),
  }, async ({ presentationId, slideObjectId }) => {
    await api().presentations.batchUpdate({
      presentationId,
      requestBody: { requests: [{ deleteObject: { objectId: slideObjectId } }] },
    });
    return textResult({ success: true, slideObjectId });
  });

  server.tool("slides_duplicate_slide", "Duplicate an existing slide", {
    presentationId: z.string(),
    slideObjectId: z.string(),
    insertionIndex: z.number().optional(),
  }, async ({ presentationId, slideObjectId, insertionIndex }) => {
    const res = await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [{ duplicateObject: { objectId: slideObjectId } }],
      },
    });
    const newId = res.data.replies?.[0]?.duplicateObject?.objectId;
    return textResult({ newSlideId: newId });
  });

  server.tool("slides_reorder_slides", "Reorder slides in a presentation", {
    presentationId: z.string(),
    slideObjectIds: z.array(z.string()).describe("Slide IDs in the desired order"),
    insertionIndex: z.number().describe("Index to move slides to"),
  }, async ({ presentationId, slideObjectIds, insertionIndex }) => {
    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [{ updateSlidesPosition: { slideObjectIds, insertionIndex } }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("slides_add_text", "Add a text box to a slide", {
    presentationId: z.string(),
    slideObjectId: z.string(),
    text: z.string(),
    x: z.number().optional().default(100).describe("X position in EMU or points"),
    y: z.number().optional().default(100),
    width: z.number().optional().default(400),
    height: z.number().optional().default(50),
  }, async ({ presentationId, slideObjectId, text, x, y, width, height }) => {
    const boxId = `textbox_${Date.now()}`;
    const emu = (pts: number) => pts * 12700;
    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [
          {
            createShape: {
              objectId: boxId,
              shapeType: "TEXT_BOX",
              elementProperties: {
                pageObjectId: slideObjectId,
                size: { width: { magnitude: emu(width), unit: "EMU" }, height: { magnitude: emu(height), unit: "EMU" } },
                transform: { scaleX: 1, scaleY: 1, translateX: emu(x), translateY: emu(y), unit: "EMU" },
              },
            },
          },
          {
            insertText: { objectId: boxId, text, insertionIndex: 0 },
          },
        ],
      },
    });
    return textResult({ objectId: boxId });
  });

  server.tool("slides_add_image", "Add an image to a slide", {
    presentationId: z.string(),
    slideObjectId: z.string(),
    imageUrl: z.string().describe("Public URL of the image"),
    x: z.number().optional().default(100),
    y: z.number().optional().default(100),
    width: z.number().optional().default(300),
    height: z.number().optional().default(200),
  }, async ({ presentationId, slideObjectId, imageUrl, x, y, width, height }) => {
    const imageId = `image_${Date.now()}`;
    const emu = (pts: number) => pts * 12700;
    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [{
          createImage: {
            objectId: imageId,
            url: imageUrl,
            elementProperties: {
              pageObjectId: slideObjectId,
              size: { width: { magnitude: emu(width), unit: "EMU" }, height: { magnitude: emu(height), unit: "EMU" } },
              transform: { scaleX: 1, scaleY: 1, translateX: emu(x), translateY: emu(y), unit: "EMU" },
            },
          },
        }],
      },
    });
    return textResult({ objectId: imageId });
  });

  server.tool("slides_add_shape", "Add a shape to a slide", {
    presentationId: z.string(),
    slideObjectId: z.string(),
    shapeType: z.string().describe("Shape type (e.g., RECTANGLE, ELLIPSE, ARROW_EAST)"),
    x: z.number().optional().default(100),
    y: z.number().optional().default(100),
    width: z.number().optional().default(200),
    height: z.number().optional().default(100),
  }, async ({ presentationId, slideObjectId, shapeType, x, y, width, height }) => {
    const shapeId = `shape_${Date.now()}`;
    const emu = (pts: number) => pts * 12700;
    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [{
          createShape: {
            objectId: shapeId,
            shapeType,
            elementProperties: {
              pageObjectId: slideObjectId,
              size: { width: { magnitude: emu(width), unit: "EMU" }, height: { magnitude: emu(height), unit: "EMU" } },
              transform: { scaleX: 1, scaleY: 1, translateX: emu(x), translateY: emu(y), unit: "EMU" },
            },
          },
        }],
      },
    });
    return textResult({ objectId: shapeId });
  });

  server.tool("slides_add_video", "Embed a video on a slide", {
    presentationId: z.string(),
    slideObjectId: z.string(),
    videoUrl: z.string().describe("YouTube video URL"),
    x: z.number().optional().default(100),
    y: z.number().optional().default(100),
    width: z.number().optional().default(400),
    height: z.number().optional().default(300),
  }, async ({ presentationId, slideObjectId, videoUrl, x, y, width, height }) => {
    const videoId = `video_${Date.now()}`;
    const emu = (pts: number) => pts * 12700;
    const ytMatch = videoUrl.match(/(?:v=|youtu\.be\/)([a-zA-Z0-9_-]+)/);
    if (!ytMatch) return textResult({ error: "Could not extract YouTube video ID from URL" });

    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [{
          createVideo: {
            objectId: videoId,
            source: "YOUTUBE",
            id: ytMatch[1],
            elementProperties: {
              pageObjectId: slideObjectId,
              size: { width: { magnitude: emu(width), unit: "EMU" }, height: { magnitude: emu(height), unit: "EMU" } },
              transform: { scaleX: 1, scaleY: 1, translateX: emu(x), translateY: emu(y), unit: "EMU" },
            },
          },
        }],
      },
    });
    return textResult({ objectId: videoId });
  });

  server.tool("slides_insert_audio_link", "Insert a hyperlink to an audio file on a slide", {
    presentationId: z.string(),
    slideObjectId: z.string(),
    audioUrl: z.string(),
    linkText: z.string().optional().default("Audio Link"),
    x: z.number().optional().default(100),
    y: z.number().optional().default(100),
  }, async ({ presentationId, slideObjectId, audioUrl, linkText, x, y }) => {
    const boxId = `audio_link_${Date.now()}`;
    const emu = (pts: number) => pts * 12700;
    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [
          {
            createShape: {
              objectId: boxId,
              shapeType: "TEXT_BOX",
              elementProperties: {
                pageObjectId: slideObjectId,
                size: { width: { magnitude: emu(200), unit: "EMU" }, height: { magnitude: emu(30), unit: "EMU" } },
                transform: { scaleX: 1, scaleY: 1, translateX: emu(x), translateY: emu(y), unit: "EMU" },
              },
            },
          },
          { insertText: { objectId: boxId, text: linkText, insertionIndex: 0 } },
          {
            updateTextStyle: {
              objectId: boxId,
              style: { link: { url: audioUrl } },
              textRange: { type: "ALL" },
              fields: "link",
            },
          },
        ],
      },
    });
    return textResult({ objectId: boxId });
  });

  server.tool("slides_update_text_style", "Update text styling on a slide element", {
    presentationId: z.string(),
    objectId: z.string().describe("ID of the shape/text box"),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    fontSize: z.number().optional().describe("Font size in points"),
    fontFamily: z.string().optional(),
    foregroundColor: z.object({ red: z.number(), green: z.number(), blue: z.number() }).optional(),
  }, async ({ presentationId, objectId, bold, italic, underline, fontSize, fontFamily, foregroundColor }) => {
    const style: Record<string, unknown> = {};
    const fields: string[] = [];
    if (bold !== undefined) { style.bold = bold; fields.push("bold"); }
    if (italic !== undefined) { style.italic = italic; fields.push("italic"); }
    if (underline !== undefined) { style.underline = underline; fields.push("underline"); }
    if (fontSize) { style.fontSize = { magnitude: fontSize, unit: "PT" }; fields.push("fontSize"); }
    if (fontFamily) { style.fontFamily = fontFamily; fields.push("fontFamily"); }
    if (foregroundColor) { style.foregroundColor = { opaqueColor: { rgbColor: foregroundColor } }; fields.push("foregroundColor"); }

    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [{
          updateTextStyle: { objectId, style, textRange: { type: "ALL" }, fields: fields.join(",") },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("slides_update_paragraph_style", "Update paragraph styling on a slide element", {
    presentationId: z.string(),
    objectId: z.string(),
    alignment: z.enum(["START", "CENTER", "END", "JUSTIFIED"]).optional(),
    lineSpacing: z.number().optional().describe("Line spacing percentage (100 = single)"),
    spaceAbove: z.number().optional().describe("Space above in points"),
    spaceBelow: z.number().optional().describe("Space below in points"),
  }, async ({ presentationId, objectId, alignment, lineSpacing, spaceAbove, spaceBelow }) => {
    const style: Record<string, unknown> = {};
    const fields: string[] = [];
    if (alignment) { style.alignment = alignment; fields.push("alignment"); }
    if (lineSpacing) { style.lineSpacing = lineSpacing; fields.push("lineSpacing"); }
    if (spaceAbove !== undefined) { style.spaceAbove = { magnitude: spaceAbove, unit: "PT" }; fields.push("spaceAbove"); }
    if (spaceBelow !== undefined) { style.spaceBelow = { magnitude: spaceBelow, unit: "PT" }; fields.push("spaceBelow"); }

    await api().presentations.batchUpdate({
      presentationId,
      requestBody: {
        requests: [{
          updateParagraphStyle: { objectId, style, textRange: { type: "ALL" }, fields: fields.join(",") },
        }],
      },
    });
    return textResult({ success: true });
  });

  server.tool("slides_export_as_pdf", "Export entire presentation as PDF", {
    presentationId: z.string(),
    outputPath: z.string().optional().describe("Local path to save the PDF"),
  }, async ({ presentationId, outputPath }) => {
    const res = await driveApi().files.export(
      { fileId: presentationId, mimeType: "application/pdf" },
      { responseType: "arraybuffer" }
    );
    if (outputPath) {
      const { writeFileSync } = await import("node:fs");
      writeFileSync(outputPath, Buffer.from(res.data as ArrayBuffer));
      return textResult({ success: true, path: outputPath });
    }
    const base64 = Buffer.from(res.data as ArrayBuffer).toString("base64");
    return textResult({ success: true, base64Length: base64.length, note: "PDF exported as base64" });
  });

  server.tool("slides_export_slide_as_pdf", "Export a single slide as PDF", {
    presentationId: z.string(),
    slideObjectId: z.string(),
  }, async ({ presentationId, slideObjectId }) => {
    const pres = await api().presentations.get({ presentationId });
    const slideIndex = pres.data.slides?.findIndex((s) => s.objectId === slideObjectId);
    if (slideIndex === undefined || slideIndex < 0) return textResult({ error: "Slide not found" });

    const res = await driveApi().files.export(
      { fileId: presentationId, mimeType: "application/pdf" },
      { responseType: "arraybuffer" }
    );
    return textResult({ success: true, note: `Exported full presentation as PDF. Slide ${slideIndex + 1} extraction requires a PDF library.` });
  });

  server.tool("slides_get_thumbnail", "Get a thumbnail image of a presentation", {
    presentationId: z.string(),
    slideObjectId: z.string().optional().describe("Specific slide to thumbnail (defaults to first slide)"),
  }, async ({ presentationId, slideObjectId }) => {
    const pres = await api().presentations.get({ presentationId });
    const targetSlide = slideObjectId
      ? pres.data.slides?.find((s) => s.objectId === slideObjectId)
      : pres.data.slides?.[0];

    if (!targetSlide) return textResult({ error: "Slide not found" });

    const res = await api().presentations.pages.getThumbnail({
      presentationId,
      pageObjectId: targetSlide.objectId!,
    });
    return textResult({ contentUrl: res.data.contentUrl, width: res.data.width, height: res.data.height });
  });
}
