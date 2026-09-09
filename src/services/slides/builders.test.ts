import { describe, expect, it } from "vitest";
import {
  ptToEmu,
  buildCreateSlideRequest,
  parseBatchRequests,
  buildTemplateReplacementRequests,
  buildStyleShapeRequest,
  buildSetBackgroundRequests,
  computeAlignedTransforms,
  textContentLength,
  chunkRequests,
  buildTextStyleRequests,
  buildTableHeaderRequests,
  buildOutlineRequests,
  buildSpeakerNotesRequests,
  normalizeBody,
  type AlignInputElement,
} from "./builders.js";

describe("ptToEmu", () => {
  it("converts points to EMU at 12700 per point", () => {
    expect(ptToEmu(1)).toBe(12700);
    expect(ptToEmu(72)).toBe(914400);
  });
});

describe("buildCreateSlideRequest", () => {
  it("maps a predefined layout + placeholders to createSlide with placeholderIdMappings", () => {
    const req = buildCreateSlideRequest({
      slideObjectId: "slide_1",
      predefinedLayout: "TITLE_AND_BODY",
      insertionIndex: 2,
      placeholders: [
        { type: "TITLE", objectId: "ph_title" },
        { type: "BODY", index: 0, objectId: "ph_body" },
      ],
    });
    expect(req).toEqual({
      createSlide: {
        objectId: "slide_1",
        insertionIndex: 2,
        slideLayoutReference: { predefinedLayout: "TITLE_AND_BODY" },
        placeholderIdMappings: [
          { layoutPlaceholder: { type: "TITLE", index: 0 }, objectId: "ph_title" },
          { layoutPlaceholder: { type: "BODY", index: 0 }, objectId: "ph_body" },
        ],
      },
    });
  });

  it("uses layoutId when no predefinedLayout is given", () => {
    const req = buildCreateSlideRequest({ layoutId: "g123", placeholders: [{ type: "TITLE", objectId: "t" }] });
    expect(req.createSlide?.slideLayoutReference).toEqual({ layoutId: "g123" });
    expect(req.createSlide?.placeholderIdMappings).toHaveLength(1);
  });

  it("drops placeholderIdMappings when there is no layout reference (API forbids them)", () => {
    const req = buildCreateSlideRequest({ placeholders: [{ type: "TITLE", objectId: "t" }] });
    expect(req.createSlide?.placeholderIdMappings).toBeUndefined();
    expect(req.createSlide?.slideLayoutReference).toBeUndefined();
  });
});

describe("parseBatchRequests", () => {
  it("passes through an array of objects", () => {
    const arr = [{ insertText: { objectId: "x", text: "hi" } }];
    expect(parseBatchRequests(arr)).toEqual(arr);
  });

  it("parses a JSON string payload", () => {
    expect(parseBatchRequests('[{"deleteObject":{"objectId":"x"}}]')).toEqual([
      { deleteObject: { objectId: "x" } },
    ]);
  });

  it("rejects invalid JSON strings", () => {
    expect(() => parseBatchRequests("not json")).toThrow(/valid JSON/);
  });

  it("rejects non-arrays", () => {
    expect(() => parseBatchRequests({ insertText: {} })).toThrow(/must be an array/);
  });

  it("rejects an empty array", () => {
    expect(() => parseBatchRequests([])).toThrow(/empty/);
  });

  it("enforces the size guard", () => {
    const many = Array.from({ length: 51 }, () => ({ deleteObject: { objectId: "x" } }));
    expect(() => parseBatchRequests(many, 50)).toThrow(/Too many/);
  });

  it("rejects non-object entries", () => {
    expect(() => parseBatchRequests([{ a: 1 }, 5])).toThrow(/requests\[1\]/);
  });
});

describe("buildTemplateReplacementRequests", () => {
  it("builds replaceAllText + replaceAllShapesWithImage requests", () => {
    const reqs = buildTemplateReplacementRequests(
      { "{{name}}": "Ada", "{{date}}": "2026" },
      { "{{logo}}": "https://ex/logo.png" }
    );
    expect(reqs).toEqual([
      { replaceAllText: { containsText: { text: "{{name}}", matchCase: false }, replaceText: "Ada" } },
      { replaceAllText: { containsText: { text: "{{date}}", matchCase: false }, replaceText: "2026" } },
      { replaceAllShapesWithImage: { containsText: { text: "{{logo}}", matchCase: false }, imageUrl: "https://ex/logo.png" } },
    ]);
  });

  it("returns [] when nothing is provided", () => {
    expect(buildTemplateReplacementRequests(undefined, undefined)).toEqual([]);
  });
});

describe("buildStyleShapeRequest", () => {
  it("builds a fill+outline+shadow+autofit request with a precise field mask", () => {
    const req = buildStyleShapeRequest("box1", {
      fillColor: { red: 1, green: 0, blue: 0 },
      outlineColor: { red: 0, green: 0, blue: 1 },
      outlineWeightPt: 2,
      shadow: true,
      autofitType: "TEXT_AUTOFIT",
    });
    expect(req.updateShapeProperties?.objectId).toBe("box1");
    expect(req.updateShapeProperties?.fields).toBe(
      "shapeBackgroundFill.solidFill.color,outline.outlineFill.solidFill.color,outline.weight,shadow.propertyState,autofit.autofitType"
    );
    expect(req.updateShapeProperties?.shapeProperties).toEqual({
      shapeBackgroundFill: { solidFill: { color: { rgbColor: { red: 1, green: 0, blue: 0 } } } },
      outline: {
        outlineFill: { solidFill: { color: { rgbColor: { red: 0, green: 0, blue: 1 } } } },
        weight: { magnitude: 25400, unit: "EMU" },
      },
      shadow: { propertyState: "RENDERED" },
      autofit: { autofitType: "TEXT_AUTOFIT" },
    });
  });

  it("sets NOT_RENDERED when shadow is false", () => {
    const req = buildStyleShapeRequest("b", { shadow: false });
    expect(req.updateShapeProperties?.shapeProperties?.shadow).toEqual({ propertyState: "NOT_RENDERED" });
    expect(req.updateShapeProperties?.fields).toBe("shadow.propertyState");
  });

  it("includes the alpha field only when fillAlpha is set", () => {
    const req = buildStyleShapeRequest("b", { fillColor: { red: 0.5 }, fillAlpha: 0.8 });
    expect(req.updateShapeProperties?.fields).toBe("shapeBackgroundFill.solidFill.color,shapeBackgroundFill.solidFill.alpha");
    expect(req.updateShapeProperties?.shapeProperties?.shapeBackgroundFill?.solidFill?.alpha).toBe(0.8);
  });

  it("throws when no properties are set (empty field mask is invalid)", () => {
    expect(() => buildStyleShapeRequest("b", {})).toThrow(/nothing to update/);
  });
});

describe("buildSetBackgroundRequests", () => {
  it("builds a solid-color updatePageProperties per slide", () => {
    const reqs = buildSetBackgroundRequests(["s1", "s2"], { solidColor: { red: 0.1, green: 0.2, blue: 0.3 } });
    expect(reqs).toHaveLength(2);
    expect(reqs[0]).toEqual({
      updatePageProperties: {
        objectId: "s1",
        pageProperties: { pageBackgroundFill: { solidFill: { color: { rgbColor: { red: 0.1, green: 0.2, blue: 0.3 } } } } },
        fields: "pageBackgroundFill.solidFill.color",
      },
    });
    expect(reqs[1].updatePageProperties?.objectId).toBe("s2");
  });

  it("builds a stretched-picture background for an image URL", () => {
    const reqs = buildSetBackgroundRequests(["s1"], { imageUrl: "https://ex/bg.png" });
    expect(reqs[0]).toEqual({
      updatePageProperties: {
        objectId: "s1",
        pageProperties: { pageBackgroundFill: { stretchedPictureFill: { contentUrl: "https://ex/bg.png" } } },
        fields: "pageBackgroundFill.stretchedPictureFill.contentUrl",
      },
    });
  });
});

describe("computeAlignedTransforms", () => {
  const el = (objectId: string, x: number, y: number, w: number, h: number, scaleX = 1, scaleY = 1): AlignInputElement => ({
    objectId,
    transform: { scaleX, scaleY, translateX: x, translateY: y, unit: "EMU" },
    width: w,
    height: h,
  });

  it("aligns left edges to the leftmost element and preserves the y axis", () => {
    const out = computeAlignedTransforms([el("a", 100, 10, 50, 20), el("b", 300, 40, 50, 20)], "left");
    expect(out.map((o) => o.transform.translateX)).toEqual([100, 100]);
    // y axis untouched
    expect(out.map((o) => o.transform.translateY)).toEqual([10, 40]);
  });

  it("aligns right edges to the rightmost edge, accounting for scaled width", () => {
    // b spans 300..300+100 = 400 (rightmost). a is 50 wide -> left becomes 350.
    const out = computeAlignedTransforms([el("a", 100, 10, 50, 20), el("b", 300, 40, 100, 20)], "right");
    const byId = Object.fromEntries(out.map((o) => [o.objectId, o.transform.translateX]));
    expect(byId["a"]).toBe(350);
    expect(byId["b"]).toBe(300);
  });

  it("centers elements on the horizontal midpoint of the bounding box", () => {
    // bbox: left 100, right max(150, 400)=400 -> center 250. a width 50 -> 225. b width 100 -> 200.
    const out = computeAlignedTransforms([el("a", 100, 0, 50, 20), el("b", 300, 0, 100, 20)], "center");
    const byId = Object.fromEntries(out.map((o) => [o.objectId, o.transform.translateX]));
    expect(byId["a"]).toBe(225);
    expect(byId["b"]).toBe(200);
  });

  it("aligns top and bottom on the y axis", () => {
    const top = computeAlignedTransforms([el("a", 0, 10, 20, 30), el("b", 0, 50, 20, 30)], "top");
    expect(top.map((o) => o.transform.translateY)).toEqual([10, 10]);
    const bottom = computeAlignedTransforms([el("a", 0, 10, 20, 30), el("b", 0, 50, 20, 40)], "bottom");
    // bbox bottom = max(10+30, 50+40)=90. a h30 -> 60. b h40 -> 50.
    const byId = Object.fromEntries(bottom.map((o) => [o.objectId, o.transform.translateY]));
    expect(byId["a"]).toBe(60);
    expect(byId["b"]).toBe(50);
  });

  it("respects element scale when computing rendered size", () => {
    // a is 50 wide at scale 2 -> rendered 100, spanning 100..200.
    // b at 300 width 100 -> 300..400. right align: a.left = 400-100 = 300.
    const out = computeAlignedTransforms([el("a", 100, 0, 50, 20, 2, 1), el("b", 300, 0, 100, 20)], "right");
    const byId = Object.fromEntries(out.map((o) => [o.objectId, o.transform.translateX]));
    expect(byId["a"]).toBe(300);
  });

  it("distributes horizontally so edge gaps are equal and endpoints stay fixed", () => {
    // widths 100 each at 0, 250, 900. bbox 0..1000, total width 300, freeSpace 700, gap 350.
    const out = computeAlignedTransforms(
      [el("a", 0, 0, 100, 20), el("mid", 250, 0, 100, 20), el("z", 900, 0, 100, 20)],
      "distribute_horizontal"
    );
    const byId = Object.fromEntries(out.map((o) => [o.objectId, o.transform.translateX]));
    expect(byId["a"]).toBe(0); // first fixed
    expect(byId["mid"]).toBe(450); // 0 + 100 + 350
    expect(byId["z"]).toBe(900); // last fixed (ends at 1000)
  });

  it("returns [] for distribute with fewer than 3 elements", () => {
    expect(computeAlignedTransforms([el("a", 0, 0, 10, 10), el("b", 50, 0, 10, 10)], "distribute_horizontal")).toEqual([]);
  });

  it("returns [] for alignment with fewer than 2 elements", () => {
    expect(computeAlignedTransforms([el("a", 0, 0, 10, 10)], "left")).toEqual([]);
  });
});

describe("textContentLength", () => {
  it("sums textRun content lengths", () => {
    expect(
      textContentLength({
        textElements: [
          { paragraphMarker: {} },
          { textRun: { content: "Hello " } },
          { textRun: { content: "world" } },
        ],
      })
    ).toBe(11);
  });

  it("returns 0 for an empty shape (only the implicit paragraph marker)", () => {
    expect(textContentLength({ textElements: [{ paragraphMarker: {} }] })).toBe(0);
    expect(textContentLength(undefined)).toBe(0);
  });
});

describe("chunkRequests", () => {
  it("splits into batches of at most `size` and keeps order", () => {
    const items = Array.from({ length: 103 }, (_, i) => i);
    const chunks = chunkRequests(items, 50);
    expect(chunks.map((c) => c.length)).toEqual([50, 50, 3]);
    expect(chunks[2]).toEqual([100, 101, 102]);
  });

  it("returns no batches for an empty list", () => {
    expect(chunkRequests([], 50)).toEqual([]);
  });
});

describe("normalizeBody", () => {
  it("joins array items with newlines and marks them as bullets", () => {
    expect(normalizeBody(["a", " b ", ""])).toEqual({ text: "a\nb", bullets: true });
  });
  it("treats a multi-line string as bullets and a single line as a paragraph", () => {
    expect(normalizeBody("one\ntwo")).toEqual({ text: "one\ntwo", bullets: true });
    expect(normalizeBody("just text")).toEqual({ text: "just text", bullets: false });
  });
});

describe("buildTextStyleRequests", () => {
  it("returns nothing when no options are set", () => {
    expect(buildTextStyleRequests("box", {})).toEqual([]);
  });

  it("emits bullets then one updateTextStyle with only the requested fields", () => {
    const reqs = buildTextStyleRequests("box", { bullets: true, fontSize: 18, bold: true, fontFamily: "Roboto" });
    expect(reqs).toEqual([
      { createParagraphBullets: { objectId: "box", textRange: { type: "ALL" }, bulletPreset: "BULLET_DISC_CIRCLE_SQUARE" } },
      {
        updateTextStyle: {
          objectId: "box",
          textRange: { type: "ALL" },
          style: { fontSize: { magnitude: 18, unit: "PT" }, bold: true, fontFamily: "Roboto" },
          fields: "fontSize,bold,fontFamily",
        },
      },
    ]);
  });

  it("honors bold=false explicitly", () => {
    const [req] = buildTextStyleRequests("box", { bold: false });
    expect(req.updateTextStyle?.style).toEqual({ bold: false });
    expect(req.updateTextStyle?.fields).toBe("bold");
  });
});

describe("buildTableHeaderRequests", () => {
  it("fills row 0 across all columns and bolds each header cell", () => {
    const reqs = buildTableHeaderRequests("tbl", 3);
    expect(reqs).toHaveLength(4);
    expect(reqs[0].updateTableCellProperties?.tableRange).toEqual({
      location: { rowIndex: 0, columnIndex: 0 }, rowSpan: 1, columnSpan: 3,
    });
    expect(reqs[0].updateTableCellProperties?.fields).toBe("tableCellBackgroundFill.solidFill.color");
    expect(reqs.slice(1).map((r) => r.updateTextStyle?.cellLocation)).toEqual([
      { rowIndex: 0, columnIndex: 0 }, { rowIndex: 0, columnIndex: 1 }, { rowIndex: 0, columnIndex: 2 },
    ]);
    expect(reqs[1].updateTextStyle?.style).toEqual({ bold: true });
  });
});

describe("buildOutlineRequests", () => {
  const makeId = (prefix: string) => `${prefix}_x`;

  it("defaults the first slide to TITLE and later slides to TITLE_AND_BODY, with placeholder mappings", () => {
    const { requests, slides } = buildOutlineRequests([
      { title: "Deck", subtitle: "Sub" },
      { title: "Agenda", body: ["One", "Two"] },
    ], makeId);

    expect(slides.map((s) => s.layout)).toEqual(["TITLE", "TITLE_AND_BODY"]);
    expect(slides[0].slideId).toBe("slide_1_x");
    expect(slides[0].placeholderIds).toEqual({ title: "slide_1_x_title", subtitle: "slide_1_x_subtitle" });
    expect(slides[1].placeholderIds).toEqual({ title: "slide_2_x_title", body: "slide_2_x_body" });

    const create0 = requests[0].createSlide!;
    expect(create0.objectId).toBe("slide_1_x");
    expect(create0.insertionIndex).toBe(0);
    expect(create0.slideLayoutReference).toEqual({ predefinedLayout: "TITLE" });
    expect(create0.placeholderIdMappings).toEqual([
      { layoutPlaceholder: { type: "CENTERED_TITLE", index: 0 }, objectId: "slide_1_x_title" },
      { layoutPlaceholder: { type: "SUBTITLE", index: 0 }, objectId: "slide_1_x_subtitle" },
    ]);

    const types = requests.map((r) => Object.keys(r)[0]);
    expect(types).toEqual([
      "createSlide", "insertText", "insertText",
      "createSlide", "insertText", "insertText", "createParagraphBullets",
    ]);
    const bodyInsert = requests[5].insertText!;
    expect(bodyInsert).toEqual({ objectId: "slide_2_x_body", text: "One\nTwo", insertionIndex: 0 });
    expect(requests[6].createParagraphBullets).toEqual({
      objectId: "slide_2_x_body", textRange: { type: "ALL" }, bulletPreset: "BULLET_DISC_CIRCLE_SQUARE",
    });
  });

  it("does not bullet a single-line body string", () => {
    const { requests } = buildOutlineRequests([{ layout: "TITLE_AND_BODY", title: "T", body: "Just a sentence." }], makeId);
    expect(requests.some((r) => r.createParagraphBullets)).toBe(false);
  });

  it("skips fields the layout has no placeholder for and reports them", () => {
    const { requests, slides } = buildOutlineRequests([{ layout: "TITLE_ONLY", title: "T", body: ["a"], subtitle: "s" }], makeId);
    expect(slides[0].skipped).toEqual(["subtitle", "body"]);
    expect(slides[0].placeholderIds).toEqual({ title: "slide_1_x_title" });
    expect(requests[0].createSlide?.placeholderIdMappings).toHaveLength(1);
  });

  it("carries notes through without emitting requests for them", () => {
    const { requests, slides } = buildOutlineRequests([{ layout: "BLANK", notes: "say hi" }], makeId);
    expect(slides[0].notes).toBe("say hi");
    expect(requests).toHaveLength(1);
    expect(requests[0].createSlide?.placeholderIdMappings).toBeUndefined();
  });
});

describe("buildSpeakerNotesRequests", () => {
  it("emits one insertText per non-empty note", () => {
    expect(buildSpeakerNotesRequests([
      { speakerNotesObjectId: "n1", notes: "hello" },
      { speakerNotesObjectId: "n2", notes: "" },
    ])).toEqual([{ insertText: { objectId: "n1", text: "hello", insertionIndex: 0 } }]);
  });
});
