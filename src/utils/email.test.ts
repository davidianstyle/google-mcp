import { describe, expect, it } from "vitest";
import { buildRawEmail, htmlToText } from "./email.js";

describe("buildRawEmail header hardening", () => {
  it("neutralizes a CRLF header-injection attempt in the subject", () => {
    const raw = buildRawEmail({
      to: ["victim@example.com"],
      subject: "Hi\r\nBcc: attacker@evil.com\r\nX-Injected: 1",
      body: "hello",
    });
    const lines = raw.split("\r\n");

    // The injected "headers" must not become real header lines of their own.
    expect(lines.some((l) => /^Bcc:/i.test(l))).toBe(false);
    expect(lines.some((l) => /^X-Injected:/i.test(l))).toBe(false);

    // The injected text should still be present, but folded into the Subject line.
    const subjectLine = lines.find((l) => l.startsWith("Subject:"));
    expect(subjectLine).toBeDefined();
    expect(subjectLine).toContain("Bcc: attacker@evil.com");
  });

  it("RFC 2047 encodes a non-ASCII subject", () => {
    const raw = buildRawEmail({ to: ["a@b.com"], subject: "Hello 🎉 World", body: "hi" });
    const subjectLine = raw.split("\r\n").find((l) => l.startsWith("Subject:"));
    expect(subjectLine).toBeDefined();

    const match = /^Subject: =\?UTF-8\?B\?(.+)\?=$/.exec(subjectLine!);
    expect(match).not.toBeNull();
    const decoded = Buffer.from(match![1], "base64").toString("utf-8");
    expect(decoded).toBe("Hello 🎉 World");
  });

  it("leaves a plain ASCII subject unencoded", () => {
    const raw = buildRawEmail({ to: ["a@b.com"], subject: "Plain subject", body: "hi" });
    expect(raw.split("\r\n")).toContain("Subject: Plain subject");
  });

  it("base64-encodes a long-line body so no raw line can exceed RFC 5322's 998-octet limit", () => {
    const longLine = "a".repeat(2000);
    const raw = buildRawEmail({ to: ["a@b.com"], subject: "test", body: longLine });

    expect(raw).toContain("Content-Transfer-Encoding: base64");
    for (const line of raw.split("\r\n")) {
      expect(line.length).toBeLessThanOrEqual(998);
    }

    const headerEnd = raw.indexOf("\r\n\r\n");
    const bodyPart = raw.slice(headerEnd + 4);
    const decodedBody = Buffer.from(bodyPart.replace(/\r\n/g, ""), "base64").toString("utf-8");
    expect(decodedBody).toBe(longLine);
  });

  it("base64-encodes both parts of a multipart/alternative message", () => {
    const raw = buildRawEmail({
      to: ["a@b.com"],
      subject: "test",
      body: "plain body",
      htmlBody: "<p>html body</p>",
      mimeType: "multipart/alternative",
    });
    const encodedCount = (raw.match(/Content-Transfer-Encoding: base64/g) || []).length;
    expect(encodedCount).toBe(2);
  });
});

describe("htmlToText", () => {
  it("strips tags, scripts, and styles, converts <br> to newlines, and decodes entities", () => {
    const html =
      "<html><head><style>body{color:red}</style></head><body>" +
      "<script>alert(1)</script><p>Hello &amp; welcome</p><br>Line2</body></html>";
    expect(htmlToText(html)).toBe("Hello & welcome\nLine2");
  });

  it("caps output length and appends a truncation note", () => {
    const html = `<p>${"a".repeat(100)}</p>`;
    const text = htmlToText(html, 20);
    expect(text.startsWith("a".repeat(20))).toBe(true);
    expect(text).toContain("truncated");
  });

  it("does not truncate content within the length cap", () => {
    const text = htmlToText("<p>short</p>", 1000);
    expect(text).toBe("short");
  });
});
