import { gmail_v1 } from "googleapis";

/** Decodes a Gmail base64url string to raw bytes (safe for binary content like attachments). */
export function decodeBase64UrlToBuffer(data: string): Buffer {
  return Buffer.from(data.replace(/-/g, "+").replace(/_/g, "/"), "base64");
}

export function decodeBase64Url(data: string): string {
  return decodeBase64UrlToBuffer(data).toString("utf-8");
}

export function encodeBase64Url(data: string): string {
  return Buffer.from(data).toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export function extractBody(payload: gmail_v1.Schema$MessagePart | undefined): { text: string; html: string } {
  if (!payload) return { text: "", html: "" };

  let text = "";
  let html = "";

  if (payload.mimeType === "text/plain" && payload.body?.data) {
    text = decodeBase64Url(payload.body.data);
  } else if (payload.mimeType === "text/html" && payload.body?.data) {
    html = decodeBase64Url(payload.body.data);
  }

  if (payload.parts) {
    for (const part of payload.parts) {
      if (part.mimeType === "text/plain" && part.body?.data) {
        text += decodeBase64Url(part.body.data);
      } else if (part.mimeType === "text/html" && part.body?.data) {
        html += decodeBase64Url(part.body.data);
      } else if (part.mimeType?.startsWith("multipart/") && part.parts) {
        const nested = extractBody(part);
        text += nested.text;
        html += nested.html;
      }
    }
  }

  return { text, html };
}

export function getHeader(headers: gmail_v1.Schema$MessagePartHeader[] | undefined, name: string): string {
  return headers?.find((h) => h.name?.toLowerCase() === name.toLowerCase())?.value || "";
}

/** Strips CR/LF out of a header value so it can't inject additional header lines (or a body separator). */
function sanitizeHeaderValue(value: string): string {
  return value.replace(/[\r\n]+/g, " ").trim();
}

/**
 * RFC 2047-encodes a header value if it contains any non-ASCII characters
 * (e.g. "=?UTF-8?B?...?="), after stripping CR/LF. ASCII-only values are
 * returned unchanged (untouched, unencoded) so the common case stays
 * human-readable in the raw message.
 */
function encodeHeaderValue(value: string): string {
  const sanitized = sanitizeHeaderValue(value);
  // eslint-disable-next-line no-control-regex
  if (!/[^\x00-\x7F]/.test(sanitized)) return sanitized;
  return `=?UTF-8?B?${Buffer.from(sanitized, "utf-8").toString("base64")}?=`;
}

/** Free-text header (Subject etc.): CRLF-sanitized and RFC 2047-encoded when non-ASCII. */
function header(name: string, value: string): string {
  return `${name}: ${encodeHeaderValue(value)}`;
}

/**
 * Address header (To/Cc/Bcc/From): CRLF-sanitized only. RFC 2047
 * encoded-words are not allowed inside an addr-spec, so base64-encoding the
 * whole value would swallow the email address into an opaque blob and break
 * parsing/delivery. Non-ASCII display names pass through as raw UTF-8,
 * which Gmail accepts (SMTPUTF8) — imperfect for strict RFC 5322 relays,
 * but strictly better than corrupting the address.
 */
function addressHeader(name: string, value: string): string {
  return `${name}: ${sanitizeHeaderValue(value)}`;
}

/** Wraps a base64 string into CRLF-separated 76-character lines (RFC 2045 §6.8). */
function wrapBase64(base64: string): string {
  const lines: string[] = [];
  for (let i = 0; i < base64.length; i += 76) lines.push(base64.slice(i, i + 76));
  return lines.join("\r\n");
}

/**
 * Base64-encodes a MIME body part so arbitrarily long lines in the source
 * text (which would otherwise violate RFC 5322's 998-octet line limit)
 * become short, safe base64 lines instead.
 */
function encodeBodyBase64(text: string): string {
  return wrapBase64(Buffer.from(text, "utf-8").toString("base64"));
}

export function buildRawEmail(opts: {
  to: string[];
  subject: string;
  body: string;
  htmlBody?: string;
  cc?: string[];
  bcc?: string[];
  from?: string;
  inReplyTo?: string;
  references?: string;
  mimeType?: string;
}): string {
  const boundary = `boundary_${Date.now()}`;
  const headers: string[] = [];

  headers.push(addressHeader("To", opts.to.join(", ")));
  if (opts.from) headers.push(addressHeader("From", opts.from));
  if (opts.cc?.length) headers.push(addressHeader("Cc", opts.cc.join(", ")));
  if (opts.bcc?.length) headers.push(addressHeader("Bcc", opts.bcc.join(", ")));
  headers.push(header("Subject", opts.subject));
  if (opts.inReplyTo) {
    // Message-ID references are ASCII identifiers; sanitize only.
    headers.push(addressHeader("In-Reply-To", opts.inReplyTo));
    headers.push(addressHeader("References", opts.references || opts.inReplyTo));
  }

  if (opts.htmlBody && opts.mimeType === "multipart/alternative") {
    headers.push(`MIME-Version: 1.0`);
    headers.push(`Content-Type: multipart/alternative; boundary="${boundary}"`);
    const parts = [
      `--${boundary}\r\nContent-Type: text/plain; charset="UTF-8"\r\nContent-Transfer-Encoding: base64\r\n\r\n${encodeBodyBase64(opts.body)}`,
      `--${boundary}\r\nContent-Type: text/html; charset="UTF-8"\r\nContent-Transfer-Encoding: base64\r\n\r\n${encodeBodyBase64(opts.htmlBody)}`,
      `--${boundary}--`,
    ];
    return headers.join("\r\n") + "\r\n\r\n" + parts.join("\r\n");
  }

  if (opts.htmlBody || opts.mimeType === "text/html") {
    headers.push(`MIME-Version: 1.0`);
    headers.push(`Content-Type: text/html; charset="UTF-8"`);
    headers.push(`Content-Transfer-Encoding: base64`);
    return headers.join("\r\n") + "\r\n\r\n" + encodeBodyBase64(opts.htmlBody || opts.body);
  }

  headers.push(`MIME-Version: 1.0`);
  headers.push(`Content-Type: text/plain; charset="UTF-8"`);
  headers.push(`Content-Transfer-Encoding: base64`);
  return headers.join("\r\n") + "\r\n\r\n" + encodeBodyBase64(opts.body);
}

export interface AttachmentInfo {
  attachmentId: string;
  filename: string;
  mimeType: string;
  size: number;
}

export function extractAttachments(payload: gmail_v1.Schema$MessagePart | undefined): AttachmentInfo[] {
  if (!payload) return [];
  const attachments: AttachmentInfo[] = [];

  function walk(part: gmail_v1.Schema$MessagePart) {
    if (part.filename && part.filename.length > 0 && part.body?.attachmentId) {
      attachments.push({
        attachmentId: part.body.attachmentId,
        filename: part.filename,
        mimeType: part.mimeType || "application/octet-stream",
        size: part.body.size || 0,
      });
    }
    if (part.parts) {
      for (const child of part.parts) {
        walk(child);
      }
    }
  }

  walk(payload);
  return attachments;
}

const NAMED_HTML_ENTITIES: Record<string, string> = {
  amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ",
  copy: "©", reg: "®", trade: "™",
  mdash: "—", ndash: "–", hellip: "…",
  lsquo: "‘", rsquo: "’", ldquo: "“", rdquo: "”",
};

function decodeHtmlEntities(input: string): string {
  return input.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (match, entity: string) => {
    if (entity[0] === "#") {
      const isHex = entity[1] === "x" || entity[1] === "X";
      const codePoint = isHex ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
      if (Number.isNaN(codePoint)) return match;
      try {
        return String.fromCodePoint(codePoint);
      } catch {
        return match;
      }
    }
    return NAMED_HTML_ENTITIES[entity] ?? match;
  });
}

const DEFAULT_HTML_TEXT_MAX_LENGTH = 50_000;

/**
 * Basic HTML -> plain text conversion used as a fallback when a message has
 * no text/plain part: drops <script>/<style> blocks entirely, turns <br>
 * into newlines, strips all remaining tags, decodes common HTML entities,
 * and caps the result length (with a note) so a huge HTML email can't blow
 * past response size limits.
 */
export function htmlToText(html: string, maxLength = DEFAULT_HTML_TEXT_MAX_LENGTH): string {
  let text = html.replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, "");
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<[^>]+>/g, "");
  text = decodeHtmlEntities(text);
  text = text
    .replace(/[ \t]+/g, " ")
    .replace(/[ \t]*\n[ \t]*/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  if (text.length > maxLength) {
    const shown = text.slice(0, maxLength);
    return `${shown}\n\n[truncated: showing ${maxLength} of ${text.length} characters]`;
  }
  return text;
}

export function formatMessage(msg: gmail_v1.Schema$Message): Record<string, unknown> {
  const headers = msg.payload?.headers;
  const body = extractBody(msg.payload);
  const attachments = extractAttachments(msg.payload);
  return {
    id: msg.id,
    threadId: msg.threadId,
    labelIds: msg.labelIds,
    snippet: msg.snippet,
    subject: getHeader(headers, "subject"),
    from: getHeader(headers, "from"),
    to: getHeader(headers, "to"),
    cc: getHeader(headers, "cc"),
    date: getHeader(headers, "date"),
    body: body.text || (body.html ? htmlToText(body.html) : ""),
    ...(attachments.length > 0 ? { attachments } : {}),
  };
}
