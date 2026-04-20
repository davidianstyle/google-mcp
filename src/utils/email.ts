import { gmail_v1 } from "googleapis";

export function decodeBase64Url(data: string): string {
  return Buffer.from(data.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf-8");
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

  headers.push(`To: ${opts.to.join(", ")}`);
  if (opts.from) headers.push(`From: ${opts.from}`);
  if (opts.cc?.length) headers.push(`Cc: ${opts.cc.join(", ")}`);
  if (opts.bcc?.length) headers.push(`Bcc: ${opts.bcc.join(", ")}`);
  headers.push(`Subject: ${opts.subject}`);
  if (opts.inReplyTo) {
    headers.push(`In-Reply-To: ${opts.inReplyTo}`);
    headers.push(`References: ${opts.references || opts.inReplyTo}`);
  }

  if (opts.htmlBody && opts.mimeType === "multipart/alternative") {
    headers.push(`MIME-Version: 1.0`);
    headers.push(`Content-Type: multipart/alternative; boundary="${boundary}"`);
    const parts = [
      `--${boundary}\r\nContent-Type: text/plain; charset="UTF-8"\r\n\r\n${opts.body}`,
      `--${boundary}\r\nContent-Type: text/html; charset="UTF-8"\r\n\r\n${opts.htmlBody}`,
      `--${boundary}--`,
    ];
    return headers.join("\r\n") + "\r\n\r\n" + parts.join("\r\n");
  }

  if (opts.htmlBody || opts.mimeType === "text/html") {
    headers.push(`MIME-Version: 1.0`);
    headers.push(`Content-Type: text/html; charset="UTF-8"`);
    return headers.join("\r\n") + "\r\n\r\n" + (opts.htmlBody || opts.body);
  }

  headers.push(`Content-Type: text/plain; charset="UTF-8"`);
  return headers.join("\r\n") + "\r\n\r\n" + opts.body;
}

export function formatMessage(msg: gmail_v1.Schema$Message): Record<string, unknown> {
  const headers = msg.payload?.headers;
  const body = extractBody(msg.payload);
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
    body: body.text || body.html,
  };
}
