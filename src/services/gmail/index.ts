import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { google } from "googleapis";
import { z } from "zod";
import { ServiceContext } from "../../types.js";
import { textResult } from "../../utils/formatting.js";
import { buildRawEmail, encodeBase64Url, extractBody, formatMessage, getHeader } from "../../utils/email.js";

const FILTER_TEMPLATES: Record<string, { criteria: Record<string, unknown>; action: Record<string, unknown> }> = {
  newsletter: { criteria: { query: "unsubscribe" }, action: { removeLabelIds: ["INBOX"] } },
  social_notifications: { criteria: { from: "notification" }, action: { removeLabelIds: ["INBOX"] } },
  auto_archive_noreply: { criteria: { from: "noreply" }, action: { removeLabelIds: ["INBOX"] } },
};

export function registerGmailTools(server: McpServer, ctx: ServiceContext): void {
  const api = () => google.gmail({ version: "v1", auth: ctx.auth });

  server.tool("gmail_search_emails", "Search emails using Gmail search syntax", {
    query: z.string().describe("Gmail search query (e.g., 'from:example@gmail.com')"),
    maxResults: z.number().optional().describe("Maximum number of results to return"),
  }, async ({ query, maxResults }) => {
    const gmail = api();
    const res = await gmail.users.messages.list({ userId: "me", q: query, maxResults: maxResults || 10 });
    if (!res.data.messages?.length) return textResult("No messages found.");

    const messages = await Promise.all(
      res.data.messages.map(async (m) => {
        const full = await gmail.users.messages.get({ userId: "me", id: m.id!, format: "metadata", metadataHeaders: ["From", "To", "Subject", "Date"] });
        return {
          id: full.data.id,
          threadId: full.data.threadId,
          snippet: full.data.snippet,
          subject: getHeader(full.data.payload?.headers, "subject"),
          from: getHeader(full.data.payload?.headers, "from"),
          date: getHeader(full.data.payload?.headers, "date"),
        };
      })
    );
    return textResult(messages);
  });

  server.tool("gmail_read_email", "Read the full content of an email", {
    messageId: z.string().describe("ID of the email message to retrieve"),
  }, async ({ messageId }) => {
    const res = await api().users.messages.get({ userId: "me", id: messageId, format: "full" });
    return textResult(formatMessage(res.data));
  });

  server.tool("gmail_send_email", "Send an email", {
    to: z.array(z.string()).describe("Recipient email addresses"),
    subject: z.string().describe("Email subject"),
    body: z.string().describe("Email body (plain text)"),
    htmlBody: z.string().optional().describe("HTML version of the email body"),
    cc: z.array(z.string()).optional().describe("CC recipients"),
    bcc: z.array(z.string()).optional().describe("BCC recipients"),
    mimeType: z.enum(["text/plain", "text/html", "multipart/alternative"]).optional().default("text/plain"),
    threadId: z.string().optional().describe("Thread ID to reply to"),
    inReplyTo: z.string().optional().describe("Message ID being replied to"),
  }, async (opts) => {
    const raw = encodeBase64Url(buildRawEmail(opts));
    const res = await api().users.messages.send({
      userId: "me",
      requestBody: { raw, threadId: opts.threadId },
    });
    return textResult({ id: res.data.id, threadId: res.data.threadId, labelIds: res.data.labelIds });
  });

  server.tool("gmail_draft_email", "Create an email draft", {
    to: z.array(z.string()).describe("Recipient email addresses"),
    subject: z.string().describe("Email subject"),
    body: z.string().describe("Email body"),
    htmlBody: z.string().optional().describe("HTML version of the email body"),
    cc: z.array(z.string()).optional().describe("CC recipients"),
    bcc: z.array(z.string()).optional().describe("BCC recipients"),
    mimeType: z.enum(["text/plain", "text/html", "multipart/alternative"]).optional().default("text/plain"),
    threadId: z.string().optional().describe("Thread ID to reply to"),
    inReplyTo: z.string().optional().describe("Message ID being replied to"),
  }, async (opts) => {
    const raw = encodeBase64Url(buildRawEmail(opts));
    const res = await api().users.drafts.create({
      userId: "me",
      requestBody: { message: { raw, threadId: opts.threadId } },
    });
    return textResult({ draftId: res.data.id, messageId: res.data.message?.id, threadId: res.data.message?.threadId });
  });

  server.tool("gmail_modify_email", "Modify email labels (add/remove)", {
    messageId: z.string().describe("ID of the message to modify"),
    addLabelIds: z.array(z.string()).optional().describe("Label IDs to add"),
    removeLabelIds: z.array(z.string()).optional().describe("Label IDs to remove"),
  }, async ({ messageId, addLabelIds, removeLabelIds }) => {
    const res = await api().users.messages.modify({
      userId: "me", id: messageId,
      requestBody: { addLabelIds: addLabelIds || [], removeLabelIds: removeLabelIds || [] },
    });
    return textResult({ id: res.data.id, labelIds: res.data.labelIds });
  });

  server.tool("gmail_delete_email", "Move an email to trash", {
    messageId: z.string().describe("ID of the message to trash"),
  }, async ({ messageId }) => {
    await api().users.messages.trash({ userId: "me", id: messageId });
    return textResult({ success: true, messageId });
  });

  server.tool("gmail_batch_delete_emails", "Permanently delete multiple emails", {
    messageIds: z.array(z.string()).describe("IDs of messages to delete"),
  }, async ({ messageIds }) => {
    await api().users.messages.batchDelete({ userId: "me", requestBody: { ids: messageIds } });
    return textResult({ success: true, count: messageIds.length });
  });

  server.tool("gmail_batch_modify_emails", "Modify labels on multiple emails", {
    messageIds: z.array(z.string()).describe("IDs of messages to modify"),
    addLabelIds: z.array(z.string()).optional().describe("Label IDs to add"),
    removeLabelIds: z.array(z.string()).optional().describe("Label IDs to remove"),
  }, async ({ messageIds, addLabelIds, removeLabelIds }) => {
    await api().users.messages.batchModify({
      userId: "me",
      requestBody: { ids: messageIds, addLabelIds: addLabelIds || [], removeLabelIds: removeLabelIds || [] },
    });
    return textResult({ success: true, count: messageIds.length });
  });

  server.tool("gmail_create_filter", "Create a Gmail filter with custom criteria and actions", {
    criteria: z.object({
      from: z.string().optional(),
      to: z.string().optional(),
      subject: z.string().optional(),
      query: z.string().optional(),
      negatedQuery: z.string().optional(),
      hasAttachment: z.boolean().optional(),
      excludeChats: z.boolean().optional(),
      size: z.number().optional(),
      sizeComparison: z.enum(["unspecified", "smaller", "larger"]).optional(),
    }).describe("Filter matching criteria"),
    action: z.object({
      addLabelIds: z.array(z.string()).optional(),
      removeLabelIds: z.array(z.string()).optional(),
      forward: z.string().optional(),
    }).describe("Actions to perform on matching emails"),
  }, async ({ criteria, action }) => {
    const res = await api().users.settings.filters.create({
      userId: "me",
      requestBody: { criteria, action },
    });
    return textResult({ id: res.data.id, criteria: res.data.criteria, action: res.data.action });
  });

  server.tool("gmail_create_filter_from_template", "Create a filter from a predefined template", {
    template: z.enum(["newsletter", "social_notifications", "auto_archive_noreply"]).describe("Template name"),
    customizations: z.object({
      from: z.string().optional(),
      to: z.string().optional(),
      subject: z.string().optional(),
      query: z.string().optional(),
      addLabelIds: z.array(z.string()).optional(),
      removeLabelIds: z.array(z.string()).optional(),
      forward: z.string().optional(),
    }).optional().describe("Customizations to apply on top of the template"),
  }, async ({ template, customizations }) => {
    const tpl = FILTER_TEMPLATES[template];
    const criteria = { ...tpl.criteria, ...customizations };
    const action: Record<string, unknown> = { ...tpl.action };
    if (customizations?.addLabelIds) action.addLabelIds = customizations.addLabelIds;
    if (customizations?.removeLabelIds) action.removeLabelIds = customizations.removeLabelIds;
    if (customizations?.forward) action.forward = customizations.forward;

    const res = await api().users.settings.filters.create({ userId: "me", requestBody: { criteria, action } });
    return textResult({ id: res.data.id, template, criteria: res.data.criteria, action: res.data.action });
  });

  server.tool("gmail_delete_filter", "Delete a Gmail filter", {
    filterId: z.string().describe("ID of the filter to delete"),
  }, async ({ filterId }) => {
    await api().users.settings.filters.delete({ userId: "me", id: filterId });
    return textResult({ success: true, filterId });
  });

  server.tool("gmail_get_filter", "Get details of a Gmail filter", {
    filterId: z.string().describe("ID of the filter to retrieve"),
  }, async ({ filterId }) => {
    const res = await api().users.settings.filters.get({ userId: "me", id: filterId });
    return textResult(res.data);
  });

  server.tool("gmail_list_filters", "List all Gmail filters", {}, async () => {
    const res = await api().users.settings.filters.list({ userId: "me" });
    return textResult(res.data.filter || []);
  });

  server.tool("gmail_create_label", "Create a Gmail label", {
    name: z.string().describe("Label name"),
    messageListVisibility: z.enum(["show", "hide"]).optional(),
    labelListVisibility: z.enum(["labelShow", "labelShowIfUnread", "labelHide"]).optional(),
  }, async ({ name, messageListVisibility, labelListVisibility }) => {
    const res = await api().users.labels.create({
      userId: "me",
      requestBody: { name, messageListVisibility, labelListVisibility },
    });
    return textResult({ id: res.data.id, name: res.data.name });
  });

  server.tool("gmail_delete_label", "Delete a Gmail label", {
    labelId: z.string().describe("ID of the label to delete"),
  }, async ({ labelId }) => {
    await api().users.labels.delete({ userId: "me", id: labelId });
    return textResult({ success: true, labelId });
  });

  server.tool("gmail_update_label", "Update a Gmail label", {
    labelId: z.string().describe("ID of the label to update"),
    name: z.string().optional().describe("New label name"),
    messageListVisibility: z.enum(["show", "hide"]).optional(),
    labelListVisibility: z.enum(["labelShow", "labelShowIfUnread", "labelHide"]).optional(),
  }, async ({ labelId, name, messageListVisibility, labelListVisibility }) => {
    const res = await api().users.labels.update({
      userId: "me", id: labelId,
      requestBody: { name, messageListVisibility, labelListVisibility },
    });
    return textResult({ id: res.data.id, name: res.data.name });
  });

  server.tool("gmail_list_labels", "List all Gmail labels", {}, async () => {
    const res = await api().users.labels.list({ userId: "me" });
    return textResult(res.data.labels?.map((l) => ({ id: l.id, name: l.name, type: l.type })) || []);
  });

  server.tool("gmail_get_or_create_label", "Get a label by name, creating it if it doesn't exist", {
    name: z.string().describe("Label name"),
  }, async ({ name }) => {
    const gmail = api();
    const labels = await gmail.users.labels.list({ userId: "me" });
    const existing = labels.data.labels?.find((l) => l.name?.toLowerCase() === name.toLowerCase());
    if (existing) return textResult({ id: existing.id, name: existing.name, created: false });

    const res = await gmail.users.labels.create({ userId: "me", requestBody: { name } });
    return textResult({ id: res.data.id, name: res.data.name, created: true });
  });

  server.tool("gmail_download_attachment", "Download an email attachment", {
    messageId: z.string().describe("ID of the message containing the attachment"),
    attachmentId: z.string().describe("ID of the attachment to download"),
  }, async ({ messageId, attachmentId }) => {
    const res = await api().users.messages.attachments.get({
      userId: "me", messageId, id: attachmentId,
    });
    return textResult({ data: res.data.data, size: res.data.size });
  });
}
