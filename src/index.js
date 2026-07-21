#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { getAccessToken, getAuthStatus, signOut } from "./auth.js";
import {
  createDraft,
  getMessage,
  getProfile,
  listCalendarEvents,
  listMessages,
  searchMessages,
  sendMail,
} from "./graph.js";

function text(data) {
  return {
    content: [
      {
        type: "text",
        text: typeof data === "string" ? data : JSON.stringify(data, null, 2),
      },
    ],
  };
}

function fail(error) {
  const message = error instanceof Error ? error.message : String(error);
  return {
    isError: true,
    content: [{ type: "text", text: message }],
  };
}

const server = new McpServer({
  name: "outlook-mcp",
  version: "1.0.0",
});

server.tool(
  "outlook_login",
  "Start Microsoft device-code login for Outlook. Returns a URL and code (also printed to stderr). Open the URL, enter the code, and sign in.",
  {
    force: z
      .boolean()
      .optional()
      .default(false)
      .describe("Force a fresh login even if a cached token exists"),
  },
  async ({ force }) => {
    try {
      await getAccessToken({ forceLogin: force });
      const status = await getAuthStatus();
      const profile = await getProfile();
      return text({ message: "Signed in successfully.", status, profile });
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_auth_status",
  "Check whether Outlook/Microsoft Graph is signed in.",
  {},
  async () => {
    try {
      return text(await getAuthStatus());
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_sign_out",
  "Clear cached Microsoft tokens for this MCP server.",
  {},
  async () => {
    try {
      return text(await signOut());
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_get_profile",
  "Get the signed-in user's Outlook / Microsoft 365 profile.",
  {},
  async () => {
    try {
      return text(await getProfile());
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_list_emails",
  "List recent emails from a mail folder (default: inbox).",
  {
    folder: z
      .string()
      .optional()
      .default("inbox")
      .describe("Mail folder id or well-known name: inbox, sentitems, drafts, deleteditems"),
    top: z.number().int().min(1).max(50).optional().default(10),
    unreadOnly: z.boolean().optional().default(false),
  },
  async (args) => {
    try {
      return text(await listMessages(args));
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_get_email",
  "Get a single email by id, including body by default.",
  {
    id: z.string().describe("Message id from outlook_list_emails or outlook_search_emails"),
    includeBody: z.boolean().optional().default(true),
  },
  async (args) => {
    try {
      return text(await getMessage(args.id, { includeBody: args.includeBody }));
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_search_emails",
  "Search mailbox messages with a free-text query.",
  {
    query: z.string().min(1).describe("Search text, e.g. from:boss@contoso.com invoice"),
    top: z.number().int().min(1).max(50).optional().default(10),
  },
  async (args) => {
    try {
      return text(await searchMessages(args));
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_send_email",
  "Send an email immediately from the signed-in mailbox.",
  {
    to: z.array(z.string().email()).min(1),
    subject: z.string().min(1),
    body: z.string().min(1),
    cc: z.array(z.string().email()).optional().default([]),
    bodyType: z.enum(["Text", "HTML"]).optional().default("Text"),
  },
  async (args) => {
    try {
      return text(await sendMail(args));
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_create_draft",
  "Create an email draft in Outlook (does not send).",
  {
    to: z.array(z.string().email()).min(1),
    subject: z.string().min(1),
    body: z.string().min(1),
    cc: z.array(z.string().email()).optional().default([]),
    bodyType: z.enum(["Text", "HTML"]).optional().default("Text"),
  },
  async (args) => {
    try {
      return text(await createDraft(args));
    } catch (error) {
      return fail(error);
    }
  }
);

server.tool(
  "outlook_list_calendar",
  "List upcoming calendar events.",
  {
    days: z.number().int().min(1).max(60).optional().default(7),
    top: z.number().int().min(1).max(50).optional().default(20),
  },
  async (args) => {
    try {
      return text(await listCalendarEvents(args));
    } catch (error) {
      return fail(error);
    }
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
