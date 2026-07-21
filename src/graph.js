import { getAccessToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

async function graphFetch(pathname, { method = "GET", body, query } = {}) {
  const token = await getAccessToken();
  const url = new URL(`${GRAPH}${pathname}`);
  if (query) {
    for (const [key, value] of Object.entries(query)) {
      if (value !== undefined && value !== null && value !== "") {
        url.searchParams.set(key, String(value));
      }
    }
  }

  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
  };
  if (body !== undefined) {
    headers["Content-Type"] = "application/json";
  }

  const response = await fetch(url, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });

  const text = await response.text();
  let data = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { raw: text };
    }
  }

  if (!response.ok) {
    const detail =
      data?.error?.message ||
      data?.error_description ||
      text ||
      response.statusText;
    throw new Error(`Graph ${method} ${pathname} failed (${response.status}): ${detail}`);
  }

  return data;
}

function summarizeMessage(message) {
  return {
    id: message.id,
    subject: message.subject ?? "(no subject)",
    from: message.from?.emailAddress
      ? `${message.from.emailAddress.name || ""} <${message.from.emailAddress.address}>`.trim()
      : null,
    receivedDateTime: message.receivedDateTime,
    isRead: message.isRead,
    hasAttachments: message.hasAttachments,
    importance: message.importance,
    preview: message.bodyPreview,
  };
}

export async function getProfile() {
  return graphFetch("/me", {
    query: {
      $select: "id,displayName,mail,userPrincipalName,jobTitle,officeLocation",
    },
  });
}

export async function listMessages({
  folder = "inbox",
  top = 10,
  unreadOnly = false,
} = {}) {
  const filter = unreadOnly ? "isRead eq false" : undefined;
  const data = await graphFetch(`/me/mailFolders/${encodeURIComponent(folder)}/messages`, {
    query: {
      $top: Math.min(Math.max(top, 1), 50),
      $orderby: "receivedDateTime desc",
      $select:
        "id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview",
      ...(filter ? { $filter: filter } : {}),
    },
  });

  return {
    count: data.value?.length ?? 0,
    messages: (data.value ?? []).map(summarizeMessage),
  };
}

export async function getMessage(id, { includeBody = true } = {}) {
  const select = includeBody
    ? "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,importance,bodyPreview,body"
    : "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,importance,bodyPreview";

  const message = await graphFetch(`/me/messages/${encodeURIComponent(id)}`, {
    query: { $select: select },
  });

  return {
    ...summarizeMessage(message),
    to: (message.toRecipients ?? []).map((r) => r.emailAddress?.address).filter(Boolean),
    cc: (message.ccRecipients ?? []).map((r) => r.emailAddress?.address).filter(Boolean),
    body:
      includeBody && message.body
        ? {
            contentType: message.body.contentType,
            content: message.body.content,
          }
        : undefined,
  };
}

export async function searchMessages({ query, top = 10 }) {
  const data = await graphFetch("/me/messages", {
    query: {
      $search: `"${query.replaceAll('"', "")}"`,
      $top: Math.min(Math.max(top, 1), 50),
      $select:
        "id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview",
    },
  });

  return {
    count: data.value?.length ?? 0,
    messages: (data.value ?? []).map(summarizeMessage),
  };
}

function recipientList(addresses) {
  return addresses.map((address) => ({
    emailAddress: { address },
  }));
}

export async function sendMail({ to, subject, body, cc = [], bodyType = "Text" }) {
  await graphFetch("/me/sendMail", {
    method: "POST",
    body: {
      message: {
        subject,
        body: {
          contentType: bodyType,
          content: body,
        },
        toRecipients: recipientList(to),
        ...(cc.length ? { ccRecipients: recipientList(cc) } : {}),
      },
      saveToSentItems: true,
    },
  });

  return { sent: true, to, subject };
}

export async function createDraft({ to, subject, body, cc = [], bodyType = "Text" }) {
  const draft = await graphFetch("/me/messages", {
    method: "POST",
    body: {
      subject,
      body: {
        contentType: bodyType,
        content: body,
      },
      toRecipients: recipientList(to),
      ...(cc.length ? { ccRecipients: recipientList(cc) } : {}),
    },
  });

  return {
    id: draft.id,
    subject: draft.subject,
    webLink: draft.webLink,
  };
}

export async function listCalendarEvents({ days = 7, top = 20 } = {}) {
  const start = new Date();
  const end = new Date(start.getTime() + days * 24 * 60 * 60 * 1000);

  const data = await graphFetch("/me/calendarView", {
    query: {
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString(),
      $top: Math.min(Math.max(top, 1), 50),
      $orderby: "start/dateTime",
      $select: "id,subject,start,end,location,isAllDay,organizer,webLink",
    },
  });

  return {
    count: data.value?.length ?? 0,
    events: (data.value ?? []).map((event) => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName ?? null,
      isAllDay: event.isAllDay,
      organizer: event.organizer?.emailAddress?.address ?? null,
      webLink: event.webLink,
    })),
  };
}
