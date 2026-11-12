// @vitest-environment jsdom

import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { act } from "react";
import type { Root } from "react-dom/client";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { TeamsAppState } from "./TeamsAppCard";

vi.mock("../lib/rpc", () => ({
  withSpaceHeaders: () => new Headers(),
  recheckSessionOnUnauthorized: () => undefined,
}));
vi.mock("@lingui/react/macro", () => {
  const t = (parts: TemplateStringsArray, ...values: unknown[]) =>
    parts.reduce((text, part, index) => `${text}${index > 0 ? values[index - 1] : ""}${part}`, "");
  return { useLingui: () => ({ t }), Trans: ({ children }: { children: ReactNode }) => children };
});

import { TEAMS_APP_POLL_MS, TeamsAppCard } from "./TeamsAppCard";

type App = NonNullable<TeamsAppState["app"]>;

function app(overrides: Partial<App> = {}): App {
  return {
    status: "creating",
    step: "creating",
    label: "Creating",
    estimated: false,
    chatCount: 0,
    chats: [],
    lastError: null,
    userMessage: null,
    userMessageKind: null,
    publishedAt: null,
    firstInstalledAt: null,
    ...overrides,
  };
}

let container: HTMLDivElement;
let root: Root;
let state: TeamsAppState;
const calls: { url: string; method: string }[] = [];

beforeEach(() => {
  vi.stubGlobal("IS_REACT_ACT_ENVIRONMENT", true);
  calls.length = 0;
  state = { enabled: true, app: null };
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const method = init?.method ?? "GET";
      calls.push({ url, method });
      return new Response(JSON.stringify(method === "GET" ? state : { accepted: true }), {
        status: method === "GET" ? 200 : 202,
      });
    }),
  );
  container = document.createElement("div");
  document.body.appendChild(container);
  root = createRoot(container);
});

afterEach(async () => {
  await act(async () => root.unmount());
  container.remove();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

async function render() {
  const client = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  await act(async () => {
    root.render(
      <QueryClientProvider client={client}>
        <TeamsAppCard botId="bot-1" botName="Tony" />
      </QueryClientProvider>,
    );
  });
  await settle();
}

async function settle() {
  for (let tick = 0; tick < 5; tick += 1) {
    await act(async () => new Promise((resolve) => setTimeout(resolve, 0)));
  }
}

const text = () => document.body.textContent ?? "";
const stepStatuses = () =>
  [...container.querySelectorAll('[data-slot="stepper-step"]')].map((step) =>
    step.getAttribute("data-status"),
  );
const button = (label: string) =>
  [...document.querySelectorAll("button")].find((candidate) =>
    candidate.textContent?.includes(label),
  );

describe("TeamsAppCard", () => {
  it("renders nothing when the server flag is off", async () => {
    state = { enabled: false, app: null };
    await render();
    expect(container.querySelector('[data-testid="teams-app-card"]')).toBeNull();
  });

  it("offers to create the app and starts provisioning", async () => {
    await render();
    const create = button("Give Tony its own Teams app");
    expect(create).toBeDefined();
    await act(async () => create?.click());
    expect(calls).toContainEqual({ url: "/api/bots/bot-1/teams-app", method: "POST" });
  });

  it.each([
    ["creating", ["current", "upcoming", "upcoming", "upcoming", "upcoming"]],
    ["registered", ["complete", "current", "upcoming", "upcoming", "upcoming"]],
    ["published", ["complete", "complete", "current", "upcoming", "upcoming"]],
  ] as const)("highlights the %s step", async (step, statuses) => {
    state = { enabled: true, app: app({ status: step, step }) };
    await render();
    expect(stepStatuses()).toEqual(statuses);
    expect(container.querySelector('[aria-current="step"]')).not.toBeNull();
  });

  it("marks @mention visibility as estimated until the first install", async () => {
    state = { enabled: true, app: app({ status: "visible", step: "visible", estimated: true }) };
    await render();
    const estimated = container.querySelector('[data-testid="teams-app-estimated"]');
    expect(estimated?.textContent).toContain("(estimated)");
    expect(text()).toContain(
      "Teams has no signal for this. Estimated from catalog publish time; confirmed when Tony is first added to a chat.",
    );
    expect(stepStatuses()).toEqual(["complete", "complete", "complete", "current", "upcoming"]);
  });

  it("drops the estimate once confirmed and lists every chat", async () => {
    state = {
      enabled: true,
      app: app({
        status: "visible",
        step: "in_chats",
        estimated: false,
        firstInstalledAt: "2026-11-11T11:30:00.000Z",
        chatCount: 2,
        chats: [
          { chatId: "chat-a", installedAt: "2026-11-11T11:30:00.000Z" },
          { chatId: "chat-b", installedAt: "2026-11-11T11:40:00.000Z", topic: "Leasing" },
        ],
      }),
    };
    await render();
    expect(container.querySelector('[data-testid="teams-app-estimated"]')).toBeNull();
    expect(text()).toContain("In 2 chats");
    const chats = container.querySelector('[data-testid="teams-app-chats"]');
    expect(chats?.textContent).toContain("chat-a");
    expect(chats?.textContent).toContain("Leasing");
    expect(stepStatuses()).toEqual(["complete", "complete", "complete", "complete", "complete"]);
  });

  it("shows the failure and retries", async () => {
    state = {
      enabled: true,
      app: app({
        status: "failed",
        step: "registered",
        lastError: "read app role assignments failed (404 Request_ResourceNotFound): gone",
        userMessage: "Microsoft is still setting up this agent's app.",
        userMessageKind: "replicating",
      }),
    };
    await render();
    expect(container.querySelector('[data-testid="teams-app-error"]')?.textContent).toBe(
      "Microsoft is still setting up Tony's app. This usually takes under a minute — Retry.",
    );
    // The raw detail sits behind a collapsed "Details" disclosure, not in the message.
    expect(container.querySelector('[data-testid="teams-app-error"]')?.textContent).not.toContain(
      "Request_ResourceNotFound",
    );
    expect(container.textContent).toContain("Details");
    await act(async () => button("Retry")?.click());
    expect(calls).toContainEqual({ url: "/api/bots/bot-1/teams-app/retry", method: "POST" });
  });

  it("shows the provisioner's failed step with its error and a retry, without a stepper", async () => {
    state = {
      enabled: true,
      app: app({
        status: "failed",
        step: "failed",
        lastError: "create Azure Bot failed (409)",
        userMessage: "Something went wrong setting up Teams. Retry, or ask an admin.",
        userMessageKind: "unknown",
      }),
    };
    await render();
    expect(stepStatuses()).toEqual([]);
    expect(container.querySelector('[data-testid="teams-app-error"]')?.textContent).toBe(
      "Something went wrong setting up Teams. Retry, or ask an admin.",
    );
    expect(button("Retry")).toBeDefined();
  });

  it("shows removal in progress and offers no second removal", async () => {
    state = { enabled: true, app: app({ status: "removing", step: "removing", chatCount: 0 }) };
    await render();
    expect(stepStatuses()).toEqual([]);
    expect(text()).toContain("Removing…");
    expect(button("Remove from Teams")).toBeUndefined();
  });

  it("offers to create the app again once removed", async () => {
    state = { enabled: true, app: app({ status: "removed", step: "removed" }) };
    await render();
    expect(container.querySelector('[data-testid="teams-app-removed"]')).not.toBeNull();
    expect(button("Remove from Teams")).toBeUndefined();
    await act(async () => button("Give Tony its own Teams app")?.click());
    expect(calls).toContainEqual({ url: "/api/bots/bot-1/teams-app", method: "POST" });
  });

  it("confirms removal in a dialog before deleting", async () => {
    state = { enabled: true, app: app({ status: "visible", step: "visible" }) };
    await render();
    await act(async () => button("Remove from Teams")?.click());
    await settle();
    expect(text()).toContain(
      "Tony leaves every Teams chat it's in; Bot joins those chats instead.",
    );
    expect(calls.some((call) => call.method === "DELETE")).toBe(false);
    const confirm = [...document.querySelectorAll("button")].find(
      (candidate) => candidate.textContent === "Remove",
    );
    await act(async () => confirm?.click());
    expect(calls).toContainEqual({ url: "/api/bots/bot-1/teams-app", method: "DELETE" });
  });

  it("polls every 10s while provisioning and stops once settled", async () => {
    vi.useFakeTimers({ shouldAdvanceTime: true });
    state = { enabled: true, app: app({ status: "published", step: "published" }) };
    await render();
    const gets = () => calls.filter((call) => call.method === "GET").length;
    const before = gets();
    await act(async () => vi.advanceTimersByTimeAsync(TEAMS_APP_POLL_MS));
    expect(gets()).toBe(before + 1);
    state = { enabled: true, app: app({ status: "visible", step: "visible" }) };
    await act(async () => vi.advanceTimersByTimeAsync(TEAMS_APP_POLL_MS));
    const settled = gets();
    await act(async () => vi.advanceTimersByTimeAsync(TEAMS_APP_POLL_MS * 3));
    expect(gets()).toBe(settled);
  });
});
