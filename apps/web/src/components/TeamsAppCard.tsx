import { Trans, useLingui } from "@lingui/react/macro";
import type { StepperStep } from "@rakazo/ui-web";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  Button,
  SectionHeading,
  Spinner,
  Stepper,
} from "@rakazo/ui-web";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { useState } from "react";
import { recheckSessionOnUnauthorized, withSpaceHeaders } from "../lib/rpc";
import { InfoTooltip } from "./InfoTooltip";

type TeamsAppStatus =
  | "creating"
  | "registered"
  | "published"
  | "visible"
  | "removing"
  | "removed"
  | "failed";

type TeamsAppStep = "creating" | "registered" | "published" | "visible" | "in_chats";

export type TeamsAppState = {
  enabled: boolean;
  app: null | {
    status: TeamsAppStatus;
    step: TeamsAppStep;
    label: string;
    estimated: boolean;
    chatCount: number;
    chats: { chatId: string; installedAt: string; topic?: string }[];
    lastError: string | null;
    publishedAt: string | null;
    firstInstalledAt: string | null;
  };
};

const STEPS: readonly TeamsAppStep[] = [
  "creating",
  "registered",
  "published",
  "visible",
  "in_chats",
];
const POLLING: ReadonlySet<TeamsAppStatus> = new Set([
  "creating",
  "registered",
  "published",
  "removing",
]);
export const TEAMS_APP_POLL_MS = 10_000;

function endpoint(botId: string, suffix = "") {
  return `/api/bots/${encodeURIComponent(botId)}/teams-app${suffix}`;
}

async function request(url: string, method = "GET"): Promise<Response> {
  const response = await fetch(url, {
    method,
    credentials: "include",
    headers: withSpaceHeaders(),
  });
  recheckSessionOnUnauthorized(response.status);
  if (!response.ok) throw new Error(`Request failed (${response.status})`);
  return response;
}

/** The agent's own chat app: create it, follow it to the @mention list, see its chats, remove it. */
export function TeamsAppCard({ botId, botName }: { botId: string; botName: string }) {
  const { t } = useLingui();
  const queryClient = useQueryClient();
  const queryKey = ["teams-app", botId];
  const [confirmRemove, setConfirmRemove] = useState(false);

  const state = useQuery({
    queryKey,
    queryFn: async (): Promise<TeamsAppState> => (await request(endpoint(botId))).json(),
    refetchInterval: (query) => {
      const status = query.state.data?.app?.status;
      return status && POLLING.has(status) ? TEAMS_APP_POLL_MS : false;
    },
  });

  const action = useMutation({
    mutationFn: async (kind: "create" | "retry" | "remove") => {
      if (kind === "create") await request(endpoint(botId), "POST");
      else if (kind === "retry") await request(endpoint(botId, "/retry"), "POST");
      else await request(endpoint(botId), "DELETE");
    },
    onSettled: () => queryClient.invalidateQueries({ queryKey }),
  });

  if (!state.data?.enabled) return null;
  const app = state.data.app;
  const pending = action.isPending ? action.variables : null;

  return (
    <section className="mt-6" data-testid="teams-app-card">
      <SectionHeading variant="label" className="mb-2">
        <Trans>Microsoft Teams</Trans>
      </SectionHeading>
      <div className="space-y-4 rounded-xl bg-muted/50 p-4">
        {app ? (
          <>
            <TeamsAppSteps app={app} botName={botName} />
            {app.status === "failed" ? (
              <div className="flex items-start justify-between gap-3" role="alert">
                <p className="text-[13px] text-destructive">{app.lastError ?? t`Setup failed`}</p>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pending !== null}
                  onClick={() => action.mutate("retry")}
                >
                  {pending === "retry" ? <Spinner aria-hidden /> : null}
                  <Trans>Retry</Trans>
                </Button>
              </div>
            ) : null}
            <TeamsAppChats chats={app.chats} chatCount={app.chatCount} />
            {app.status === "removing" ? (
              <p className="text-[13px] text-muted-foreground text-shimmer">
                <Trans>Removing…</Trans>
              </p>
            ) : (
              <Button
                size="sm"
                variant="outline"
                disabled={pending !== null}
                onClick={() => setConfirmRemove(true)}
              >
                <Trans>Remove from Teams</Trans>
              </Button>
            )}
          </>
        ) : (
          <Button size="sm" disabled={pending !== null} onClick={() => action.mutate("create")}>
            {pending === "create" ? <Spinner aria-hidden /> : null}
            <Trans>Give {botName} its own Teams app</Trans>
          </Button>
        )}
        {action.isError ? (
          <p className="text-[13px] text-destructive">
            <Trans>Could not continue</Trans>
          </p>
        ) : null}
      </div>
      <AlertDialog open={confirmRemove} onOpenChange={setConfirmRemove}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>
              <Trans>Remove {botName} from Teams?</Trans>
            </AlertDialogTitle>
            <AlertDialogDescription>
              <Trans>
                {botName} leaves every Teams chat it's in; Bot joins those chats instead.
              </Trans>
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>
              <Trans>Cancel</Trans>
            </AlertDialogCancel>
            <AlertDialogAction
              variant="destructive"
              disabled={pending !== null}
              onClick={() => {
                setConfirmRemove(false);
                action.mutate("remove");
              }}
            >
              <Trans>Remove</Trans>
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </section>
  );
}

export function TeamsAppSteps({
  app,
  botName,
}: {
  app: NonNullable<TeamsAppState["app"]>;
  botName: string;
}) {
  const { t } = useLingui();
  const current = STEPS.indexOf(app.step);
  const titles: Record<TeamsAppStep, string> = {
    creating: t`Creating`,
    registered: t`Registered`,
    published: t`Published to catalog`,
    visible: t`Visible in @mentions`,
    in_chats: t`In ${app.chatCount} chats`,
  };
  const steps: StepperStep[] = STEPS.map((step, index) => ({
    id: step,
    status:
      // Once the agent is in a chat every step is done; the last one stays highlighted.
      index < current || (step === "in_chats" && app.chatCount > 0 && current === index)
        ? "complete"
        : index === current
          ? "current"
          : "upcoming",
    busy: index === current && app.status !== "failed" && POLLING.has(app.status),
    title:
      step === "visible" && app.estimated && index <= current ? (
        <span className="inline-flex items-center gap-1" data-testid="teams-app-estimated">
          {titles[step]}{" "}
          <span className="text-muted-foreground">
            <Trans>(estimated)</Trans>
          </span>
          <InfoTooltip label={titles[step]}>
            {t`Teams has no signal for this. Estimated from catalog publish time; confirmed when ${botName} is first added to a chat.`}
          </InfoTooltip>
        </span>
      ) : (
        titles[step]
      ),
  }));
  return <Stepper steps={steps} completeLabel={t`Done`} data-testid="teams-app-steps" />;
}

function TeamsAppChats({
  chats,
  chatCount,
}: {
  chats: NonNullable<TeamsAppState["app"]>["chats"];
  chatCount: number;
}) {
  if (chatCount === 0) return null;
  return (
    <div data-testid="teams-app-chats">
      <p className="text-[13px] font-medium">
        <Trans>In {chatCount} chats</Trans>
      </p>
      <ul className="mt-1 space-y-0.5 text-[13px] text-muted-foreground">
        {chats.map((chat) => (
          <li key={chat.chatId} className="truncate">
            {chat.topic ?? chat.chatId}
          </li>
        ))}
      </ul>
    </div>
  );
}
