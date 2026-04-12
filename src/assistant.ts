import type { OpenACPCore, Session } from "@openacp/plugin-sdk";

export async function spawnAssistant(
  core: OpenACPCore,
  threadId: string,
): Promise<{ session: Session; ready: Promise<void> }> {
  const readyResolve: (() => void)[] = [];
  const ready = new Promise<void>((resolve) => {
    readyResolve.push(resolve);
  });

  const session = await core.sessionManager.createSession(
    "teams",
    core.configManager.get().defaultAgent ?? "openacp",
    process.cwd(),
    core.agentManager,
  );

  // Set threadId so messages route to this session
  session.threadId = threadId;
  session.threadIds.set("teams", threadId);

  // Mark as assistant session
  (session as Session & { isAssistant?: boolean }).isAssistant = true;

  // Resolve ready when session reaches a terminal or usable state
  const onStatus = (_from: string, to: string) => {
    if (to === "active" || to === "finished" || to === "error" || to === "cancelled") {
      session.off("status_change", onStatus);
      session.off("error", onError);
      readyResolve.forEach((r) => r());
    }
  };
  const onError = () => {
    session.off("status_change", onStatus);
    session.off("error", onError);
    readyResolve.forEach((r) => r());
  };
  session.on("status_change", onStatus);
  session.on("error", onError);

  return { session, ready };
}
