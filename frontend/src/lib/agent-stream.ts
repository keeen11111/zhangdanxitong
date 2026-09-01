export type AgentStreamEvent = {
  id?: string;
  type: string;
  payload: Record<string, unknown>;
};

export function parseSseFrames(buffer: string): {
  events: AgentStreamEvent[];
  remaining: string;
} {
  const normalized = buffer.replace(/\r\n/g, "\n");
  const frames = normalized.split("\n\n");
  const unfinished = frames.pop() || "";
  const remaining = unfinished.trim() ? unfinished : "";
  const events: AgentStreamEvent[] = [];

  for (const frame of frames) {
    const lines = frame.split("\n");
    const eventLine = lines.find((line) => line.startsWith("event:"));
    const idLine = lines.find((line) => line.startsWith("id:"));
    const dataLines = lines.filter((line) => line.startsWith("data:"));
    if (!eventLine || !dataLines.length) continue;
    try {
      const payload = JSON.parse(dataLines.map((line) => line.slice("data:".length).trimStart()).join("\n"));
      if (payload && typeof payload === "object" && !Array.isArray(payload)) {
        events.push({
          ...(idLine ? { id: idLine.slice("id:".length).trim() } : {}),
          type: eventLine.slice("event:".length).trim(),
          payload: payload as Record<string, unknown>,
        });
      }
    } catch {
      // Ignore malformed frames: the next complete server event is still usable.
    }
  }

  return { events, remaining };
}
