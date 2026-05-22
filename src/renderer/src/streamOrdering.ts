import type { StreamItem } from "../../shared/types";

export function orderStreamItemsForDisplay(items: StreamItem[]): StreamItem[] {
  const ordered: StreamItem[] = [];
  let turnItems: StreamItem[] = [];

  const flushTurn = (): void => {
    if (!turnItems.length) return;

    const userMessages: StreamItem[] = [];
    const toolOutputs: StreamItem[] = [];
    const assistantMessages: StreamItem[] = [];

    for (const item of turnItems) {
      if (item.kind === "message" && item.role === "user") {
        userMessages.push(item);
      } else if (item.kind === "message" && item.role === "assistant") {
        assistantMessages.push(item);
      } else {
        toolOutputs.push(item);
      }
    }

    ordered.push(...userMessages, ...toolOutputs, ...assistantMessages);
    turnItems = [];
  };

  for (const item of items) {
    if (item.kind === "message" && item.role === "user" && turnItems.length) {
      flushTurn();
    }
    turnItems.push(item);
  }

  flushTurn();
  return ordered;
}
