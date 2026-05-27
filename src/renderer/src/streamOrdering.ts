import type { StreamItem } from "../../shared/types";

export type ProcessStreamItem = Extract<StreamItem, { kind: "tool" | "file" | "stage" }>;
export type MessageStreamItem = Extract<StreamItem, { kind: "message" }>;

export type StreamDisplayBlock =
  | {
      id: string;
      kind: "message";
      item: MessageStreamItem;
    }
  | {
      id: string;
      kind: "process";
      items: ProcessStreamItem[];
    };

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

  for (const item of items.filter(shouldDisplayStreamItem)) {
    if (item.kind === "message" && item.role === "user" && turnItems.length) {
      flushTurn();
    }
    turnItems.push(item);
  }

  flushTurn();
  return ordered;
}

export function groupStreamItemsForDisplay(items: StreamItem[]): StreamDisplayBlock[] {
  const blocks: StreamDisplayBlock[] = [];
  let processItems: ProcessStreamItem[] = [];

  const flushProcessItems = (): void => {
    if (!processItems.length) return;
    blocks.push({
      id: `process_${processItems[0].id}`,
      kind: "process",
      items: processItems
    });
    processItems = [];
  };

  for (const item of orderStreamItemsForDisplay(items)) {
    if (item.kind === "message") {
      flushProcessItems();
      blocks.push({
        id: item.id,
        kind: "message",
        item
      });
      continue;
    }

    if (item.kind === "tool" || item.kind === "file" || item.kind === "stage") {
      processItems.push(item);
    }
  }

  flushProcessItems();
  return blocks;
}

function shouldDisplayStreamItem(item: StreamItem): boolean {
  if (item.kind === "tool" && item.toolName === "openai.chat.tools") return false;
  if (item.kind === "scheme_progress") return false;
  if (item.kind === "message" && item.role === "assistant" && item.isFinished && !item.content.trim()) {
    return false;
  }
  return true;
}
