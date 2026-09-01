import type { AgentWorkItem } from "./api";

export type AgentItemGroup = {
  key: string;
  label: string;
  items: AgentWorkItem[];
};

export function agentItemGroupKey(item: AgentWorkItem): string {
  const category = String(item.category || "").trim();
  const field = String(item.field || "").trim();
  if (category) return category;
  if (field) return field;
  return "其他待确认事项";
}

export function groupAgentItems(items: AgentWorkItem[]): AgentItemGroup[] {
  const groups = new Map<string, AgentItemGroup>();
  for (const item of items) {
    const key = agentItemGroupKey(item);
    const existing = groups.get(key);
    if (existing) existing.items.push(item);
    else groups.set(key, { key, label: key, items: [item] });
  }
  return Array.from(groups.values());
}

export function selectedGroupItemIds(groups: AgentItemGroup[], key: string): string[] {
  return groups.find((group) => group.key === key)?.items.map((item) => item.item_id) || [];
}
