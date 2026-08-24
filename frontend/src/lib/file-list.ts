export function mergeFileRecords<T extends { id: string }>(current: T[], incoming: T[]): T[] {
  const byId = new Map(current.map((record) => [record.id, record]));
  incoming.forEach((record) => byId.set(record.id, record));
  return Array.from(byId.values());
}
