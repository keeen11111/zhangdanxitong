export function isAgentProcessingInstruction(message: string): boolean {
  const normalized = message.trim();
  if (!normalized || /[？?]$/.test(normalized) || /(为什么|什么|哪些|是否|吗)$/.test(normalized)) return false;
  if (/^(继续|恢复|续跑)$/.test(normalized)) return true;
  const startsWithAction = /^(?:请|麻烦)?(?:继续|开始|重新|立即|直接|跳过|不要|无需|不必)?(?:进行|完成|处理|执行|更新|修改|写入|补充|删除|调整|清空|导入|同步|核对|生成|把|将|按)/.test(normalized);
  const skipAndAct = /(?:跳过|不要|无需|不必).*(?:直接|继续|进行|处理|执行|更新|修改|写入)/.test(normalized);
  return startsWithAction || skipAndAct;
}
