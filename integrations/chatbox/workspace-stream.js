// Some compatible providers split tool-call metadata across initial SSE chunks.
// Buffer only the incomplete header; never guess a name or execute partial input.
const gridpilotToolHeaders = new WeakMap();
function gridpilotToolDelta(delta, state, index) {
  if (state[index] != null) return delta.function ? delta : { ...delta, function: { arguments: '' } };
  let pending = gridpilotToolHeaders.get(state);
  if (!pending) { pending = new Map(); gridpilotToolHeaders.set(state, pending); }
  const previous = pending.get(index) || {};
  if (previous.id && delta.id && previous.id !== delta.id) throw Error('工具调用标识在流式传输中发生变化，请重新发送。');
  const combined = { ...previous, ...delta, id: delta.id || previous.id,
    function: { name: (previous.function?.name || '') + (delta.function?.name || ''),
      arguments: (previous.function?.arguments || '') + (delta.function?.arguments || '') } };
  if (combined.function.arguments.length > 1024 * 1024 || combined.function.name.length > 256 || pending.size > 64) throw Error('工具调用流超出处理上限。');
  if (combined.id && combined.function.name && combined.function.arguments.length) { pending.delete(index); return combined; }
  pending.set(index, combined); return null;
}
function gridpilotFinishToolDeltas(state) {
  const pending = gridpilotToolHeaders.get(state);
  gridpilotToolHeaders.delete(state);
  if (pending?.size) throw Error('模型返回了不完整的工具调用，未执行该操作，请重新发送。');
}
