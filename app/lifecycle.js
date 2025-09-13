import { EVENTS, COMMANDS } from '../core/constants.js';
import { logger } from '../core/logger.js';

export function bootstrap({ host, plugin, selection, events, commands }) {
  // Asc 插件 init → 通知宿主已就绪
  plugin.onInit(() => {
    events.emit(EVENTS.PLUGIN_INITIALIZED, { guid: window.Asc?.plugin?.guid });
  });

  // 选择变化 → 回传当前 SDT 信息（供宿主显示）
  selection.bindSelectionChange((res) => {
    logger.info('active SDT:', res);
    events.emit(EVENTS.SELECTION_CHANGED_FIRED);
    events.emit(EVENTS.ACTIVE_SDT, res);
  });

  // 宿主 → 内部命令（如 insertText）
  host.onInternalCommand(async (msg) => {
    // 约束：只处理我们关注的命令
    if (!msg?.command) return;
    //if (![COMMANDS.INSERT_TEXT, COMMANDS.REPORT_ACTIVE_SDT].includes(msg.command)) return;

    // 兼容字符串与 {text:'...'} 入参
    const data = msg.data;
    await commands.dispatch({ command: msg.command, data });
    // 可根据需要回执
    host.sendInfo('pluginAck', { op: msg.command, data });
  });
}
