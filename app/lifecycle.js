import { EVENTS, COMMANDS } from '../core/constants.js';
import { logger } from '../core/logger.js';

export function bootstrap({ host, plugin, selection, events, commands }) {
  // Asc 插件 init → 通知宿主已就绪
  plugin.onInit(() => {
    events.emit(EVENTS.PLUGIN_INITIALIZED, { guid: window.Asc?.plugin?.guid });
  });

  // 选择变化 → 回传当前 SDT 信息（供宿主显示）
  selection.bindSelectionChange(async (res) => {
    logger.info('=== Selection changed ===');
    logger.info('Selection result:', res);
    events.emit(EVENTS.SELECTION_CHANGED_FIRED);
    events.emit(EVENTS.ACTIVE_SDT, res);

    // 添加延迟，确保选择已经稳定
    setTimeout(async () => {
      try {
        // 检查链接控件
        logger.info('Checking for link control after selection change...');
        const linkResult = await commands.dispatch({ command: COMMANDS.LINK_CLICKED });
        logger.info('Link check result:', linkResult);

        if (linkResult.ok && linkResult.data) {
          logger.info('Link control detected, sending to host:', linkResult.data);
          host.sendInfo('linkClicked', linkResult.data);
        } else {
          logger.info('No link control found at current selection');
        }

        // 检查表格点击
        logger.info('Checking for table click after selection change...');
        const tableResult = await commands.dispatch({ command: COMMANDS.TABLE_CLICKED });
        logger.info('Table check result:', tableResult);

        if (tableResult.ok && tableResult.data) {
          logger.info('Table click detected, sending to host:', tableResult.data);
          host.sendInfo('tableClicked', tableResult.data);
        } else {
          logger.info('No table found at current selection');
        }
      } catch (e) {
        logger.info('Selection change processing error:', e.message);
      }
    }, 100); // 100ms 延迟
  });

  // 超链接点击事件处理
  plugin.onHyperLinkClick(async (data) => {
    logger.info('hyperlink clicked:', data);
    // 当点击超链接时，检查是否是我们的数据绑定链接
    const result = await commands.dispatch({ command: COMMANDS.LINK_CLICKED });
    if (result.ok && result.data) {
      // 将链接数据发送回宿主页面
      host.sendInfo('linkClicked', result.data);
    }
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
