import { logger } from '../core/logger.js';

export class PluginBridge {
  onInit(cb) {
    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};
    window.Asc.plugin.init = () => {
      logger.info('plugin init');
      cb?.();
    };
  }

  callCommand(fn, isAsync = false, callback) {
    return window.Asc?.plugin?.callCommand?.(fn, isAsync, callback);
  }

  onSelectionChanged(cb) {
    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    window.Asc.plugin.event_onClick = function(isSelectionUse) {
      logger.info('onClick event triggered!', { isSelectionUse });

      // 获取当前内容控件属性
      window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function(obj) {
        logger.info('Current content control properties:', obj);
        if (obj && obj.Tag && obj.Tag.startsWith('link-data:')) {
          logger.info('Link content control clicked!');
        }
        cb?.();
      });
    };

    // 使用 onSelectionChanged 事件（用于光标移动等）
    window.Asc.plugin.event_onSelectionChanged = () => {
      logger.info('Selection changed event triggered - running element detection');
      cb?.();
    };

    // 如果有 connector API，也使用它
    if (window.connector && typeof window.connector.attachEvent === 'function') {
      logger.info('Using connector.attachEvent for enhanced selection tracking');

      // 监听内容控件变化事件
      window.connector.attachEvent("onChangeContentControl", (obj) => {
        logger.info('Content control changed via connector:', obj);
        cb?.();
      });

      // 监听目标位置变化（光标移动）
      window.connector.attachEvent("onTargetPositionChanged", (obj) => {
        logger.info('Target position changed via connector:', obj);
        cb?.();
      });
    }

    logger.info('✅ Selection change handlers configured');
    logger.info('onClick handler set:', typeof window.Asc.plugin.event_onClick);
    logger.info('Selection change handler set:', typeof window.Asc.plugin.event_onSelectionChanged);
  }

  onHyperLinkClick(cb) {
    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};
    window.Asc.plugin.event_onHyperLinkClick = (data) => {
      logger.info('hyperlink click detected:', data);
      cb?.(data);
    };
  }
}
