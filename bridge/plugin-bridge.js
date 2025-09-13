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
    // 1. 首先尝试使用 onClick 事件（根据官方文档）
    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    window.Asc.plugin.event_onClick = function(isSelectionUse) {
      logger.info('🔥 onClick event triggered!', { isSelectionUse });

      // 获取当前内容控件属性
      window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function(obj) {
        logger.info('Current content control properties:', obj);
        if (obj && obj.Tag && obj.Tag.startsWith('link-data:')) {
          logger.info('Link content control clicked!');
        }
        cb?.();
      });
    };

    // 2. 使用正确的OnlyOffice API - connector.attachEvent
    if (window.connector && typeof window.connector.attachEvent === 'function') {
      logger.info('Using connector.attachEvent for content control changes');

      // 监听内容控件变化事件
      window.connector.attachEvent("onChangeContentControl", (obj) => {
        logger.info('🔥 Content control changed via connector:', obj);
        cb?.();
      });

      // 也监听目标位置变化（光标移动）
      window.connector.attachEvent("onTargetPositionChanged", (obj) => {
        logger.info('🔥 Target position changed via connector:', obj);
        cb?.();
      });

    } else {
      // 3. 降级到原有的事件机制
      logger.info('Connector not available, using legacy event handlers');

      window.Asc.plugin.event_onSelectionChanged = () => {
        logger.info('🔥 Selection changed event triggered!');
        cb?.();
      };

      window.Asc.plugin.event_onContentControlClick = (data) => {
        logger.info('🔥 Content control clicked!', data);
        cb?.();
      };

      // 确保事件处理器已正确设置
      logger.info('onClick handler set:', typeof window.Asc.plugin.event_onClick);
      logger.info('Selection change handler set:', typeof window.Asc.plugin.event_onSelectionChanged);
      logger.info('Content control click handler set:', typeof window.Asc.plugin.event_onContentControlClick);
    }
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
