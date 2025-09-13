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
    window.Asc.plugin.event_onSelectionChanged = () => cb?.();
  }
}
