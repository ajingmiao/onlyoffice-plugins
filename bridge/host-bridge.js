import { logger } from '../core/logger.js';
import { EVENTS } from '../core/constants.js';

export class HostBridge {
  constructor(gateway = tryGetGateway()) {
    this.gateway = gateway;
  }

  sendInfo(command, data) {
    try {
      this.gateway?.sendInfo?.({ command, data });
    } catch (e) {  }
  }

  onInternalCommand(cb) {
    try {
      this.gateway?.on?.('internalcommand', (payload) => {
        cb?.(payload);
      });
    } catch (e) {
      logger.warn('Gateway.on unavailable');
    }
  }
}

function tryGetGateway() {
  try { return window.parent?.Common?.Gateway || null; } catch (e) { return null; }
}
