import { HostBridge } from '../bridge/host-bridge.js';
import { PluginBridge } from '../bridge/plugin-bridge.js';
import { EditorService } from '../services/editor-service.js';
import { SdtService } from '../services/sdt-service.js';
import { SelectionService } from '../services/selection-service.js';
import { CommandBus } from './command-bus.js';
import { EventBus } from './event-bus.js';

export function createContainer() {
  const host = new HostBridge();
  const plugin = new PluginBridge();

  const editor = new EditorService(plugin);
  const sdt = new SdtService(editor);
  const selection = new SelectionService(plugin, sdt);

  const events = new EventBus({ hostBridge: host });
  const commands = new CommandBus({ sdtService: sdt });

  return { host, plugin, editor, sdt, selection, events, commands };
}
