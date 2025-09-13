export class SelectionService {
    constructor(pluginBridge, sdtService) {
      this.plugin = pluginBridge;
      this.sdt = sdtService;
    }
  
    bindSelectionChange(handler) {
      this.plugin.onSelectionChanged(async () => {
        const res = await this.sdt.detectActiveSdt();
        handler?.(res);
      });
    }
  }
  