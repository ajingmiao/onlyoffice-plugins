import { COMMANDS } from '../core/constants.js';

export class CommandBus {
  constructor({ sdtService }) {
    this.sdt = sdtService;
  }

  async dispatch(cmd) {
    const { command, data } = cmd || {};
    switch (command) {
      case COMMANDS.INSERT_TEXT: {
        // 兼容字符串或 {text:'...'}，将其视为 tagKey
        const payload = data;
        const text = (typeof payload === 'string') ? payload : (payload && payload.text) || '';
        const tagKey = text || 'customer_name';

        await this.sdt.insertInlineSdt({
          tagKey,
          title: '客户姓名',
          placeholder: `{{${tagKey}}}`,
        });
        return { ok: true };
      }

      case COMMANDS.REPORT_ACTIVE_SDT: {
        const info = await this.sdt.detectActiveSdt();
        return { ok: true, data: info };
      }

      default:
        return { ok: false, error: 'Unknown command' };
    }
  }
}
