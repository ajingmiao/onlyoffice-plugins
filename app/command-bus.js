import { COMMANDS } from '../core/constants.js';

export class CommandBus {
  constructor({ sdtService,linkService }) {
    this.sdt = sdtService;
    this.link = linkService;
  }

  async dispatch(cmd) {
    console.log('[CommandBus] 收到 INSERT_LINK 指令，参数 =', cmd);
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


      case COMMANDS.INSERT_LINK: {
        // 允许 data 是字符串（显示文本），或对象 { text, url, color, bold, underline }
        const text = (typeof data === 'string') ? data : (data?.text || '点击这里');
        const url  = (typeof data === 'object' && data?.url !== undefined) ? data.url : ''; // 允许空
        await this.link.insertLinkInline({
          text,
          url,
          style: {
            bold: (typeof data === 'object' && data?.bold !== undefined) ? !!data.bold : true,
            underline: (typeof data === 'object' && data?.underline !== undefined) ? !!data.underline : true,
            // 标准蓝：#1976D2（亦可传入 data.color）
            color: (typeof data === 'object' && data?.color) || '#1976D2'
          }
        });
        return { ok: true };
      }


      default:
        return { ok: false, error: 'Unknown command' };
    }
  }
}
