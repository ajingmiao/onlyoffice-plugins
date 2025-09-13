export const COMMANDS = {
    INSERT_TEXT: 'insertText',      // 宿主下发：插入占位文本/控件
    REPORT_ACTIVE_SDT: 'reportActiveSdt'
  };
  
  export const EVENTS = {
    PLUGIN_INITIALIZED: 'pluginInitialized',
    ACTIVE_SDT: 'activeSdt',
    SELECTION_CHANGED_FIRED: 'onSelectionChanged-fired'
  };
  
  export const TAG_PREFIX = 'bind:'; // 统一为数据绑定控件加前缀
  