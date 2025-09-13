export class EditorService {
    constructor(pluginBridge) {
      this.plugin = pluginBridge;
    }
  
    /**
     * 在文档上下文执行：提供一个回调函数，其内可调用 Api.*
     */
    runInDoc(fn, { async = false, cb } = {}) {
      return this.plugin.callCommand(fn, async, cb);
    }
  
    /** 读取当前选择的 Range（在文档上下文内调用）*/
    // 注意：此方法需放在 runInDoc 的回调里执行 Api.*，此处仅做封装示意
  }
  