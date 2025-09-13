sdkjs-plugins/demo-clean/
├─ config.json
├─ index.html
├─ plugin.js                  // 入口：仅负责装配与启动
├─ core/
│  ├─ logger.js               // 统一日志
│  ├─ constants.js            // 命令字、事件名、Key等
│  ├─ utils.js                // 小工具（防抖、tryWrap等）
├─ bridge/
│  ├─ host-bridge.js          // 宿主通信（Common.Gateway）
│  └─ plugin-bridge.js        // Asc 插件运行时适配（init/callCommand/事件）
├─ app/
│  ├─ command-bus.js          // 命令总线（派发 insertText 等）
│  ├─ event-bus.js            // 事件总线（选择变化等，供 UI/宿主订阅）
│  ├─ container.js            // 依赖注入/装配（极简）
│  └─ lifecycle.js            // 启动/销毁流程
├─ services/
│  ├─ editor-service.js       // 文档通用操作（GetDocument/Range等）
│  ├─ sdt-service.js          // 内容控件相关（创建/查找/标记）
│  └─ selection-service.js    // 选择变更感知与信息提取
└─ handlers/
   ├─ insert-text.js          // 用例：插入占位符（行内控件）
   └─ report-active-sdt.js    // 用例：回传当前激活SDT信息