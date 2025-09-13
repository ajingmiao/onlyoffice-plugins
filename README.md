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



每个模块只做一件事，例如 sdt-service 只处理内容控件读写。

依赖倒置：应用层（handlers）依赖服务接口，不直接依赖底层 API；Asc API 统一通过 plugin-bridge 暴露。

显式边界：宿主通信集中在 host-bridge，避免业务代码到处 window.parent...。

可测性：handlers 是纯函数（或接近），方便未来做单元测试（可传入 mock service）。

无副作用启动：入口只做装配与订阅；真正业务都通过 CommandBus 驱动。




constants：新增命令常量

services：新建 xxx-service.js，只负责文档端 callCommand

container：装配并注入 new XxxService(editor)

command-bus：新增 case，把命令路由到新 Service

host（可选）：新增按钮或调用 serviceCommandSafe('yourCommand', data)

日志/回执：插件端打印日志并用 sendInfo('pluginAck' | 'pluginError') 回宿主

