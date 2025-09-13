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




修改总结

  1. 光标位置插入链接

  - 修改了 link-service.js:insertLinkInline() 方法，优化了链接插入逻辑
  - 使用 doc.GetRangeBySelect() 获取当前选择范围或光标位置
  - 确保链接在光标位置插入而非文档末尾

  2. JSON数据绑定功能

  - 在链接服务中添加了JSON数据支持，存储在内容控件的tag中
  - 为空链接创建内联级别的内容控件（SDT），使用 link-data: 前缀存储JSON数据
  - 更新了命令总线以传递JSON数据参数

  3. 链接点击事件处理

  - 在 constants.js 中添加了 LINK_CLICKED 命令
  - 在 plugin-bridge.js 中添加了 onHyperLinkClick() 方法监听超链接点击事件
  - 在 link-service.js 中添加了 handleLinkClick() 方法提取链接的JSON数据
  - 在 lifecycle.js 中设置了超链接点击监听器
  - 在 command-bus.js 中添加了链接点击命令处理

  4. 宿主页面通信

  - 更新了 App.vue 以监听 linkClicked 事件
  - 添加了 handleLinkClick() 函数处理从插件传回的JSON数据
  - 显示接收到的数据并可扩展为其他操作

  使用方式

  1. 插入链接: 点击"插入链接"按钮会在光标位置插入带有JSON数据绑定的可点击链接
  2. 点击链接: 点击文档中的链接会触发事件，将绑定的JSON数据传递回宿主页面
  3. 数据处理: 宿主页面会显示接收到的数据，你可以根据需要进行处理

  现在链接会在光标位置正确插入，点击时会将绑定的JSON数据（如 {name:'gm'}）传递回宿主页面进行处理。
  