# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## OnlyOffice Plugin Architecture

This is a OnlyOffice Document Editor plugin with a clean, modular architecture. The plugin follows dependency inversion principles with clear separation of concerns.

### Core Architecture

The plugin uses a service-oriented architecture with these key layers:

1. **Entry Point** (`plugin.js`): Minimal bootstrap that creates DI container and starts lifecycle
2. **App Layer** (`app/`):
   - `container.js`: Dependency injection container
   - `lifecycle.js`: Plugin startup/shutdown flow
   - `command-bus.js`: Central command dispatcher
   - `event-bus.js`: Event system for UI/host communication
3. **Services** (`services/`):
   - `editor-service.js`: Document operations (GetDocument/Range)
   - `sdt-service.js`: Content control management (create/find/mark)
   - `selection-service.js`: Selection change detection
   - `link-service.js`: Link insertion functionality
4. **Bridge Layer** (`bridge/`):
   - `host-bridge.js`: Host communication via `Common.Gateway`
   - `plugin-bridge.js`: Asc plugin runtime adapter
5. **Core Utilities** (`core/`):
   - `constants.js`: Command names, event names, constants
   - `logger.js`: Unified logging system
   - `utils.js`: Utility functions (debounce, tryWrap, etc.)

### Key Design Patterns

- **Command Pattern**: All operations flow through `CommandBus.dispatch()`
- **Dependency Injection**: Services injected via simple container
- **Event-Driven**: Selection changes and plugin events via event bus
- **Bridge Pattern**: Host communication abstracted through bridges

### Available Commands

Commands are defined in `core/constants.js`:
- `INSERT_TEXT`: Insert placeholder content controls
- `REPORT_ACTIVE_SDT`: Report active content control info
- `INSERT_LINK`: Insert formatted hyperlinks

### Development Workflow

1. **Adding New Commands**:
   - Add constant to `core/constants.js`
   - Create service in `services/` if needed
   - Register service in `app/container.js`
   - Add case to `app/command-bus.js`
   - Optional: Add host button or call `serviceCommandSafe()`

2. **Plugin Communication**:
   - Commands flow: Host → `host-bridge.js` → `command-bus.js` → Services
   - Responses: Services → `command-bus.js` → `host-bridge.js` → Host
   - Use `sendInfo('pluginAck' | 'pluginError')` for host feedback

### Plugin Configuration

- **config.json**: Plugin metadata with GUID, version, supported editors
- **index.html**: Plugin UI container that loads the main module
- **Events**: Listens to `onSelectionChanged` and `onDocumentContentReady`

### Content Control System

The plugin specializes in content controls (SDT) with:
- Tag-based binding using `TAG_PREFIX = 'bind:'`
- Inline and block-level content controls
- Placeholder text with `{{tagKey}}` format
- Detection of active content controls

This architecture enables easy extension while maintaining clear boundaries between host communication, business logic, and document operations.