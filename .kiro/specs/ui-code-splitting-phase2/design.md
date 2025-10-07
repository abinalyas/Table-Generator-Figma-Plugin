# Design Document

## Overview

Phase 2 of the UI code splitting will focus on extracting the remaining large functions from ui.ts to achieve the target of under 2,000 lines. Based on analysis, ui.ts currently has 3,843 lines with several major areas that can be modularized:

1. **Message handling** (~400 lines) - Plugin communication logic
2. **Property rendering** (~800 lines) - Complex UI rendering functions  
3. **Event management** (~200 lines) - DOM event setup and handling
4. **State management** (~300 lines) - Validation and update functions
5. **Remaining utilities** (~200 lines) - Miscellaneous helper functions

The design will create 5 new components and refactor ui.ts to serve as a lightweight coordinator.

## Architecture

### Current State
```
ui.ts (3,843 lines)
├── Initialization logic
├── Message handling (onmessage, postMessage calls)
├── Property rendering functions
├── Event setup functions  
├── State management functions
├── Validation functions
├── Utility functions
└── Component coordination
```

### Target State
```
ui.ts (< 2,000 lines)
├── Component initialization
├── Basic coordination logic
└── Entry point setup

New Components:
├── MessageHandler.ts - Plugin communication
├── PropertyRenderer.ts - UI rendering logic
├── EventManager.ts - DOM event management
├── StateManager.ts - State updates and validation
└── RemainingUtils.ts - Miscellaneous utilities
```

## Components and Interfaces

### 1. MessageHandler Component

**Purpose:** Centralize all plugin communication logic

**Responsibilities:**
- Handle incoming messages from Figma plugin
- Route messages to appropriate components
- Send outgoing messages to plugin
- Manage message queuing and error handling

**Interface:**
```typescript
class MessageHandler {
  constructor(state: State, elements: Elements)
  
  // Message routing
  handleIncomingMessage(message: any): void
  sendMessage(type: string, data?: any): void
  
  // Specific message handlers
  private handleScanTable(data: any): void
  private handleComponentInfo(data: any): void
  private handleFakeData(data: any): void
  private handleWatsonxData(data: any): void
  
  // Message registration
  registerMessageHandler(type: string, handler: Function): void
}
```

**Functions to Extract:**
- `window.onmessage` handler (~130 lines)
- All `parent.postMessage` calls
- Message processing logic
- API key synchronization logic

### 2. PropertyRenderer Component

**Purpose:** Handle all property field rendering and UI generation

**Responsibilities:**
- Render dynamic property fields for different cell types
- Handle property field updates and validation
- Manage property field visibility and state
- Generate appropriate UI controls for different property types

**Interface:**
```typescript
class PropertyRenderer {
  constructor(state: State, elements: Elements)
  
  // Main rendering methods
  renderDynamicFields(availableProps: string[], propertyTypes: any, props: any): void
  renderBodyCellProperties(availableProps: string[], propertyTypes: any, props: any): void
  renderFromModel(model: any[], props: any): void
  
  // Field creation
  private createTextField(propName: string, value: any): HTMLElement
  private createSelectField(propName: string, value: any, options: string[]): HTMLElement
  private createCheckboxField(propName: string, value: boolean): HTMLElement
  
  // Utility methods
  updateStaticFieldsVisibility(availableProps: string[], propertyTypes: any, props: any): void
  findMatchingProperty(availableProperties: string[], uiPropName: string): string | null
}
```

**Functions to Extract:**
- `renderDynamicPropertyFields` (~165 lines)
- `renderDynamicPropertyFieldsFromModel` (~128 lines)
- `renderBodyCellProperties` (~69 lines)
- `updateStaticFieldsVisibilityAndValues` (~102 lines)
- `findMatchingProperty` (~25 lines)

### 3. EventManager Component

**Purpose:** Centralize DOM event setup and management

**Responsibilities:**
- Set up all DOM event listeners
- Manage event delegation and routing
- Handle event cleanup and memory management
- Coordinate between different event sources

**Interface:**
```typescript
class EventManager {
  constructor(state: State, elements: Elements)
  
  // Event setup
  setupAllEventListeners(): void
  setupFileUploadEvents(): void
  setupResizeEvents(): void
  setupApiKeyEvents(): void
  
  // Event handlers
  handleApplyOptionClick(element: HTMLElement): void
  
  // Event management
  addEventListener(element: HTMLElement, event: string, handler: Function): void
  removeAllEventListeners(): void
}
```

**Functions to Extract:**
- `setupEventListeners` (~122 lines)
- `setupFileUploadControls` (~19 lines)
- `setupResizeCorner` (~61 lines)
- `setupApiKeySync` (~61 lines)
- `handleApplyOptionClick` (~24 lines)

### 4. StateManager Component

**Purpose:** Handle state updates, validation, and UI synchronization

**Responsibilities:**
- Manage application state updates
- Validate state changes
- Synchronize UI with state changes
- Handle mode transitions and visibility updates

**Interface:**
```typescript
class StateManager {
  constructor(state: State, elements: Elements)
  
  // State management
  setMode(newMode: State['mode']): void
  updateModeDependentVisibility(): void
  getCellState(key: string): any
  
  // Reset and cleanup
  resetTableProperties(): void
  markChangesForReset(): void
  clearChangesForReset(): void
  
  // UI updates
  updateCreateButtonState(): void
  updateResetButtonState(): void
  updateCellVisuals(): void
  
  // Property editor management
  openPropertyEditor(key: string): void
  closePropertyEditor(): void
}
```

**Functions to Extract:**
- `setMode` (~13 lines)
- `updateModeDependentVisibility` (~11 lines)
- `getCellState` (~6 lines)
- `resetTableProperties` (~38 lines)
- `updateCreateButtonState` (~3 lines)
- `updateResetButtonState` (~6 lines)
- `updateCellVisuals` (~206 lines)
- `openPropertyEditor` (~37 lines)
- `openPropertyEditorInternal` (~518 lines)
- `closePropertyEditor` (~8 lines)

### 5. RemainingUtils Component

**Purpose:** House remaining utility functions that don't fit other categories

**Responsibilities:**
- Property name cleaning and formatting
- Grid scaffolding and setup
- Header/footer grid rendering
- Miscellaneous helper functions

**Interface:**
```typescript
class RemainingUtils {
  // Property utilities
  static cleanPropName(name: string): string
  
  // Grid utilities
  static ensurePropertyEditorScaffold(): void
  static renderHeaderFooterGrids(): void
  
  // Other utilities
  static createGrid(): void
}
```

**Functions to Extract:**
- `cleanPropName` (~2 lines)
- `ensurePropertyEditorScaffold` (~356 lines)
- `renderHeaderFooterGrids` (~68 lines)
- `createGrid` (~67 lines)

## Data Models

### Component Communication
Components will communicate through:
1. **Event system** - Custom events for loose coupling
2. **Shared state** - Read-only access to global state
3. **Method calls** - Direct calls for tight coupling where needed

### State Management
- State remains centralized in the main state object
- Components get references to state but don't modify directly
- State changes go through StateManager component
- UI updates triggered by state change events

## Error Handling

### Component Isolation
- Each component handles its own errors
- Failed components don't crash the entire application
- Graceful degradation when components fail

### Message Handling
- Message validation before processing
- Error responses for invalid messages
- Timeout handling for long-running operations

### UI Error States
- Loading states for async operations
- Error messages for user-facing failures
- Fallback UI when components fail to load

## Testing Strategy

### Unit Testing
- Each component will be independently testable
- Mock dependencies for isolated testing
- Test component interfaces and public methods

### Integration Testing
- Test component communication
- Test message flow between UI and plugin
- Test state synchronization

### Performance Testing
- Measure plugin loading time before and after
- Monitor memory usage of components
- Test with large datasets

### Regression Testing
- Ensure all existing functionality works
- Test edge cases and error conditions
- Validate UI appearance and behavior

## Implementation Plan

### Phase 2A: Extract MessageHandler
1. Create MessageHandler component
2. Move all message handling logic
3. Update ui.ts to use MessageHandler
4. Test message flow

### Phase 2B: Extract PropertyRenderer  
1. Create PropertyRenderer component
2. Move property rendering functions
3. Update PropertyEditor to use PropertyRenderer
4. Test property rendering

### Phase 2C: Extract EventManager
1. Create EventManager component
2. Move event setup functions
3. Update ui.ts initialization
4. Test event handling

### Phase 2D: Extract StateManager
1. Create StateManager component
2. Move state management functions
3. Update components to use StateManager
4. Test state synchronization

### Phase 2E: Extract RemainingUtils
1. Create RemainingUtils component
2. Move remaining utility functions
3. Update imports across codebase
4. Final testing and cleanup

### Phase 2F: Final Optimization
1. Remove unused code from ui.ts
2. Optimize imports and dependencies
3. Performance testing and tuning
4. Documentation updates