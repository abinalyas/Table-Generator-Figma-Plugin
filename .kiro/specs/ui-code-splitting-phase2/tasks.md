# Implementation Plan

- [x] 1. Create MessageHandler component and extract message handling logic
  - Create MessageHandler.ts with centralized message routing
  - Move window.onmessage handler from ui.ts to MessageHandler
  - Extract all parent.postMessage calls into MessageHandler methods
  - Update ui.ts to initialize and use MessageHandler
  - _Requirements: 2.1, 2.2, 2.3, 2.4_

- [x] 1.1 Implement MessageHandler class structure
  - Define MessageHandler class with constructor taking state and elements
  - Create handleIncomingMessage method for routing incoming messages
  - Create sendMessage method for outgoing plugin communication
  - Add message type constants and interfaces
  - _Requirements: 2.1, 2.2_

- [x] 1.2 Extract window.onmessage handler logic
  - Move the large window.onmessage function from ui.ts to MessageHandler
  - Implement message routing to appropriate handler methods
  - Add error handling and validation for incoming messages
  - _Requirements: 2.2, 2.4_

- [x] 1.3 Extract all postMessage calls
  - Move all parent.postMessage calls to MessageHandler methods
  - Create specific methods for each message type (scanTable, saveApiKey, etc.)
  - Update calling code to use MessageHandler instead of direct postMessage
  - _Requirements: 2.3, 2.4_

- [ ]* 1.4 Add unit tests for MessageHandler
  - Test message routing functionality
  - Test error handling for invalid messages
  - Mock plugin communication for testing
  - _Requirements: 2.1, 2.2, 2.3, 2.4_

- [x] 2. Create PropertyRenderer component and extract property rendering logic
  - Create PropertyRenderer.ts with UI rendering methods
  - Move renderDynamicPropertyFields and related functions from ui.ts
  - Update PropertyEditor to use PropertyRenderer
  - Test property field rendering functionality
  - _Requirements: 3.1, 3.2, 3.3, 3.4_

- [x] 2.1 Implement PropertyRenderer class structure
  - Define PropertyRenderer class with rendering methods
  - Create interfaces for property field types and configurations
  - Add methods for different field types (text, select, checkbox)
  - _Requirements: 3.1, 3.2_

- [x] 2.2 Extract renderDynamicPropertyFields function
  - Move renderDynamicPropertyFields (~165 lines) to PropertyRenderer
  - Move renderDynamicPropertyFieldsFromModel (~128 lines) to PropertyRenderer
  - Update function signatures to work with class methods
  - _Requirements: 3.2, 3.3_

- [x] 2.3 Extract renderBodyCellProperties function
  - Move renderBodyCellProperties (~69 lines) to PropertyRenderer
  - Move updateStaticFieldsVisibilityAndValues (~102 lines) to PropertyRenderer
  - Move findMatchingProperty (~25 lines) to PropertyRenderer
  - _Requirements: 3.2, 3.3_

- [x] 2.4 Update PropertyEditor integration
  - Update PropertyEditor component to use PropertyRenderer
  - Remove duplicate rendering logic from PropertyEditor
  - Test property rendering with different cell types
  - _Requirements: 3.1, 3.3_

- [ ]* 2.5 Add unit tests for PropertyRenderer
  - Test field creation for different property types
  - Test property field updates and validation
  - Test integration with PropertyEditor component
  - _Requirements: 3.1, 3.2, 3.3, 3.4_

- [x] 3. Create EventManager component and extract event handling logic
  - Create EventManager.ts with centralized event management
  - Move setupEventListeners and related functions from ui.ts
  - Update ui.ts initialization to use EventManager
  - Test all event handling functionality
  - _Requirements: 4.1, 4.2, 4.3, 4.4_

- [x] 3.1 Implement EventManager class structure
  - Define EventManager class with event setup methods
  - Create methods for different event categories (file upload, resize, API keys)
  - Add event cleanup and memory management
  - _Requirements: 4.1, 4.2_

- [x] 3.2 Extract setupEventListeners function
  - Move setupEventListeners (~122 lines) to EventManager
  - Move setupFileUploadControls (~19 lines) to EventManager
  - Move setupResizeCorner (~61 lines) to EventManager
  - _Requirements: 4.2, 4.4_

- [x] 3.3 Extract API key and option handling
  - Move setupApiKeySync (~61 lines) to EventManager
  - Move handleApplyOptionClick (~24 lines) to EventManager
  - Update event handler references in ui.ts
  - _Requirements: 4.2, 4.4_

- [x] 3.4 Update ui.ts event initialization
  - Replace direct event setup calls with EventManager usage
  - Remove extracted event functions from ui.ts
  - Test all event handling works correctly
  - _Requirements: 4.1, 4.3_

- [ ]* 3.5 Add unit tests for EventManager
  - Test event listener setup and cleanup
  - Test event delegation and routing
  - Mock DOM elements for testing
  - _Requirements: 4.1, 4.2, 4.3, 4.4_

- [x] 4. Create StateManager component and extract state management logic
  - Create StateManager.ts with state update methods
  - Move state management functions from ui.ts to StateManager
  - Update components to use StateManager for state changes
  - Test state synchronization and UI updates
  - _Requirements: 5.1, 5.2, 5.3, 5.4_

- [x] 4.1 Implement StateManager class structure
  - Define StateManager class with state management methods
  - Create methods for mode transitions and visibility updates
  - Add validation methods for state changes
  - _Requirements: 5.1, 5.2_

- [x] 4.2 Extract mode and visibility functions
  - Move setMode (~13 lines) to StateManager
  - Move updateModeDependentVisibility (~11 lines) to StateManager
  - Move getCellState (~6 lines) to StateManager
  - _Requirements: 5.2, 5.4_

- [x] 4.3 Extract reset and update functions
  - Move resetTableProperties (~38 lines) to StateManager
  - Move updateCreateButtonState (~3 lines) to StateManager
  - Move updateResetButtonState (~6 lines) to StateManager
  - _Requirements: 5.1, 5.4_

- [x] 4.4 Extract property editor state management
  - Move openPropertyEditor (~37 lines) to StateManager
  - Move openPropertyEditorInternal (~518 lines) to StateManager
  - Move closePropertyEditor (~8 lines) to StateManager
  - Move updateCellVisuals (~206 lines) to StateManager
  - _Requirements: 5.2, 5.4_

- [x] 4.5 Update component integration
  - Update components to use StateManager for state changes
  - Remove direct state manipulation from components
  - Test state synchronization across components
  - _Requirements: 5.1, 5.3_

- [ ]* 4.6 Add unit tests for StateManager
  - Test state validation and updates
  - Test mode transitions and UI synchronization
  - Test property editor state management
  - _Requirements: 5.1, 5.2, 5.3, 5.4_

- [x] 5. Create RemainingUtils component and extract utility functions
  - Create RemainingUtils.ts with remaining utility functions
  - Move cleanPropName and grid utility functions from ui.ts
  - Update imports across codebase to use RemainingUtils
  - Test utility function functionality
  - _Requirements: 6.1, 6.2, 6.3, 6.4_

- [x] 5.1 Implement RemainingUtils class structure
  - Define RemainingUtils class with static utility methods
  - Create property name cleaning utilities
  - Add grid scaffolding and setup utilities
  - _Requirements: 6.1, 6.2_

- [x] 5.2 Extract property and grid utilities
  - Move cleanPropName (~2 lines) to RemainingUtils
  - Move ensurePropertyEditorScaffold (~356 lines) to RemainingUtils
  - Move renderHeaderFooterGrids (~68 lines) to RemainingUtils
  - _Requirements: 6.2, 6.3_

- [x] 5.3 Extract remaining grid functions
  - Move createGrid (~67 lines) to RemainingUtils or appropriate component
  - Update GridManager to use utility functions where appropriate
  - Remove duplicate grid creation logic
  - _Requirements: 6.2, 6.3_

- [x] 5.4 Update imports and references
  - Update all files that use extracted utility functions
  - Remove extracted functions from ui.ts
  - Test that all utility functions work correctly
  - _Requirements: 6.1, 6.4_

- [ ]* 5.5 Add unit tests for RemainingUtils
  - Test property name cleaning functionality
  - Test grid utility functions
  - Test integration with other components
  - _Requirements: 6.1, 6.2, 6.3, 6.4_

- [x] 6. Final optimization and cleanup of ui.ts
  - Remove all extracted functions from ui.ts
  - Optimize imports and component initialization
  - Verify ui.ts is under 2,000 lines
  - Perform comprehensive testing of all functionality
  - _Requirements: 1.1, 1.2, 1.3, 1.4, 6.1, 6.2, 6.3, 6.4_

- [x] 6.1 Clean up ui.ts structure
  - Remove all functions that have been moved to components
  - Organize remaining code into logical sections
  - Update imports to use new components
  - _Requirements: 6.1, 6.2_

- [x] 6.2 Optimize component initialization
  - Streamline component initialization in ui.ts
  - Add proper error handling for component creation
  - Ensure proper component lifecycle management
  - _Requirements: 6.2, 6.3_

- [x] 6.3 Verify line count and performance
  - Confirm ui.ts is under 2,000 lines
  - Test plugin loading performance
  - Measure improvement in loading time
  - _Requirements: 1.1, 1.2_

- [x] 6.4 Comprehensive functionality testing
  - Test all existing features work correctly
  - Test component communication and coordination
  - Verify no regressions in functionality
  - _Requirements: 1.3, 1.4_

- [ ]* 6.5 Add integration tests
  - Test component interaction and communication
  - Test plugin message flow end-to-end
  - Test state synchronization across components
  - _Requirements: 1.3, 1.4_

- [x] 6.6 Update documentation and build
  - Update component documentation
  - Verify build process works correctly
  - Update any configuration files if needed
  - _Requirements: 6.4_