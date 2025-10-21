import { SCAN_DEBOUNCE_MS, getLastScanTime, setLastScanTime } from './state';
import { isGeneratedTable, readTableSettings } from './metadata';

export function initSelectionHandlers(): void {
  figma.on('selectionchange', async () => {
    const selection = figma.currentPage.selection;
    console.log(`[selectionchange] Selection changed - count: ${selection.length}, types: ${selection.map(s => s.type).join(', ')}, names: ${selection.map(s => s.name).join(', ')}`);

    // Track if we found a valid selection
    let foundValidSelection = false;
    
    // Apply debouncing AFTER we check for valid selection
    const now = Date.now();
    const last = getLastScanTime();
    if (now - last < SCAN_DEBOUNCE_MS) {
      console.log(`[selectionchange] Debouncing - ${now - last}ms since last scan`);
      // Still set foundValidSelection to true if we have a selection to prevent selection-cleared
      if (selection.length > 0) {
        foundValidSelection = true;
      }
      return;
    }
    setLastScanTime(now);

    if (selection.length === 1) {
      const selectedNode = selection[0];

      // Case 1: Existing "Generated Table" frame or component is selected for update
      console.log(`[selectionchange] Checking for generated table: type=${selectedNode.type}, name=${selectedNode.name}, isGeneratedTable=${isGeneratedTable(selectedNode as any)}`);
      if ((selectedNode.type === 'FRAME' || selectedNode.type === 'COMPONENT') && isGeneratedTable(selectedNode as any) && (selectedNode.name === 'Generated Table' || selectedNode.name.startsWith('Generated Table ('))) {
        const tableFrame = selectedNode as FrameNode | ComponentNode;
        const parsedSettings = readTableSettings(tableFrame as any);
        if (parsedSettings) {
          try {
            figma.ui.postMessage({
              type: 'edit-existing-table',
              settings: parsedSettings,
              tableId: tableFrame.id
            });
            foundValidSelection = true;
            return;
          } catch (e) {
            console.error('Error parsing table settings from plugin data', e);
            figma.ui.postMessage({ type: 'generated-table-selected-no-settings', tableId: tableFrame.id });
            foundValidSelection = true;
            return;
          }
        } else {
          console.log('Generated table selected but no settings found');
          figma.ui.postMessage({ type: 'generated-table-selected-no-settings', tableId: tableFrame.id });
          foundValidSelection = true;
          return;
        }
      }

      // Case 1.5: "Data table" component is selected
      if (selectedNode.type === 'FRAME' && selectedNode.name.includes('Data table')) {
        figma.ui.postMessage({ type: 'table-selected', tableId: selectedNode.id });
        foundValidSelection = true;
        return;
      }

      // Case 2: Data table component (Frame, Component, ComponentSet) selected for scanning
      if ((selectedNode.type === 'FRAME' || selectedNode.type === 'COMPONENT' || selectedNode.type === 'COMPONENT_SET') && selectedNode.name.includes('Data table')) {
        figma.ui.postMessage({ type: 'table-selected', tableId: selectedNode.id });
        foundValidSelection = true;
        return;
      }

      // Case 3: Instance is selected
      if (selectedNode.type === 'INSTANCE') {
        try {
          const mainComponent = await selectedNode.getMainComponentAsync();
          console.log(`Selected Instance: ${selectedNode.name}, Main Component: ${mainComponent?.name}`);

          // Ignore individual cell items
          if ((mainComponent && mainComponent.name === 'Data table row cell item') || selectedNode.name === 'Data table body row item' || selectedNode.name.includes('row cell item')) {
            console.log(`[selectionchange] Ignoring selection of individual cell item: ${selectedNode.name}`);
            return;
          }

          // Data table instance by main component name
          if (mainComponent && mainComponent.name.includes('Data table')) {
            console.log('[selectionchange] Found Data table instance, sending table-selected');
            figma.ui.postMessage({ type: 'table-selected', tableId: selectedNode.id });
            foundValidSelection = true;
            return;
          }

          // Fallback by node name
          if (selectedNode.name.includes('Data table') && !selectedNode.name.includes('row cell item')) {
            console.log('[selectionchange] Found Data table instance by node name, sending table-selected');
            figma.ui.postMessage({ type: 'table-selected', tableId: selectedNode.id });
            foundValidSelection = true;
            return;
          }
        } catch (error) {
          console.error('❌ Error getting main component or component name:', error);
        }
      }
    }

    // Only send selection-cleared if we didn't find a valid selection
    if (!foundValidSelection) {
      console.log('[selectionchange] Clearing selection - no valid component found');
      setTimeout(() => {
        const currentSelection = figma.currentPage.selection;
        if (currentSelection.length === 0) {
          figma.ui.postMessage({ type: 'selection-cleared', isValidComponent: false, clearUI: true });
        }
      }, 50);
    }
  });
}


