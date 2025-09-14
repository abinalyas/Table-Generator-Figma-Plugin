const faker = require('faker');

// Plugin default dimensions
const pluginDefaultWidth = 500;
const pluginDefaultHeight = 750;
const pluginMaxWidth = 1200;
const pluginMaxHeight = 1000;
const pluginMinWidth = 300;
const pluginMinHeight = 400;

// Helper function to sort column data based on header properties
function sortColumnData(cellProps: any, cols: number, rows: number): any {
    // Create a copy of the cellProps to avoid modifying the original
    const sortedCellProps = {...cellProps};
    
    // Check each column for sorting properties
    for (let c = 1; c <= cols; c++) {
        const headerKey = `header-${c}`;
        const headerData = cellProps[headerKey];
        
        // Check if sorting is enabled for this header
        if (headerData && headerData.properties) {
            const sortable = headerData.properties['Sortable'];
            const sorted = headerData.properties['Sorted'];
            
            // If sorting is enabled, sort the column data
            if (sortable === 'True' && (sorted === 'Ascending' || sorted === 'Descending')) {
                // Extract column data
                const columnData: { rowIndex: number; cellData: any; value: string }[] = [];
                
                for (let r = 0; r < rows; r++) {
                    const cellKey = `${r}-${c - 1}`; // 0-based indexing for cellProps
                    const cellData = cellProps[cellKey];
                    
                    // Get the cell text value for sorting
                    let cellValue = '';
                    if (cellData && cellData.properties) {
                        // Try to find text property
                        const textProp = Object.keys(cellData.properties).find(
                            prop => prop.toLowerCase().includes('text') && 
                                   !prop.toLowerCase().includes('second') &&
                                   typeof cellData.properties[prop] === 'string'
                        );
                        
                        if (textProp) {
                            cellValue = String(cellData.properties[textProp]);
                        } else if (cellData.properties['Cell text#12234:32']) {
                            cellValue = String(cellData.properties['Cell text#12234:32']);
                        }
                    }
                    
                    columnData.push({
                        rowIndex: r,
                        cellData: cellData,
                        value: cellValue
                    });
                }
                
                // Sort the column data
                columnData.sort((a, b) => {
                    // Try to parse as numbers first
                    const numA = parseFloat(a.value);
                    const numB = parseFloat(b.value);
                    
                    if (!isNaN(numA) && !isNaN(numB)) {
                        // Numeric sort
                        return sorted === 'Ascending' ? numA - numB : numB - numA;
                    } else {
                        // Alphabetic sort
                        return sorted === 'Ascending' 
                            ? a.value.localeCompare(b.value) 
                            : b.value.localeCompare(a.value);
                    }
                });
                
                // Reassign sorted data back to cellProps
                for (let r = 0; r < rows; r++) {
                    const originalRowIndex = columnData[r].rowIndex;
                    const newCellKey = `${r}-${c - 1}`; // New position
                    const originalCellKey = `${originalRowIndex}-${c - 1}`; // Original position
                    
                    // Move the data from original position to new position
                    sortedCellProps[newCellKey] = columnData[r].cellData;
                }
            }
        }
    }
    
    return sortedCellProps;
}

figma.showUI(__html__, { 
    width: pluginDefaultWidth, 
    height: pluginDefaultHeight,
    themeColors: true // Enable theme colors for better integration
});

// Restore previous size when reopening the plugin
figma.clientStorage.getAsync('pluginSize').then(size => {
    if (size && size.w && size.h) {
        figma.ui.resize(size.w, size.h);
        console.log(`Restored plugin size: ${size.w}×${size.h}`);
    }
}).catch(err => {
    console.log('No previous size found, using defaults');
});

let selectedComponent: ComponentNode | null = null;
const cellInstanceMap = new Map<string, InstanceNode>();
let lastScanTime = 0;
const SCAN_DEBOUNCE_MS = 500;

interface ScanResult {
    headerCell: ComponentNode | null;
    headerRowComponent: ComponentNode | null;
    bodyCell: ComponentNode | null;
    bodyRowComponent: ComponentNode | null;
    footer: ComponentNode | null;
    numCols: number;
    selectCellComponent: ComponentNode | null;
    expandCellComponent: ComponentNode | null;
    dividerComponent: ComponentNode | null;
}

let lastScanResult: ScanResult | undefined;
let isCreatingTable = false;

figma.on("selectionchange", async () => {
    const now = Date.now();
    if (now - lastScanTime < SCAN_DEBOUNCE_MS) {
        console.log(`[selectionchange] Debouncing - ${now - lastScanTime}ms since last scan`);
        return;
    }
    
    const selection = figma.currentPage.selection;
    console.log(`[selectionchange] Selection changed - count: ${selection.length}, types: ${selection.map(s => s.type).join(', ')}, names: ${selection.map(s => s.name).join(', ')}`);

    if (selection.length === 1) {
        const selectedNode = selection[0];

        // Case 1: Existing "Generated Table" frame or component is selected for update
        console.log(`[selectionchange] Checking for generated table: type=${selectedNode.type}, name=${selectedNode.name}, isGeneratedTable=${selectedNode.getPluginData('isGeneratedTable')}`);
        if ((selectedNode.type === "FRAME" || selectedNode.type === "COMPONENT") && selectedNode.getPluginData('isGeneratedTable') === 'true' && (selectedNode.name === 'Generated Table' || selectedNode.name.startsWith('Generated Table ('))) {
            const tableFrame = selectedNode as FrameNode | ComponentNode;
            const settings = tableFrame.getPluginData('tableSettings');
            if (settings) {
                try {
                    // Note: We'll rely on the existing lastScanResult or let the user scan again if needed
                    // The rescan logic was causing memory issues, so we'll keep it simple
                    const parsedSettings = JSON.parse(settings);
                    figma.ui.postMessage({
                        type: 'edit-existing-table',
                        settings: parsedSettings,
                        tableId: tableFrame.id
                    });
                    return;
                } catch (e) {
                    console.error("Error parsing table settings from plugin data", e);
                    // Even if parsing fails, this is still a valid generated table
                    // Send a message to show the table is selected but settings couldn't be loaded
                    figma.ui.postMessage({
                        type: 'generated-table-selected-no-settings',
                        tableId: tableFrame.id
                    });
                    return;
                }
            } else {
                // Generated table without settings - still valid, just no stored configuration
                console.log("Generated table selected but no settings found");
                figma.ui.postMessage({
                    type: 'generated-table-selected-no-settings',
                    tableId: tableFrame.id
                });
                return;
            }
        }
        // Case 1.5: "Data table" component is selected
        else if (selectedNode.type === "FRAME" && selectedNode.name.includes('Data table')) {
            figma.ui.postMessage({
                type: "table-selected",
                tableId: selectedNode.id
            });
            return;
        }
        // Case 2: "Data table" component (Frame, Component, ComponentSet) or instance of one is selected for scanning
        else if (
            (selectedNode.type === "FRAME" || selectedNode.type === "COMPONENT" || selectedNode.type === "COMPONENT_SET") && selectedNode.name.includes('Data table')
        ) {
            figma.ui.postMessage({
                type: "table-selected",
                tableId: selectedNode.id
            });
            return;
        }
        // Case 3: Instance is selected (could be a Data table instance or a Data table row cell item)
        else if (selectedNode.type === "INSTANCE") {
            try {
                const mainComponent = await selectedNode.getMainComponentAsync();
                console.log(`Selected Instance: ${selectedNode.name}, Main Component: ${mainComponent?.name}`);

                // FIRST: Check if it's a "Data table row cell item" - ignore these
                // Check both the main component name and the node name for cell items
                if ((mainComponent && mainComponent.name === "Data table row cell item") || 
                    selectedNode.name === "Data table body row item" ||
                    selectedNode.name.includes("row cell item")) {
                    console.log(`[selectionchange] Ignoring selection of individual cell item: ${selectedNode.name}`);
                    return;
                }
                
                // SECOND: Check if it's a "Data table" instance (main component name contains "Data table")
                if (mainComponent && mainComponent.name.includes('Data table')) {
                    console.log(`[selectionchange] Found Data table instance, sending table-selected`);
        figma.ui.postMessage({
                        type: "table-selected",
                        tableId: selectedNode.id
                    });
                    return; // IMPORTANT: Return here to prevent further processing
                }
                
                // THIRD: Fallback - check if the selected node name contains "Data table" (but not "row cell item")
                if (selectedNode.name.includes('Data table') && !selectedNode.name.includes('row cell item')) {
                    console.log(`[selectionchange] Found Data table instance by node name, sending table-selected`);
      figma.ui.postMessage({
                        type: "table-selected",
                        tableId: selectedNode.id
                    });
                    return; // IMPORTANT: Return here to prevent further processing
                }
    } catch (error) {
                console.error('❌ Error getting main component or component name:', error);
            }
        }
    }

    // If none of the above conditions are met, clear selection
    console.log(`[selectionchange] Clearing selection - no valid component found`);
      selectedComponent = null;
    lastScanResult = undefined;
    
    // Add a small delay to prevent race conditions with subsequent selections
    setTimeout(() => {
        // Double-check that selection is still empty before clearing
        const currentSelection = figma.currentPage.selection;
        if (currentSelection.length === 0) {
      figma.ui.postMessage({
        type: "selection-cleared",
                isValidComponent: false,
                clearUI: true
      });
    }
    }, 50);
    });



// Helper function to find matching property
function findMatchingProperty(availableProperties: string[], uiPropName: string): string | null {
    // First try exact match
    const exactMatch = availableProperties.find((prop: string) => prop === uiPropName);
    if (exactMatch) {
        return exactMatch;
    }
    // Try match at start of property name
    const startsWithMatch = availableProperties.find((prop: string) => prop.toLowerCase().startsWith(uiPropName.toLowerCase()));
    if (startsWithMatch) {
        return startsWithMatch;
    }
    // Try contains match
    const containsMatch = availableProperties.find((prop: string) => prop.toLowerCase().includes(uiPropName.toLowerCase()));
    if (containsMatch) {
        return containsMatch;
    }
    
    return null;
}

// Helper function to extract color variable from footer divider
function extractFooterDividerColorVariable(footer: ComponentNode): VariableAlias | null {
    try {
        console.log('🔍 Extracting color variable from footer divider:', footer.name);
        
        // Look for divider elements in the footer component
        if ('children' in footer) {
            for (const child of footer.children) {
                // Check if this child is a divider (rectangle, line, or has divider-like name)
                const isDividerChild = child.name.toLowerCase().includes('divider') ||
                                     child.name.toLowerCase().includes('line') ||
                                     child.name.toLowerCase().includes('border') ||
                                     child.name.toLowerCase().includes('separator') ||
                                     child.type === 'LINE' ||
                                     (child.type === 'RECTANGLE' && child.height <= 2); // Thin rectangles are likely dividers
                
                if (isDividerChild && 'fills' in child && child.fills && Array.isArray(child.fills)) {
                    for (const fill of child.fills) {
                        if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                            console.log('✅ Found color variable in footer divider:', fill.boundVariables.color);
                            return fill.boundVariables.color;
                        }
                    }
                }
                
                // Also check nested children (in case divider is inside a frame)
                if ('children' in child) {
                    for (const grandChild of child.children) {
                        const isDividerGrandChild = grandChild.name.toLowerCase().includes('divider') ||
                                                  grandChild.name.toLowerCase().includes('line') ||
                                                  grandChild.name.toLowerCase().includes('border') ||
                                                  grandChild.name.toLowerCase().includes('separator') ||
                                                  grandChild.type === 'LINE' ||
                                                  (grandChild.type === 'RECTANGLE' && grandChild.height <= 2);
                        
                        if (isDividerGrandChild && 'fills' in grandChild && grandChild.fills && Array.isArray(grandChild.fills)) {
                            for (const fill of grandChild.fills) {
                                if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                                    console.log('✅ Found color variable in footer nested divider:', fill.boundVariables.color);
                                    return fill.boundVariables.color;
                                }
                            }
                        }
                    }
                }
            }
        }
        
        console.log('⚠️ No color variable found in footer divider elements');
        return null;
    } catch (error) {
        console.error('❌ Error extracting footer divider color variable:', error);
        return null;
    }
}

// Helper function to find border color variable
function findBorderColorVariable(): VariableAlias | null {
    try {
        // Get all variables from all libraries (including local and external)
        const allVariableCollections = figma.variables.getLocalVariableCollections();
        const allVariables: Variable[] = [];
        
        // Collect all variables from all collections
        for (const collection of allVariableCollections) {
            const collectionVariables = collection.variableIds.map(id => figma.variables.getVariableById(id)).filter(v => v !== null) as Variable[];
            allVariables.push(...collectionVariables);
        }
        
        console.log('🔍 Available variables from all libraries:', allVariables.map(v => `${v.name} (ID: ${v.id})`));
        
        // Debug: Print all variable names that contain 'border'
        const borderVariables = allVariables.filter(v => v.name.toLowerCase().includes('border'));
        console.log('🔍 Border-related variables found:', borderVariables.map(v => `${v.name} (ID: ${v.id})`));
        
        // Debug: Print all variable names that contain 'subtle'
        const subtleVariables = allVariables.filter(v => v.name.toLowerCase().includes('subtle'));
        console.log('🔍 Subtle-related variables found:', subtleVariables.map(v => `${v.name} (ID: ${v.id})`));
        
        // Look for common border variable names, prioritizing border-subtle-01
        const borderVariableNames = ['border-subtle-01', 'border-subtle-02', 'border-subtle', 'border-01', 'border-subtle-1'];
        let borderVariable = null;
        
        for (const variableName of borderVariableNames) {
            borderVariable = allVariables.find(v => v.name === variableName);
            if (borderVariable) {
                console.log(`🎨 Found border variable from libraries: ${variableName}`);
                break;
            }
        }
        
        if (borderVariable) {
            console.log('✅ Using border color variable from libraries');
            return {
                id: borderVariable.id,
                type: 'VARIABLE_ALIAS'
            };
        } else {
            console.log('⚠️ No border color variable found in any library');
            return null;
        }
    } catch (error) {
        console.error('❌ Error finding border color variable:', error);
        return null;
    }
}

// Removed createCheckboxComponent function - no longer creating manual components

// Function to perform table scanning on a given component
async function performTableScan(tableFrame: SceneNode): Promise<ScanResult | null> {
    console.log('🔄 Performing table scan on:', tableFrame.name);
    
    try {
        // Check if this is a generated table first
        const isGeneratedTable = 'getPluginData' in tableFrame && tableFrame.getPluginData('isGeneratedTable') === 'true';
        console.log('🔍 Is generated table:', isGeneratedTable);
        
        if (isGeneratedTable) {
            console.log('🔄 Scanning generated table with fallback method...');
            return await scanGeneratedTable(tableFrame as FrameNode | ComponentNode);
        }
        
        // --- Custom scan for specific instances/components ---
        let expandCellFound = false;
        let selectCellFound = false;
        const allInstanceNames: string[] = [];
        let allInstances: (InstanceNode | ComponentNode)[] = [];
        let expandCellComponent: ComponentNode | null = null;
        let selectCellComponent: ComponentNode | null = null;
        let dividerComponent: ComponentNode | null = null;

        if ('findOne' in tableFrame && typeof tableFrame.findOne === 'function') {
            // Find header row (INSTANCE)
            const headerRow = tableFrame.findOne(n =>
                n.name?.includes('header row') && n.type === 'INSTANCE'
            ) as InstanceNode | undefined;
            // Find body group (FRAME)
            const bodyGroup = tableFrame.findOne(n => n.name === 'Body' && n.type === 'FRAME');
            // Find footer bar (INSTANCE)
            const footerBar = tableFrame.findOne(n => n.name === 'Pagination - Table bar' && n.type === 'INSTANCE');
            // FIRST INSTANCE - Variable declarations for component extraction
            let headerCell: InstanceNode | null = null;
            let bodyCell: InstanceNode | null = null;
            let footer: InstanceNode | null = null;
            let numCols = 0;
            let headerRowComponent: ComponentNode | null = null;
            let bodyRowComponent: ComponentNode | null = null;
            
            // Extract header cell template
            if (headerRow && 'children' in headerRow) {
                headerRowComponent = await headerRow.getMainComponentAsync();
                
                const headerCellInstance = headerRow.children.find((n: SceneNode) =>
                    n.type === 'INSTANCE' && n.name?.toLowerCase().includes('col 1')
                ) as InstanceNode | undefined;
                if (headerCellInstance) {
                    headerCell = headerCellInstance;
                    // Count only instances that are visible data columns
                    numCols = headerRow.children.filter((n: SceneNode) => 
                        n.type === 'INSTANCE' && 
                        n.name?.toLowerCase().includes('col') &&
                        !n.name?.toLowerCase().includes('select') &&
                        !n.name?.toLowerCase().includes('expand') &&
                        n.visible
                    ).length;
                }
            }
            
            // Extract body cell template
            if (bodyGroup && 'children' in bodyGroup) {
                const bodyRowInstance = bodyGroup.children.find((n: SceneNode) => n.type === 'INSTANCE' && n.name === 'Data table body row item') as InstanceNode | undefined;
                if (bodyRowInstance) {
                    bodyRowComponent = await bodyRowInstance.getMainComponentAsync();
                    
                    if ('children' in bodyRowInstance) {
                        const dataTableRow = bodyRowInstance.children.find(child => child.type === 'FRAME' && child.name === 'Data table row') as FrameNode | undefined;
                        if (dataTableRow && 'children' in dataTableRow) {
                            // Find the first data cell instance whose name includes 'Col 1'
                            const bodyCellInstance = dataTableRow.children.find((cell: SceneNode) =>
                                cell.type === 'INSTANCE' && cell.name?.toLowerCase().includes('col 1')
                            ) as InstanceNode | undefined;
                            if (bodyCellInstance) {
                                // REVERTING - Keep as InstanceNode for performTableScan
                                bodyCell = bodyCellInstance;
                                console.log('✅ Extracted body cell component from instance:', bodyCell?.name);
                            }
                        }
                    }
                }
            }
            
            // Extract footer template
            if (footerBar && footerBar.type === 'INSTANCE') {
                footer = footerBar;
            }
            
            // --- Enhanced divider scanning - look in the specific path first ---
            console.log('🔍 Scanning for divider rectangles in specific path: Data table > Body > Data table body row item > Divider...');
            let foundDividerRect: RectangleNode | null = null;
            
            if (bodyGroup && 'children' in bodyGroup) {
                const bodyRowInstance = bodyGroup.children.find((n: SceneNode) => n.type === 'INSTANCE' && n.name === 'Data table body row item') as InstanceNode | undefined;
                if (bodyRowInstance && 'children' in bodyRowInstance) {
                    console.log('🔍 Found Data table body row item, searching for Divider rectangle inside...');
                    
                    // Look for the Divider rectangle directly in the body row item
                    const dividerInBodyRow = bodyRowInstance.children.find((child: SceneNode) => 
                        child.type === 'RECTANGLE' && child.name === 'Divider'
                    ) as RectangleNode | undefined;
                    
                    if (dividerInBodyRow) {
                        foundDividerRect = dividerInBodyRow;
                        console.log('✅ Found Divider rectangle in Data table body row item:', dividerInBodyRow.name);
                        console.log('🎨 Divider rectangle properties:', {
                            width: dividerInBodyRow.width,
                            height: dividerInBodyRow.height,
                            fills: dividerInBodyRow.fills,
                            locked: dividerInBodyRow.locked
                        });
                    } else {
                        console.log('⚠️ No Divider rectangle found directly in Data table body row item, checking nested children...');
                        
                        // Check nested children in case the divider is deeper in the hierarchy
                        for (const child of bodyRowInstance.children) {
                            if ('children' in child) {
                                const nestedDivider = child.children.find((grandChild: SceneNode) => 
                                    grandChild.type === 'RECTANGLE' && grandChild.name === 'Divider'
                                ) as RectangleNode | undefined;
                                
                                if (nestedDivider) {
                                    foundDividerRect = nestedDivider;
                                    console.log('✅ Found Divider rectangle in nested structure:', nestedDivider.name);
                                    console.log('🎨 Nested divider rectangle properties:', {
                                        width: nestedDivider.width,
                                        height: nestedDivider.height,
                                        fills: nestedDivider.fills,
                                        locked: nestedDivider.locked
                                    });
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            
            // Store the found divider rectangle
            if (foundDividerRect) {
                (lastScanResult as any).originalDividerRect = foundDividerRect;
                console.log('✅ Stored original divider rectangle for direct use');
            }
            
            // --- Scan for select/expand cell components and divider ---
            if ('findAll' in tableFrame && typeof tableFrame.findAll === 'function') {
                // Fallback: scan all rectangles if specific path didn't work
                if (!foundDividerRect) {
                    console.log('🔍 Specific path search failed, falling back to general rectangle scan...');
                    const allRectangles = tableFrame.findAll(n => n.type === 'RECTANGLE') as RectangleNode[];
                    console.log(`🔍 Found ${allRectangles.length} total rectangles`);
                    console.log('🔍 All rectangle names:', allRectangles.map(r => `"${r.name}" (${r.width}x${r.height})`));
                    
                    const dividerRects = tableFrame.findAll(n => n.name === 'Divider' && n.type === 'RECTANGLE') as RectangleNode[];
                    console.log(`🔍 Found ${dividerRects.length} rectangle elements named exactly 'Divider'`);
                    
                    if (dividerRects.length > 0) {
                        foundDividerRect = dividerRects[0];
                        console.log('✅ Found divider rectangle via fallback search:', foundDividerRect.name);
                        (lastScanResult as any).originalDividerRect = foundDividerRect;
                        console.log('✅ Stored fallback divider rectangle for direct use');
                    }
                }
                
                allInstances = tableFrame.findAll(n => n.type === 'INSTANCE' || n.type === 'COMPONENT') as (InstanceNode | ComponentNode)[];
                for (const node of allInstances) {
                    if (node.name === 'Data table expand cell item') {
                      expandCellFound = true;
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        expandCellComponent = node.mainComponent;
                      }
                    }
                    // Check for select/checkbox components with various common names
                    const isSelectCell = node.name === 'Data table select cell item' ||
                                       node.name.includes('select') ||
                                       node.name.includes('checkbox') ||
                                       node.name.includes('check') ||
                                       node.name.includes('radio') ||
                                       node.name.toLowerCase().includes('select') ||
                                       node.name.toLowerCase().includes('checkbox');
                    
                    if (isSelectCell && !selectCellFound) {
                      selectCellFound = true;
                      console.log('🔍 Found select cell component:', node.name, node.type);
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        selectCellComponent = node.mainComponent;
                        console.log('✅ Using select cell from instance:', selectCellComponent.name);
                      } else if (node.type === 'COMPONENT') {
                        selectCellComponent = node;
                        console.log('✅ Using select cell component directly:', selectCellComponent.name);
                      }
                    }
                    
                    // Check for divider components - prioritize exact name match and skip AI labels
                    if (node.name.includes('AI label') || node.name.includes('Select menu')) {
                        continue; // Skip AI labels and select menus to improve performance
                    }
                    
                    const isDivider = node.name === 'Divider' || 
                                    (node.name.toLowerCase().includes('divider') && !node.name.includes('AI')) ||
                                    node.name.includes('line') ||
                                    node.name.includes('border') ||
                                    node.name.includes('separator');
                    
                    if (isDivider) {
                      console.log('🔍 Found potential divider component:', node.name, node.type);
                      
                      // Check if this component is actually a visual divider (not text-based)
                      let isValidDivider = false;
                      
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        if ('children' in node.mainComponent) {
                          for (const child of node.mainComponent.children) {
                            if (child.type === 'RECTANGLE' || child.type === 'LINE' || child.type === 'VECTOR') {
                              isValidDivider = true;
                              console.log('✅ Valid divider found with visual element:', child.type);
                              break;
                            } else if (child.type === 'TEXT') {
                              console.log('⚠️ Skipping text-based component:', node.name);
                              break;
                            }
                          }
                        }
                      } else if (node.type === 'COMPONENT') {
                        if ('children' in node) {
                          for (const child of node.children) {
                            if (child.type === 'RECTANGLE' || child.type === 'LINE' || child.type === 'VECTOR') {
                              isValidDivider = true;
                              console.log('✅ Valid divider found with visual element:', child.type);
                              break;
                            } else if (child.type === 'TEXT') {
                              console.log('⚠️ Skipping text-based component:', node.name);
                              break;
                            }
                          }
                        }
                      }
                      
                      if (isValidDivider) {
                        if (node.type === 'INSTANCE' && node.mainComponent) {
                          dividerComponent = node.mainComponent;
                          console.log('✅ Using divider from instance:', dividerComponent.name);
                        } else if (node.type === 'COMPONENT') {
                          dividerComponent = node;
                          console.log('✅ Using divider component directly:', dividerComponent.name);
                        }
                      } else {
                        console.log('⚠️ Skipping invalid divider component:', node.name);
                      }
                    }
                    allInstanceNames.push(node.name);
                }
            }
            
            // If select cell component is still not found, log a warning but don't create one
            if (!selectCellComponent) {
                console.warn('⚠️ Select cell component not found in scanned table. Selectable functionality will be disabled.');
            }
            
            // Log divider component status
            if (!dividerComponent) {
                console.log('⚠️ No divider component found, will use original divider rectangle if available');
            }
            
            // Return the scan result
            return {
                headerCell: headerCell as ComponentNode | null,
                headerRowComponent: headerRowComponent as ComponentNode | null,
                bodyCell: bodyCell as ComponentNode | null,
                bodyRowComponent: bodyRowComponent as ComponentNode | null,
                footer: footer as ComponentNode | null,
                numCols: numCols || 5,
                selectCellComponent: selectCellComponent as ComponentNode | null,
                expandCellComponent: expandCellComponent as ComponentNode | null,
                dividerComponent: dividerComponent as ComponentNode | null,
            };
        }
        
        return null;
        
    } catch (error) {
        console.error('❌ Error during table scan:', error);
        return null;
    }
}

// Function to scan generated tables with their specific structure
async function scanGeneratedTable(tableFrame: FrameNode | ComponentNode): Promise<ScanResult | null> {
    console.log('🔄 Scanning generated table:', tableFrame.name);
    
    try {
        // Try to restore from stored plugin data first
        const scanData = tableFrame.getPluginData('tableGeneratorScan');
        if (scanData) {
            try {
                const storedScan = JSON.parse(scanData);
                console.log('🔄 Found stored scan data, attempting to restore components...');
                
                let headerCell = null, bodyCell = null, footer = null;
                let selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
                
                // Restore components by ID
                if (storedScan.bodyCellId) {
                    const bodyCellNode = figma.getNodeById(storedScan.bodyCellId);
                    if (bodyCellNode && bodyCellNode.type === 'COMPONENT') {
                        bodyCell = bodyCellNode;
                        console.log('✅ Restored body cell from stored ID:', bodyCell.name);
                    }
                }
                
                if (storedScan.headerCellId) {
                    const headerCellNode = figma.getNodeById(storedScan.headerCellId);
                    if (headerCellNode && headerCellNode.type === 'COMPONENT') {
                        headerCell = headerCellNode;
                        console.log('✅ Restored header cell from stored ID:', headerCell.name);
                    }
                }
                
                if (storedScan.footerId) {
                    const footerNode = figma.getNodeById(storedScan.footerId);
                    if (footerNode && footerNode.type === 'COMPONENT') {
                        footer = footerNode;
                        console.log('✅ Restored footer from stored ID:', footer.name);
                    }
                }
                
                if (storedScan.selectCellId) {
                    const selectCellNode = figma.getNodeById(storedScan.selectCellId);
                    if (selectCellNode && selectCellNode.type === 'COMPONENT') {
                        selectCellComponent = selectCellNode;
                        console.log('✅ Restored select cell from stored ID:', selectCellComponent.name);
                    }
                }
                
                if (storedScan.expandCellId) {
                    const expandCellNode = figma.getNodeById(storedScan.expandCellId);
                    if (expandCellNode && expandCellNode.type === 'COMPONENT') {
                        expandCellComponent = expandCellNode;
                        console.log('✅ Restored expand cell from stored ID:', expandCellComponent.name);
                    }
                }
                
                if (storedScan.dividerId) {
                    const dividerNode = figma.getNodeById(storedScan.dividerId);
                    if (dividerNode && dividerNode.type === 'COMPONENT') {
                        dividerComponent = dividerNode;
                        console.log('✅ Restored divider from stored ID:', dividerComponent.name);
                    }
                }
                
                // If we have the essential components, return the scan result
                if (bodyCell) {
                    console.log('✅ Successfully restored components from stored scan data');
                    return {
                        headerCell,
                        headerRowComponent: null,
                        bodyCell,
                        bodyRowComponent: null,
                        footer,
                        numCols: storedScan.numCols || 5,
                        selectCellComponent,
                        expandCellComponent,
                        dividerComponent
                    };
                }
            } catch (error) {
                console.log('⚠️ Could not restore from stored scan data:', error);
            }
        }
        
        // Fallback: scan the generated table structure
        console.log('🔄 Scanning generated table structure manually...');
        
        let headerCell = null, bodyCell = null, footer = null;
        let selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
        let numCols = 0;
        
        // Find all instances in the generated table
        if ('findAll' in tableFrame) {
            const allInstances = tableFrame.findAll(n => n.type === 'INSTANCE') as InstanceNode[];
            console.log(`🔍 Found ${allInstances.length} instances in generated table`);
            
            // Group instances by their main component names
            const componentMap = new Map<string, ComponentNode>();
            
            for (const instance of allInstances) {
                if (instance.mainComponent) {
                    const compName = instance.mainComponent.name.toLowerCase();
                    
                    // Categorize components
                    if (!bodyCell && compName.includes('cell') && !compName.includes('header') && !compName.includes('footer') && !compName.includes('select') && !compName.includes('expand')) {
                        bodyCell = instance.mainComponent;
                        console.log(`✅ Found body cell: ${bodyCell.name}`);
                    } else if (!headerCell && compName.includes('header') && compName.includes('cell')) {
                        headerCell = instance.mainComponent;
                        console.log(`✅ Found header cell: ${headerCell.name}`);
                    } else if (!footer && (compName.includes('footer') || compName.includes('pagination'))) {
                        footer = instance.mainComponent;
                        console.log(`✅ Found footer: ${footer.name}`);
                    } else if (!selectCellComponent && compName.includes('select') && compName.includes('cell')) {
                        selectCellComponent = instance.mainComponent;
                        console.log(`✅ Found select cell: ${selectCellComponent.name}`);
                    } else if (!expandCellComponent && compName.includes('expand') && compName.includes('cell')) {
                        expandCellComponent = instance.mainComponent;
                        console.log(`✅ Found expand cell: ${expandCellComponent.name}`);
                    } else if (!dividerComponent && compName.includes('divider')) {
                        dividerComponent = instance.mainComponent;
                        console.log(`✅ Found divider: ${dividerComponent.name}`);
                    }
                    
                    componentMap.set(instance.mainComponent.id, instance.mainComponent);
                }
            }
            
            // Count columns by looking at the first row
            const headerRow = tableFrame.findOne(n => n.name === 'Header Row' && n.type === 'FRAME');
            if (headerRow && 'children' in headerRow) {
                numCols = headerRow.children.filter(child => 
                    child.type === 'INSTANCE' && 
                    child.name && 
                    !child.name.toLowerCase().includes('select') && 
                    !child.name.toLowerCase().includes('expand')
                ).length;
                console.log(`🔢 Counted ${numCols} columns from header row`);
            } else {
                // Fallback: count from first body row
                const bodyFrame = tableFrame.findOne(n => n.name === 'Body' && n.type === 'FRAME');
                if (bodyFrame && 'children' in bodyFrame) {
                    const firstBodyRow = bodyFrame.children[0];
                    if (firstBodyRow && 'children' in firstBodyRow) {
                        const dataRow = firstBodyRow.children.find(child => child.name && child.name.startsWith('Row'));
                        if (dataRow && 'children' in dataRow) {
                            numCols = dataRow.children.filter(child => 
                                child.type === 'INSTANCE' && 
                                child.name && 
                                !child.name.toLowerCase().includes('select') && 
                                !child.name.toLowerCase().includes('expand')
                            ).length;
                            console.log(`🔢 Counted ${numCols} columns from first body row`);
                        }
                    }
                }
            }
            
            // If we still don't have a body cell, this scan failed
            if (!bodyCell) {
                console.log('❌ Could not extract body cell from generated table');
                return null;
            }
            
            console.log('✅ Successfully scanned generated table structure');
            return {
                headerCell,
                headerRowComponent: null,
                bodyCell,
                bodyRowComponent: null,
                footer,
                numCols: numCols || 5,
                selectCellComponent,
                expandCellComponent,
                dividerComponent
            };
        }
        
        return null;
        
    } catch (error) {
        console.error('❌ Error scanning generated table:', error);
        return null;
    }
}

// Function to automatically scan for a valid "Data table" component on the page
async function autoScanForDataTable(): Promise<ScanResult | null> {
    console.log('🔄 Auto-scanning for Data table components on the page...');
    
    try {
        // Search for any "Data table" components on the current page
        const dataTableComponents = figma.currentPage.findAll(node => 
            node.type === 'INSTANCE' && 
            node.name.toLowerCase().includes('data table') &&
            !node.getPluginData('isGeneratedTable') // Exclude generated tables
        ) as InstanceNode[];
        
        console.log(`🔍 Found ${dataTableComponents.length} potential Data table components`);
        
        if (dataTableComponents.length === 0) {
            console.log('❌ No Data table components found on the page');
            return null;
        }
        
        // Use the first valid Data table component found
        const selectedComponent = dataTableComponents[0];
        console.log('✅ Using Data table component for auto-scan:', selectedComponent.name);
        
        // Simulate the scan process on this component by reusing the existing scan logic
        const scanResult = await performTableScan(selectedComponent);
        
        if (scanResult && scanResult.bodyCell) {
            console.log('✅ Auto-scan successful, found valid components');
            return scanResult;
        } else {
            console.log('❌ Auto-scan failed, no valid components found');
            return null;
        }
        
    } catch (error) {
        console.error('❌ Error during auto-scan:', error);
        return null;
    }
}

// Helper function to restore full scan result from stored data
async function restoreFullScanResult(storedScan: any): Promise<ScanResult | null> {
    try {
        let headerCell = null, bodyCell = null, footer = null;
        let selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
        let headerRowComponent = null, bodyRowComponent = null;
        
        // Restore components by ID
        if (storedScan.bodyCellId) {
            const bodyCellNode = figma.getNodeById(storedScan.bodyCellId);
            if (bodyCellNode && bodyCellNode.type === 'COMPONENT') {
                bodyCell = bodyCellNode;
            }
        }
        
        if (storedScan.headerCellId) {
            const headerCellNode = figma.getNodeById(storedScan.headerCellId);
            if (headerCellNode && headerCellNode.type === 'COMPONENT') {
                headerCell = headerCellNode;
            }
        }
        
        if (storedScan.footerId) {
            const footerNode = figma.getNodeById(storedScan.footerId);
            if (footerNode && footerNode.type === 'COMPONENT') {
                footer = footerNode;
            }
        }
        
        if (storedScan.selectCellId) {
            const selectCellNode = figma.getNodeById(storedScan.selectCellId);
            if (selectCellNode && selectCellNode.type === 'COMPONENT') {
                selectCellComponent = selectCellNode;
            }
        }
        
        if (storedScan.expandCellId) {
            const expandCellNode = figma.getNodeById(storedScan.expandCellId);
            if (expandCellNode && expandCellNode.type === 'COMPONENT') {
                expandCellComponent = expandCellNode;
            }
        }
        
        if (storedScan.dividerId) {
            const dividerNode = figma.getNodeById(storedScan.dividerId);
            if (dividerNode && dividerNode.type === 'COMPONENT') {
                dividerComponent = dividerNode;
            }
        }
        
        if (storedScan.headerRowComponentId) {
            const headerRowNode = figma.getNodeById(storedScan.headerRowComponentId);
            if (headerRowNode && headerRowNode.type === 'COMPONENT') {
                headerRowComponent = headerRowNode;
            }
        }
        
        if (storedScan.bodyRowComponentId) {
            const bodyRowNode = figma.getNodeById(storedScan.bodyRowComponentId);
            if (bodyRowNode && bodyRowNode.type === 'COMPONENT') {
                bodyRowComponent = bodyRowNode;
            }
        }
        
        if (bodyCell) {
            return {
                headerCell,
                headerRowComponent,
                bodyCell,
                bodyRowComponent,
                footer,
                numCols: storedScan.numCols || 5,
                selectCellComponent,
                expandCellComponent,
                dividerComponent
            };
        }
        
        return null;
    } catch (error) {
        console.error('Error restoring full scan result:', error);
        return null;
    }
}

// Helper function to update body row properties (placeholder for now)
async function updateBodyRowProperties(bodyRowInstance: InstanceNode): Promise<boolean> {
    try {
        // This function would update the body row to enable selectable/expandable functionality
        // For now, just return true to indicate success
        console.log('🔄 Updating body row properties for:', bodyRowInstance.name);
        return true;
    } catch (error) {
        console.error('❌ Error updating body row properties:', error);
        return false;
    }
}

// Helper function to check if variant combination is valid
function isValidVariantCombination(componentSet: ComponentSetNode, properties: { [key: string]: any }): boolean {
    try {
        // This is a simplified check - in practice you'd validate against the component set's variants
        return true;
    } catch (error) {
        console.error('❌ Error validating variant combination:', error);
        return false;
    }
}

// Helper function to map property names
function mapPropertyNames(uiProperties: { [key: string]: any }, componentProperties: { [key: string]: any }): { [key: string]: any } {
    const mappedProps: { [key: string]: any } = {};
    
    for (const [uiPropName, uiPropValue] of Object.entries(uiProperties)) {
        const availableProperties = Object.keys(componentProperties);
        const matchedProp = findMatchingProperty(availableProperties, uiPropName);
        
        if (matchedProp) {
            mappedProps[matchedProp] = uiPropValue;
        } else {
            // Try direct mapping as fallback
            if (componentProperties[uiPropName]) {
                mappedProps[uiPropName] = uiPropValue;
            }
        }
    }
    
    return mappedProps;
}




figma.ui.onmessage = async (msg: any) => {
    // Handle drag resize from corner
    if (msg.type === "resize") {
        const { size } = msg;
        let { w, h } = size;
        
        // Apply constraints
        if (h > pluginMaxHeight) {
            h = pluginMaxHeight;
        } else if (h < pluginMinHeight) {
            h = pluginMinHeight;
        }
        
        if (w > pluginMaxWidth) {
            w = pluginMaxWidth;
        } else if (w < pluginMinWidth) {
            w = pluginMinWidth;
        }
        
        // Resize the plugin
        figma.ui.resize(w, h);
        
        // Save the size for next time
        figma.clientStorage.setAsync('pluginSize', { w, h }).catch(err => {
            console.log('Failed to save plugin size:', err);
        });
        
        console.log(`Plugin resized to ${w}×${h}`);
        return;
    }
    
    // Handle programmatic resize requests (keeping for backward compatibility)
    if (msg.type === "resize-ui") {
        const { width, height } = msg;
        // Constrain dimensions for safety
        const constrainedWidth = Math.max(pluginMinWidth, Math.min(pluginMaxWidth, width || pluginDefaultWidth));
        const constrainedHeight = Math.max(pluginMinHeight, Math.min(pluginMaxHeight, height || pluginDefaultHeight));
        
        figma.ui.resize(constrainedWidth, constrainedHeight);
        
        // Save the size for next time
        figma.clientStorage.setAsync('pluginSize', { w: constrainedWidth, h: constrainedHeight }).catch(err => {
            console.log('Failed to save plugin size:', err);
        });
        
        console.log(`Plugin resized to ${constrainedWidth}×${constrainedHeight}`);
        return;
    }
    
    if (msg.type === "scan-selected-table") {
        const tableNode = figma.getNodeById(msg.tableId);
        console.log(`[scan-selected-table] Node found: ${tableNode ? tableNode.type : 'null'}, name: ${tableNode?.name}`);
        if (tableNode && (tableNode.type === "FRAME" || tableNode.type === "COMPONENT" || tableNode.type === "INSTANCE")) {
            figma.currentPage.selection = [tableNode];
            console.log('[scan-selected-table] Triggering selectionchange for table scan');
            // Trigger the selectionchange handler which will handle the scan
        } else {
            console.warn(`[scan-selected-table] Invalid node type for scanning: ${tableNode?.type}`);
        }
        return;
    }
    else if (msg.type === "generate-fake-data") {
        const { dataType, count } = msg;
        const data = [];
        for (let i = 0; i < (count || 10); i++) {
            try {
                const [category, method] = dataType.split('.');
                if (faker[category] && typeof faker[category][method] === 'function') {
                    const value = faker[category][method]();
                    data.push(value);
                } else {
                    data.push('Invalid Method');
                }
            } catch (e) {
                data.push('Error');
            }
        }
        figma.ui.postMessage({
            type: 'fake-data-response',
            data
        });
        return;

    } else if (msg.type === "generate-watsonx-data") {
        try {
            const { prompt, endpoint, apiKey, accessToken: uiAccessToken, useAccessToken, useProxy, proxyUrl, count, remember } = msg as { prompt: string, endpoint: string, apiKey: string, accessToken?: string, useAccessToken?: boolean, useProxy?: boolean, proxyUrl?: string, count: number, remember?: boolean };
            let endpointToUse = endpoint;
            let apiKeyToUse = apiKey;
            // Fallback to stored values if missing
            if (!endpointToUse) {
                const storedEndpoint = await figma.clientStorage.getAsync('watsonx.endpoint');
                if (storedEndpoint) endpointToUse = String(storedEndpoint);
            }
            if (!apiKeyToUse) {
                const storedKey = await figma.clientStorage.getAsync('watsonx.apiKey');
                if (storedKey) apiKeyToUse = String(storedKey);
            }
            if (remember) {
                try {
                    if (endpoint) await figma.clientStorage.setAsync('watsonx.endpoint', endpoint);
                    // Store API key securely in clientStorage (still local to the user’s device)
                    if (apiKey) await figma.clientStorage.setAsync('watsonx.apiKey', apiKey);
                } catch (e) {
                    console.warn('Failed to persist watsonx settings', e);
                }
            }
            if (!endpointToUse || (!useAccessToken && !apiKeyToUse && !uiAccessToken)) {
                figma.ui.postMessage({ type: 'watsonx-data-response', data: [] });
                figma.notify('Missing watsonx endpoint or credentials');
                return;
            }
            // token + generation flow via proxy or direct
            let accessToken = (useAccessToken && uiAccessToken) ? uiAccessToken : '';
            if (useProxy && proxyUrl) {
                console.log('Using proxy server:', proxyUrl);
                // Use proxy endpoints
                if (!accessToken) {
                    console.log('Getting token via proxy...');
                    const tokenRes = await fetch(`${proxyUrl.replace(/\/$/, '')}/token`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ apiKey: apiKeyToUse })
                    });
                    if (!tokenRes.ok) throw new Error('Proxy token fetch failed');
                    const tokenJson = await tokenRes.json();
                    accessToken = tokenJson.access_token as string;
                    console.log('Got token via proxy');
                }
                console.log('Generating text via proxy...');
                const genRes = await fetch(`${proxyUrl.replace(/\/$/, '')}/generate`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ endpoint: endpointToUse, accessToken, prompt, count })
                });
                if (!genRes.ok) {
                    const t = await genRes.text();
                    throw new Error(`proxy watsonx error: ${genRes.status} ${t}`);
                }
                const genJson = await genRes.json();
                console.log('Proxy response:', genJson);
                const lines = genJson.data || [];
                figma.ui.postMessage({ type: 'watsonx-data-response', data: lines });
            } else {
                console.log('Using direct API calls (may fail due to CORS)');
                // Direct endpoints (may CORS fail depending on Figma env)
                if (!accessToken) {
                    const tokenRes = await fetch('https://iam.cloud.ibm.com/identity/token', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Accept': 'application/json' },
                        body: `grant_type=urn:ibm:params:oauth:grant-type:apikey&apikey=${encodeURIComponent(apiKeyToUse)}`
                    });
                    if (!tokenRes.ok) throw new Error('Failed to fetch IAM token');
                    const tokenJson = await tokenRes.json();
                    accessToken = tokenJson.access_token as string;
                }
                const genUrl = `${endpointToUse.replace(/\/$/, '')}/ml/v1/text/chat?version=2023-05-29`;
                const wxBody = {
                    input: `${prompt}\nReturn ${count} short values as a plain list, one per line, no numbering.`,
                    parameters: { decoding_method: 'greedy', max_new_tokens: 16, stop_sequences: ['\n\n'] }
                } as any;
                const genRes = await fetch(genUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${accessToken}` }, body: JSON.stringify(wxBody) });
                if (!genRes.ok) {
                    const t = await genRes.text();
                    throw new Error(`watsonx error: ${genRes.status} ${t}`);
                }
                const genJson = await genRes.json();
                let text = '';
                if (Array.isArray(genJson.results) && genJson.results[0]?.generated_text) text = genJson.results[0].generated_text as string;
                else if (genJson.generated_text) text = genJson.generated_text as string;
                else text = String(genJson.output || '');
                const lines = text.split('\n').map((s: string) => s.replace(/^[-*\d\.\)\s]+/, '').trim()).filter((s: string) => s.length > 0).slice(0, count);
                figma.ui.postMessage({ type: 'watsonx-data-response', data: lines });
            }
        } catch (err: any) {
            figma.ui.postMessage({ type: 'watsonx-data-response', data: [] });
            figma.notify('watsonx.ai request failed');
            console.error('watsonx.ai error', err);
        }
        return;

    } else if (msg.type === 'generate-table-with-ai') {
        try {
            const { prompt, apiKey, rows, cols } = msg as { prompt: string, apiKey: string, rows: number, cols: number };
            const proxyUrl = 'http://localhost:3000';
            const endpoint = 'https://us-south.ml.cloud.ibm.com';

            const tokenRes = await fetch(`${proxyUrl}/token`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ apiKey })
            });
            if (!tokenRes.ok) {
                const t = await tokenRes.text();
                throw new Error(`proxy token error: ${tokenRes.status} ${t}`);
            }
            const { access_token } = await tokenRes.json();

            const tableRes = await fetch(`${proxyUrl}/generateTable`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ endpoint, accessToken: access_token, prompt, rows, cols })
            });
            if (!tableRes.ok) {
                const t = await tableRes.text();
                throw new Error(`proxy generateTable error: ${tableRes.status} ${t}`);
            }
            const tableJson = await tableRes.json();
            const headers: string[] = Array.isArray(tableJson.headers) ? tableJson.headers.map(String).slice(0, cols) : Array(cols).fill('').map((_, i) => `Column ${i+1}`);
            const rowsData: string[][] = Array.isArray(tableJson.rows) ? tableJson.rows.map((r: any) => Array.isArray(r) ? r.map(String) : []) : [];
            figma.ui.postMessage({ type: 'ai-table-response', headers, rows: rowsData });
        } catch (e) {
            console.error('generate-table-with-ai error', e);
            figma.ui.postMessage({ type: 'ai-table-response', headers: [], rows: [] });
            figma.notify('Failed to generate table with AI');
        }
        return;

    } else if (msg.type === "load-instance-by-id") {
        const nodeId = msg.nodeId;
        const node = figma.getNodeById(nodeId);

        if (!node) {
            figma.ui.postMessage({
                type: "creation-error",
                message: "Node not found."
            });
            return;
        }
        if (node && (node.type === "INSTANCE" || node.type === "COMPONENT" || node.type === "COMPONENT_SET")) {
            try {
                let componentNode: ComponentNode | ComponentSetNode | null = null;
                if (node.type === "INSTANCE") {
                    componentNode = await node.getMainComponentAsync();
                } else {
                    componentNode = node;
                }

                if (!componentNode) throw new Error("Component not found");

                if (componentNode.type === "COMPONENT_SET") {
                    componentNode = componentNode.children[0] as ComponentNode;
                }

                const tempInstance = (componentNode as ComponentNode).createInstance();
                const propertyValues: { [key: string]: any } = {};
                const propertyTypes: { [key: string]: "TEXT" | "BOOLEAN" | "INSTANCE_SWAP" | "VARIANT" } = {};
                const availableProperties = Object.keys(tempInstance.componentProperties);
                for (const propName of availableProperties) {
                    const prop = tempInstance.componentProperties[propName];
                    propertyValues[propName] = prop.value;
                    propertyTypes[propName] = prop.type;
                }
                tempInstance.remove();

                figma.ui.postMessage({
                    type: "component-selected",
                    componentName: componentNode.name || "Table Cell",
                    componentId: componentNode.id || null,
                    properties: propertyValues,
                    availableProperties: availableProperties,
                    propertyTypes: propertyTypes,
                    isValidComponent: true
                });
            } catch (error) {
                figma.ui.postMessage({
                    type: "creation-error",
                    message: "Failed to load component properties."
                });
            }
        } else {
            figma.ui.postMessage({
                type: "creation-error",
                message: "Node not found or not a valid component/instance."
            });
        }
        return;
    
    } else if (msg.type === "update-cell-properties") {
        const { cellKey, properties } = msg as { cellKey: string, properties: { [key: string]: string } };
        try {
            const cell = cellInstanceMap.get(cellKey);
            if (cell) {
                let mainComponentSet = null;
                if (cell.mainComponent && cell.mainComponent.parent && cell.mainComponent.parent.type === "COMPONENT_SET") {
                    mainComponentSet = cell.mainComponent.parent as ComponentSetNode;
                }
                
                if (mainComponentSet && isValidVariantCombination(mainComponentSet, properties)) {
                    try {
                await cell.setProperties(properties);
                figma.notify('✅ Cell properties updated');
                    } catch (error) {
                        console.error('Error in update-cell-properties setProperties:', error);
                        figma.notify('❌ Error updating cell properties');
                    }
                } else {
                    console.warn('Invalid variant combination for cell update:', properties);
                    figma.notify('⚠️ Invalid property combination');
                }
            } else {
                figma.notify('Error: Cell not found');
            }
        } catch (error: any) {
            figma.notify(`Error: ${error.message}`);
        }
    
    } else if (msg.type === "save-watsonx-api-key") {
        try {
            const { apiKey } = msg;
            if (apiKey) {
                await figma.clientStorage.setAsync('watsonx.apiKey', apiKey);
                console.log('API key saved successfully');
            }
        } catch (e) {
            console.warn('Failed to save watsonx API key', e);
        }
        return;
        
    } else if (msg.type === "load-watsonx-settings") {
        try {
            const endpoint = await figma.clientStorage.getAsync('watsonx.endpoint');
            const apiKey = await figma.clientStorage.getAsync('watsonx.apiKey');
            const masked = apiKey ? `${String(apiKey).slice(0,4)}••••${String(apiKey).slice(-4)}` : '';
            
            // Send both masked version for display and full key for auto-population
            figma.ui.postMessage({ 
                type: 'watsonx-settings', 
                endpoint, 
                apiKeyMasked: masked,
                fullApiKey: apiKey ? String(apiKey) : null
            });
        } catch (e) {
            figma.ui.postMessage({ type: 'watsonx-settings', endpoint: '', apiKeyMasked: '' });
        }
        return;

    } else if (msg.type === "request-selection-state") {
        const selection = figma.currentPage.selection;
        
        if (selection.length === 1) {
            const selectedNode = selection[0];
            
            // Check if it's a Data table component
            if (selectedNode.type === "FRAME" && selectedNode.name.includes('Data table')) {
                figma.ui.postMessage({
                    type: "table-selected",
                    tableId: selectedNode.id
                });
                return;
            }
            
            // Check if it's a Data table instance
            if (selectedNode.type === "INSTANCE") {
                try {
                    const mainComponent = await selectedNode.getMainComponentAsync();
                    if (mainComponent && mainComponent.name.includes('Data table')) {
        figma.ui.postMessage({
                            type: "table-selected",
                            tableId: selectedNode.id
                        });
                        return;
                    }
                } catch (error) {
                    console.error('Error getting main component:', error);
                }
            }
            
            // Check if it's a Data table row cell item
            if (selectedNode.type === "INSTANCE") {
                try {
                    const mainComponent = await selectedNode.getMainComponentAsync();
                    if (mainComponent && mainComponent.name === "Data table row cell item") {
                        // Don't trigger scan for individual cell items - just ignore the selection
                        console.log(`[request-selection-state] Ignoring selection of individual cell item: ${selectedNode.name}`);
                        return;
                    }
                } catch (error) {
                    console.error('Error getting main component:', error);
                }
            }
        }
        
        // If no valid selection found, send selection cleared message
        figma.ui.postMessage({
            type: "selection-cleared",
            isValidComponent: false,
            clearUI: true
        });
    
    } else if (msg.type === "request-component-info") {
        console.log('[Backend] Received request-component-info');
        
        let tableId = msg.tableId;
        if (!tableId) {
            const selection = figma.currentPage.selection;
            if (selection.length === 1 && selection[0].getPluginData('isGeneratedTable') === 'true') {
                tableId = selection[0].id;
            }
        }
        
        // If we have a table ID, scan it directly to find body cell instances
        if (tableId) {
            const tableNode = figma.getNodeById(tableId);
            if (tableNode && 'findOne' in tableNode) {
                console.log('[Backend] Scanning generated table for body cell instances');
                
                // Find any body cell instance in the generated table
                const bodyCellInstance = tableNode.findOne(n => 
                    n.type === 'INSTANCE' && 
                    n.mainComponent !== null &&
                    !n.name.toLowerCase().includes('header') &&
                    !n.name.toLowerCase().includes('footer') &&
                    !n.name.toLowerCase().includes('select') &&
                    !n.name.toLowerCase().includes('expand')
                ) as InstanceNode | null;
                
                if (bodyCellInstance && bodyCellInstance.mainComponent) {
                    console.log('[Backend] Found body cell instance in table:', bodyCellInstance.mainComponent.name);
                    
                    const bodyCell = bodyCellInstance.mainComponent;
                    const propertyValues: { [key: string]: any } = {};
                    const propertyTypes: { [key: string]: "TEXT" | "BOOLEAN" | "INSTANCE_SWAP" | "VARIANT" } = {};
                    const availableProperties = Object.keys(bodyCellInstance.componentProperties);
                    
                    for (const propName of availableProperties) {
                        const prop = bodyCellInstance.componentProperties[propName];
                        propertyValues[propName] = prop.value;
                        propertyTypes[propName] = prop.type;
                    }
                    
                    figma.ui.postMessage({
                        type: "component-info",
                        component: {
                            id: bodyCell.id,
                            name: bodyCell.name,
                            width: bodyCellInstance.width,
                            properties: propertyValues,
                            availableProperties,
                            propertyTypes,
                        }
                    });
                    return;
                }
            }
        }
        
        // Fallback to lastScanResult if available
        if (lastScanResult && lastScanResult.bodyCell) {
            const bodyCell = lastScanResult.bodyCell;
            const tempInstance = bodyCell.createInstance();
            const propertyValues: { [key: string]: any } = {};
            const propertyTypes: { [key: string]: "TEXT" | "BOOLEAN" | "INSTANCE_SWAP" | "VARIANT" } = {};
            const availableProperties = Object.keys(tempInstance.componentProperties);
            for (const propName of availableProperties) {
                const prop = tempInstance.componentProperties[propName];
                propertyValues[propName] = prop.value;
                propertyTypes[propName] = prop.type;
            }
            const instanceWidth = tempInstance.width;
            tempInstance.remove();
            
            figma.ui.postMessage({
                type: "component-info",
                component: {
                    id: bodyCell.id,
                    name: bodyCell.name,
                    width: instanceWidth,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes,
                }
            });
            return;
        }
        
        // Final fallback
        figma.ui.postMessage({
            type: "component-info",
            component: {
                id: 'fallback',
                name: 'Fallback Component',
                width: 100,
                properties: { 'Cell text#12234:16': 'Sample Text' },
                availableProperties: ['Cell text#12234:16'],
                propertyTypes: { 'Cell text#12234:16': 'TEXT' as const }
            }
        });
        
    } else if (msg.type === "clear-cell-instances") {
        cellInstanceMap.forEach(instance => instance.remove());

    } else if (msg.type === 'scan-table') {
        const now = Date.now();
        if (now - lastScanTime < SCAN_DEBOUNCE_MS) {
            console.log(`[scan-table] Debouncing scan - ${now - lastScanTime}ms since last scan`);
            return;
        }
        lastScanTime = now;
        
        const selection = figma.currentPage.selection;
        console.log(`[scan-table] Selection check - count: ${selection.length}, types: ${selection.map(s => s.type).join(', ')}, names: ${selection.map(s => s.name).join(', ')}`);
        
        if (selection.length !== 1 || !['FRAME', 'COMPONENT', 'INSTANCE'].includes(selection[0].type)) {
            console.log(`[scan-table] Invalid selection - count: ${selection.length}, type: ${selection[0]?.type}`);
            figma.ui.postMessage({
                type: 'scan-table-result',
                success: false,
                message: 'Please select a single table component (Frame, Component or Instance) in Figma.'
            });
            return;
        }
        console.log(`[scan-table] Valid selection - type: ${selection[0].type}, name: ${selection[0].name}`);

        const tableFrame = selection[0];
        
        // Check if the selected item is a cell item that should be ignored
        if (tableFrame.type === "INSTANCE") {
            try {
                const mainComponent = await tableFrame.getMainComponentAsync();
                if ((mainComponent && mainComponent.name === "Data table row cell item") || 
                    tableFrame.name === "Data table body row item" ||
                    tableFrame.name.includes("row cell item")) {
                    console.log(`[scan-table] Ignoring scan for individual cell item: ${tableFrame.name}`);
                    figma.ui.postMessage({
                        type: 'scan-table-result',
                        success: false,
                        message: 'Please select a main table component, not individual cells.'
                    });
                    return;
                }
            } catch (error) {
                console.error('Error checking main component in scan-table:', error);
            }
        }
        
        // --- Set the 'type' property to 'Expandable + Selectable' if possible ---
        if ('componentProperties' in tableFrame && 'setProperties' in tableFrame && typeof tableFrame.setProperties === 'function') {
          const typeProp = tableFrame.componentProperties['type'];
          
          // Find the correct property name for the type/variant property
          let typePropertyName = null;
          for (const [propName, prop] of Object.entries(tableFrame.componentProperties)) {
            if (prop.type === 'VARIANT') {
              // Check if this property has expandable/selectable options
              if ('mainComponent' in tableFrame && tableFrame.mainComponent && 'variantProperties' in tableFrame.mainComponent) {
                const variantProps = tableFrame.mainComponent.variantProperties as any;
                const variantKey = Object.keys(variantProps).find(k => k.startsWith(propName));
                if (variantKey && variantProps[variantKey] && typeof variantProps[variantKey] === 'object' && 'values' in variantProps[variantKey] && Array.isArray(variantProps[variantKey].values)) {
                  const possibleValues = variantProps[variantKey].values as string[];
                  // Find a case-insensitive match for expandable + selectable
                  const found = possibleValues.find(v => v.trim().toLowerCase().includes('expandable') && v.trim().toLowerCase().includes('selectable'));
                  if (found) {
                    typePropertyName = propName;
                  }
                } else {
                }
              } else {
              }
            }
          }
          
          
          if (typePropertyName && tableFrame.mainComponent) {
            // Find the correct variantKey again for setting
            const variantProps = tableFrame.mainComponent.variantProperties as any;
            const variantKey = Object.keys(variantProps).find(k => k.startsWith(typePropertyName));
            let targetValue = 'Expandable + Selectable';
            let variantValue = targetValue;
            if (variantKey && variantProps[variantKey] && typeof variantProps[variantKey] === 'object' && 'values' in variantProps[variantKey] && Array.isArray(variantProps[variantKey].values)) {
              const possibleValues = variantProps[variantKey].values as string[];
              // Find a case-insensitive match
              const found = possibleValues.find(v => v.trim().toLowerCase().includes('expandable') && v.trim().toLowerCase().includes('selectable'));
              if (found) {
                variantValue = found;
              } else {
              }
            }
            const propertiesToSet: any = {};
            propertiesToSet[typePropertyName] = variantValue;
            
            // Check if the combination is valid before setting
            let mainComponentSet = null;
            if (tableFrame.mainComponent && tableFrame.mainComponent.parent && tableFrame.mainComponent.parent.type === "COMPONENT_SET") {
                mainComponentSet = tableFrame.mainComponent.parent as ComponentSetNode;
            }
            
            if (mainComponentSet && isValidVariantCombination(mainComponentSet, propertiesToSet)) {
                try {
            tableFrame.setProperties(propertiesToSet);
            // Wait for Figma to update the instance tree
            await Promise.resolve();
            // Add additional delay to ensure the instance tree is updated
            await new Promise(resolve => setTimeout(resolve, 100));
                } catch (error) {
                    console.error('Error in tableFrame setProperties:', error);
                }
            } else {
                console.warn('Invalid variant combination for table frame:', propertiesToSet);
            }
          } else {
          }
        } else {
        }
        // --- Custom scan for specific instances/components ---
        let expandCellFound = false;
        let selectCellFound = false;
        const allInstanceNames: string[] = [];
        let allInstances: (InstanceNode | ComponentNode)[] = [];
        let expandCellComponent: ComponentNode | null = null;
        let selectCellComponent: ComponentNode | null = null;

        // --- Check for 'type' property with value 'expandable + selectable' ---
        let hasExpandableSelectableType = false;
        if ('componentProperties' in tableFrame) {
          const typeProp = tableFrame.componentProperties['type'];
          if (typeProp && typeProp.value && typeof typeProp.value === 'string') {
            if (typeProp.value.trim().toLowerCase() === 'expandable + selectable') {
              hasExpandableSelectableType = true;
            }
          }
        }

        // --- Print all componentProperties and their possible values ---
        if ('componentProperties' in tableFrame) {
          const props = tableFrame.componentProperties;
          let variantValues: { [key: string]: string[] } = {};
          if ('mainComponent' in tableFrame && tableFrame.mainComponent && 'variantProperties' in tableFrame.mainComponent) {
            // Figma API: mainComponent.variantProperties[key].values
            for (const [k, v] of Object.entries(tableFrame.mainComponent.variantProperties || {})) {
              if (v && typeof v === 'object' && 'values' in v && Array.isArray((v as any).values)) {
                variantValues[k] = (v as any).values;
              }
            }
          }
          for (const [key, prop] of Object.entries(props)) {
            let possibleValues = undefined;
            if (prop.type === 'VARIANT' && variantValues[key]) {
              possibleValues = variantValues[key];
            } else if (prop.type === 'BOOLEAN') {
              possibleValues = [true, false];
            }
          }
        }

        // --- Generate unique instances/components in a preview frame ---
        // Preview frame creation removed as requested

        if ('children' in tableFrame) {
        }
        if ('findOne' in tableFrame && typeof tableFrame.findOne === 'function') {
            // Find header row (INSTANCE)
            const headerRow = tableFrame.findOne(n =>
                n.name?.includes('header row') && n.type === 'INSTANCE'
            ) as InstanceNode | undefined;
            // Find body group (FRAME)
            const bodyGroup = tableFrame.findOne(n => n.name === 'Body' && n.type === 'FRAME');
            // Find footer bar (INSTANCE)
            const footerBar = tableFrame.findOne(n => n.name === 'Pagination - Table bar' && n.type === 'INSTANCE');
            // FIRST INSTANCE - Variable declarations for component extraction
            let headerCell: InstanceNode | null = null;
            let bodyCell: InstanceNode | null = null;
            let footer: InstanceNode | null = null;
            let numCols = 0;
            let headerRowComponent: ComponentNode | null = null;
            let bodyRowComponent: ComponentNode | null = null;
            // Extract header cell template
            if (headerRow && 'children' in headerRow) {
                headerRowComponent = await headerRow.getMainComponentAsync();
                
                console.log('🔄 Updating header row properties for:', headerRow.name);
                await updateBodyRowProperties(headerRow);
                console.log('✅ Header row properties updated successfully');
                
                const headerCellInstance = headerRow.children.find((n: SceneNode) =>
                    n.type === 'INSTANCE' && n.name?.toLowerCase().includes('col 1')
                ) as InstanceNode | undefined;
                if (headerCellInstance) {
                    headerCell = headerCellInstance;
                    // Count only instances that are visible data columns (name includes 'col' but not select/expand cells)
                    numCols = headerRow.children.filter((n: SceneNode) => 
                        n.type === 'INSTANCE' && 
                        n.name?.toLowerCase().includes('col') &&
                        !n.name?.toLowerCase().includes('select') &&
                        !n.name?.toLowerCase().includes('expand') &&
                        n.visible
                    ).length;
                }
            }
            // Extract body cell template
            if (bodyGroup && 'children' in bodyGroup) {
                const bodyRowInstance = bodyGroup.children.find((n: SceneNode) => n.type === 'INSTANCE' && n.name === 'Data table body row item') as InstanceNode | undefined;
                if (bodyRowInstance) {
                    bodyRowComponent = await bodyRowInstance.getMainComponentAsync();
                    
                    console.log('🔄 Updating body row properties for:', bodyRowInstance.name);
                    await updateBodyRowProperties(bodyRowInstance);
                    console.log('✅ Body row properties updated successfully');
                    
                    if ('children' in bodyRowInstance) {
                        const dataTableRow = bodyRowInstance.children.find(child => child.type === 'FRAME' && child.name === 'Data table row') as FrameNode | undefined;
                        if (dataTableRow && 'children' in dataTableRow) {
                            // Find the first data cell instance whose name includes 'Col 1'
                            const bodyCellInstance = dataTableRow.children.find((cell: SceneNode) =>
                                cell.type === 'INSTANCE' && cell.name?.toLowerCase().includes('col 1')
                            ) as InstanceNode | undefined;
                            if (bodyCellInstance) {
                                // REVERTING - Keep as InstanceNode for performTableScan
                                bodyCell = bodyCellInstance;
                                console.log('✅ Extracted body cell component from instance:', bodyCell?.name);
                            }
                        }
                    }
                }
            }
            // Extract footer template
            if (footerBar && footerBar.type === 'INSTANCE') {
                footer = footerBar;
            }
            
            // --- Enhanced divider scanning - look in the specific path first ---
            console.log('🔍 Scanning for divider rectangles in specific path: Data table > Body > Data table body row item > Divider...');
            let foundDividerRect: RectangleNode | null = null;
            
            if (bodyGroup && 'children' in bodyGroup) {
                const bodyRowInstance = bodyGroup.children.find((n: SceneNode) => n.type === 'INSTANCE' && n.name === 'Data table body row item') as InstanceNode | undefined;
                if (bodyRowInstance && 'children' in bodyRowInstance) {
                    console.log('🔍 Found Data table body row item, searching for Divider rectangle inside...');
                    
                    // Look for the Divider rectangle directly in the body row item
                    const dividerInBodyRow = bodyRowInstance.children.find((child: SceneNode) => 
                        child.type === 'RECTANGLE' && child.name === 'Divider'
                    ) as RectangleNode | undefined;
                    
                    if (dividerInBodyRow) {
                        foundDividerRect = dividerInBodyRow;
                        console.log('✅ Found Divider rectangle in Data table body row item:', dividerInBodyRow.name);
                        console.log('🎨 Divider rectangle properties:', {
                            width: dividerInBodyRow.width,
                            height: dividerInBodyRow.height,
                            fills: dividerInBodyRow.fills,
                            locked: dividerInBodyRow.locked
                        });
                    } else {
                        console.log('⚠️ No Divider rectangle found directly in Data table body row item, checking nested children...');
                        
                        // Check nested children in case the divider is deeper in the hierarchy
                        for (const child of bodyRowInstance.children) {
                            if ('children' in child) {
                                const nestedDivider = child.children.find((grandChild: SceneNode) => 
                                    grandChild.type === 'RECTANGLE' && grandChild.name === 'Divider'
                                ) as RectangleNode | undefined;
                                
                                if (nestedDivider) {
                                    foundDividerRect = nestedDivider;
                                    console.log('✅ Found Divider rectangle in nested structure:', nestedDivider.name);
                                    console.log('🎨 Nested divider rectangle properties:', {
                                        width: nestedDivider.width,
                                        height: nestedDivider.height,
                                        fills: nestedDivider.fills,
                                        locked: nestedDivider.locked
                                    });
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            
            // --- Scan for select/expand cell components and divider AFTER updating body row properties ---
            let dividerComponent: ComponentNode | null = null;
            if ('findAll' in tableFrame && typeof tableFrame.findAll === 'function') {
                allInstances = tableFrame.findAll(n => n.type === 'INSTANCE' || n.type === 'COMPONENT') as (InstanceNode | ComponentNode)[];
                for (const node of allInstances) {
                    if (node.name === 'Data table expand cell item') {
                      expandCellFound = true;
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        expandCellComponent = node.mainComponent;
                      }
                    }
                    // Check for select/checkbox components with various common names
                    const isSelectCell = node.name === 'Data table select cell item' ||
                                       node.name.includes('select') ||
                                       node.name.includes('checkbox') ||
                                       node.name.includes('check') ||
                                       node.name.includes('radio') ||
                                       node.name.toLowerCase().includes('select') ||
                                       node.name.toLowerCase().includes('checkbox');
                    
                    if (isSelectCell && !selectCellFound) {
                      selectCellFound = true;
                      console.log('🔍 Found select cell component:', node.name, node.type);
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        selectCellComponent = node.mainComponent;
                        console.log('✅ Using select cell from instance:', selectCellComponent.name);
                      } else if (node.type === 'COMPONENT') {
                        selectCellComponent = node;
                        console.log('✅ Using select cell component directly:', selectCellComponent.name);
                      }
                    }
                    // Check for divider components - prioritize exact name match and skip AI labels
                    if (node.name.includes('AI label') || node.name.includes('Select menu')) {
                        continue; // Skip AI labels and select menus to improve performance
                    }
                    
                    const isDivider = node.name === 'Divider' || 
                                    (node.name.toLowerCase().includes('divider') && !node.name.includes('AI')) ||
                                    node.name.includes('line') ||
                                    node.name.includes('border') ||
                                    node.name.includes('separator');
                    
                    if (isDivider) {
                      console.log('🔍 Found potential divider component:', node.name, node.type);
                      
                      // Check if this component is actually a visual divider (not text-based)
                      let isValidDivider = false;
                      
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        // Check the main component's children
                        if ('children' in node.mainComponent) {
                          for (const child of node.mainComponent.children) {
                            if (child.type === 'RECTANGLE' || child.type === 'LINE' || child.type === 'VECTOR') {
                              isValidDivider = true;
                              console.log('✅ Valid divider found with visual element:', child.type);
                              break;
                            } else if (child.type === 'TEXT') {
                              console.log('⚠️ Skipping text-based component:', node.name);
                              break;
                            }
                          }
                        }
                      } else if (node.type === 'COMPONENT') {
                        // Check the component's children
                        if ('children' in node) {
                          for (const child of node.children) {
                            if (child.type === 'RECTANGLE' || child.type === 'LINE' || child.type === 'VECTOR') {
                              isValidDivider = true;
                              console.log('✅ Valid divider found with visual element:', child.type);
                              break;
                            } else if (child.type === 'TEXT') {
                              console.log('⚠️ Skipping text-based component:', node.name);
                              break;
                            }
                          }
                        }
                      }
                      
                      if (isValidDivider) {
                        if (node.type === 'INSTANCE' && node.mainComponent) {
                          dividerComponent = node.mainComponent;
                          console.log('✅ Using divider from instance:', dividerComponent.name);
                        } else if (node.type === 'COMPONENT') {
                          dividerComponent = node;
                          console.log('✅ Using divider component directly:', dividerComponent.name);
                        }
                      } else {
                        console.log('⚠️ Skipping invalid divider component:', node.name);
                      }
                    }
                    allInstanceNames.push(node.name);
                }
            }
            
            // If select cell component is still not found, log a warning but don't create one
            if (!selectCellComponent) {
                console.warn('⚠️ Select cell component not found in scanned table. Selectable functionality will be disabled.');
            }
            
            // Log divider component status
            if (!dividerComponent) {
                console.log('⚠️ No divider component found, will use original divider rectangle if available');
            }
            
            // Get properties for the bodyCell template to send to UI
            let bodyCellComponent = null;
            if (bodyCell) {
                try {
                const mainComponent = bodyCell.mainComponent;
                    if (mainComponent) {
                        const instance = bodyCell;
                const propertyValues: { [key: string]: any } = {};
                const propertyTypes: { [key: string]: any } = {};
                const availableProperties = Object.keys(instance.componentProperties);
                for (const propName of availableProperties) {
                    const prop = instance.componentProperties[propName];
                    propertyValues[propName] = prop.value;
                    propertyTypes[propName] = prop.type;
                }
                        let instanceWidth = instance?.width || 100; // Use optional chaining
                bodyCellComponent = {
                    id: mainComponent.id,
                    name: mainComponent.name,
                            width: instanceWidth,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes,
                };
                    }
                } catch (error) {
                    console.error('Error accessing bodyCell mainComponent:', error);
                }
            }
            // Get properties for the headerCell template to send to UI
            let headerCellComponent = null;
            if (headerCell) {
                try {
                const mainComponent = headerCell.mainComponent;
                    if (mainComponent) {
                        const instance = headerCell;
                const propertyValues: { [key: string]: any } = {};
                const propertyTypes: { [key: string]: any } = {};
                const availableProperties = Object.keys(instance.componentProperties);
                for (const propName of availableProperties) {
                    const prop = instance.componentProperties[propName];
                    propertyValues[propName] = prop.value;
                    propertyTypes[propName] = prop.type;
                }
                        let instanceWidth = instance?.width || 100; // Use optional chaining
                headerCellComponent = {
                    id: mainComponent.id,
                    name: mainComponent.name,
                            width: instanceWidth,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes,
                };
                    }
                } catch (error) {
                    console.error('Error accessing headerCell mainComponent:', error);
                }
            }
            // Get properties for the footer template to send to UI
            let footerComponent = null;
            if (footer) {
                try {
                const mainComponent = footer.mainComponent;
                    if (mainComponent) {
                        const instance = footer;
                const propertyValues: { [key: string]: any } = {};
                const propertyTypes: { [key: string]: any } = {};
                const availableProperties = Object.keys(instance.componentProperties);
                for (const propName of availableProperties) {
                    const prop = instance.componentProperties[propName];
                    propertyValues[propName] = prop.value;
                    propertyTypes[propName] = prop.type;
                }
                        let instanceWidth = instance?.width || 100; // Use optional chaining
                footerComponent = {
                    id: mainComponent.id,
                    name: mainComponent.name,
                            width: instanceWidth,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes,
                };
                    }
                } catch (error) {
                    console.error('Error accessing footer mainComponent:', error);
                }
            }
            // Store header row properties for applying to generated header row
            let headerRowProperties: { [key: string]: any } = {};
            if (headerRow) {
                console.log('🔍 Header row properties available:', {
                    fills: headerRow.fills,
                    strokes: headerRow.strokes,
                    strokeWeight: headerRow.strokeWeight,
                    cornerRadius: headerRow.cornerRadius,
                    effects: headerRow.effects
                });
                
                // Store the actual visual properties from the header row
                headerRowProperties = {
                    fills: headerRow.fills,
                    strokes: headerRow.strokes,
                    strokeWeight: headerRow.strokeWeight,
                    cornerRadius: headerRow.cornerRadius,
                    effects: headerRow.effects
                };
                
                // Also store component properties if available
                if ('componentProperties' in headerRow) {
                    for (const [key, prop] of Object.entries(headerRow.componentProperties)) {
                        headerRowProperties[`component_${key}`] = prop.value;
                    }
                }
            }
            
            // Store table properties for applying to generated table frame
            let tableProperties: { [key: string]: any } = {};
            if (tableFrame && 'fills' in tableFrame) {
                const tableFrameWithFills = tableFrame as FrameNode | ComponentNode | InstanceNode;
                console.log('🔍 Table properties available:', {
                    fills: tableFrameWithFills.fills,
                    strokes: tableFrameWithFills.strokes,
                    strokeWeight: tableFrameWithFills.strokeWeight,
                    cornerRadius: tableFrameWithFills.cornerRadius,
                    effects: tableFrameWithFills.effects
                });
                
                // Store the actual visual properties from the table frame
                tableProperties = {
                    fills: tableFrameWithFills.fills,
                    strokes: tableFrameWithFills.strokes,
                    strokeWeight: tableFrameWithFills.strokeWeight,
                    cornerRadius: tableFrameWithFills.cornerRadius,
                    effects: tableFrameWithFills.effects
                };
                
                // Also store component properties if available
                if ('componentProperties' in tableFrame) {
                    for (const [key, prop] of Object.entries(tableFrame.componentProperties)) {
                        tableProperties[`component_${key}`] = prop.value;
                    }
                }
            }
            
            // Store divider properties for applying to generated dividers
            let dividerProperties: { [key: string]: any } = {};
            if (dividerComponent) {
                console.log('🔍 Divider component properties available:', {
                    fills: dividerComponent.fills,
                    strokes: dividerComponent.strokes,
                    strokeWeight: dividerComponent.strokeWeight,
                    cornerRadius: dividerComponent.cornerRadius,
                    effects: dividerComponent.effects
                });
                
                // Store the actual visual properties from the divider component
                dividerProperties = {
                    fills: dividerComponent.fills,
                    strokes: dividerComponent.strokes,
                    strokeWeight: dividerComponent.strokeWeight,
                    cornerRadius: dividerComponent.cornerRadius,
                    effects: dividerComponent.effects,
                    height: dividerComponent.height // Capture the height
                };
                
                // Check for color variables in main component fills
                if (dividerComponent.fills && Array.isArray(dividerComponent.fills) && dividerComponent.fills.length > 0) {
                    for (const fill of dividerComponent.fills) {
                        if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                            console.log('🎨 Found color variable in main component:', fill.boundVariables.color);
                            dividerProperties.colorVariable = fill.boundVariables.color;
                        }
                    }
                }
                
                // Also store component properties if available
                if ('componentProperties' in dividerComponent && dividerComponent.componentProperties) {
                    for (const [key, prop] of Object.entries(dividerComponent.componentProperties as any)) {
                        if (prop && typeof prop === 'object' && 'value' in prop) {
                            dividerProperties[`component_${key}`] = (prop as any).value;
                        }
                    }
                }
                
                // Try to create a temporary instance to get the actual visual properties
                try {
                    const tempDividerInstance = dividerComponent.createInstance();
                    // Store instance properties efficiently
                    dividerProperties.instanceFills = tempDividerInstance.fills;
                    dividerProperties.instanceStrokes = tempDividerInstance.strokes;
                    dividerProperties.instanceStrokeWeight = tempDividerInstance.strokeWeight;
                    dividerProperties.instanceHeight = tempDividerInstance.height;
                    
                    // Check for color variables in instance fills
                    if (tempDividerInstance.fills && Array.isArray(tempDividerInstance.fills)) {
                        for (const fill of tempDividerInstance.fills) {
                            if (fill.type === 'SOLID' && fill.boundVariables?.color) {
                                dividerProperties.instanceColorVariable = fill.boundVariables.color;
                                break;
                            }
                        }
                    }
                    
                    // Check first child for properties
                    if ('children' in tempDividerInstance && tempDividerInstance.children.length > 0) {
                        const child = tempDividerInstance.children[0];
                        if ('fills' in child && child.fills && Array.isArray(child.fills) && child.fills.length > 0) {
                            dividerProperties.instanceChildFills = child.fills;
                            for (const fill of child.fills) {
                                if (fill.type === 'SOLID' && fill.boundVariables?.color) {
                                    dividerProperties.instanceChildColorVariable = fill.boundVariables.color;
                                    break;
                                }
                            }
                        }
                        if ('strokes' in child && child.strokes) dividerProperties.instanceChildStrokes = child.strokes;
                        if ('strokeWeight' in child) dividerProperties.instanceChildStrokeWeight = child.strokeWeight;
                        if ('height' in child) dividerProperties.instanceChildHeight = child.height;
                    }
                    
                    tempDividerInstance.remove();
                } catch (error) {
                    console.error('❌ Error creating temporary divider instance:', error);
                }
                
                // Find and store properties from divider's child elements (like rectangles, lines, etc.)
                if ('children' in dividerComponent) {
                    console.log('🔍 Divider children found:', dividerComponent.children.length);
                    for (const child of dividerComponent.children) {
                        console.log('🔍 Examining divider child:', child.name, child.type);
                        
                        // Check if child has fills
                        if ('fills' in child) {
                            console.log('🔍 Divider child with fills:', child.name, child.type, {
                                fills: child.fills,
                                strokes: 'strokes' in child ? child.strokes : undefined,
                                strokeWeight: 'strokeWeight' in child ? child.strokeWeight : undefined
                            });
                            
                            // Store the first child's properties as the main divider properties
                            if (child.fills && Array.isArray(child.fills) && child.fills.length > 0) {
                                dividerProperties.childFills = child.fills;
                                console.log('✅ Found child fills:', child.fills);
                                
                                // Check for color variables in fills
                                for (const fill of child.fills) {
                                    if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                                        console.log('🎨 Found color variable in child:', fill.boundVariables.color);
                                        // Debug: Get variable name
                                        try {
                                            const variable = figma.variables.getVariableById(fill.boundVariables.color.id);
                                            if (variable) {
                                                console.log('🎨 Child color variable name:', variable.name);
                                            }
                                        } catch (e) {
                                            console.log('⚠️ Could not get variable name for child color variable');
                                        }
                                        dividerProperties.childColorVariable = fill.boundVariables.color;
                                    }
                                }
                            }
                            if ('strokes' in child && child.strokes && Array.isArray(child.strokes) && child.strokes.length > 0) {
                                dividerProperties.childStrokes = child.strokes;
                                console.log('✅ Found child strokes:', child.strokes);
                            }
                            if ('strokeWeight' in child && child.strokeWeight !== undefined) {
                                dividerProperties.childStrokeWeight = child.strokeWeight;
                                console.log('✅ Found child stroke weight:', child.strokeWeight);
                            }
                            if ('height' in child && child.height !== undefined) {
                                dividerProperties.childHeight = child.height;
                                console.log('✅ Found child height:', child.height);
                            }
                            break; // Use the first child's properties
                        }
                        
                        // Also check if child has children (nested structure)
                        if ('children' in child && child.children.length > 0) {
                            console.log('🔍 Divider child has nested children:', child.children.length);
                            for (const grandChild of child.children) {
                                console.log('🔍 Examining grandchild:', grandChild.name, grandChild.type);
                                if ('fills' in grandChild) {
                                    console.log('🔍 Grandchild with fills:', grandChild.name, grandChild.type, {
                                        fills: grandChild.fills,
                                        strokes: 'strokes' in grandChild ? grandChild.strokes : undefined,
                                        strokeWeight: 'strokeWeight' in grandChild ? grandChild.strokeWeight : undefined
                                    });
                                    
                                    if (grandChild.fills && Array.isArray(grandChild.fills) && grandChild.fills.length > 0) {
                                        dividerProperties.childFills = grandChild.fills;
                                        console.log('✅ Found grandchild fills:', grandChild.fills);
                                    }
                                    if ('strokes' in grandChild && grandChild.strokes && Array.isArray(grandChild.strokes) && grandChild.strokes.length > 0) {
                                        dividerProperties.childStrokes = grandChild.strokes;
                                        console.log('✅ Found grandchild strokes:', grandChild.strokes);
                                    }
                                    if ('strokeWeight' in grandChild && grandChild.strokeWeight !== undefined) {
                                        dividerProperties.childStrokeWeight = grandChild.strokeWeight;
                                        console.log('✅ Found grandchild stroke weight:', grandChild.strokeWeight);
                                    }
                                    break; // Use the first grandchild's properties
                                }
                            }
                            break; // Use the first child with children
                        }
                    }
                }
            }
            
            lastScanResult = {
                headerCell: headerCell ? headerCell.mainComponent : null,
                headerRowComponent,
                bodyCell: bodyCell ? bodyCell.mainComponent : null,
                bodyRowComponent,
                footer: footer ? footer.mainComponent : null,
                numCols,
                selectCellComponent,
                expandCellComponent,
                dividerComponent
            };
            
            // Store the found divider rectangle in lastScanResult for direct use
            if (foundDividerRect) {
                (lastScanResult as any).originalDividerRect = foundDividerRect;
                console.log('✅ Stored original divider rectangle in lastScanResult for direct use');
            } else {
                console.log('⚠️ No divider rectangle found in specific path, will use component fallback');
            }
            
            // Store header row properties in lastScanResult for later use
            (lastScanResult as any).headerRowProperties = headerRowProperties;
            // Store table properties in lastScanResult for later use
            (lastScanResult as any).tableProperties = tableProperties;
            // Store divider properties in lastScanResult for later use
            (lastScanResult as any).dividerProperties = dividerProperties;
            
            // Store scan result in table frame's plugin data for persistence
            try {
                if (tableFrame && 'setPluginData' in tableFrame) {
                    const scanData = {
                        headerCellId: headerCell ? headerCell.id : null,
                        bodyCellId: bodyCell ? bodyCell.id : null,
                        footerId: footer ? footer.id : null,
                        selectCellId: selectCellComponent?.id || null,
                        expandCellId: expandCellComponent?.id || null,
                        dividerId: dividerComponent?.id || null,
                        numCols,
                        timestamp: Date.now()
                    };
                    tableFrame.setPluginData('tableGeneratorScan', JSON.stringify(scanData));
                    console.log('✅ Stored scan result in table frame plugin data');
                }
            } catch (error) {
                console.log('⚠️ Could not store scan result in plugin data:', error);
            }
       
            // Send summary to UI with updated body row component
            figma.ui.postMessage({
                type: 'scan-table-result',
                success: true,
                message: `Tap on the cells below to personalize your table.`,
                details: {
                    footer: !!footer,
                    numCols,
                    bodyCellComponent,
                    headerCellComponent,
                    footerComponent,
                    expandCellComponent,
                    selectCellComponent,
                    bodyRowComponent: bodyRowComponent
                }
            });
        }
    
    } else if (msg.type === 'create-table-from-ai') {
        // Build a table directly from provided headers + rows
        try {
            const headers: string[] = (msg.headers || []).map(String);
            const rowsData: string[][] = Array.isArray(msg.rows) ? msg.rows.map((r: any) => Array.isArray(r) ? r.map(String) : []) : [];
            const includeHeader = msg.includeHeader !== false;
            const includeFooter = msg.includeFooter === true;
            const includeSelectable = msg.includeSelectable === true;
            const includeExpandable = msg.includeExpandable === true;
            const rows = rowsData.length;
            const cols = headers.length || (rowsData[0]?.length || 0);

            // Reuse lastScanResult to create structure
            if (!lastScanResult) {
                figma.notify('❌ No scanned template available. Scan a table first.');
                return;
            }

            // Prepare cellProps mapping for create-table-from-scan
            const cellProps: { [key: string]: any } = {};
            // Header properties
            for (let c = 1; c <= cols; c++) {
                const key = `header-${c}`;
                cellProps[key] = { properties: { 'Cell text#': headers[c - 1] || `Column ${c}` } };
            }
            // Body properties
            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const key = `${r}-${c}`;
                    cellProps[key] = { properties: { 'Cell text#': rowsData[r][c] || '' } };
                }
            }

            // Delegate to existing flow
            figma.ui.postMessage({ type: 'ai-table-mapped' });
            figma.ui.postMessage({
                type: 'create-table-from-scan',
                includeHeader,
                includeFooter,
                includeSelectable,
                includeExpandable,
                rows,
                cols,
                cellProps
            });
        } catch (e) {
            console.error('create-table-from-ai error', e);
            figma.notify('Failed to build table from AI data');
        }
        return;

    } else if (msg.type === 'create-table-from-scan') {
        if (isCreatingTable) {
            figma.notify("Already creating a table. Please wait.");
            return;
        }
        isCreatingTable = true;
        try {
            if (!lastScanResult) {
                figma.notify('❌ No scan result available. Please scan a table first.');
                return;
            }

            const { headerCell, bodyCell, footer, numCols } = lastScanResult;
            figma.ui.postMessage({ type: 'show-loader', message: 'Generating table...' });

            const includeHeader = msg.includeHeader !== false;
            const includeFooter = msg.includeFooter === true;
            const includeSelectable = msg.includeSelectable === true;
            const includeExpandable = msg.includeExpandable === true;
            const rows = msg.rows || 3;
            const cols = msg.cols || numCols || 3;
            let cellProps = msg.cellProps || {};

            // Apply sorting if enabled
            cellProps = sortColumnData(cellProps, cols, rows);

            if (includeSelectable && !lastScanResult.selectCellComponent) {
                figma.notify('⚠️ Selectable cells not available');
            }
            if (includeExpandable && !lastScanResult.expandCellComponent) {
                figma.notify('⚠️ Expandable cells not available');
            }

            const columnWidths: number[] = [];
            let totalTableWidth = 0;

            if (includeExpandable && lastScanResult.expandCellComponent) {
                try {
                    const expandCellInstance = lastScanResult.expandCellComponent.createInstance();
                    totalTableWidth += expandCellInstance.width;
                    expandCellInstance.remove();
                } catch (error) {
                    totalTableWidth += 52;
                }
            }
            if (includeSelectable && lastScanResult.selectCellComponent) {
                try {
                    const selectCellInstance = lastScanResult.selectCellComponent.createInstance();
                    totalTableWidth += selectCellInstance.width;
                    selectCellInstance.remove();
                } catch (error) {
                    totalTableWidth += 52;
                }
            }

            for (let c = 0; c < cols; c++) {
                let widthFound = false;
                for (let r = 0; r < rows; r++) {
                    const key = `${r}-${c}`;
                    const cellData = cellProps[key];
                    if (cellData && cellData.colWidth) {
                        columnWidths[c] = cellData.colWidth;
                        widthFound = true;
                        break;
                    }
                }
                if (!widthFound) {
                    columnWidths[c] = 120;
                }
                totalTableWidth += columnWidths[c];
            }

            const tableFrame = figma.createFrame();
        tableFrame.name = "Generated Table";
        tableFrame.layoutMode = "VERTICAL";
        tableFrame.counterAxisSizingMode = "AUTO";
        tableFrame.primaryAxisSizingMode = "AUTO";
        tableFrame.itemSpacing = 0;
            tableFrame.paddingLeft = 0;
            tableFrame.paddingRight = 0;
            tableFrame.paddingTop = 0;
            tableFrame.paddingBottom = 0;
        // Clear default fill to make it theme-compatible
        tableFrame.fills = [];
        
        // Apply table properties from scanned table to maintain color variables
        if (lastScanResult && (lastScanResult as any).tableProperties) {
            const tableProps = (lastScanResult as any).tableProperties;
            console.log('🔄 Applying table properties to generated table frame:', tableProps);
            
            try {
                // Apply fills if present in the original table
                if (tableProps.fills && Array.isArray(tableProps.fills) && tableProps.fills.length > 0) {
                    console.log('🎨 Applying table fills:', tableProps.fills);
                    tableFrame.fills = tableProps.fills;
                }
                
                // Apply strokes if present in the original table
                if (tableProps.strokes && Array.isArray(tableProps.strokes) && tableProps.strokes.length > 0) {
                    console.log('🎨 Applying table strokes:', tableProps.strokes);
                    tableFrame.strokes = tableProps.strokes;
                }
                
                // Apply stroke weight if present
                if (tableProps.strokeWeight !== undefined && tableProps.strokeWeight > 0) {
                    console.log('🎨 Applying table stroke weight:', tableProps.strokeWeight);
                    tableFrame.strokeWeight = tableProps.strokeWeight;
                }
                
                // Apply corner radius if present
                if (tableProps.cornerRadius !== undefined && tableProps.cornerRadius > 0) {
                    console.log('🎨 Applying table corner radius:', tableProps.cornerRadius);
                    tableFrame.cornerRadius = tableProps.cornerRadius;
                }
                
                // Apply effects if present
                if (tableProps.effects && Array.isArray(tableProps.effects) && tableProps.effects.length > 0) {
                    console.log('🎨 Applying table effects:', tableProps.effects);
                    tableFrame.effects = tableProps.effects;
                }
                
                console.log('✅ Table properties applied successfully');
            } catch (error) {
                console.error('❌ Error applying table properties:', error);
            }
        } else {
            console.log('⚠️ No table properties available to apply');
        }

            if (includeHeader && headerCell) {
                try {
                    const headerRowFrame = figma.createFrame();
            headerRowFrame.name = "Header Row";
            headerRowFrame.layoutMode = "HORIZONTAL";
            headerRowFrame.primaryAxisSizingMode = "AUTO";
            headerRowFrame.counterAxisSizingMode = "AUTO";
            headerRowFrame.itemSpacing = 0;
                    headerRowFrame.paddingLeft = 0;
                    headerRowFrame.paddingRight = 0;
                    headerRowFrame.paddingTop = 0;
                    headerRowFrame.paddingBottom = 0;
            // Clear default fill to make it theme-compatible
            headerRowFrame.fills = [];
            
            // Apply header row properties from scanned table to maintain color variables
            if (lastScanResult && (lastScanResult as any).headerRowProperties) {
                const headerRowProps = (lastScanResult as any).headerRowProperties;
                console.log('🔄 Applying header row properties to generated header row:', headerRowProps);
                
                try {
                    // Apply fills if present in the original header row
                    if (headerRowProps.fills && Array.isArray(headerRowProps.fills) && headerRowProps.fills.length > 0) {
                        console.log('🎨 Applying fills:', headerRowProps.fills);
                        headerRowFrame.fills = headerRowProps.fills;
                    }
                    
                    // Apply strokes if present in the original header row
                    if (headerRowProps.strokes && Array.isArray(headerRowProps.strokes) && headerRowProps.strokes.length > 0) {
                        console.log('🎨 Applying strokes:', headerRowProps.strokes);
                        headerRowFrame.strokes = headerRowProps.strokes;
                    }
                    
                    // Apply stroke weight if present
                    if (headerRowProps.strokeWeight !== undefined && headerRowProps.strokeWeight > 0) {
                        console.log('🎨 Applying stroke weight:', headerRowProps.strokeWeight);
                        headerRowFrame.strokeWeight = headerRowProps.strokeWeight;
                    }
                    
                    // Apply corner radius if present
                    if (headerRowProps.cornerRadius !== undefined && headerRowProps.cornerRadius > 0) {
                        console.log('🎨 Applying corner radius:', headerRowProps.cornerRadius);
                        headerRowFrame.cornerRadius = headerRowProps.cornerRadius;
                    }
                    
                    // Apply effects if present
                    if (headerRowProps.effects && Array.isArray(headerRowProps.effects) && headerRowProps.effects.length > 0) {
                        console.log('🎨 Applying effects:', headerRowProps.effects);
                        headerRowFrame.effects = headerRowProps.effects;
                    }
                    
                    console.log('✅ Header row properties applied successfully');
                } catch (error) {
                    console.error('❌ Error applying header row properties:', error);
                }
            } else {
                console.log('⚠️ No header row properties available to apply');
            }
                
                    if (includeExpandable && lastScanResult.expandCellComponent) {
                        try {
                            const expandCell = lastScanResult.expandCellComponent.createInstance();
                            headerRowFrame.appendChild(expandCell);
                        } catch (error) { console.error('Error creating header expand cell'); }
                    }

                    if (includeSelectable && lastScanResult.selectCellComponent) {
                        try {
                            const selectCell = lastScanResult.selectCellComponent.createInstance();
                            headerRowFrame.appendChild(selectCell);
                        } catch (error) { console.error('Error creating header select cell'); }
            }

            for (let c = 1; c <= cols; c++) {
                        const hCell = headerCell.createInstance();
                hCell.layoutSizingHorizontal = 'FIXED';
                hCell.resize(columnWidths[c - 1], hCell.height);
                headerRowFrame.appendChild(hCell);

                const key = `header-${c}`;
                const cellData = cellProps[key];
                if (cellData && cellData.properties) {
                    try {
                        const validProps = mapPropertyNames(cellData.properties, hCell.componentProperties);
                        
                        // Try to set all properties at once first
                        try {
                            hCell.setProperties(validProps);
                        } catch (variantError) {
                            console.warn('Variant combination failed for header cell', key, 'trying individual properties');
                            // Fallback: set properties one by one
                            for (const [propName, propValue] of Object.entries(validProps)) {
                                try {
                                    hCell.setProperties({ [propName]: propValue });
                                } catch (individualError) {
                                    console.warn(`Could not set individual property ${propName} for header cell ${key}:`, individualError);
                                }
                            }
                        }
                    } catch (e) { console.warn('Could not set properties for header cell', key, e); }
                }
            }
            tableFrame.appendChild(headerRowFrame);
                } catch (error) {
                    console.error('Error creating header row:', error);
                    figma.notify('⚠️ Header cell component is no longer available');
                }
            }

            if (bodyCell) {
                try {
                    const bodyWrapperFrame = figma.createFrame();
            bodyWrapperFrame.name = 'Body';
            bodyWrapperFrame.layoutMode = 'VERTICAL';
            bodyWrapperFrame.primaryAxisSizingMode = 'AUTO';
            bodyWrapperFrame.counterAxisSizingMode = 'AUTO';
            bodyWrapperFrame.itemSpacing = 0;
                    bodyWrapperFrame.paddingLeft = 0;
                    bodyWrapperFrame.paddingRight = 0;
                    bodyWrapperFrame.paddingTop = 0;
                    bodyWrapperFrame.paddingBottom = 0;
            // Clear default fill to make it theme-compatible
            bodyWrapperFrame.fills = [];
            tableFrame.appendChild(bodyWrapperFrame);

            for (let r = 0; r < rows; r++) {
                // Create a body row item frame (similar to "Data table body row item")
                const bodyRowItemFrame = figma.createFrame();
                bodyRowItemFrame.name = `Body Row Item ${r + 1}`;
                bodyRowItemFrame.layoutMode = "VERTICAL";
                bodyRowItemFrame.primaryAxisSizingMode = "AUTO";
                bodyRowItemFrame.counterAxisSizingMode = "AUTO";
                bodyRowItemFrame.itemSpacing = 0;
                bodyRowItemFrame.paddingLeft = 0;
                bodyRowItemFrame.paddingRight = 0;
                bodyRowItemFrame.paddingTop = 0;
                bodyRowItemFrame.paddingBottom = 0;
                // Clear default fill to make it theme-compatible
                bodyRowItemFrame.fills = [];
                
                // Create the data row frame (similar to "Data table row")
                const rowFrame = figma.createFrame();
                rowFrame.name = `Row ${r + 1}`;
                rowFrame.layoutMode = "HORIZONTAL";
                rowFrame.primaryAxisSizingMode = "AUTO";
                rowFrame.counterAxisSizingMode = "AUTO";
                rowFrame.itemSpacing = 0;
                        rowFrame.paddingLeft = 0;
                        rowFrame.paddingRight = 0;
                        rowFrame.paddingTop = 0;
                        rowFrame.paddingBottom = 0;
                // Clear default fill to make it theme-compatible
                rowFrame.fills = [];

                        if (includeExpandable && lastScanResult.expandCellComponent) {
                            try {
                                const expandCell = lastScanResult.expandCellComponent.createInstance();
                                rowFrame.appendChild(expandCell);
                            } catch (error) { console.error('Error creating expand cell instance'); }
                        }

                        if (includeSelectable && lastScanResult.selectCellComponent) {
                            try {
                    const selectCell = lastScanResult.selectCellComponent.createInstance();
                    rowFrame.appendChild(selectCell);
                            } catch (error) { console.error('Error creating select cell instance'); }
                        }

                        for (let c = 0; c < cols; c++) {
                            const cell = bodyCell.createInstance();
                            cell.layoutSizingHorizontal = 'FIXED';
                            cell.resize(columnWidths[c], cell.height);
                            rowFrame.appendChild(cell);

                            const key = `${r}-${c}`;
                            const cellData = cellProps[key];
                            if (cellData && cellData.properties) {
                                try {
                                    const validProps = mapPropertyNames(cellData.properties, cell.componentProperties);
                                    cell.setProperties(validProps);
                                } catch (e) { console.warn('Could not set properties for cell', key, e); }
                            }
                        }
                        
                        // Add the data row to the body row item frame
                        bodyRowItemFrame.appendChild(rowFrame);
                        
                        // Add divider after the data row (including the last row)
                        {
                            // Use original divider rectangle if available, otherwise fallback to component
                            if ((lastScanResult as any).originalDividerRect) {
                                try {
                                    const divider = ((lastScanResult as any).originalDividerRect as RectangleNode).clone();
                                    divider.resize(totalTableWidth, divider.height);
                                    bodyRowItemFrame.appendChild(divider);
                                    console.log('✅ Added cloned divider rectangle with theme-aware colors');
                                } catch (error) {
                                    console.error('❌ Error cloning divider rectangle:', error);
                                }
                            } else if (lastScanResult.dividerComponent) {
                                try {
                                    const divider = lastScanResult.dividerComponent.createInstance();
                                    
                                    // Apply divider properties from scanned table
                                    if (lastScanResult && (lastScanResult as any).dividerProperties) {
                                        const dividerProps = (lastScanResult as any).dividerProperties;
                                        console.log('🔄 Applying divider properties:', dividerProps);
                                        
                                        try {
                                                                                // Priority order: 1) Scanned divider color variable, 2) Hardcoded border-subtle-01, 3) Scanned divider fills, 4) Fallback
                                    let appliedColor = false;
                                    
                                    // First priority: Apply color variable from scanned divider
                                    const scannedColorVariable = dividerProps.colorVariable || dividerProps.instanceColorVariable || dividerProps.instanceChildColorVariable || dividerProps.childColorVariable;
                                    if (scannedColorVariable) {
                                        console.log('🎨 Applying scanned divider color variable:', scannedColorVariable);
                                        try {
                                            const colorVariableFill: Paint = {
                                                type: 'SOLID',
                                                color: { r: 0, g: 0, b: 0 }, // Default color (will be overridden by variable)
                                                boundVariables: {
                                                    color: scannedColorVariable
                                                }
                                            };
                                            divider.fills = [colorVariableFill];
                                            appliedColor = true;
                                            console.log('✅ Applied scanned divider color variable');
                                        } catch (error) {
                                            console.error('❌ Error applying scanned divider color variable:', error);
                                        }
                                    }
                                    
                                    // Second priority: Use hardcoded border-subtle-01 color variable
                                    if (!appliedColor) {
                                        console.log('🎨 Applying hardcoded border-subtle-01 color variable');
                                        try {
                                            // Find the border-subtle-01 variable
                                            const localVariables = figma.variables.getLocalVariables();
                                            const borderVariable = localVariables.find(v => v.name === 'border-subtle-01');
                                            
                                            if (borderVariable) {
                                                const colorVariableFill: Paint = {
                                                    type: 'SOLID',
                                                    color: { r: 0, g: 0, b: 0 }, // Default color (will be overridden by variable)
                                                    boundVariables: {
                                                        color: {
                                                            id: borderVariable.id,
                                                            type: 'VARIABLE_ALIAS'
                                                        }
                                                    }
                                                };
                                                divider.fills = [colorVariableFill];
                                                appliedColor = true;
                                                console.log('✅ Applied hardcoded border-subtle-01 color variable to divider');
                                            } else {
                                                console.log('⚠️ border-subtle-01 variable not found in local variables');
                                            }
                                        } catch (error) {
                                            console.error('❌ Error applying hardcoded border-subtle-01 color variable:', error);
                                        }
                                    }
                                    
                                    // Second priority: Apply fills from scanned divider child
                                    if (!appliedColor && dividerProps.childFills && Array.isArray(dividerProps.childFills) && dividerProps.childFills.length > 0) {
                                        console.log('🎨 Applying scanned divider child fills:', dividerProps.childFills);
                                        divider.fills = dividerProps.childFills;
                                        appliedColor = true;
                                        console.log('✅ Applied scanned divider child fills');
                                    }
                                    
                                    // Third priority: Apply fills from scanned divider
                                    if (!appliedColor && dividerProps.fills && Array.isArray(dividerProps.fills) && dividerProps.fills.length > 0) {
                                        console.log('🎨 Applying scanned divider fills:', dividerProps.fills);
                                        divider.fills = dividerProps.fills;
                                        appliedColor = true;
                                    }
                                    
                                    // Last resort: Use scanned divider color or fallback gray
                                    if (!appliedColor) {
                                        if (dividerProps.childFills && Array.isArray(dividerProps.childFills) && dividerProps.childFills.length > 0) {
                                            console.log('🎨 Using scanned divider child color:', dividerProps.childFills[0]);
                                            divider.fills = dividerProps.childFills;
                                            appliedColor = true;
                                        } else {
                                            console.log('⚠️ No color variable found, applying fallback gray color');
                                            divider.fills = [{
                                                type: 'SOLID',
                                                color: { r: 0.9, g: 0.9, b: 0.9 } // Light gray
                                            }];
                                        }
                                        console.log('✅ Applied divider color');
                                    }
                                            
                                            // Apply strokes if present in the original divider
                                            if (dividerProps.strokes && Array.isArray(dividerProps.strokes) && dividerProps.strokes.length > 0) {
                                                console.log('🎨 Applying divider strokes:', dividerProps.strokes);
                                                divider.strokes = dividerProps.strokes;
                                            }
                                            
                                            // Apply stroke weight if present
                                            if (dividerProps.strokeWeight !== undefined && dividerProps.strokeWeight > 0) {
                                                console.log('🎨 Applying divider stroke weight:', dividerProps.strokeWeight);
                                                divider.strokeWeight = dividerProps.strokeWeight;
                                            }
                                            
                                            // Apply corner radius if present
                                            if (dividerProps.cornerRadius !== undefined && dividerProps.cornerRadius > 0) {
                                                console.log('🎨 Applying divider corner radius:', dividerProps.cornerRadius);
                                                divider.cornerRadius = dividerProps.cornerRadius;
                                            }
                                            
                                            // Apply effects if present
                                            if (dividerProps.effects && Array.isArray(dividerProps.effects) && dividerProps.effects.length > 0) {
                                                console.log('🎨 Applying divider effects:', dividerProps.effects);
                                                divider.effects = dividerProps.effects;
                                            }
                                            
                                            // Apply height if present
                                            if (dividerProps.height !== undefined && dividerProps.height > 0) {
                                                console.log('📏 Applying divider height:', dividerProps.height);
                                                divider.resize(divider.width, dividerProps.height);
                                            }
                                            
                                            // Apply instance properties first (these are from the actual instance)
                                            if (dividerProps.instanceFills || dividerProps.instanceStrokes || dividerProps.instanceStrokeWeight) {
                                        console.log('🎨 Applying instance properties to divider');
                                        if (dividerProps.instanceFills && Array.isArray(dividerProps.instanceFills) && dividerProps.instanceFills.length > 0) {
                                            console.log('🎨 Applying instance fills to divider');
                                            divider.fills = dividerProps.instanceFills;
                                        } else {
                                            // If instance fills are empty, try to apply border-subtle-01 color variable
                                            console.log('🎨 Instance fills are empty, checking for border color variable');
                                            const borderVariable = findBorderColorVariable();
                                            if (borderVariable) {
                                                try {
                                                    const colorVariableFill: Paint = {
                                                        type: 'SOLID',
                                                        color: { r: 0, g: 0, b: 0 },
                                                        boundVariables: {
                                                            color: borderVariable
                                                        }
                                                    };
                                                    divider.fills = [colorVariableFill];
                                                    console.log('✅ Applied border-subtle-01 color variable to divider instance');
                                                } catch (error) {
                                                    console.error('❌ Error applying border color variable to divider instance:', error);
                                                }
                                            }
                                        }
                                        if (dividerProps.instanceColorVariable) {
                                            console.log('🎨 Applying instance color variable to divider:', dividerProps.instanceColorVariable);
                                            try {
                                                const colorVariableFill: Paint = {
                                                    type: 'SOLID',
                                                    color: { r: 0, g: 0, b: 0 },
                                                    boundVariables: {
                                                        color: dividerProps.instanceColorVariable
                                                    }
                                                };
                                                divider.fills = [colorVariableFill];
                                            } catch (error) {
                                                console.error('❌ Error applying instance color variable to divider:', error);
                                            }
                                        }
                                        if (dividerProps.instanceStrokes && Array.isArray(dividerProps.instanceStrokes) && dividerProps.instanceStrokes.length > 0) {
                                            console.log('🎨 Applying instance strokes to divider');
                                            divider.strokes = dividerProps.instanceStrokes;
                                        }
                                        if (dividerProps.instanceStrokeWeight !== undefined) {
                                            console.log('🎨 Applying instance stroke weight to divider');
                                            divider.strokeWeight = dividerProps.instanceStrokeWeight;
                                        }
                                        if (dividerProps.instanceHeight !== undefined) {
                                            console.log('📏 Applying instance height to divider');
                                            divider.resize(divider.width, dividerProps.instanceHeight);
                                        }
                                    }
                                    
                                    // Apply instance child properties to divider's children
                                    if (dividerProps.instanceChildFills || dividerProps.instanceChildStrokes || dividerProps.instanceChildStrokeWeight) {
                                        console.log('🎨 Applying instance child properties to divider children');
                                        if ('children' in divider) {
                                            console.log('🎨 Divider has', divider.children.length, 'children');
                                            for (const child of divider.children) {
                                                console.log('🎨 Processing divider child:', child.name, child.type);
                                                
                                                if ('fills' in child && dividerProps.instanceChildFills && Array.isArray(dividerProps.instanceChildFills) && dividerProps.instanceChildFills.length > 0) {
                                                    console.log('🎨 Applying instance child fills to:', child.name);
                                                    child.fills = dividerProps.instanceChildFills;
                                                } else if ('fills' in child) {
                                                    // If instance child fills are empty, try to apply border-subtle-01 color variable
                                                    console.log('🎨 Instance child fills are empty, checking for border color variable');
                                                    const borderVariable = findBorderColorVariable();
                                                    if (borderVariable) {
                                                        try {
                                                            const colorVariableFill: Paint = {
                                                                type: 'SOLID',
                                                                color: { r: 0, g: 0, b: 0 },
                                                                boundVariables: {
                                                                    color: borderVariable
                                                                }
                                                            };
                                                            child.fills = [colorVariableFill];
                                                            console.log('✅ Applied border-subtle-01 color variable to divider child');
                                                        } catch (error) {
                                                            console.error('❌ Error applying border color variable to divider child:', error);
                                                        }
                                                    }
                                                }
                                                if ('fills' in child && dividerProps.instanceChildColorVariable) {
                                                    console.log('🎨 Applying instance child color variable to:', child.name, dividerProps.instanceChildColorVariable);
                                                    try {
                                                        const colorVariableFill: Paint = {
                                                            type: 'SOLID',
                                                            color: { r: 0, g: 0, b: 0 },
                                                            boundVariables: {
                                                                color: dividerProps.instanceChildColorVariable
                                                            }
                                                        };
                                                        child.fills = [colorVariableFill];
                                                    } catch (error) {
                                                        console.error('❌ Error applying instance child color variable:', error);
                                                    }
                                                }
                                                if ('strokes' in child && dividerProps.instanceChildStrokes) {
                                                    console.log('🎨 Applying instance child strokes to:', child.name);
                                                    child.strokes = dividerProps.instanceChildStrokes;
                                                }
                                                if ('strokeWeight' in child && dividerProps.instanceChildStrokeWeight !== undefined) {
                                                    console.log('🎨 Applying instance child stroke weight to:', child.name);
                                                    child.strokeWeight = dividerProps.instanceChildStrokeWeight;
                                                }
                                                if ('height' in child && dividerProps.instanceChildHeight !== undefined && 'resize' in child) {
                                                    console.log('📏 Applying instance child height to:', child.name);
                                                    (child as any).resize(child.width, dividerProps.instanceChildHeight);
                                                }
                                                break; // Apply to first child only
                                            }
                                        }
                                    }
                                    
                                    // Apply child properties to divider's children
                                    if (dividerProps.childFills || dividerProps.childStrokes || dividerProps.childStrokeWeight) {
                                        console.log('🎨 Applying child properties to divider children');
                                        if ('children' in divider) {
                                            console.log('🎨 Divider has', divider.children.length, 'children');
                                            for (const child of divider.children) {
                                                console.log('🎨 Processing divider child:', child.name, child.type);
                                                
                                                // Apply to direct children
                                                if ('fills' in child && dividerProps.childFills && Array.isArray(dividerProps.childFills) && dividerProps.childFills.length > 0) {
                                                    console.log('🎨 Applying child fills to:', child.name);
                                                    child.fills = dividerProps.childFills;
                                                } else if ('fills' in child) {
                                                    // If child fills are empty, try to apply border-subtle-01 color variable
                                                    console.log('🎨 Child fills are empty, checking for border color variable');
                                                    const borderVariable = findBorderColorVariable();
                                                    if (borderVariable) {
                                                        try {
                                                            const colorVariableFill: Paint = {
                                                                type: 'SOLID',
                                                                color: { r: 0, g: 0, b: 0 },
                                                                boundVariables: {
                                                                    color: borderVariable
                                                                }
                                                            };
                                                            child.fills = [colorVariableFill];
                                                            console.log('✅ Applied border-subtle-01 color variable to child');
                                                        } catch (error) {
                                                            console.error('❌ Error applying border color variable to child:', error);
                                                        }
                                                    }
                                                }
                                                if ('fills' in child && dividerProps.childColorVariable) {
                                                    console.log('🎨 Applying child color variable to:', child.name, dividerProps.childColorVariable);
                                                    try {
                                                        const colorVariableFill: Paint = {
                                                            type: 'SOLID',
                                                            color: { r: 0, g: 0, b: 0 },
                                                            boundVariables: {
                                                                color: dividerProps.childColorVariable
                                                            }
                                                        };
                                                        child.fills = [colorVariableFill];
                                                    } catch (error) {
                                                        console.error('❌ Error applying child color variable:', error);
                                                    }
                                                }
                                                if ('strokes' in child && dividerProps.childStrokes) {
                                                    console.log('🎨 Applying child strokes to:', child.name);
                                                    child.strokes = dividerProps.childStrokes;
                                                }
                                                if ('strokeWeight' in child && dividerProps.childStrokeWeight !== undefined) {
                                                    console.log('🎨 Applying child stroke weight to:', child.name);
                                                    child.strokeWeight = dividerProps.childStrokeWeight;
                                                }
                                                if ('height' in child && dividerProps.childHeight !== undefined && 'resize' in child) {
                                                    console.log('📏 Applying child height to:', child.name);
                                                    (child as any).resize(child.width, dividerProps.childHeight);
                                                }
                                                
                                                // Also check for nested children
                                                if ('children' in child && child.children.length > 0) {
                                                    console.log('🎨 Processing nested children in:', child.name);
                                                    for (const grandChild of child.children) {
                                                        console.log('🎨 Processing grandchild:', grandChild.name, grandChild.type);
                                                        if ('fills' in grandChild && dividerProps.childFills && Array.isArray(dividerProps.childFills) && dividerProps.childFills.length > 0) {
                                                            console.log('🎨 Applying child fills to grandchild:', grandChild.name);
                                                            grandChild.fills = dividerProps.childFills;
                                                        } else if ('fills' in grandChild) {
                                                            // If grandchild fills are empty, try to apply border-subtle-01 color variable
                                                            console.log('🎨 Grandchild fills are empty, checking for border color variable');
                                                            const borderVariable = findBorderColorVariable();
                                                            if (borderVariable) {
                                                                try {
                                                                    const colorVariableFill: Paint = {
                                                                        type: 'SOLID',
                                                                        color: { r: 0, g: 0, b: 0 },
                                                                        boundVariables: {
                                                                            color: borderVariable
                                                                        }
                                                                    };
                                                                    grandChild.fills = [colorVariableFill];
                                                                    console.log('✅ Applied border-subtle-01 color variable to grandchild');
                                                                } catch (error) {
                                                                    console.error('❌ Error applying border color variable to grandchild:', error);
                                                                }
                                                            }
                                                        }
                                                        if ('strokes' in grandChild && dividerProps.childStrokes) {
                                                            console.log('🎨 Applying child strokes to grandchild:', grandChild.name);
                                                            grandChild.strokes = dividerProps.childStrokes;
                                                        }
                                                        if ('strokeWeight' in grandChild && dividerProps.childStrokeWeight !== undefined) {
                                                            console.log('🎨 Applying child stroke weight to grandchild:', grandChild.name);
                                                            grandChild.strokeWeight = dividerProps.childStrokeWeight;
                                                        }
                                                        break; // Apply to first grandchild only
                                                    }
                                                }
                                                break; // Apply to first child only
                                            }
                                        }
                                    }
                                            
                                            console.log('✅ Divider properties applied successfully');
                                        } catch (error) {
                                            console.error('❌ Error applying divider properties:', error);
                                        }
                                    }
                                    
                                    // Set divider width to match the table width
                                    divider.resize(totalTableWidth, divider.height);
                                    console.log('📏 Set divider width to:', totalTableWidth);
                                    
                                    bodyRowItemFrame.appendChild(divider);
                                    console.log('✅ Added divider after row', r + 1);
                                } catch (error) { 
                                    console.error('Error creating divider instance:', error);
                                }
                            } else {
                                console.log('⚠️ No divider component available for row', r + 1, '(update)');
                            }
                        }
                        
                        bodyWrapperFrame.appendChild(bodyRowItemFrame);
                    }
                } catch (error) {
                    console.error('Error creating body rows:', error);
                    figma.notify('❌ Cannot generate table: body cell component is no longer available');
                    return;
                }
            }

            if (includeFooter && footer) {
                try {
                    const footerClone = footer.createInstance();
                    const footerCellData = cellProps['footer'];
                    if (footerCellData && footerCellData.properties) {
                        const validProps = mapPropertyNames(footerCellData.properties, footerClone.componentProperties);
                        try {
                            footerClone.setProperties(validProps);
                        } catch (error) { console.error('Error in footer setProperties:', error); }
                    }
                    if (footerClone) {
                        footerClone.layoutSizingHorizontal = 'FIXED';
                        footerClone.resize(totalTableWidth, footerClone.height);
                        tableFrame.appendChild(footerClone);
                    }
                } catch (error) {
                    console.error('Error creating footer:', error);
                    figma.notify('⚠️ Footer component is no longer available');
                }
            }

            tableFrame.resize(totalTableWidth, tableFrame.height);
            
            tableFrame.setPluginData('isGeneratedTable', 'true');
            const tableSettings = {
                columns: msg.cols,
                rows: msg.rows,
                includeHeader: msg.includeHeader,
                includeFooter: msg.includeFooter,
                includeSelectable: msg.includeSelectable,
                includeExpandable: msg.includeExpandable,
                cellProperties: msg.cellProps,
            };
            console.log('[Backend] Saving table settings:', tableSettings);
            console.log('[Backend] cellProperties keys:', Object.keys(msg.cellProps || {}));
            tableFrame.setPluginData('tableSettings', JSON.stringify(tableSettings));
            
            // Store the scan result in the table frame for future updates
            if (lastScanResult) {
                try {
                    const scanData = {
                        headerCellId: lastScanResult.headerCell?.id || null,
                        bodyCellId: lastScanResult.bodyCell?.id || null,
                        footerId: lastScanResult.footer?.id || null,
                        selectCellId: lastScanResult.selectCellComponent?.id || null,
                        expandCellId: lastScanResult.expandCellComponent?.id || null,
                        dividerId: lastScanResult.dividerComponent?.id || null,
                        numCols: lastScanResult.numCols,
                        timestamp: Date.now(),
                        // Store additional metadata for better restoration
                        headerRowComponentId: lastScanResult.headerRowComponent?.id || null,
                        bodyRowComponentId: lastScanResult.bodyRowComponent?.id || null,
                        // Store header row properties for fill variables
                        headerRowProperties: (lastScanResult as any).headerRowProperties || null,
                        // Store divider properties for fill variables
                        dividerProperties: (lastScanResult as any).dividerProperties || null,
                        // Store original divider rectangle ID for direct cloning
                        originalDividerRectId: (lastScanResult as any).originalDividerRect?.id || null,
                        // Store component names for debugging
                        componentNames: {
                            headerCell: lastScanResult.headerCell?.name || null,
                            bodyCell: lastScanResult.bodyCell?.name || null,
                            footer: lastScanResult.footer?.name || null,
                            selectCell: lastScanResult.selectCellComponent?.name || null,
                            expandCell: lastScanResult.expandCellComponent?.name || null,
                            divider: lastScanResult.dividerComponent?.name || null
                        }
                    };
                    tableFrame.setPluginData('tableGeneratorScan', JSON.stringify(scanData));
                    console.log('✅ Stored scan result in generated table for future updates');
                    console.log('Stored scan data:', scanData);
                } catch (error) {
                    console.log('⚠️ Could not store scan result in generated table:', error);
                }
            }

            // Convert the table frame to a component for reusability
            try {
                const tableComponent = figma.createComponent();
                tableComponent.name = `Generated Table (${cols}×${rows})`;
                tableComponent.resize(tableFrame.width, tableFrame.height);
                
                // Move all children from the frame to the component
                const children = [...tableFrame.children];
                for (const child of children) {
                    tableComponent.appendChild(child);
                }
                
                // Copy all properties from the frame to the component
                tableComponent.layoutMode = tableFrame.layoutMode;
                tableComponent.counterAxisSizingMode = tableFrame.counterAxisSizingMode;
                tableComponent.primaryAxisSizingMode = tableFrame.primaryAxisSizingMode;
                tableComponent.itemSpacing = tableFrame.itemSpacing;
                tableComponent.paddingLeft = tableFrame.paddingLeft;
                tableComponent.paddingRight = tableFrame.paddingRight;
                tableComponent.paddingTop = tableFrame.paddingTop;
                tableComponent.paddingBottom = tableFrame.paddingBottom;
                tableComponent.fills = tableFrame.fills;
                tableComponent.strokes = tableFrame.strokes;
                tableComponent.strokeWeight = tableFrame.strokeWeight;
                tableComponent.cornerRadius = tableFrame.cornerRadius;
                tableComponent.effects = tableComponent.effects;
                
                // Copy plugin data
                tableComponent.setPluginData('isGeneratedTable', tableFrame.getPluginData('isGeneratedTable'));
                tableComponent.setPluginData('tableSettings', tableFrame.getPluginData('tableSettings'));
                tableComponent.setPluginData('tableGeneratorScan', tableFrame.getPluginData('tableGeneratorScan'));
                
                // Remove the original frame
                tableFrame.remove();
                
                // Add the component to the page
                figma.currentPage.appendChild(tableComponent);
                figma.viewport.scrollAndZoomIntoView([tableComponent]);
                
                console.log('✅ Table converted to component successfully');
                figma.notify('Table component created successfully!');
                figma.ui.postMessage({ type: 'table-created', success: true, isComponent: true });
            } catch (error) {
                console.error('❌ Error converting table to component:', error);
                // Fallback: use the original frame
                figma.currentPage.appendChild(tableFrame);
                figma.viewport.scrollAndZoomIntoView([tableFrame]);
                figma.notify('Table created successfully! (as frame)');
                figma.ui.postMessage({ type: 'table-created', success: true, isComponent: false });
            }

        } catch (error) {
            console.error('Error in create-table-from-scan:', error);
            figma.notify('❌ An unexpected error occurred. Please check the console for details.');
        } finally {
            isCreatingTable = false;
        }
    } else if (msg.type === 'update-table') {
        if (isCreatingTable) {
            figma.notify("Already updating a table. Please wait.");
            return;
        }
        isCreatingTable = true;
        
        // Make this async to support font loading
        (async () => {
        try {
            // Get the table frame or component first
            const tableNode = figma.getNodeById(msg.tableId);
            if (!tableNode || (tableNode.type !== 'FRAME' && tableNode.type !== 'COMPONENT')) {
                figma.notify("❌ Table to update not found.");
                return;
            }
            
            // Cast to the appropriate type for processing
            const tableFrame = tableNode as FrameNode | ComponentNode;
            
            // For updates, try to restore lastScanResult from table frame plugin data first
            if (!lastScanResult) {
                try {
                    const scanData = tableFrame.getPluginData('tableGeneratorScan');
                    if (scanData) {
                        const storedScan = JSON.parse(scanData);
                            console.log('🔄 Attempting to restore scan result from table frame plugin data...');
                        console.log('Stored scan data:', storedScan);
                        
                        // Try to restore components by their IDs
                        let headerCell = null, bodyCell = null, footer = null, selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
                        let headerRowComponent = null, bodyRowComponent = null;
                        
                        // Restore body cell (most critical)
                        if (storedScan.bodyCellId) {
                            try {
                                const bodyCellNode = figma.getNodeById(storedScan.bodyCellId);
                                if (bodyCellNode && bodyCellNode.type === 'COMPONENT') {
                                    bodyCell = bodyCellNode;
                                    console.log('✅ Restored body cell from stored ID:', bodyCell.name);
                                } else {
                                    console.log('❌ Body cell node not found or not a component:', storedScan.bodyCellId);
                                }
                            } catch (error) {
                                console.log('Could not restore body cell from ID:', error);
                            }
                        } else {
                            console.log('❌ No body cell ID stored in scan data');
                        }
                        
                        // Restore header cell
                        if (storedScan.headerCellId) {
                            try {
                                const headerCellNode = figma.getNodeById(storedScan.headerCellId);
                                if (headerCellNode && headerCellNode.type === 'COMPONENT') {
                                    headerCell = headerCellNode;
                                    console.log('✅ Restored header cell from stored ID:', headerCell.name);
                                } else {
                                    console.log('❌ Header cell node not found or not a component:', storedScan.headerCellId);
                                }
                            } catch (error) {
                                console.log('Could not restore header cell from ID:', error);
                            }
                        }
                        
                        // Restore footer
                        if (storedScan.footerId) {
                            try {
                                const footerNode = figma.getNodeById(storedScan.footerId);
                                if (footerNode && footerNode.type === 'COMPONENT') {
                                    footer = footerNode;
                                    console.log('✅ Restored footer from stored ID:', footer.name);
                                }
                            } catch (error) {
                                console.log('Could not restore footer from ID:', error);
                            }
                        }
                        
                        // Restore select cell
                        if (storedScan.selectCellId) {
                            try {
                                const selectCellNode = figma.getNodeById(storedScan.selectCellId);
                                if (selectCellNode && selectCellNode.type === 'COMPONENT') {
                                    selectCellComponent = selectCellNode;
                                    console.log('✅ Restored select cell from stored ID:', selectCellComponent.name);
                                }
                            } catch (error) {
                                console.log('Could not restore select cell from ID:', error);
                            }
                        }
                        
                        // Restore expand cell
                        if (storedScan.expandCellId) {
                            try {
                                const expandCellNode = figma.getNodeById(storedScan.expandCellId);
                                if (expandCellNode && expandCellNode.type === 'COMPONENT') {
                                    expandCellComponent = expandCellNode;
                                    console.log('✅ Restored expand cell from stored ID:', expandCellComponent.name);
                                }
                            } catch (error) {
                                console.log('Could not restore expand cell from ID:', error);
                            }
                        }
                        
                        // Restore divider
                        if (storedScan.dividerId) {
                            try {
                                const dividerNode = figma.getNodeById(storedScan.dividerId);
                                if (dividerNode && dividerNode.type === 'COMPONENT') {
                                    dividerComponent = dividerNode;
                                    console.log('✅ Restored divider from stored ID:', dividerComponent.name);
                                }
                            } catch (error) {
                                console.log('Could not restore divider from ID:', error);
                            }
                        }
                        
                        // Restore row components
                        if (storedScan.headerRowComponentId) {
                            try {
                                const headerRowNode = figma.getNodeById(storedScan.headerRowComponentId);
                                if (headerRowNode && headerRowNode.type === 'COMPONENT') {
                                    headerRowComponent = headerRowNode;
                                    console.log('✅ Restored header row component from stored ID:', headerRowComponent.name);
                                }
                            } catch (error) {
                                console.log('Could not restore header row component from ID:', error);
                            }
                        }
                        
                        if (storedScan.bodyRowComponentId) {
                            try {
                                const bodyRowNode = figma.getNodeById(storedScan.bodyRowComponentId);
                                if (bodyRowNode && bodyRowNode.type === 'COMPONENT') {
                                    bodyRowComponent = bodyRowNode;
                                    console.log('✅ Restored body row component from stored ID:', bodyRowComponent.name);
                                }
                            } catch (error) {
                                console.log('Could not restore body row component from ID:', error);
                            }
                        }
                        
                        // If we successfully restored the body cell, create a lastScanResult
                        if (bodyCell) {
                            lastScanResult = {
                                headerCell,
                                headerRowComponent,
                                bodyCell,
                                bodyRowComponent,
                                footer,
                                numCols: storedScan.numCols || 3,
                                selectCellComponent,
                                expandCellComponent,
                                dividerComponent
                            };
                            
                            // Restore header row properties if available
                            if (storedScan.headerRowProperties) {
                                (lastScanResult as any).headerRowProperties = storedScan.headerRowProperties;
                                console.log('✅ Restored header row properties:', storedScan.headerRowProperties);
                            }
                            
                            // Restore divider properties if available
                            if (storedScan.dividerProperties) {
                                (lastScanResult as any).dividerProperties = storedScan.dividerProperties;
                                console.log('✅ Restored divider properties:', storedScan.dividerProperties);
                            }
                            
                            // Restore original divider rectangle if available
                            if (storedScan.originalDividerRectId) {
                                try {
                                    const dividerRectNode = figma.getNodeById(storedScan.originalDividerRectId);
                                    if (dividerRectNode && dividerRectNode.type === 'RECTANGLE') {
                                        (lastScanResult as any).originalDividerRect = dividerRectNode;
                                        console.log('✅ Restored original divider rectangle:', dividerRectNode.name);
                                    }
                                } catch (error) {
                                    console.log('⚠️ Could not restore original divider rectangle:', error);
                                }
                            }
                            
                            console.log('✅ Successfully restored lastScanResult from stored data');
                            console.log('Restored components:', {
                                headerCell: headerCell?.name,
                                bodyCell: bodyCell?.name,
                                footer: footer?.name,
                                selectCell: selectCellComponent?.name,
                                expandCell: expandCellComponent?.name,
                                divider: dividerComponent?.name
                            });
                        } else {
                            console.log('❌ Could not restore body cell - this is critical for table updates');
                            if (storedScan.componentNames) {
                                console.log('Expected component names from scan:', storedScan.componentNames);
                                
                                // Try to find components by name as a fallback with more flexible matching
                                console.log('🔍 Attempting to find components by name with flexible matching...');
                                const allComponents = figma.root.findAll(n => n.type === 'COMPONENT' || n.type === 'COMPONENT_SET') as (ComponentNode | ComponentSetNode)[];
                                const componentNodes = allComponents.filter(comp => comp.type === 'COMPONENT') as ComponentNode[];
                                
                                console.log(`Found ${componentNodes.length} total components in document`);
                                
                                // More flexible name matching for body cell
                                if (storedScan.componentNames.bodyCell) {
                                    const targetName = storedScan.componentNames.bodyCell.toLowerCase();
                                    let foundBodyCell = componentNodes.find(comp => 
                                        comp.name.toLowerCase() === targetName
                                    );
                                    
                                    // If exact match not found, try partial matches
                                    if (!foundBodyCell) {
                                        foundBodyCell = componentNodes.find(comp => 
                                            comp.name.toLowerCase().includes('cell') && 
                                            !comp.name.toLowerCase().includes('header') && 
                                            !comp.name.toLowerCase().includes('footer') &&
                                            !comp.name.toLowerCase().includes('select') &&
                                            !comp.name.toLowerCase().includes('expand')
                                        );
                                    }
                                    
                                    if (foundBodyCell) {
                                        bodyCell = foundBodyCell;
                                        console.log('✅ Found body cell by name (flexible):', foundBodyCell.name);
                                    }
                                }
                                
                                // More flexible name matching for header cell
                                if (storedScan.componentNames.headerCell) {
                                    const targetName = storedScan.componentNames.headerCell.toLowerCase();
                                    let foundHeaderCell = componentNodes.find(comp => 
                                        comp.name.toLowerCase() === targetName
                                    );
                                    
                                    // If exact match not found, try partial matches
                                    if (!foundHeaderCell) {
                                        foundHeaderCell = componentNodes.find(comp => 
                                            comp.name.toLowerCase().includes('header') && 
                                            comp.name.toLowerCase().includes('cell')
                                        );
                                    }
                                    
                                    if (foundHeaderCell) {
                                        headerCell = foundHeaderCell;
                                        console.log('✅ Found header cell by name (flexible):', foundHeaderCell.name);
                                    }
                                }
                                
                                // More flexible name matching for footer
                                if (storedScan.componentNames.footer) {
                                    const targetName = storedScan.componentNames.footer.toLowerCase();
                                    let foundFooter = componentNodes.find(comp => 
                                        comp.name.toLowerCase() === targetName
                                    );
                                    
                                    // If exact match not found, try partial matches
                                    if (!foundFooter) {
                                        foundFooter = componentNodes.find(comp => 
                                            comp.name.toLowerCase().includes('footer')
                                        );
                                    }
                                    
                                    if (foundFooter) {
                                        footer = foundFooter;
                                        console.log('✅ Found footer by name (flexible):', foundFooter.name);
                                    }
                                }
                                
                                // More flexible name matching for divider
                                if (storedScan.componentNames.divider) {
                                    const targetName = storedScan.componentNames.divider.toLowerCase();
                                    let foundDivider = componentNodes.find(comp => 
                                        comp.name.toLowerCase() === targetName
                                    );
                                    
                                    // If exact match not found, try partial matches
                                    if (!foundDivider) {
                                        foundDivider = componentNodes.find(comp => 
                                            comp.name.toLowerCase().includes('divider')
                                        );
                                    }
                                    
                                    if (foundDivider) {
                                        dividerComponent = foundDivider;
                                        console.log('✅ Found divider by name (flexible):', foundDivider.name);
                                    }
                                }
                                
                                // If we found the body cell by name, create the lastScanResult
                                if (bodyCell) {
                                    lastScanResult = {
                                        headerCell,
                                        headerRowComponent,
                                        bodyCell,
                                        bodyRowComponent,
                                        footer,
                                        numCols: storedScan.numCols || 3,
                                        selectCellComponent,
                                        expandCellComponent,
                                        dividerComponent
                                    };
                                    
                                    // Restore header row properties if available
                                    if (storedScan.headerRowProperties) {
                                        (lastScanResult as any).headerRowProperties = storedScan.headerRowProperties;
                                        console.log('✅ Restored header row properties (name-based):', storedScan.headerRowProperties);
                                    }
                                    
                                    // Restore divider properties if available
                                    if (storedScan.dividerProperties) {
                                        (lastScanResult as any).dividerProperties = storedScan.dividerProperties;
                                        console.log('✅ Restored divider properties (name-based):', storedScan.dividerProperties);
                                    }
                                    
                                    // Restore original divider rectangle if available
                                    if (storedScan.originalDividerRectId) {
                                        try {
                                            const dividerRectNode = figma.getNodeById(storedScan.originalDividerRectId);
                                            if (dividerRectNode && dividerRectNode.type === 'RECTANGLE') {
                                                (lastScanResult as any).originalDividerRect = dividerRectNode;
                                                console.log('✅ Restored original divider rectangle (name-based):', dividerRectNode.name);
                                            }
                                        } catch (error) {
                                            console.log('⚠️ Could not restore original divider rectangle (name-based):', error);
                                        }
                                    }
                                    
                                    console.log('✅ Successfully restored lastScanResult using name-based fallback');
                                } else {
                                    console.log('❌ Could not find any suitable body cell component');
                                    console.log('Available components for debugging:', componentNodes.map(c => c.name).slice(0, 20)); // Show first 20
                                }
                            }
                        }
                    } else {
                        console.log('❌ No scan data found in table frame plugin data');
                    }
                } catch (error) {
                    console.log('Could not restore scan result from plugin data:', error);
                }
            }
            
            // Find components dynamically if lastScanResult is still not available
            if (!lastScanResult) {
                console.log('No lastScanResult available, finding components dynamically...');
                
                // More intelligent component search that prioritizes original components over instances
                let headerCell = null, bodyCell = null, footer = null, selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
                
                // First, try to find components by looking at the existing table instances
                // This is more reliable than searching the entire document
                try {
                    const tableInstances = tableFrame.findAll(n => n.type === 'INSTANCE') as InstanceNode[];
                    console.log(`Found ${tableInstances.length} instances in the table frame`);
                    
                    // Group instances by their main component to avoid duplicates
                    const componentMap = new Map<string, ComponentNode>();
                    
                    for (const instance of tableInstances) {
                        if (instance.mainComponent) {
                            const compName = instance.mainComponent.name.toLowerCase();
                            console.log(`Found table instance: ${instance.mainComponent.name}`);
                            
                            // Categorize components based on their names
                            if (!bodyCell && compName.includes('cell') && !compName.includes('header') && !compName.includes('footer') && !compName.includes('select') && !compName.includes('expand')) {
                                bodyCell = instance.mainComponent;
                                console.log(`✅ Found body cell from table instance: ${bodyCell.name}`);
                            } else if (!headerCell && compName.includes('header') && compName.includes('cell')) {
                                headerCell = instance.mainComponent;
                                console.log(`✅ Found header cell from table instance: ${headerCell.name}`);
                            } else if (!footer && compName.includes('footer')) {
                                footer = instance.mainComponent;
                                console.log(`✅ Found footer from table instance: ${footer.name}`);
                            } else if (!selectCellComponent && compName.includes('select') && compName.includes('cell')) {
                                selectCellComponent = instance.mainComponent;
                                console.log(`✅ Found select cell from table instance: ${selectCellComponent.name}`);
                            } else if (!expandCellComponent && compName.includes('expand') && compName.includes('cell')) {
                                expandCellComponent = instance.mainComponent;
                                console.log(`✅ Found expand cell from table instance: ${expandCellComponent.name}`);
                            } else if (!dividerComponent && compName.includes('divider')) {
                                dividerComponent = instance.mainComponent;
                                console.log(`✅ Found divider from table instance: ${dividerComponent.name}`);
                            }
                            
                            // Store all unique components
                            componentMap.set(instance.mainComponent.id, instance.mainComponent);
                        }
                    }
                    
                    console.log(`Found ${componentMap.size} unique components from table instances`);
                    
                    // If we didn't find essential components from instances, try a broader search
                    if (!bodyCell || !headerCell || !footer) {
                        console.log('🔍 Some components not found in table instances, searching more broadly...');
                        
                        // Search for components in the current page and other pages
                        const allComponents = figma.root.findAll(n => n.type === 'COMPONENT' || n.type === 'COMPONENT_SET') as (ComponentNode | ComponentSetNode)[];
                        const componentNodes = allComponents.filter(comp => comp.type === 'COMPONENT') as ComponentNode[];
                        
                        console.log(`Found ${componentNodes.length} total components in document`);
                        
                        for (const comp of componentNodes) {
                    const compName = comp.name.toLowerCase();
                    
                            // More specific search patterns for body cell
                    if (!bodyCell) {
                                if (compName.includes('body') && compName.includes('cell')) {
                            bodyCell = comp;
                                    console.log(`✅ Found body cell (body cell): ${comp.name}`);
                                } else if (compName.includes('data table') && compName.includes('cell') && !compName.includes('header') && !compName.includes('footer')) {
                            bodyCell = comp;
                                    console.log(`✅ Found body cell (data table cell): ${comp.name}`);
                                } else if (compName.includes('table') && compName.includes('cell') && !compName.includes('header') && !compName.includes('footer') && !compName.includes('select') && !compName.includes('expand')) {
                            bodyCell = comp;
                            console.log(`✅ Found body cell (table cell): ${comp.name}`);
                        } else if (compName.includes('row') && compName.includes('cell') && !compName.includes('header')) {
                            bodyCell = comp;
                            console.log(`✅ Found body cell (row cell): ${comp.name}`);
                        }
                    }
                    
                    // Header cell search
                            if (!headerCell && compName.includes('header') && compName.includes('cell')) {
                        headerCell = comp;
                        console.log(`✅ Found header cell: ${comp.name}`);
                    }
                    
                    // Footer search
                    if (!footer && compName.includes('footer')) {
                        footer = comp;
                        console.log(`✅ Found footer: ${comp.name}`);
                    }
                    
                    // Select cell search
                    if (!selectCellComponent && compName.includes('select') && compName.includes('cell')) {
                        selectCellComponent = comp;
                        console.log(`✅ Found select cell: ${comp.name}`);
                    }
                    
                    // Expand cell search
                    if (!expandCellComponent && compName.includes('expand') && compName.includes('cell')) {
                        expandCellComponent = comp;
                        console.log(`✅ Found expand cell: ${comp.name}`);
                    }
                            
                            // Divider search
                            if (!dividerComponent && compName.includes('divider')) {
                                dividerComponent = comp;
                                console.log(`✅ Found divider: ${comp.name}`);
                            }
                        }
                    }
                    
                } catch (error) {
                    console.log('Error searching for components:', error);
                }
                
                // Validate that we found at least the essential components
                if (!bodyCell) {
                    console.error('❌ Could not find body cell component for table update');
                    
                    // Last resort: create a simple fallback component
                    console.log('🛠️ Creating fallback body cell component...');
                    try {
                        // Load font first
                        await figma.loadFontAsync({ family: "Inter", style: "Regular" });
                        
                        const fallbackComponent = figma.createComponent();
                        fallbackComponent.name = "Fallback Body Cell";
                        fallbackComponent.resize(100, 40);
                        
                        const textNode = figma.createText();
                        textNode.characters = "Cell";
                        textNode.x = 8;
                        textNode.y = 8;
                        fallbackComponent.appendChild(textNode);
                        
                        bodyCell = fallbackComponent;
                        console.log('✅ Created fallback body cell component');
                        figma.notify('⚠️ Created fallback component. Consider re-scanning the table for better results.');
                    } catch (error) {
                        console.error('❌ Could not create fallback component:', error);
                            figma.notify('❌ Could not find any suitable components. Please re-scan the table or ensure table components are available.');
                            return;
                    }
                }
                
                console.log('✅ Dynamic component search results:', {
                    headerCell: headerCell?.name || 'Not found',
                    bodyCell: bodyCell?.name || 'Not found',
                    footer: footer?.name || 'Not found',
                    selectCellComponent: selectCellComponent?.name || 'Not found',
                    expandCellComponent: expandCellComponent?.name || 'Not found'
                });
                
                // Create a temporary lastScanResult
                lastScanResult = {
                    headerCell: headerCell as ComponentNode | null,
                    headerRowComponent: null,
                    bodyCell: bodyCell as ComponentNode | null,
                    bodyRowComponent: null,
                    footer: footer as ComponentNode | null,
                    numCols: msg.cols || 5,
                    selectCellComponent: selectCellComponent as ComponentNode | null,
                    expandCellComponent: expandCellComponent as ComponentNode | null,
                    dividerComponent: dividerComponent as ComponentNode | null
                };
                
                // Try to extract divider properties from the found divider component
                if (dividerComponent) {
                    console.log('🔍 Extracting divider properties from found component:', dividerComponent.name);
                    const dividerProperties: { [key: string]: any } = {};
                    
                    try {
                        // Store the actual visual properties from the divider component
                        dividerProperties.fills = dividerComponent.fills;
                        dividerProperties.strokes = dividerComponent.strokes;
                        dividerProperties.strokeWeight = dividerComponent.strokeWeight;
                        dividerProperties.cornerRadius = dividerComponent.cornerRadius;
                        dividerProperties.effects = dividerComponent.effects;
                        dividerProperties.height = dividerComponent.height;
                        
                        // Check for color variables in main component fills
                        if (dividerComponent.fills && Array.isArray(dividerComponent.fills) && dividerComponent.fills.length > 0) {
                            for (const fill of dividerComponent.fills) {
                                if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                                    console.log('🎨 Found color variable in divider component:', fill.boundVariables.color);
                                    dividerProperties.colorVariable = fill.boundVariables.color;
                                    break;
                                }
                            }
                        }
                        
                        // Check child elements for properties
                        if ('children' in dividerComponent && dividerComponent.children.length > 0) {
                            const child = dividerComponent.children[0];
                            if ('fills' in child && child.fills && Array.isArray(child.fills) && child.fills.length > 0) {
                                dividerProperties.childFills = child.fills;
                                for (const fill of child.fills) {
                                    if (fill.type === 'SOLID' && fill.boundVariables?.color) {
                                        dividerProperties.childColorVariable = fill.boundVariables.color;
                                        break;
                                    }
                                }
                            }
                            if ('strokes' in child && child.strokes) dividerProperties.childStrokes = child.strokes;
                            if ('strokeWeight' in child) dividerProperties.childStrokeWeight = child.strokeWeight;
                            if ('height' in child) dividerProperties.childHeight = child.height;
                        }
                        
                        // Store divider properties in lastScanResult
                        (lastScanResult as any).dividerProperties = dividerProperties;
                        console.log('✅ Extracted and stored divider properties for update');
                    } catch (error) {
                        console.error('❌ Error extracting divider properties:', error);
                    }
                }
                
                // Try to find original divider rectangle in the table frame
                if (tableFrame && 'findOne' in tableFrame) {
                    try {
                        const dividerRect = tableFrame.findOne(n => 
                            n.type === 'RECTANGLE' && n.name === 'Divider'
                        ) as RectangleNode | null;
                        
                        if (dividerRect) {
                            (lastScanResult as any).originalDividerRect = dividerRect;
                            console.log('✅ Found original divider rectangle in table frame:', dividerRect.name);
                        }
                    } catch (error) {
                        console.log('⚠️ Could not find original divider rectangle in table frame:', error);
                    }
                }
            }

            console.log('✅ Found table frame to update:', tableFrame.name);
            console.log('📊 Update parameters:', {
                rows: msg.rows,
                cols: msg.cols,
                includeHeader: msg.includeHeader,
                includeFooter: msg.includeFooter,
                includeSelectable: msg.includeSelectable,
                includeExpandable: msg.includeExpandable,
                cellPropsKeys: Object.keys(msg.cellProps)
            });

            // Clear existing content
            const existingChildren = tableFrame.children.slice();
            console.log(`🗑️ Removing ${existingChildren.length} existing children from table frame`);
            existingChildren.forEach(child => child.remove());

            // Ensure table frame is properly configured
            tableFrame.layoutMode = "VERTICAL";
            tableFrame.counterAxisSizingMode = "AUTO";
            tableFrame.primaryAxisSizingMode = "AUTO";
            tableFrame.itemSpacing = 0;
            tableFrame.paddingLeft = 0;
            tableFrame.paddingRight = 0;
            tableFrame.paddingTop = 0;
            tableFrame.paddingBottom = 0;
            console.log('✅ Configured table frame layout properties');

            // Re-use the creation logic, but target the existing frame
            const { headerCell, bodyCell, footer, numCols } = lastScanResult!;
            const { rows, cols, includeHeader, includeFooter, includeSelectable, includeExpandable } = msg;
            let { cellProps } = msg;
            
            // Get the original table settings to preserve header/footer state
            let originalIncludeHeader = includeHeader;
            let originalIncludeFooter = includeFooter;
            let originalIncludeSelectable = includeSelectable;
            let originalIncludeExpandable = includeExpandable;
            
            // Apply sorting if enabled
            cellProps = sortColumnData(cellProps, cols, rows);
            
            try {
                const tableSettings = tableFrame.getPluginData('tableSettings');
                if (tableSettings) {
                    const settings = JSON.parse(tableSettings);
                    console.log('📋 Original table settings:', settings);
                    
                    // Use original settings if not provided in update message
                    if (originalIncludeHeader === undefined) {
                        originalIncludeHeader = settings.includeHeader !== false; // Default to true if not explicitly false
                        console.log('🔄 Using original includeHeader setting:', originalIncludeHeader);
                    }
                    if (originalIncludeFooter === undefined) {
                        originalIncludeFooter = settings.includeFooter === true; // Default to false if not explicitly true
                        console.log('🔄 Using original includeFooter setting:', originalIncludeFooter);
                    }
                    if (originalIncludeSelectable === undefined) {
                        originalIncludeSelectable = settings.includeSelectable === true;
                        console.log('🔄 Using original includeSelectable setting:', originalIncludeSelectable);
                    }
                    if (originalIncludeExpandable === undefined) {
                        originalIncludeExpandable = settings.includeExpandable === true;
                        console.log('🔄 Using original includeExpandable setting:', originalIncludeExpandable);
                    }
                }
            } catch (error) {
                console.log('⚠️ Could not read original table settings:', error);
            }
            
            // Validate that essential components are available
            if (!lastScanResult?.bodyCell) {
                console.error('❌ Body cell component not found for table update');
                console.log('🔄 Attempting to auto-scan for Data table components on the page...');
                
                // Try to auto-scan for a valid Data table component
                const autoScanResult = await autoScanForDataTable();
                
                if (autoScanResult && autoScanResult.bodyCell) {
                    console.log('✅ Auto-scan successful! Using found components for table update.');
                    figma.notify('✅ Found Data table components on the page. Using them for table update.');
                    
                    // Update the lastScanResult with the auto-scanned components
                    lastScanResult = autoScanResult;
                    
                    // Continue with the updated components
                    if (!autoScanResult.bodyCell) {
                        console.error('❌ Body cell component still not found after auto-scan');
                        figma.notify('❌ Cannot update table: No valid Data table components found on the page.');
                        return;
                    }
                    
                } else {
                    console.error('❌ No suitable Data table components found on the page');
                    figma.notify('❌ Cannot update table: No suitable Data table components found on the page. Please add a Data table component and try again.');
                    return;
                }
            }
            
            // Warn about missing optional components but continue
            if (originalIncludeHeader && !lastScanResult?.headerCell) {
                console.warn('⚠️ Header cell component not found, header will be skipped');
            }
            if (originalIncludeFooter && !lastScanResult?.footer) {
                console.warn('⚠️ Footer component not found, footer will be skipped');
            }
            if (originalIncludeSelectable && !lastScanResult?.selectCellComponent) {
                console.warn('⚠️ Select cell component not found, selectable functionality will be disabled');
            }
            if (originalIncludeExpandable && !lastScanResult?.expandCellComponent) {
                console.warn('⚠️ Expand cell component not found, expandable functionality will be disabled');
            }

            // (The entire table generation logic from 'create-table-from-scan' will be duplicated here,
            // but instead of creating a new tableFrame, it will append children to the existing one.)
            
            // --- Rebuild logic starts here ---
            const columnWidths: number[] = [];
            let totalTableWidth = 0;
            
            if (includeExpandable && lastScanResult.expandCellComponent) {
                try {
                    const expandCellInstance = lastScanResult.expandCellComponent.createInstance();
                    totalTableWidth += expandCellInstance.width;
                    expandCellInstance.remove();
                } catch (error) {
                    totalTableWidth += 52;
                }
            }
            if (includeSelectable && lastScanResult.selectCellComponent) {
                try {
                    const selectCellInstance = lastScanResult.selectCellComponent.createInstance();
                    totalTableWidth += selectCellInstance.width;
                    selectCellInstance.remove();
                } catch (error) {
                    totalTableWidth += 52;
                }
            }
            
            for (let c = 0; c < cols; c++) {
                let widthFound = false;
                for (let r = 0; r < rows; r++) {
                    const key = `${r}-${c}`;
                    const cellData = cellProps[key];
                    if (cellData && cellData.colWidth) {
                        columnWidths[c] = cellData.colWidth;
                        widthFound = true;
                        break;
                    }
                }
                if (!widthFound) {
                    // Use default width of 120 for initial table generation
                    columnWidths[c] = 120;
                }
                totalTableWidth += columnWidths[c];
            }
            
            if (originalIncludeHeader && lastScanResult?.headerCell) {
                try {
                    const headerRowFrame = figma.createFrame();
                    headerRowFrame.name = "Header Row";
                    headerRowFrame.layoutMode = "HORIZONTAL";
                    headerRowFrame.primaryAxisSizingMode = "AUTO";
                    headerRowFrame.counterAxisSizingMode = "AUTO";
                    headerRowFrame.itemSpacing = 0;
                    headerRowFrame.paddingLeft = 0;
                    headerRowFrame.paddingRight = 0;
                    headerRowFrame.paddingTop = 0;
                    headerRowFrame.paddingBottom = 0;
                    // Clear default fill to make it theme-compatible
                    headerRowFrame.fills = [];
                    
                    // Apply header row properties from scanned table to maintain color variables
                    if (lastScanResult && (lastScanResult as any).headerRowProperties) {
                        const headerRowProps = (lastScanResult as any).headerRowProperties;
                        console.log('🔄 Applying header row properties to generated header row (update):', headerRowProps);
                        
                        try {
                            // Apply fills if present in the original header row
                            if (headerRowProps.fills && Array.isArray(headerRowProps.fills) && headerRowProps.fills.length > 0) {
                                console.log('🎨 Applying fills (update):', headerRowProps.fills);
                                headerRowFrame.fills = headerRowProps.fills;
                            }
                            
                            // Apply strokes if present in the original header row
                            if (headerRowProps.strokes && Array.isArray(headerRowProps.strokes) && headerRowProps.strokes.length > 0) {
                                console.log('🎨 Applying strokes (update):', headerRowProps.strokes);
                                headerRowFrame.strokes = headerRowProps.strokes;
                            }
                            
                            // Apply stroke weight if present
                            if (headerRowProps.strokeWeight !== undefined && headerRowProps.strokeWeight > 0) {
                                console.log('🎨 Applying stroke weight (update):', headerRowProps.strokeWeight);
                                headerRowFrame.strokeWeight = headerRowProps.strokeWeight;
                            }
                            
                            // Apply corner radius if present
                            if (headerRowProps.cornerRadius !== undefined && headerRowProps.cornerRadius > 0) {
                                console.log('🎨 Applying corner radius (update):', headerRowProps.cornerRadius);
                                headerRowFrame.cornerRadius = headerRowProps.cornerRadius;
                            }
                            
                            // Apply effects if present
                            if (headerRowProps.effects && Array.isArray(headerRowProps.effects) && headerRowProps.effects.length > 0) {
                                console.log('🎨 Applying effects (update):', headerRowProps.effects);
                                headerRowFrame.effects = headerRowProps.effects;
                            }
                            
                            console.log('✅ Header row properties applied successfully (update)');
                        } catch (error) {
                            console.error('❌ Error applying header row properties (update):', error);
                        }
                    } else {
                        console.log('⚠️ No header row properties available to apply (update)');
                    }
                
                    if (originalIncludeExpandable && lastScanResult.expandCellComponent) {
                        try {
                            const expandCell = lastScanResult.expandCellComponent.createInstance();
                            headerRowFrame.appendChild(expandCell);
                        } catch (error) { console.error('Error creating header expand cell'); }
                    }

                    if (originalIncludeSelectable && lastScanResult.selectCellComponent) {
                        try {
                            const selectCell = lastScanResult.selectCellComponent.createInstance();
                            headerRowFrame.appendChild(selectCell);
                        } catch (error) { console.error('Error creating header select cell'); }
                    }

                    for (let c = 1; c <= cols; c++) {
                        const hCell = lastScanResult.headerCell.createInstance();
                        hCell.layoutSizingHorizontal = 'FIXED';
                        hCell.resize(columnWidths[c - 1], hCell.height);
                        headerRowFrame.appendChild(hCell);

                        const key = `header-${c}`;
                        const cellData = cellProps[key];
                        if (cellData && cellData.properties) {
                            try {
                                const validProps = mapPropertyNames(cellData.properties, hCell.componentProperties);
                                
                                // Try to set all properties at once first
                                try {
                                    hCell.setProperties(validProps);
                                } catch (variantError) {
                                    console.warn('Variant combination failed for header cell', key, 'trying individual properties');
                                    // Fallback: set properties one by one
                                    for (const [propName, propValue] of Object.entries(validProps)) {
                                        try {
                                            hCell.setProperties({ [propName]: propValue });
                                        } catch (individualError) {
                                            console.warn(`Could not set individual property ${propName} for header cell ${key}:`, individualError);
                                        }
                                    }
                                }
                            } catch (e) { console.warn('Could not set properties for header cell', key, e); }
                        }
                    }
                    tableFrame.appendChild(headerRowFrame);
                } catch (error) {
                    console.error('Error creating header row:', error);
                    figma.notify('⚠️ Header cell component is no longer available');
                }
            }

            if (lastScanResult?.bodyCell) {
                try {
                    const bodyWrapperFrame = figma.createFrame();
                    bodyWrapperFrame.name = 'Body';
                    bodyWrapperFrame.layoutMode = 'VERTICAL';
                    bodyWrapperFrame.primaryAxisSizingMode = 'AUTO';
                    bodyWrapperFrame.counterAxisSizingMode = 'AUTO';
                    bodyWrapperFrame.itemSpacing = 0;
                    bodyWrapperFrame.paddingLeft = 0;
                    bodyWrapperFrame.paddingRight = 0;
                    bodyWrapperFrame.paddingTop = 0;
                    bodyWrapperFrame.paddingBottom = 0;
                    // Clear default fill to make it theme-compatible
                    bodyWrapperFrame.fills = [];
                    tableFrame.appendChild(bodyWrapperFrame);

                    for (let r = 0; r < rows; r++) {
                        // Create a body row item frame (similar to "Data table body row item")
                        const bodyRowItemFrame = figma.createFrame();
                        bodyRowItemFrame.name = `Body Row Item ${r + 1}`;
                        bodyRowItemFrame.layoutMode = "VERTICAL";
                        bodyRowItemFrame.primaryAxisSizingMode = "AUTO";
                        bodyRowItemFrame.counterAxisSizingMode = "AUTO";
                        bodyRowItemFrame.itemSpacing = 0;
                        bodyRowItemFrame.paddingLeft = 0;
                        bodyRowItemFrame.paddingRight = 0;
                        bodyRowItemFrame.paddingTop = 0;
                        bodyRowItemFrame.paddingBottom = 0;
                        // Clear default fill to make it theme-compatible
                        bodyRowItemFrame.fills = [];
                        
                        // Create the data row frame (similar to "Data table row")
                        const rowFrame = figma.createFrame();
                        rowFrame.name = `Row ${r + 1}`;
                        rowFrame.layoutMode = "HORIZONTAL";
                        rowFrame.primaryAxisSizingMode = "AUTO";
                        rowFrame.counterAxisSizingMode = "AUTO";
                        rowFrame.itemSpacing = 0;
                        rowFrame.paddingLeft = 0;
                        rowFrame.paddingRight = 0;
                        rowFrame.paddingTop = 0;
                        rowFrame.paddingBottom = 0;
                        // Clear default fill to make it theme-compatible
                        rowFrame.fills = [];

                        if (originalIncludeExpandable && lastScanResult.expandCellComponent) {
                            try {
                    const expandCell = lastScanResult.expandCellComponent.createInstance();
                    rowFrame.appendChild(expandCell);
                            } catch (error) { console.error('Error creating expand cell instance'); }
                        }

                        if (originalIncludeSelectable && lastScanResult.selectCellComponent) {
                            try {
                                const selectCell = lastScanResult.selectCellComponent.createInstance();
                                rowFrame.appendChild(selectCell);
                            } catch (error) { console.error('Error creating select cell instance'); }
                }

                for (let c = 0; c < cols; c++) {
                            const cell = lastScanResult.bodyCell.createInstance();
                    cell.layoutSizingHorizontal = 'FIXED';
                    cell.resize(columnWidths[c], cell.height);
                    rowFrame.appendChild(cell);

                    const key = `${r}-${c}`;
                    const cellData = cellProps[key];
                    if (cellData && cellData.properties) {
                        try {
                            const validProps = mapPropertyNames(cellData.properties, cell.componentProperties);
                            cell.setProperties(validProps);
                                } catch (e) { console.warn('Could not set properties for cell', key, e); }
                    }
                }
                
                // Add the data row to the body row item frame
                bodyRowItemFrame.appendChild(rowFrame);
                
                // Add divider after the data row (including the last row)
                {
                    console.log('🔍 Checking divider component for row', r + 1, '(update):', {
                        hasDividerComponent: !!lastScanResult.dividerComponent,
                        dividerComponentName: lastScanResult.dividerComponent?.name,
                        hasDividerProperties: !!(lastScanResult as any).dividerProperties,
                        hasOriginalDividerRect: !!(lastScanResult as any).originalDividerRect
                    });
                    
                    if (lastScanResult.dividerComponent) {
                        try {
                            const divider = lastScanResult.dividerComponent.createInstance();
                            
                            // Apply divider properties from scanned table
                            if (lastScanResult && (lastScanResult as any).dividerProperties) {
                                const dividerProps = (lastScanResult as any).dividerProperties;
                                console.log('🔄 Applying divider properties (update):', dividerProps);
                                
                                try {
                                    // Priority order for color: 1) Scanned divider color variable, 2) Hardcoded border-subtle-01, 3) Scanned divider fills, 4) Fallback
                                    let appliedColor = false;
                                    
                                    // First priority: Apply color variable from scanned divider
                                    const scannedColorVariable = dividerProps.colorVariable || dividerProps.instanceColorVariable || dividerProps.instanceChildColorVariable || dividerProps.childColorVariable;
                                    if (scannedColorVariable) {
                                        console.log('🎨 Applying scanned divider color variable (update):', scannedColorVariable);
                                        try {
                                            const colorVariableFill: Paint = {
                                                type: 'SOLID',
                                                color: { r: 0, g: 0, b: 0 }, // Default color (will be overridden by variable)
                                                boundVariables: {
                                                    color: scannedColorVariable
                                                }
                                            };
                                            divider.fills = [colorVariableFill];
                                            appliedColor = true;
                                            console.log('✅ Applied scanned divider color variable (update)');
                                        } catch (error) {
                                            console.error('❌ Error applying scanned divider color variable (update):', error);
                                        }
                                    }
                                    
                                    // Second priority: Use hardcoded border-subtle-01 color variable
                                    if (!appliedColor) {
                                        console.log('🎨 Applying hardcoded border-subtle-01 color variable (update)');
                                        try {
                                            // Find the border-subtle-01 variable
                                            const localVariables = figma.variables.getLocalVariables();
                                            const borderVariable = localVariables.find(v => v.name === 'border-subtle-01');
                                            
                                            if (borderVariable) {
                                                const colorVariableFill: Paint = {
                                                    type: 'SOLID',
                                                    color: { r: 0, g: 0, b: 0 }, // Default color (will be overridden by variable)
                                                    boundVariables: {
                                                        color: {
                                                            id: borderVariable.id,
                                                            type: 'VARIABLE_ALIAS'
                                                        }
                                                    }
                                                };
                                                divider.fills = [colorVariableFill];
                                                appliedColor = true;
                                                console.log('✅ Applied hardcoded border-subtle-01 color variable to divider (update)');
                                            } else {
                                                console.log('⚠️ border-subtle-01 variable not found in local variables (update)');
                                            }
                                        } catch (error) {
                                            console.error('❌ Error applying hardcoded border-subtle-01 color variable (update):', error);
                                        }
                                    }
                                    
                                    // Third priority: Apply fills from scanned divider
                                    if (!appliedColor && dividerProps.fills && Array.isArray(dividerProps.fills) && dividerProps.fills.length > 0) {
                                        console.log('🎨 Applying divider fills (update):', dividerProps.fills);
                                        divider.fills = dividerProps.fills;
                                        appliedColor = true;
                                    }
                                    
                                    // Also apply instance fills if no color variable was applied
                                    if (!appliedColor && dividerProps.instanceFills && Array.isArray(dividerProps.instanceFills) && dividerProps.instanceFills.length > 0) {
                                        console.log('🎨 Applying instance fills to divider (update)');
                                        divider.fills = dividerProps.instanceFills;
                                        appliedColor = true;
                                    }
                                    
                                    // Last resort: Apply fallback gray color
                                    if (!appliedColor) {
                                        console.log('⚠️ No color variable found, applying fallback gray color (update)');
                                        divider.fills = [{
                                            type: 'SOLID',
                                            color: { r: 0.9, g: 0.9, b: 0.9 } // Light gray
                                        }];
                                        console.log('✅ Applied fallback gray color to divider (update)');
                                    }
                                    
                                    // Apply other properties (strokes, stroke weight, etc.)
                                    if (dividerProps.strokes && Array.isArray(dividerProps.strokes) && dividerProps.strokes.length > 0) {
                                        console.log('🎨 Applying divider strokes (update):', dividerProps.strokes);
                                        divider.strokes = dividerProps.strokes;
                                    }
                                    
                                    if (dividerProps.strokeWeight !== undefined && dividerProps.strokeWeight > 0) {
                                        console.log('🎨 Applying divider stroke weight (update):', dividerProps.strokeWeight);
                                        divider.strokeWeight = dividerProps.strokeWeight;
                                    }
                                    
                                    if (dividerProps.cornerRadius !== undefined && dividerProps.cornerRadius > 0) {
                                        console.log('🎨 Applying divider corner radius (update):', dividerProps.cornerRadius);
                                        divider.cornerRadius = dividerProps.cornerRadius;
                                    }
                                    
                                    if (dividerProps.effects && Array.isArray(dividerProps.effects) && dividerProps.effects.length > 0) {
                                        console.log('🎨 Applying divider effects (update):', dividerProps.effects);
                                        divider.effects = dividerProps.effects;
                                    }
                                    
                                    if (dividerProps.height !== undefined && dividerProps.height > 0) {
                                        console.log('📏 Applying divider height:', dividerProps.height);
                                        divider.resize(divider.width, dividerProps.height);
                                    }
                                    
                                    // Apply instance strokes and stroke weight
                                    if (dividerProps.instanceStrokes && Array.isArray(dividerProps.instanceStrokes) && dividerProps.instanceStrokes.length > 0) {
                                        console.log('🎨 Applying instance strokes to divider');
                                        divider.strokes = dividerProps.instanceStrokes;
                                    }
                                    if (dividerProps.instanceStrokeWeight !== undefined) {
                                        console.log('🎨 Applying instance stroke weight to divider');
                                        divider.strokeWeight = dividerProps.instanceStrokeWeight;
                                    }
                                    if (dividerProps.instanceHeight !== undefined) {
                                        console.log('📏 Applying instance height to divider');
                                        divider.resize(divider.width, dividerProps.instanceHeight);
                                    }
                                    
                                    // Apply instance child properties to divider's children
                                    if (dividerProps.instanceChildFills || dividerProps.instanceChildStrokes || dividerProps.instanceChildStrokeWeight) {
                                        console.log('🎨 Applying instance child properties to divider children');
                                        if ('children' in divider) {
                                            console.log('🎨 Divider has', divider.children.length, 'children');
                                            for (const child of divider.children) {
                                                console.log('🎨 Processing divider child:', child.name, child.type);
                                                
                                                if ('fills' in child && dividerProps.instanceChildFills && Array.isArray(dividerProps.instanceChildFills) && dividerProps.instanceChildFills.length > 0) {
                                                    console.log('🎨 Applying instance child fills to:', child.name);
                                                    child.fills = dividerProps.instanceChildFills;
                                                } else if ('fills' in child && dividerProps.instanceChildColorVariable) {
                                                    console.log('🎨 Applying instance child color variable to:', child.name, dividerProps.instanceChildColorVariable);
                                                    try {
                                                        const colorVariableFill: Paint = {
                                                            type: 'SOLID',
                                                            color: { r: 0, g: 0, b: 0 },
                                                            boundVariables: {
                                                                color: dividerProps.instanceChildColorVariable
                                                            }
                                                        };
                                                        child.fills = [colorVariableFill];
                                                    } catch (error) {
                                                        console.error('❌ Error applying instance child color variable:', error);
                                                    }
                                                }
                                                if ('strokes' in child && dividerProps.instanceChildStrokes) {
                                                    console.log('🎨 Applying instance child strokes to:', child.name);
                                                    child.strokes = dividerProps.instanceChildStrokes;
                                                }
                                                if ('strokeWeight' in child && dividerProps.instanceChildStrokeWeight !== undefined) {
                                                    console.log('🎨 Applying instance child stroke weight to:', child.name);
                                                    child.strokeWeight = dividerProps.instanceChildStrokeWeight;
                                                }
                                                if ('height' in child && dividerProps.instanceChildHeight !== undefined && 'resize' in child) {
                                                    console.log('📏 Applying instance child height to:', child.name);
                                                    (child as any).resize(child.width, dividerProps.instanceChildHeight);
                                                }
                                                break; // Apply to first child only
                                            }
                                        }
                                    }
                                    
                                    // Apply child properties to divider's children
                                    if (dividerProps.childFills || dividerProps.childStrokes || dividerProps.childStrokeWeight) {
                                        console.log('🎨 Applying child properties to divider children (update)');
                                        if ('children' in divider) {
                                            console.log('🎨 Divider has', divider.children.length, 'children (update)');
                                            for (const child of divider.children) {
                                                console.log('🎨 Processing divider child:', child.name, child.type, '(update)');
                                                
                                                // Apply to direct children
                                                if ('fills' in child && dividerProps.childFills && Array.isArray(dividerProps.childFills) && dividerProps.childFills.length > 0) {
                                                    console.log('🎨 Applying child fills to:', child.name, '(update)');
                                                    child.fills = dividerProps.childFills;
                                                }
                                                if ('fills' in child && dividerProps.childColorVariable) {
                                                    console.log('🎨 Applying child color variable to:', child.name, dividerProps.childColorVariable);
                                                    try {
                                                        const colorVariableFill: Paint = {
                                                            type: 'SOLID',
                                                            color: { r: 0, g: 0, b: 0 },
                                                            boundVariables: {
                                                                color: dividerProps.childColorVariable
                                                            }
                                                        };
                                                        child.fills = [colorVariableFill];
                                                    } catch (error) {
                                                        console.error('❌ Error applying child color variable:', error);
                                                    }
                                                }
                                                if ('strokes' in child && dividerProps.childStrokes) {
                                                    console.log('🎨 Applying child strokes to:', child.name, '(update)');
                                                    child.strokes = dividerProps.childStrokes;
                                                }
                                                if ('strokeWeight' in child && dividerProps.childStrokeWeight !== undefined) {
                                                    console.log('🎨 Applying child stroke weight to:', child.name, '(update)');
                                                    child.strokeWeight = dividerProps.childStrokeWeight;
                                                }
                                                if ('height' in child && dividerProps.childHeight !== undefined && 'resize' in child) {
                                                    console.log('📏 Applying child height to:', child.name);
                                                    (child as any).resize(child.width, dividerProps.childHeight);
                                                }
                                                
                                                // Also check for nested children
                                                if ('children' in child && child.children.length > 0) {
                                                    console.log('🎨 Processing nested children in:', child.name, '(update)');
                                                    for (const grandChild of child.children) {
                                                        console.log('🎨 Processing grandchild:', grandChild.name, grandChild.type, '(update)');
                                                        if ('fills' in grandChild && dividerProps.childFills && Array.isArray(dividerProps.childFills) && dividerProps.childFills.length > 0) {
                                                            console.log('🎨 Applying child fills to grandchild:', grandChild.name, '(update)');
                                                            grandChild.fills = dividerProps.childFills;
                                                        } else if ('fills' in grandChild) {
                                                            // If grandchild fills are empty, try to apply border-subtle-01 color variable
                                                            console.log('🎨 Grandchild fills are empty, checking for border color variable');
                                                            const borderVariable = findBorderColorVariable();
                                                            if (borderVariable) {
                                                                try {
                                                                    const colorVariableFill: Paint = {
                                                                        type: 'SOLID',
                                                                        color: { r: 0, g: 0, b: 0 },
                                                                        boundVariables: {
                                                                            color: borderVariable
                                                                        }
                                                                    };
                                                                    grandChild.fills = [colorVariableFill];
                                                                    console.log('✅ Applied border-subtle-01 color variable to grandchild');
                                                                } catch (error) {
                                                                    console.error('❌ Error applying border color variable to grandchild:', error);
                                                                }
                                                            }
                                                        }
                                                        if ('strokes' in grandChild && dividerProps.childStrokes) {
                                                            console.log('🎨 Applying child strokes to grandchild:', grandChild.name);
                                                            grandChild.strokes = dividerProps.childStrokes;
                                                        }
                                                        if ('strokeWeight' in grandChild && dividerProps.childStrokeWeight !== undefined) {
                                                            console.log('🎨 Applying child stroke weight to grandchild:', grandChild.name);
                                                            grandChild.strokeWeight = dividerProps.childStrokeWeight;
                                                        }
                                                        break; // Apply to first grandchild only
                                                    }
                                                }
                                                break; // Apply to first child only
                                            }
                                        }
                                    }
                                    
                                    console.log('✅ Divider properties applied successfully (update)');
                                } catch (error) {
                                    console.error('❌ Error applying divider properties (update):', error);
                                }
                            }
                            
                            // Set divider width to match the table width
                            divider.resize(totalTableWidth, divider.height);
                            console.log('📏 Set divider width to:', totalTableWidth, '(update)');
                            
                            bodyRowItemFrame.appendChild(divider);
                            console.log('✅ Added divider after row', r + 1, '(update)');
                        } catch (error) { 
                            console.error('Error creating divider instance (update):', error);
                        }
                    } else if ((lastScanResult as any).originalDividerRect) {
                        try {
                            console.log('🔄 Using original divider rectangle for row', r + 1, '(update)');
                            const divider = ((lastScanResult as any).originalDividerRect as RectangleNode).clone();
                            divider.resize(totalTableWidth, divider.height);
                            bodyRowItemFrame.appendChild(divider);
                            console.log('✅ Added cloned divider rectangle for row', r + 1, '(update)');
                        } catch (error) {
                            console.error('❌ Error cloning divider rectangle for update:', error);
                            console.log('⚠️ No divider available for row', r + 1, '(update)');
                        }
                    } else {
                        console.log('⚠️ No divider component or rectangle available for row', r + 1, '(update)');
                    }
                }
                
                bodyWrapperFrame.appendChild(bodyRowItemFrame);
                    }
                } catch (error) {
                    console.error('Error creating body rows:', error);
                    figma.notify('❌ Cannot generate table: body cell component is no longer available');
                    return;
                }
            }

            if (originalIncludeFooter && lastScanResult?.footer) {
                try {
                    const footerClone = lastScanResult.footer.createInstance();
            const footerCellData = cellProps['footer'];
            if (footerCellData && footerCellData.properties) {
                const validProps = mapPropertyNames(footerCellData.properties, footerClone.componentProperties);
                        try {
                footerClone.setProperties(validProps);
                        } catch (error) { console.error('Error in footer setProperties:', error); }
            }
                    if (footerClone) {
            footerClone.layoutSizingHorizontal = 'FIXED';
            footerClone.resize(totalTableWidth, footerClone.height);
            tableFrame.appendChild(footerClone);
                    }
                } catch (error) {
                    console.error('Error creating footer:', error);
                    figma.notify('⚠️ Footer component is no longer available');
                }
        }

        tableFrame.resize(totalTableWidth, tableFrame.height);

            // Update plugin data
            const tableSettings = {
                columns: msg.cols,
                rows: msg.rows,
                includeHeader: originalIncludeHeader,
                includeFooter: originalIncludeFooter,
                includeSelectable: originalIncludeSelectable,
                includeExpandable: originalIncludeExpandable,
                cellProperties: cellProps, // Use the sorted cellProps instead of msg.cellProps
            };
            tableFrame.setPluginData('tableSettings', JSON.stringify(tableSettings));

            figma.notify('✅ Table updated successfully!');
            figma.ui.postMessage({ type: 'table-updated', success: true });

        } catch (error) {
            console.error('Error in update-table:', error);
            figma.notify('❌ An unexpected error occurred while updating the table.');
        } finally {
            isCreatingTable = false;
        }
        })();
    } else if (msg.type === 'request-component-info') {
        console.log(`[Backend] Received request-component-info, lastScanResult:`, lastScanResult);
        console.log('[Backend] Request includes tableId:', msg.tableId);
        
        // Make this async to support performTableScan
        (async () => {
        let bodyCell: ComponentNode | null = null;
        
        if (lastScanResult && lastScanResult.bodyCell) {
            bodyCell = lastScanResult.bodyCell;
        } else {
            // If lastScanResult is not available, try to find the body cell component dynamically
            console.log(`[Backend] lastScanResult not available, searching for body cell component...`);
            
            // First, try to find a "Data table" component and extract the body cell from it
            console.log(`[Backend] Searching for Data table components...`);
            
            // Search for any "Data table" components on the current page
            const dataTableComponents = figma.currentPage.findAll(node => 
                node.type === 'INSTANCE' && 
                node.name.toLowerCase().includes('data table') &&
                !node.getPluginData('isGeneratedTable') // Exclude generated tables
            ) as InstanceNode[];
            
            console.log(`[Backend] Found ${dataTableComponents.length} Data table components:`, dataTableComponents.map(c => c.name));
            
            if (dataTableComponents.length > 0) {
                // Use the first Data table component to extract the body cell
                const selectedComponent = dataTableComponents[0];
                console.log(`[Backend] Using Data table component for scan:`, selectedComponent.name);
                
                try {
                    // Use the existing scan logic to extract components from the Data table
                    const scanResult = await performTableScan(selectedComponent);
                    
                    if (scanResult && scanResult.bodyCell) {
                        // The scanResult.bodyCell is an InstanceNode, we need to get its main component
                        try {
                            bodyCell = await (scanResult.bodyCell as any).getMainComponentAsync();
                            console.log(`[Backend] Successfully extracted body cell component from Data table:`, bodyCell?.name);
                        } catch (error) {
                            console.log(`[Backend] Could not get main component from body cell:`, error);
                        }
                    } else {
                        console.log(`[Backend] Could not extract body cell from Data table`);
                    }
                } catch (error) {
                    console.error(`[Backend] Error scanning Data table:`, error);
                }
            }
            
            // If we still don't have a body cell, try the broader search
            if (!bodyCell) {
                console.log(`[Backend] No Data table found or extraction failed, trying broader search...`);
            
            // Search for components that might be the body cell component
                // Search in all pages, not just the current page
                const allComponents: ComponentNode[] = [];
                
                console.log(`[Backend] Searching for components...`);
                console.log(`[Backend] Root has ${figma.root.children.length} children`);
                
                // Search in all pages
                figma.root.children.forEach((page, index) => {
                    console.log(`[Backend] Checking page ${index}: ${page.name} (type: ${page.type})`);
                    if (page.type === "PAGE") {
                        const pageComponents = page.findAll(node => 
                            node.type === "COMPONENT"
            ) as ComponentNode[];
                        console.log(`[Backend] Page ${page.name} has ${pageComponents.length} components:`, pageComponents.map(c => c.name));
                        allComponents.push(...pageComponents);
                    }
                });
                
                // Also search in the current page specifically
                const currentPageComponents = figma.currentPage.findAll(node => 
                    node.type === "COMPONENT"
                ) as ComponentNode[];
                console.log(`[Backend] Current page has ${currentPageComponents.length} components:`, currentPageComponents.map(c => c.name));
                allComponents.push(...currentPageComponents);
                
                // Remove duplicates based on ID
                const uniqueComponents = allComponents.filter((comp, index, self) => 
                    index === self.findIndex(c => c.id === comp.id)
                );
                
                console.log(`[Backend] Found ${uniqueComponents.length} total components in all pages:`, uniqueComponents.map(c => c.name));
                
                // Look for the most likely body cell component with more comprehensive search
                bodyCell = uniqueComponents.find(comp => 
                comp.name.includes("row") && comp.name.includes("cell") && comp.name.includes("item")
                ) || uniqueComponents.find(comp => 
                comp.name.includes("Data table") && comp.name.includes("cell")
                ) || uniqueComponents.find(comp => 
                comp.name.includes("table") && comp.name.includes("cell")
                ) || uniqueComponents.find(comp => 
                    comp.name.includes("cell") && !comp.name.includes("header") && !comp.name.includes("footer")
                ) || uniqueComponents.find(comp => 
                    comp.name.includes("body") && comp.name.includes("cell")
            ) || null;
            
            if (bodyCell) {
                console.log(`[Backend] Found body cell component:`, bodyCell.name);
            } else {
                    console.log(`[Backend] Could not find body cell component. Available components:`, uniqueComponents.map(c => c.name));
                    
                    // Try to find any component that has Slot properties
                    const componentWithSlot = uniqueComponents.find(comp => {
                        try {
                            const tempInstance = comp.createInstance();
                            const hasSlot = Object.keys(tempInstance.componentProperties).some(prop => 
                                prop.toLowerCase().includes('slot')
                            );
                            tempInstance.remove();
                            return hasSlot;
                        } catch (error) {
                            return false;
                        }
                    });
                    
                    if (componentWithSlot) {
                        console.log(`[Backend] Found component with Slot properties:`, componentWithSlot.name);
                        bodyCell = componentWithSlot;
                    } else {
                        console.log(`[Backend] No component with Slot properties found. Creating fallback...`);
                
                // If we can't find the component, we'll create a minimal component info
                // This allows the UI to still work with the saved properties
                figma.ui.postMessage({
                    type: 'component-info',
                    component: {
                        id: 'fallback',
                        name: 'Fallback Component',
                        width: 100,
                        properties: {},
                        availableProperties: ['Cell text#12234:16', 'Show text#12234:68', 'State', 'Size'],
                        propertyTypes: {
                            'Cell text#12234:16': 'TEXT',
                            'Show text#12234:68': 'BOOLEAN',
                            'State': 'VARIANT',
                            'Size': 'VARIANT'
                        }
                    }
                });
                return;
                    }
                }
            }
        }
        
        if (bodyCell) {
            // Create a temporary instance to get component properties
            const tempInstance = bodyCell.createInstance();
            const properties = tempInstance.componentProperties;
            const availableProperties = Object.keys(properties);
            const propertyTypes: { [key: string]: any } = {};
            
            for (const propName of availableProperties) {
                propertyTypes[propName] = properties[propName].type;
            }
            
            // Remove the temporary instance
            tempInstance.remove();
            
            console.log(`[Backend] Sending component-info with ${availableProperties.length} properties`);
            figma.ui.postMessage({
                type: 'component-info',
                component: {
                    id: bodyCell.id,
                    name: bodyCell.name,
                    width: bodyCell.width,
                    properties: properties,
                    availableProperties: availableProperties,
                    propertyTypes: propertyTypes
                }
            });
        }
        })();
    }
};

function createComponentInstanceWithProps(
  component: ComponentNode,
  cellData: { [key: string]: any },
  componentProps: { [key: string]: ComponentProperty },
  allProps: { [key: string]: any }
): SceneNode {
  console.log('🔧 Creating component instance with props:', { component: component.name, cellData, componentProps });
  
  // Create instance
  const instance = component.createInstance();
  
  // Map properties from cellData to component properties
  const mappedProps = mapPropertyNames(cellData, componentProps);
  console.log('🔧 Mapped properties:', mappedProps);
  
  // Apply the mapped properties to the instance
  for (const [propName, propValue] of Object.entries(mappedProps)) {
    try {
      console.log(`🔧 Setting property: ${propName} = ${propValue}`);
      instance.setProperties({ [propName]: propValue });
    } catch (error) {
      console.error(`❌ Error setting property ${propName}:`, error);
      // Try alternative method for setting properties
      try {
        (instance as any)[propName] = propValue;
        console.log(`🔧 Set property via direct assignment: ${propName} = ${propValue}`);
      } catch (directError) {
        console.error(`❌ Error setting property via direct assignment ${propName}:`, directError);
      }
    }
  }
  
  // Handle any additional properties that weren't mapped, especially slot properties
  for (const [originalKey, originalValue] of Object.entries(cellData)) {
    // Skip already mapped properties
    if (Object.keys(mappedProps).some(mappedKey => mappedKey.split('#')[0] === originalKey.split('#')[0])) {
      continue;
    }
    
    // Handle unmapped properties, especially slot properties
    // Ensure the value is of a valid type before setting
    if (typeof originalValue === 'string' || typeof originalValue === 'boolean') {
      try {
        console.log(`🔧 Setting unmapped property: ${originalKey} = ${originalValue}`);
        instance.setProperties({ [originalKey]: originalValue });
      } catch (error) {
        console.error(`❌ Error setting unmapped property ${originalKey}:`, error);
      }
    } else if (originalValue && typeof originalValue === 'object' && 'value' in originalValue) {
      // Handle case where the value is an object with a value property
      const valueToSet = originalValue.value;
      if (typeof valueToSet === 'string' || typeof valueToSet === 'boolean') {
        try {
          console.log(`🔧 Setting unmapped property from object: ${originalKey} = ${valueToSet}`);
          instance.setProperties({ [originalKey]: valueToSet });
        } catch (error) {
          console.error(`❌ Error setting unmapped property from object ${originalKey}:`, error);
        }
      }
    }
  }
  
  return instance;
}

function getFakerValue(type: string) {
  const t = type.toLowerCase();
  let result;
  switch (t) {
    case 'people name':
    case 'name':
      result = faker.name.findName();
      break;
    case 'first name':
      result = faker.name.firstName();
      break;
    case 'last name':
      result = faker.name.lastName();
      break;
    case 'brand name':
    case 'company':
      result = faker.company.companyName();
      break;
    case 'mobile number':
      result = faker.phone.phoneNumber();
      break;
    case 'date':
      result = faker.date.recent().toLocaleDateString();
      break;
    case 'random number':
    case 'number':
      result = faker.datatype.number({ min: 1, max: 1000 });
      break;
    case 'price':
      result = faker.commerce.price();
      break;
    case 'email':
      result = faker.internet.email();
      break;
    case 'product':
      result = faker.commerce.productName();
      break;
    case 'color':
      result = faker.commerce.color();
      break;
    default:
      result = faker.lorem.words(2);
  }
  return result;
}

interface ComponentProperty {
    type: 'BOOLEAN' | 'TEXT' | 'VARIANT' | 'INSTANCE_SWAP';
    value: any;
}

// Function to clean up any components created outside the generated table
function cleanupExternalComponents() {
    try {
        // Find and remove any "Simple Divider" components that might be outside the table
        const simpleDividers = figma.root.findAll(node => 
            node.type === "COMPONENT" && node.name === "Simple Divider"
        );
        
        simpleDividers.forEach(divider => {
            // Only remove if it's not inside a generated table
            let parent = divider.parent;
            let shouldRemove = true;
            
            while (parent) {
                if (parent.getPluginData('isGeneratedTable') === 'true') {
                    shouldRemove = false;
                    break;
                }
                parent = parent.parent;
            }
            
            if (shouldRemove) {
                divider.remove();
                console.log('🧹 Cleaned up external Simple Divider component');
            }
        });
        
        // Find and remove any "Data table select cell item" component sets that might be outside the table
        const checkboxComponents = figma.root.findAll(node => 
            node.type === "COMPONENT_SET" && node.name === "Data table select cell item"
        );
        
        checkboxComponents.forEach(componentSet => {
            // Only remove if it's not inside a generated table
            let parent = componentSet.parent;
            let shouldRemove = true;
            
            while (parent) {
                if (parent.getPluginData('isGeneratedTable') === 'true') {
                    shouldRemove = false;
                    break;
                }
                parent = parent.parent;
            }
            
            if (shouldRemove) {
                componentSet.remove();
                console.log('🧹 Cleaned up external checkbox component set');
            }
        });
    } catch (error) {
        console.error('❌ Error cleaning up external components:', error);
    }
}

// Call cleanup when the plugin starts
cleanupExternalComponents();

 