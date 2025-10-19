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
    
    // Explicitly preserve all header cell properties
    for (let c = 1; c <= cols; c++) {
        const headerKey = `header-${c}`;
        if (cellProps[headerKey]) {
            sortedCellProps[headerKey] = { ...cellProps[headerKey] };
            console.log(`[sortColumnData] Preserved header ${c} properties:`, sortedCellProps[headerKey]);
        }
    }
    
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
                
                // IMPORTANT: Preserve the header cell properties (Sortable, Sorted, etc.)
                // The header properties should remain unchanged after sorting
                console.log(`[sortColumnData] Preserving header properties for column ${c}:`, headerData.properties);
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
    toolbar?: ComponentNode | null;
    numCols: number;
    selectCellComponent: ComponentNode | null;
    expandCellComponent: ComponentNode | null;
    dividerComponent: ComponentNode | null;
}

let lastScanResult: ScanResult | undefined;
let isCreatingTable = false;

// Moved selection listener into core/selection.ts



import { initSelectionHandlers } from './core/selection';

// Initialize core listeners
initSelectionHandlers();

// Helper function to find matching property
function findMatchingProperty(availableProperties: string[], uiPropName: string): string | null {
    console.log(`🔍 [findMatchingProperty] Looking for: "${uiPropName}"`);
    console.log(`🔍 [findMatchingProperty] Available:`, availableProperties.slice(0, 5)); // Show first 5 for brevity
    
    // First try exact match
    const exactMatch = availableProperties.find((prop: string) => prop === uiPropName);
    if (exactMatch) {
        console.log(`✅ [findMatchingProperty] Exact match found: "${exactMatch}"`);
        return exactMatch;
    }
    
    // Try match at start of property name
    const startsWithMatch = availableProperties.find((prop: string) => prop.toLowerCase().startsWith(uiPropName.toLowerCase()));
    if (startsWithMatch) {
        console.log(`✅ [findMatchingProperty] StartsWith match found: "${startsWithMatch}"`);
        return startsWithMatch;
    }
    
    // Try contains match
    const containsMatch = availableProperties.find((prop: string) => prop.toLowerCase().includes(uiPropName.toLowerCase()));
    if (containsMatch) {
        console.log(`✅ [findMatchingProperty] Contains match found: "${containsMatch}"`);
        return containsMatch;
    }
    
    console.log(`❌ [findMatchingProperty] No match found for: "${uiPropName}"`);
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
            let toolbarComponent: ComponentNode | null = null;
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
                console.log(`🔍 Found ${allInstances.length} total instances in performTableScan`);
                
                // Log all component names for debugging
                console.log('🔍 All component names found in performTableScan:');
                allInstances.forEach(node => {
                    if (node.type === 'INSTANCE' && node.mainComponent) {
                        console.log(`  - ${node.mainComponent.name}`);
                        if (node.mainComponent.name.toLowerCase().includes('toolbar')) {
                            console.log(`    🎯 TOOLBAR FOUND: ${node.mainComponent.name}`);
                        }
                    } else if (node.type === 'COMPONENT') {
                        console.log(`  - ${node.name}`);
                        if (node.name.toLowerCase().includes('toolbar')) {
                            console.log(`    🎯 TOOLBAR FOUND: ${node.name}`);
                        }
                    }
                });
                
                for (const node of allInstances) {
                    // Check for toolbar components
                    if (node.name && node.name.toLowerCase().includes('toolbar')) {
                        console.log(`🎯 Found toolbar node: ${node.name} (type: ${node.type})`);
                        if (node.type === 'INSTANCE' && node.mainComponent && !toolbarComponent) {
                            toolbarComponent = node.mainComponent;
                            console.log(`✅ Found toolbar component: ${toolbarComponent.name}`);
                        }
                    }
                    
                    if (node.name === 'Data table expand cell item') {
                      expandCellFound = true;
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        expandCellComponent = node.mainComponent;
                        console.log('✅ Found Data table expand cell item from instance:', expandCellComponent.name);
                      } else if (node.type === 'COMPONENT') {
                        expandCellComponent = node;
                        console.log('✅ Found Data table expand cell item component directly:', expandCellComponent.name);
                      }
                    }
                    // Check for select cell components - only exact name match
                    if (node.name === 'Data table select cell item' && !selectCellFound) {
                      selectCellFound = true;
                      console.log('🔍 Found Data table select cell item:', node.name, node.type);
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        selectCellComponent = node.mainComponent;
                        console.log('✅ Using Data table select cell item from instance:', selectCellComponent.name);
                      } else if (node.type === 'COMPONENT') {
                        selectCellComponent = node;
                        console.log('✅ Using Data table select cell item component directly:', selectCellComponent.name);
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
                
                let headerCell = null, bodyCell = null, footer = null, toolbarComponent = null;
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
                
                if (storedScan.toolbarId) {
                    const toolbarNode = figma.getNodeById(storedScan.toolbarId);
                    if (toolbarNode && toolbarNode.type === 'COMPONENT') {
                        toolbarComponent = toolbarNode;
                        console.log('✅ Restored toolbar from stored ID:', toolbarComponent.name);
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
        
        let headerCell = null, bodyCell = null, footer = null, toolbarComponent = null;
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
                    console.log(`🔍 Checking component: ${instance.mainComponent.name} (${compName})`);
                    
                    // Check specifically for toolbar-related components
                    if (compName.includes('toolbar') || compName.includes('wrapper')) {
                        console.log(`🎯 TOOLBAR CANDIDATE: ${instance.mainComponent.name} - contains 'toolbar' or 'wrapper'`);
                    }
                    
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
                    } else if (!toolbarComponent && compName.includes('toolbar')) {
                        toolbarComponent = instance.mainComponent;
                        console.log(`✅ Found toolbar: ${toolbarComponent.name}`);
                    } else if (compName.includes('toolbar')) {
                        console.log(`🔍 Found toolbar component but already have one: ${instance.mainComponent.name}`);
                    } else if (compName.includes('toolbar') || compName.includes('wrapper')) {
                        console.log(`🔍 Component contains 'toolbar' or 'wrapper': ${instance.mainComponent.name} (${compName})`);
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
        let headerCell = null, bodyCell = null, footer = null, toolbarComponent = null;
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
        
        if (storedScan.toolbarId) {
            const toolbarNode = figma.getNodeById(storedScan.toolbarId);
            if (toolbarNode && toolbarNode.type === 'COMPONENT') {
                toolbarComponent = toolbarNode;
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

// ============================================================
// SMART SLOT DETECTION - Automatically detect content type and suggest slot components
// ============================================================

// Helper function to extract proper initials from a name
function extractInitials(name: string): string {
    if (!name || name.trim().length === 0) return '';
    
    const trimmedName = name.trim();
    const nameParts = trimmedName.split(/\s+/); // Split by whitespace
    
    if (nameParts.length === 1) {
        // Only first name - use first letter only
        return nameParts[0].charAt(0).toUpperCase();
    } else if (nameParts.length >= 2) {
        // First name + last name - use first letter of each
        const firstInitial = nameParts[0].charAt(0).toUpperCase();
        const lastInitial = nameParts[nameParts.length - 1].charAt(0).toUpperCase();
        return firstInitial + lastInitial;
    }
    
    return '';
}

interface SmartSlotSuggestion {
    columnIndex: number;
    columnName: string;
    contentType: 'status' | 'user' | 'action' | 'editDelete' | 'singleEdit' | 'singleDelete' | 'boolean' | 'url' | 'date' | 'number' | 'label' | 'text';
    suggestedComponent: 'statusIcon' | 'slotGroup' | 'tag' | 'avatar' | 'overflow' | 'checkbox' | 'link' | 'edit' | 'delete' | null;
    confidence: number; // 0-1
    samples: string[];
    componentId?: string;  // Added for auto-apply functionality
    componentName?: string;  // Added for auto-apply functionality
    actionType?: 'edit' | 'delete';  // For single action detection
}

function analyzeColumnContent(columnData: string[], columnName: string): SmartSlotSuggestion['contentType'] {
    console.log(`🔬 [analyzeColumnContent] Analyzing column: "${columnName}"`);
    
    // Remove empty values for analysis
    const nonEmptyData = columnData.filter(val => val && val.trim() !== '');
    if (nonEmptyData.length === 0) return 'text';
    
    // Action column detection - MUST come first to prevent misclassification
    // Check for column names with "action", "option", "menu", etc.
    if (columnName.toLowerCase().includes('action') || columnName.toLowerCase().includes('option') || columnName.toLowerCase().includes('menu')) {
        // Check if it's edit/delete actions (1-2 actions) vs overflow (3+ actions)
        const actionKeywords = ['edit', 'delete', 'remove', 'update', 'modify', 'trash', 'view', 'details', 'download', 'export', 'import', 'copy', 'duplicate'];
        const actionValues = nonEmptyData.filter(val => 
            actionKeywords.some(keyword => val.toLowerCase().includes(keyword))
        );
        
        const uniqueValues = new Set(nonEmptyData).size;
        console.log(`🔍 Action column analysis: ${columnName}`);
        console.log(`  📊 Total values: ${nonEmptyData.length}, Unique values: ${uniqueValues}`);
        console.log(`  🎯 Action keywords found: ${actionValues.length}/${nonEmptyData.length}`);
        console.log(`  📝 Values:`, nonEmptyData);
        
        // If we have action keywords
        if (actionValues.length > 0) {
            // FIRST: Check unique value count - if 3+, always use overflow
            if (uniqueValues >= 3) {
                console.log(`  ✅ Detected overflow actions (${uniqueValues} unique values)`);
                return 'action';
            }
            
            // For 1-2 unique values, check if cells contain combined actions
            // (e.g., "Edit/Delete", "Edit,Delete", "Edit & Delete")
            const hasCombinedAction = nonEmptyData.some(value => {
                const lowerVal = value.toLowerCase();
                const hasEdit = lowerVal.includes('edit') || lowerVal.includes('update') || lowerVal.includes('modify');
                const hasDelete = lowerVal.includes('delete') || lowerVal.includes('remove') || lowerVal.includes('trash');
                return hasEdit && hasDelete;
            });
            
            if (hasCombinedAction && uniqueValues === 1) {
                console.log(`  ✅ Detected combined edit/delete actions in single cell (e.g., "Edit/Delete")`);
                return 'editDelete';
            }
            
            // Single action type: use direct icon swap (edit OR delete only)
            if (uniqueValues === 1) {
                const firstValue = nonEmptyData[0].toLowerCase();
                const hasEdit = firstValue.includes('edit') || firstValue.includes('update') || firstValue.includes('modify');
                const hasDelete = firstValue.includes('delete') || firstValue.includes('remove') || firstValue.includes('trash');
                
                if (hasEdit) {
                    console.log(`  ✅ Detected single EDIT action`);
                    return 'singleEdit';
                } else if (hasDelete) {
                    console.log(`  ✅ Detected single DELETE action`);
                    return 'singleDelete';
                }
            }
            
            // Two different actions: check if they should use slot group or overflow
            if (uniqueValues === 2) {
                // If both values are combinations or if it's edit+delete, use slot group
                if (hasCombinedAction) {
                    console.log(`  ✅ Detected combined edit/delete actions (2 unique values with combinations)`);
                    return 'editDelete';
                }
                // Otherwise use overflow for mixed actions
                console.log(`  ✅ Detected edit/delete actions (2 unique values)`);
                return 'editDelete';
            }
        }
        
        // For 3+ different actions, use overflow
        console.log(`  ✅ Detected overflow actions (${uniqueValues} unique values)`);
        return 'action';
    }
    
    // Status detection - Carbon Design System status icon values + common status keywords
    const carbonStatusValues = ['failed', 'succeeded', 'pending', 'in-progress', 'not started', 'incomplete', 'unknown', 'normal', 'informative'];
    const statusKeywords = [
        ...carbonStatusValues, 
        'active', 'inactive', 
        'approved', 'rejected', 
        'accepted', 'declined',  // Added: Common approval/rejection status
        'completed', 'cancelled', 
        'draft', 'published', 
        'archived', 'enabled', 'disabled', 
        'success', 'error', 'warning',
        'open', 'closed',  // Added: Common ticket/issue status
        'confirmed', 'denied',  // Added: Common confirmation status
        'resolved', 'unresolved'  // Added: Common resolution status
    ];
    const statusMatch = nonEmptyData.filter(val => 
        statusKeywords.some(keyword => val.toLowerCase().includes(keyword.toLowerCase()))
    ).length;
    if (statusMatch / nonEmptyData.length > 0.5) return 'status';
    
    // Name detection - only for actual name columns (not emails, usernames, etc.)
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const namePattern = /^[A-Z][a-z]+ [A-Z][a-z]+$/; // First Last
    const usernamePattern = /^[a-zA-Z0-9_]+$/; // Username pattern
    
    // Check if this is actually a name column (not email, username, etc.)
    const isNameColumn = columnName.toLowerCase().includes('name') && 
                        !columnName.toLowerCase().includes('user') && 
                        !columnName.toLowerCase().includes('email');
    
    const emailMatch = nonEmptyData.filter(val => emailPattern.test(val)).length;
    const nameMatch = nonEmptyData.filter(val => namePattern.test(val)).length;
    const usernameMatch = nonEmptyData.filter(val => usernamePattern.test(val) && val.length <= 20).length;
    
    // Only suggest slotGroup for actual name columns with proper name patterns
    if (isNameColumn && nameMatch / nonEmptyData.length > 0.6) {
        return 'user';
    }
    
    // If it's an email column, don't suggest slotGroup
    if (emailMatch / nonEmptyData.length > 0.5) {
        return 'text'; // Treat emails as regular text
    }
    
    // If it's a username column, don't suggest slotGroup
    if (usernameMatch / nonEmptyData.length > 0.7) {
        return 'text'; // Treat usernames as regular text
    }
    
    // Boolean detection
    const booleanValues = ['true', 'false', 'yes', 'no', '0', '1'];
    const booleanMatch = nonEmptyData.filter(val => 
        booleanValues.includes(val.toLowerCase())
    ).length;
    if (booleanMatch / nonEmptyData.length > 0.7) return 'boolean';
    
    // URL detection
    const urlPattern = /^https?:\/\//;
    const urlMatch = nonEmptyData.filter(val => urlPattern.test(val)).length;
    if (urlMatch / nonEmptyData.length > 0.5) return 'url';
    
    // Date detection
    const datePattern = /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/;
    const dateMatch = nonEmptyData.filter(val => datePattern.test(val)).length;
    if (dateMatch / nonEmptyData.length > 0.5) return 'date';
    
    // Number detection
    const numberPattern = /^\d+(\.\d+)?$/;
    const numberMatch = nonEmptyData.filter(val => numberPattern.test(val)).length;
    if (numberMatch / nonEmptyData.length > 0.7) return 'number';
    
    // Tag detection - strict criteria for multiple categories that need differentiation
    const hasCommas = nonEmptyData.filter(val => val.includes(',')).length;
    const hasMultipleSeparators = nonEmptyData.filter(val => /[,;|&\/]/.test(val)).length;
    
    // Strong tag indicators - column names that clearly indicate categorization
    const strongTagKeywords = ['tag', 'category', 'type', 'label', 'status', 'priority', 'level', 'grade', 'class', 'role', 'tier', 'rank', 'group', 'department', 'team'];
    const isStrongTagColumn = strongTagKeywords.some(keyword => columnName.toLowerCase().includes(keyword));
    
    // Check for multiple distinct values that represent categories
    const uniqueValues = new Set(nonEmptyData.map(val => val.trim().toLowerCase()));
    const hasMultipleCategories = uniqueValues.size >= 3 && uniqueValues.size <= 10; // 3-10 distinct categories
    
    // Check for comma-separated values (multiple tags in one cell)
    const hasCommaSeparatedTags = hasCommas / nonEmptyData.length > 0.4; // Higher threshold: 40% of values have commas
    
    // Check for phone numbers - exclude from tag detection
    const phonePattern = /^[\d\s()+\-\.]+$/; // Matches phone number patterns
    const phoneMatch = nonEmptyData.filter(val => {
        const trimmed = val.trim();
        // Check if it looks like a phone number (mostly digits with some formatting)
        const digitCount = (trimmed.match(/\d/g) || []).length;
        return phonePattern.test(trimmed) && digitCount >= 7; // At least 7 digits for a phone number
    }).length;
    
    // If majority are phone numbers, don't suggest tags
    if (phoneMatch / nonEmptyData.length > 0.5) {
        return 'text';
    }
    
    // Check for values that look like categories (short, distinct, non-numeric)
    const categoryLikeValues = nonEmptyData.filter(val => {
        const trimmed = val.trim();
        const words = trimmed.split(/\s+/);
        return words.length <= 2 && 
               trimmed.length <= 15 && 
               !/^\d+$/.test(trimmed) && // Not just numbers
               !/^[a-z]+@/.test(trimmed.toLowerCase()) && // Not email-like
               !phonePattern.test(trimmed); // Not phone-number-like
    }).length;
    
    const hasCategoryLikeValues = categoryLikeValues / nonEmptyData.length > 0.6; // 60% of values look like categories
    
    // Only suggest tags if we have strong indicators of categorization
    if ((isStrongTagColumn && hasMultipleCategories) || 
        (hasCommaSeparatedTags && hasMultipleCategories) ||
        (hasCategoryLikeValues && hasMultipleCategories && uniqueValues.size >= 4)) {
        return 'label';
    }
    
    return 'text';
}

function suggestSlotComponent(contentType: SmartSlotSuggestion['contentType']): SmartSlotSuggestion['suggestedComponent'] {
    const componentMap: Record<SmartSlotSuggestion['contentType'], SmartSlotSuggestion['suggestedComponent']> = {
        'status': 'statusIcon', // Status Icon component for status values (Failed, Succeeded, etc.)
        'user': 'slotGroup', // Slot group with Avatar + Text for user names
        'action': 'overflow', // Overflow Menu for 3+ actions
        'editDelete': 'slotGroup', // Slot group with Edit + Delete icons for 2 actions
        'singleEdit': 'edit', // Single Edit icon for single edit action
        'singleDelete': 'delete', // Single Delete icon for single delete action
        'boolean': 'checkbox',
        'url': 'link',
        'label': 'tag', // Tag component for labels, roles, types, comma-separated values
        'date': null,
        'number': null,
        'text': null
    };
    return componentMap[contentType];
}

function analyzeTableDataForSmartSlots(gridData: string[][], headers: string[]): SmartSlotSuggestion[] {
    console.log('🔍 Analyzing table data for smart slot suggestions...');
    const suggestions: SmartSlotSuggestion[] = [];
    
    // Validate input
    if (!gridData || !Array.isArray(gridData) || gridData.length === 0) {
        console.log('⚠️ No grid data to analyze');
        return suggestions;
    }
    
    const numColumns = headers.length;
    
    for (let colIndex = 0; colIndex < numColumns; colIndex++) {
        // Convert all values to strings and filter out empty ones
        const columnData = gridData.map(row => {
            const value = row[colIndex];
            // Convert any value (including booleans) to string
            return value != null ? String(value) : '';
        });
        const columnName = headers[colIndex] || `Column ${colIndex + 1}`;
        
        const contentType = analyzeColumnContent(columnData, columnName);
        const suggestedComponent = suggestSlotComponent(contentType);
        
        if (suggestedComponent) {
            const samples = columnData.filter(val => val && val.trim() !== '' && val !== 'true' && val !== 'false').slice(0, 3);
            
            // Only add suggestion if we have actual text samples (not just 'true'/'false')
            if (samples.length > 0) {
                suggestions.push({
                    columnIndex: colIndex,
                    columnName,
                    contentType,
                    suggestedComponent,
                    confidence: 0.8, // Could be calculated based on matching percentage
                    samples
                });
                
                console.log(`💡 Column "${columnName}" (${colIndex}): ${contentType} → suggest ${suggestedComponent}`);
                console.log(`   Samples: ${samples.join(', ')}`);
            }
        }
    }
    
    console.log(`✅ Found ${suggestions.length} smart slot suggestions`);
    return suggestions;
}

// Function to discover all available components from libraries for swap slot
async function discoverAvailableComponents(): Promise<ComponentNode[]> {
    console.log('🔍 Discovering available components from libraries...');
    const allComponents: ComponentNode[] = [];
    
    try {
        // First, find all components that are already in the file (including library components that have been used)
        const fileComponents = figma.root.findAll(node => node.type === 'COMPONENT') as ComponentNode[];
        console.log(`📦 Found ${fileComponents.length} components in file`);
        
        // Also find components from component sets that are already in the file
        const componentSets = figma.root.findAll(node => node.type === 'COMPONENT_SET') as ComponentSetNode[];
        console.log(`📦 Found ${componentSets.length} component sets in file`);
        
        // Add all components from component sets
        for (const componentSet of componentSets) {
            const variants = componentSet.children.filter(child => child.type === 'COMPONENT') as ComponentNode[];
            allComponents.push(...variants);
            console.log(`📦 Added ${variants.length} variants from component set: ${componentSet.name}`);
        }
        
        // Add standalone components
        allComponents.push(...fileComponents);
        
        // Remove duplicates based on ID
        const uniqueComponents = allComponents.filter((component, index, self) => 
            index === self.findIndex(c => c.id === component.id)
        );
        
        console.log(`✅ Total unique components discovered: ${uniqueComponents.length}`);
        
        // Categorize components by source
        const pageComponents = figma.currentPage.findAll(node => node.type === 'COMPONENT') as ComponentNode[];
        const pageComponentIds = new Set(pageComponents.map(c => c.id));
        const libraryComponents = uniqueComponents.filter(c => !pageComponentIds.has(c.id));
        
        console.log(`📋 Components breakdown:`);
        console.log(`  - On current page: ${pageComponents.length}`);
        console.log(`  - From libraries: ${libraryComponents.length}`);
        
        // Log some examples from libraries
        if (libraryComponents.length > 0) {
            console.log(`📋 Sample library components:`);
            libraryComponents.slice(0, 10).forEach((comp, index) => {
                console.log(`  ${index + 1}. ${comp.name} (ID: ${comp.id}, Key: ${comp.key})`);
            });
        }
        
        // Try to find Overflow component using library component discovery
        console.log('🔍 Searching for Overflow component from library components...');
        try {
            // Method 1: Try to find from existing component sets on pages
            const componentSets = figma.root.findAll(node => node.type === 'COMPONENT_SET') as ComponentSetNode[];
            console.log(`📦 Found ${componentSets.length} component sets already on pages`);
            
            for (const componentSet of componentSets) {
                const setName = componentSet.name.toLowerCase();
                if (setName.includes('overflow') || setName.includes('menu') || setName.includes('action')) {
                    console.log(`🎯 Found potential Overflow component set on page: ${componentSet.name}`);
                    
                    const existingVariants = componentSet.children.filter(child => child.type === 'COMPONENT') as ComponentNode[];
                    if (existingVariants.length > 0) {
                        console.log(`✅ Component set already accessible with ${existingVariants.length} variants`);
                        for (const variant of existingVariants) {
                            if (!uniqueComponents.find(c => c.id === variant.id)) {
                                uniqueComponents.push(variant);
                                console.log(`✅ Added existing variant: ${variant.name}`);
                            }
                        }
                    }
                }
            }
            
            // Method 2: Try to discover from library components using known Carbon keys
            console.log('🔍 Attempting to discover Overflow from Carbon library components...');
            const carbonOverflowKeys = [
                // These are common Carbon Overflow component keys - we'll try to import them
                'overflow-menu-key-1', // Placeholder - we'll need real keys
                'overflow-menu-key-2', // Placeholder - we'll need real keys
            ];
            
            // Get any collected keys from previous sessions
            let collectedKeys = [];
            try {
                const existingKeys = figma.currentPage.getPluginData('collectedCarbonKeys');
                if (existingKeys) {
                    collectedKeys = JSON.parse(existingKeys);
                    console.log(`📦 Found ${collectedKeys.length} previously collected Carbon keys`);
                }
            } catch (e) {
                console.log('No previously collected Carbon keys found');
            }
            
            // Try to import components using collected keys
            for (const keyInfo of collectedKeys) {
                try {
                    console.log(`🔄 Trying to import component with key: ${keyInfo.key}`);
                    const importedComponent = await figma.importComponentByKeyAsync(keyInfo.key);
                    if (importedComponent) {
                        // Check if this is an Overflow component
                        const componentName = importedComponent.name.toLowerCase();
                        if (componentName.includes('overflow') || componentName.includes('menu')) {
                            console.log(`✅ Found Overflow component from library: ${importedComponent.name}`);
                            if (!uniqueComponents.find(c => c.id === importedComponent.id)) {
                                uniqueComponents.push(importedComponent);
                                
                                // Store this Overflow component for future use
                                const overflowInfo = {
                                    name: importedComponent.name,
                                    key: keyInfo.key,
                                    id: importedComponent.id,
                                    description: importedComponent.description || '',
                                    width: importedComponent.width,
                                    height: importedComponent.height,
                                    timestamp: Date.now()
                                };
                                figma.currentPage.setPluginData('foundOverflowComponent', JSON.stringify(overflowInfo));
                            }
                        } else if (!uniqueComponents.find(c => c.id === importedComponent.id)) {
                            // Add any other useful component
                            uniqueComponents.push(importedComponent);
                            console.log(`✅ Added library component: ${importedComponent.name}`);
                        }
                    }
                } catch (importError) {
                    console.log(`ℹ️ Could not import component with key ${keyInfo.key}: ${importError instanceof Error ? importError.message : 'Unknown error'}`);
                }
            }
            
        } catch (searchError) {
            console.log('ℹ️ Library component search completed with errors:', searchError);
        }

        // Auto-discover common Carbon table components
        console.log('🔍 Auto-discovering common Carbon table components...');
        try {
            // Get collected Carbon keys from plugin data
            let collectedKeys = [];
            try {
                const existingKeys = figma.currentPage.getPluginData('collectedCarbonKeys');
                if (existingKeys) {
                    collectedKeys = JSON.parse(existingKeys);
                    console.log(`📦 Found ${collectedKeys.length} collected Carbon component keys`);
                }
            } catch (e) {
                console.log('No collected Carbon keys found');
            }
            
            // ============================================================
            // HARDCODED CARBON DESIGN SYSTEM COMPONENT KEYS
            // ============================================================
            // These keys are STABLE across all files that have the Carbon library enabled.
            // You can hardcode keys here and they will work for ALL USERS!
            // 
            // To add more keys:
            // 1. Select the component in the Carbon library file
            // 2. Run in Figma console: figma.currentPage.selection[0].key
            // 3. Add the key here with a descriptive comment
            // ============================================================
            // ============================================================
            // SMART SLOT COMPONENTS - For automatic slot detection
            // ============================================================
            // These components will be automatically added to cells based on content type
            // Get keys by selecting the component instance and clicking "Get Component Keys"
            const smartSlotComponents = {
                overflow: '608ab39a49081fa1a460e71af33de459e6154ee7', // Overflow Menu - for 3+ action buttons
                statusIcon: 'd94da074d575aa0ae337fc054f4f19bc311049b8', // Status icon - for status values (Failed, Succeeded, Pending, In-progress, etc.)
                slotGroup: '761fd2270cc978c88b9cce64926b4c8245831bdc', // Slot group (Direction=Horizontal) - for name columns with avatar + text, or action columns with edit + delete
                avatar: 'd80f0d175851756c4601e87e6e0abeb5539b24e8', // Avatar - for user profile images
                text: 'e73c62eb16dcb7f54df1384f176fc7a3c0f64df5', // Text/Label - for user names
                tag: 'c95a2fb5332d75515d2ed90ac487bc864fcb09b5', // Tag - for labels, roles, types, comma-separated values
                edit: 'a4ba4c4aa1f2b0f0a5206341aafbb7d7eafa47e6', // Edit icon - for edit actions
                delete: '84a7c6755b83b8e88ca803851c280d1e06255b93', // Trash-can icon - for delete actions
                // Add more component keys here:
                // checkbox: 'CHECKBOX_KEY_HERE', // Checkbox - for boolean values
                // link: 'LINK_KEY_HERE', // Link - for URLs
            };
            
            const knownCarbonKeys = [
                'eb0c6fa56e00a94f56d7f6f1cfe9cde4006c5c20', // Carbon Table (Expandable + Selectable)
                smartSlotComponents.overflow, // Overflow Menu
                smartSlotComponents.statusIcon, // Status icon
                // Add all smart slot component keys to the import list
            ];
            
            // Combine hardcoded keys with user-collected keys
            const allKeysToTry = [
                ...knownCarbonKeys,
                ...collectedKeys.map((k: any) => k.key)
            ];
            
            console.log(`🧪 Testing ${allKeysToTry.length} potential Carbon component keys...`);
            
            for (const key of allKeysToTry) {
                try {
                    const importedComponent = await figma.importComponentByKeyAsync(key);
                    if (importedComponent && !uniqueComponents.find(c => c.id === importedComponent.id)) {
                        uniqueComponents.push(importedComponent);
                        console.log(`✅ Auto-discovered: ${importedComponent.name} (Key: ${key})`);
                        
                        // Check if this is the Overflow component we're looking for
                        if (importedComponent.name.toLowerCase().includes('overflow')) {
                            console.log(`🎯 Found Overflow component: ${importedComponent.name}`);
                        }
                        
                        // Check for other useful table components
                        const componentName = importedComponent.name.toLowerCase();
                        if (componentName.includes('button') || 
                            componentName.includes('action') || 
                            componentName.includes('menu') ||
                            componentName.includes('dropdown') ||
                            componentName.includes('icon')) {
                            console.log(`🎯 Found useful table component: ${importedComponent.name}`);
                        }
                    }
                } catch (importError) {
                    // Component with this key doesn't exist or isn't accessible - this is normal
                }
            }
        } catch (testError) {
            console.log('ℹ️ Auto-discovery completed');
        }
        
        return uniqueComponents;
        
    } catch (error) {
        console.error('❌ Error discovering components:', error);
        return [];
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
            // When using proxy, API key is not required (it's in environment variables)
            if (!endpointToUse || (!useAccessToken && !useProxy && !apiKeyToUse && !uiAccessToken)) {
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
                        body: JSON.stringify({}) // API key is now in environment variables
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
            const { prompt, rows, cols } = msg as { prompt: string, rows: number, cols: number };
            const proxyUrl = 'https://application-e9.21hwt6k1vujm.us-east.codeengine.appdomain.cloud';
            const endpoint = 'https://us-south.ml.cloud.ibm.com';

            const tokenRes = await fetch(`${proxyUrl}/token`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({}) // API key is now in environment variables
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
        
        // If we have a table ID, try to load metadata and body cell component
        if (tableId) {
            const tableNode = figma.getNodeById(tableId);
            if (tableNode) {
                // Load saved metadata
                const tableSettingsData = tableNode.getPluginData('tableSettings');
                let savedCellProperties = null;
                let actualRows: number | undefined;
                let actualCols: number | undefined;
                let bodyCellComponentId: string | null = null;
                
                if (tableSettingsData) {
                    try {
                        const tableSettings = JSON.parse(tableSettingsData);
                        savedCellProperties = tableSettings.cellProperties;
                        actualRows = tableSettings.rows;
                        actualCols = tableSettings.columns;
                        bodyCellComponentId = tableSettings.bodyCellComponentId;
                        console.log('[Backend] Loaded metadata: bodyCellComponentId =', bodyCellComponentId);
                    } catch (e) {
                        console.warn('[Backend] Could not parse table settings:', e);
                    }
                }
                
                // Try to get body cell component from saved ID
                let bodyCell: ComponentNode | null = null;
                if (bodyCellComponentId) {
                    const node = figma.getNodeById(bodyCellComponentId);
                    if (node && node.type === 'COMPONENT') {
                        bodyCell = node as ComponentNode;
                        console.log('[Backend] Found body cell component from saved ID:', bodyCell.name);
                    }
                }
                
                // If we have the body cell component, create a temp instance to get properties
                if (bodyCell) {
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
                        },
                        savedCellProperties: savedCellProperties,
                        actualRows: actualRows,
                        actualCols: actualCols
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
        
        // No component info available - send error message
        console.error('[Backend] ❌ No component info available - neither table body cell instance nor lastScanResult found');
        figma.ui.postMessage({
            type: "component-info-unavailable",
            message: "No component information available. Please scan a Carbon Data Table component first."
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
        
        // --- Set the 'type' property to 'Expandable + Selectable' and Toolbar to enable all necessary components ---
        console.log(`[scan-table] Setting table properties to enable all necessary components`);
        
        if ('componentProperties' in tableFrame && 'setProperties' in tableFrame && typeof tableFrame.setProperties === 'function') {
          console.log(`[scan-table] Table has componentProperties, attempting to set properties...`);
          
          // Log all available properties for debugging
          console.log(`[scan-table] 📋 Available properties:`, Object.keys(tableFrame.componentProperties));
          for (const [propName, prop] of Object.entries(tableFrame.componentProperties)) {
            console.log(`[scan-table]   "${propName}" = "${(prop as any).value}" (type: ${(prop as any).type})`);
          }
          
          try {
            // Try to set the 'type' property to 'Expandable + Selectable'
            const propertiesToSet: any = {};
            propertiesToSet['type'] = 'Expandable + Selectable';
            
            console.log(`[scan-table] Attempting to set Type property:`, propertiesToSet);
            tableFrame.setProperties(propertiesToSet);
            console.log(`[scan-table] ✅ Successfully set Type to 'Expandable + Selectable'`);
            
          } catch (error) {
            console.log(`[scan-table] Direct type setting failed, trying alternative approach...`);
            
            // Try alternative property names and values
            const alternativeNames = ['Type', 'Variant', 'Style', 'Mode'];
            const alternativeValues = [
              'Expandable + Selectable',
              'Expandable+Selectable', 
              'Selectable + Expandable',
              'Both'
            ];
            
            let success = false;
            for (const propName of alternativeNames) {
              for (const value of alternativeValues) {
                try {
                  const altProperties: any = {};
                  altProperties[propName] = value;
                  console.log(`[scan-table] Trying ${propName} = ${value}`);
                  tableFrame.setProperties(altProperties);
                  console.log(`[scan-table] ✅ Successfully set ${propName} to '${value}'`);
                  success = true;
                  break;
                } catch (altError) {
                  // Try next combination
                }
              }
              if (success) break;
            }
            
            if (!success) {
              console.warn(`[scan-table] Could not set table type property`);
            }
          }
          
          // Set Toolbar property to enable toolbar elements
          try {
            // Find toolbar property (might be Toolbar#62100:0 or similar)
            for (const [propName, prop] of Object.entries(tableFrame.componentProperties)) {
              if (propName.toLowerCase().startsWith('toolbar')) {
                const toolbarProps: any = {};
                if ((prop as any).type === 'BOOLEAN') {
                  toolbarProps[propName] = true;
                  console.log(`[scan-table] Setting BOOLEAN ${propName} to true`);
                } else {
                  toolbarProps[propName] = 'True';
                  console.log(`[scan-table] Setting ${propName} to 'True'`);
                }
                tableFrame.setProperties(toolbarProps);
                console.log(`[scan-table] ✅ Successfully enabled Toolbar property`);
                break;
              }
            }
          } catch (toolbarError) {
            console.log(`[scan-table] ℹ️ Toolbar property not available (this is OK)`);
          }
          
          // Wait for Figma to update the instance tree
          await Promise.resolve();
          await new Promise(resolve => setTimeout(resolve, 150));
          
        } else {
          console.log(`[scan-table] Table does not have componentProperties or setProperties method`);
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
            let toolbarComponent: ComponentNode | null = null;
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
                    // Check for toolbar components
                    if (node.name && node.name.toLowerCase().includes('toolbar')) {
                        console.log(`🎯 Found toolbar node in second scan: ${node.name} (type: ${node.type})`);
                        if (node.type === 'INSTANCE' && node.mainComponent && !toolbarComponent) {
                            toolbarComponent = node.mainComponent;
                            console.log(`✅ Found toolbar component in second scan: ${toolbarComponent.name}`);
                        }
                    }
                    
                    if (node.name === 'Data table expand cell item') {
                      expandCellFound = true;
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        expandCellComponent = node.mainComponent;
                        console.log('✅ Found Data table expand cell item from instance:', expandCellComponent.name);
                      } else if (node.type === 'COMPONENT') {
                        expandCellComponent = node;
                        console.log('✅ Found Data table expand cell item component directly:', expandCellComponent.name);
                      }
                    }
                    // Check for select cell components - only exact name match
                    if (node.name === 'Data table select cell item' && !selectCellFound) {
                      selectCellFound = true;
                      console.log('🔍 Found Data table select cell item:', node.name, node.type);
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        selectCellComponent = node.mainComponent;
                        console.log('✅ Using Data table select cell item from instance:', selectCellComponent.name);
                      } else if (node.type === 'COMPONENT') {
                        selectCellComponent = node;
                        console.log('✅ Using Data table select cell item component directly:', selectCellComponent.name);
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
                toolbar: toolbarComponent,
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
                        toolbarId: toolbarComponent ? (toolbarComponent as any).id : null,
                        selectCellId: selectCellComponent?.id || null,
                        expandCellId: expandCellComponent?.id || null,
                        dividerId: dividerComponent?.id || null,
                        numCols,
                        timestamp: Date.now()
                    };
                    tableFrame.setPluginData('tableGeneratorScan', JSON.stringify(scanData));
                    console.log('✅ Stored scan result in table frame plugin data');
            console.log('📊 Scan result summary:', {
                headerCell: (headerCell as any)?.name || 'Not found',
                bodyCell: (bodyCell as any)?.name || 'Not found', 
                footer: (footer as any)?.name || 'Not found',
                toolbar: (toolbarComponent as any)?.name || 'Not found',
                selectCell: (selectCellComponent as any)?.name || 'Not found',
                expandCell: (expandCellComponent as any)?.name || 'Not found',
                divider: (dividerComponent as any)?.name || 'Not found'
            });
            
            // Debug: Search specifically for toolbar components
            console.log('🔍 Searching specifically for toolbar components...');
            const allNodes = tableFrame.findAll(n => true);
            console.log(`🔍 Found ${allNodes.length} total nodes in table`);
            
            allNodes.forEach(node => {
                if (node.name && node.name.toLowerCase().includes('toolbar')) {
                    console.log(`🎯 FOUND TOOLBAR NODE: ${node.name} (type: ${node.type})`);
                    if (node.type === 'INSTANCE' && (node as InstanceNode).mainComponent) {
                        console.log(`🎯 TOOLBAR INSTANCE: ${node.name} -> ${(node as InstanceNode).mainComponent?.name}`);
                    }
                }
            });
                }
            } catch (error) {
                console.log('⚠️ Could not store scan result in plugin data:', error);
            }
       
            // Use saved properties instead of extracting from table
            console.log(`[scan-table] Using saved properties for table: ${tableFrame.name}`);
            const actualTableData = null; // Force use of saved properties
            
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
                    bodyRowComponent: bodyRowComponent,
                    actualTableData: actualTableData
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

            const { headerCell, bodyCell, footer, toolbar, numCols } = lastScanResult;
            console.log('🔍 Using scan result for table generation:', {
                headerCell: headerCell?.name || 'Not found',
                bodyCell: bodyCell?.name || 'Not found',
                footer: footer?.name || 'Not found', 
                toolbar: toolbar?.name || 'Not found',
                numCols
            });
            figma.ui.postMessage({ type: 'show-loader', message: 'Generating table...' });

            const includeHeader = msg.includeHeader !== false;
            const includeFooter = msg.includeFooter === true;
            const includeToolbar = msg.includeToolbar === true;
            const includeSelectable = msg.includeSelectable === true;
            const includeExpandable = msg.includeExpandable === true;
            
            console.log('🔍 Table generation flags:', {
                includeHeader,
                includeFooter,
                includeToolbar,
                includeSelectable,
                includeExpandable,
                msgIncludeToolbar: msg.includeToolbar
            });
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
            headerRowFrame.counterAxisAlignItems = "CENTER"; // Vertically center header content
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
                rowFrame.counterAxisAlignItems = "CENTER"; // Vertically center cell content
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
                                    // Check if slot is enabled using the dedicated slot variable
                                    const isSlotEnabled = cellData.slot === true;
                                    console.log(`🔄 Cell ${key} slot state: ${isSlotEnabled} (from cellData.slot)`);
                                    
                                    // Filter out slot properties if slot is disabled
                                    let propertiesToApply = { ...cellData.properties };
                                    if (!isSlotEnabled) {
                                        // Remove slot-related properties if slot is disabled
                                        const slotProps = Object.keys(propertiesToApply).filter(prop => 
                                            prop.toLowerCase().includes('slot') || 
                                            prop.toLowerCase().includes('swap')
                                        );
                                        slotProps.forEach(prop => {
                                            delete propertiesToApply[prop];
                                            console.log(`🔄 Removed slot property ${prop} for cell ${key} (slot disabled)`);
                                        });
                                    }
                                    
                                    const validProps = mapPropertyNames(propertiesToApply, cell.componentProperties);
                                    
                                    // Log if this cell has a Swap slot property
                                    const swapSlotKey = Object.keys(validProps).find(key => key.toLowerCase().includes('swap') && key.toLowerCase().includes('slot'));
                                    if (swapSlotKey) {
                                        console.log(`🔄 Cell ${key} has Swap slot: ${swapSlotKey} = ${validProps[swapSlotKey]}`);
                                    }
                                    
                                    cell.setProperties(validProps);
                                    
                                    // If this cell has slot component properties (e.g., for Status Icon label)
                                    // Only apply slot component properties if slot is enabled
                                    const hasSwapSlot = Object.keys(validProps).some(key => key.toLowerCase().includes('swap') && key.toLowerCase().includes('slot'));
                                    const swapComponentId = swapSlotKey ? validProps[swapSlotKey] : null;
                                    
                                    if (cellData.slotComponentProps && isSlotEnabled && hasSwapSlot && swapComponentId) {
                                        // Store the cell reference and config for the async operation
                                        const cellRef = cell;
                                        const slotPropsConfig = cellData.slotComponentProps;
                                        const cellKeyForLog = key;
                                        
                                        // Check if this is a nested slot case (Slot group with Avatar + Text)
                                        const isNestedSlot = slotPropsConfig.nestedSlots === true;
                                        
                                        // Use setTimeout to allow Figma to complete the component swap
                                        setTimeout(async () => {
                                            try {
                                                console.log(`🔧 Configuring slot component for cell ${cellKeyForLog}:`, slotPropsConfig);
                                                
                                                if (isNestedSlot) {
                                                    // Handle Slot Group with Avatar + Text OR Edit + Delete
                                                    let avatarComponent: ComponentNode | null = null;
                                                    let textComponent: ComponentNode | null = null;
                                                    let editComponent: ComponentNode | null = null;
                                                    let deleteComponent: ComponentNode | null = null;
                                                    
                                                    if (slotPropsConfig.userName) {
                                                        // User name case: Avatar + Text
                                                        console.log(`👥 Processing nested slots for user: ${slotPropsConfig.userName}`);
                                                        
                                                        // Import the Avatar and Text components
                                                        avatarComponent = await figma.importComponentByKeyAsync(slotPropsConfig.avatarKey);
                                                        textComponent = await figma.importComponentByKeyAsync(slotPropsConfig.textKey);
                                                        console.log(`  ✅ Imported Avatar and Text components`);
                                                    } else if (slotPropsConfig.editKey && slotPropsConfig.deleteKey) {
                                                        // Edit/Delete case: Edit + Delete icons
                                                        console.log(`⚡ Processing nested slots for edit/delete actions`);
                                                        
                                                        // Import the Edit and Delete components
                                                        editComponent = await figma.importComponentByKeyAsync(slotPropsConfig.editKey);
                                                        deleteComponent = await figma.importComponentByKeyAsync(slotPropsConfig.deleteKey);
                                                        console.log(`  ✅ Imported Edit and Delete components`);
                                                    }
                                                    
                                                    // Find the Slot group component that was just swapped in
                                                    const slotGroupComponents = cellRef.findAll(node => {
                                                        if (node.type === 'INSTANCE') {
                                                            const mainComp = (node as InstanceNode).mainComponent;
                                                            return mainComp !== null && mainComp.id === swapComponentId;
                                                        }
                                                        return false;
                                                    }) as InstanceNode[];
                                                    
                                                    if (slotGroupComponents.length > 0) {
                                                        const slotGroup = slotGroupComponents[0];
                                                        console.log(`  📦 Found Slot Group: ${slotGroup.name}`);
                                                        console.log(`  📋 Slot Group children count:`, slotGroup.children.length);
                                                        
                                                        // Find the actual slot instances inside the Slot Group
                                                        // Slots are children of the Slot Group, not properties
                                                        const slotInstances = slotGroup.children.filter(child => 
                                                            child.type === 'INSTANCE' && 
                                                            child.name.toLowerCase().includes('slot')
                                                        ) as InstanceNode[];
                                                        
                                                        console.log(`  🎰 Found ${slotInstances.length} slot instances:`, slotInstances.map(s => s.name));
                                                        
                                                        if (slotInstances.length >= 2) {
                                                            const firstSlot = slotInstances[0];
                                                            const secondSlot = slotInstances[1];
                                                            
                                                            if (slotPropsConfig.userName && avatarComponent && textComponent) {
                                                                // User name case: Avatar + Text
                                                                console.log(`  🔄 Swapping ${firstSlot.name} with Avatar...`);
                                                                console.log(`  🔄 Swapping ${secondSlot.name} with Text...`);
                                                                console.log(`  📏 Before swap - firstSlot layoutSizingHorizontal:`, firstSlot.layoutSizingHorizontal);
                                                                console.log(`  📏 Before swap - slotGroup layoutMode:`, slotGroup.layoutMode);
                                                                
                                                                firstSlot.swapComponent(avatarComponent);
                                                                secondSlot.swapComponent(textComponent);
                                                                
                                                                console.log(`  ✅ Nested slots swapped for user!`);
                                                            } else if (slotPropsConfig.editKey && slotPropsConfig.deleteKey && editComponent && deleteComponent) {
                                                                // Edit/Delete case: Edit + Delete icons
                                                                console.log(`  🔄 Swapping ${firstSlot.name} with Edit...`);
                                                                console.log(`  🔄 Swapping ${secondSlot.name} with Delete...`);
                                                                
                                                                firstSlot.swapComponent(editComponent);
                                                                secondSlot.swapComponent(deleteComponent);
                                                                
                                                                console.log(`  ✅ Nested slots swapped for edit/delete!`);
                                                            }
                                                            
                                                            // Now configure components (only for user names)
                                                            if (slotPropsConfig.userName && avatarComponent && textComponent) {
                                                                setTimeout(() => {
                                                                    // Configure Avatar component
                                                                    const avatarInstances = slotGroup.findAll(node => {
                                                                        if (node.type === 'INSTANCE') {
                                                                            const mainComp = (node as InstanceNode).mainComponent;
                                                                            return mainComp !== null && mainComp.id === avatarComponent.id;
                                                                        }
                                                                        return false;
                                                                    }) as InstanceNode[];
                                                                
                                                                if (avatarInstances.length > 0) {
                                                                    const avatarInstance = avatarInstances[0];
                                                                    console.log(`  👤 Found Avatar instance, configuring...`);
                                                                    console.log(`  📋 Avatar properties:`, Object.keys(avatarInstance.componentProperties));
                                                                    
                                                                    // Extract proper initials from the name
                                                                    const userName = slotPropsConfig.userName || '';
                                                                    const initials = extractInitials(userName);
                                                                    
                                                                    // Find the properties
                                                                    const typeProp = Object.keys(avatarInstance.componentProperties).find(prop => 
                                                                        prop.toLowerCase().includes('type')
                                                                    );
                                                                    const sizeProp = Object.keys(avatarInstance.componentProperties).find(prop => 
                                                                        prop.toLowerCase().includes('size')
                                                                    );
                                                                    const initialsProp = Object.keys(avatarInstance.componentProperties).find(prop => 
                                                                        prop.toLowerCase().includes('initial')
                                                                    );
                                                                    
                                                                    const avatarProps: any = {};
                                                                    if (typeProp) avatarProps[typeProp] = 'Initials';
                                                                    if (sizeProp) avatarProps[sizeProp] = 'Medium';
                                                                    if (initialsProp) avatarProps[initialsProp] = initials;
                                                                    
                                                                    if (Object.keys(avatarProps).length > 0) {
                                                                        avatarInstance.setProperties(avatarProps);
                                                                        console.log(`  ✅ Set Avatar properties:`, avatarProps);
                                                                    }
                                                                    
                                                                    // Set Avatar layout properties to HUG
                                                                    avatarInstance.layoutSizingHorizontal = 'HUG';
                                                                    avatarInstance.layoutSizingVertical = 'HUG';
                                                                    avatarInstance.layoutGrow = 0;
                                                                    console.log(`  📏 Set Avatar layout to HUG`);
                                                                    
                                                                    // Set fixed dimensions for Avatar (36px x 36px)
                                                                    avatarInstance.resize(36, 36);
                                                                    console.log(`  📏 Set Avatar dimensions: 36px x 36px`);
                                                                }
                                                                
                                                                // Configure Text component
                                                                const textInstances = slotGroup.findAll(node => {
                                                                    if (node.type === 'INSTANCE') {
                                                                        const mainComp = (node as InstanceNode).mainComponent;
                                                                        return mainComp !== null && mainComp.id === textComponent.id;
                                                                    }
                                                                    return false;
                                                                }) as InstanceNode[];
                                                                
                                                                if (textInstances.length > 0) {
                                                                    const textInstance = textInstances[0];
                                                                    console.log(`  📝 Found Text instance, setting text to: "${slotPropsConfig.userName}"`);
                                                                    console.log(`  📋 Text instance properties:`, Object.keys(textInstance.componentProperties));
                                                                    
                                                                    // Find the text property
                                                                    const textPropKey = Object.keys(textInstance.componentProperties).find(prop => 
                                                                        prop.toLowerCase().includes('text') || prop.toLowerCase().includes('label')
                                                                    );
                                                                    
                                                                    if (textPropKey) {
                                                                        textInstance.setProperties({
                                                                            [textPropKey]: slotPropsConfig.userName
                                                                        });
                                                                        console.log(`  ✅ Set text property: ${textPropKey} = "${slotPropsConfig.userName}"`);
                                                                    } else {
                                                                        console.warn(`  ⚠️ No text property found in Text instance`);
                                                                    }
                                                                } else {
                                                                    console.warn(`  ⚠️ Text instance not found after swap`);
                                                                }
                                                            }, 100); // Additional delay for nested component swap
                                                                }
                                                        } else {
                                                            console.warn(`  ⚠️ Expected 2 slot instances, found ${slotInstances.length}`);
                                                        }
                                                    }
                                                } else if (slotPropsConfig.suggestedComponent === 'tag') {
                                                    // Handle Tag component configuration
                                                    console.log(`  🏷️ Configuring Tag component...`);
                                                    
                                                    // Get the cell value for tag text from slotPropsConfig
                                                    const cellValue = slotPropsConfig.tagText || cellData.text || '';
                                                    console.log(`  📝 Cell value for tags: "${cellValue}"`);
                                                    
                                                    // Split by multiple separators: comma, slash, pipe, semicolon, ampersand
                                                    const tagValues = cellValue
                                                        .split(/[,\/\|;&]/)  // Split by comma, slash, pipe, semicolon, or ampersand
                                                        .map((v: string) => v.trim())
                                                        .filter((v: string) => v.length > 0);
                                                    console.log(`  🏷️ Tag values after split:`, tagValues);
                                                    
                                                    // Available colors for distinct tags
                                                    const availableColors = [
                                                        'Blue', 'Cyan', 'Teal', 'Green', 'Purple', 
                                                        'Magenta', 'Red', 'Gray', 'Cool gray', 'Warm gray'
                                                    ];
                                                    
                                                    // Find the Tag set component that was just swapped in
                                                    const tagSetComponents = cellRef.findAll(node => {
                                                        if (node.type === 'INSTANCE') {
                                                            const mainComp = (node as InstanceNode).mainComponent;
                                                            return mainComp !== null && mainComp.id === swapComponentId;
                                                        }
                                                        return false;
                                                    }) as InstanceNode[];
                                                    
                                                    if (tagSetComponents.length > 0) {
                                                        const tagSet = tagSetComponents[0];
                                                        console.log(`  🏷️ Found Tag set component: ${tagSet.name}`);
                                                        
                                                        // Disable Tag overflow property if it exists
                                                        const tagSetProps = tagSet.componentProperties || {};
                                                        const overflowProp = Object.keys(tagSetProps).find(prop => 
                                                            prop.toLowerCase().includes('overflow') || 
                                                            prop.toLowerCase().includes('tag overflow')
                                                        );
                                                        if (overflowProp) {
                                                            const overflowProps: any = {};
                                                            overflowProps[overflowProp] = false;
                                                            tagSet.setProperties(overflowProps);
                                                            console.log(`  🚫 Disabled Tag overflow property: ${overflowProp}`);
                                                        }
                                                        
                                                        // Debug: Log all children of the tag set
                                                        console.log(`  🔍 Tag set children:`, tagSet.children.map(child => ({
                                                            name: child.name,
                                                            type: child.type,
                                                            visible: child.visible
                                                        })));
                                                        
                                                        // Debug: Log all instances within the tag set
                                                        const allInstances = tagSet.findAll(node => node.type === 'INSTANCE') as InstanceNode[];
                                                        console.log(`  🔍 All instances in tag set:`, allInstances.map(instance => ({
                                                            name: instance.name,
                                                            mainComponentName: instance.mainComponent?.name,
                                                            visible: instance.visible
                                                        })));
                                                        
                                                        // Find all Tag - Read-only instances within the Tag set
                                                        // Look for instances with the name "Tag - Read-only"
                                                        const tagInstances = tagSet.findAll(node => {
                                                            if (node.type === 'INSTANCE') {
                                                                return node.name === 'Tag - Read-only';
                                                            }
                                                            return false;
                                                        }) as InstanceNode[];
                                                        
                                                        console.log(`  🏷️ Found ${tagInstances.length} tag instances`);
                                                        
                                                        // Configure each tag instance
                                                        tagInstances.forEach((tagInstance, index) => {
                                                            if (index < tagValues.length) {
                                                                // Show this tag and set its properties
                                                                const tagValue = tagValues[index];
                                                                const color = availableColors[index % availableColors.length];
                                                                
                                                                console.log(`  🏷️ Configuring tag ${index + 1}: "${tagValue}" with color "${color}"`);
                                                                
                                                                // Set tag text
                                                                const textProp = Object.keys(tagInstance.componentProperties).find(prop => 
                                                                    prop.toLowerCase().includes('text')
                                                                );
                                                                
                                                                // Set tag color
                                                                const colorProp = Object.keys(tagInstance.componentProperties).find(prop => 
                                                                    prop.toLowerCase().includes('color') || prop.toLowerCase().includes('variant')
                                                                );
                                                                
                                                                const tagProps: any = {};
                                                                if (textProp) tagProps[textProp] = tagValue;
                                                                if (colorProp) tagProps[colorProp] = color;
                                                                
                                                                if (Object.keys(tagProps).length > 0) {
                                                                    tagInstance.setProperties(tagProps);
                                                                    console.log(`  ✅ Set tag properties:`, tagProps);
                                                                }
                                                                
                                                                // Make tag visible
                                                                tagInstance.visible = true;
                                                            } else {
                                                                // Hide unused tags
                                                                console.log(`  👁️ Hiding unused tag ${index + 1}`);
                                                                tagInstance.visible = false;
                                                            }
                                                        });
                                                        
                                                        console.log(`  ✅ Tag configuration complete!`);
                                                    } else {
                                                        console.warn(`  ⚠️ Tag set component not found in cell ${cellKeyForLog}`);
                                                    }
                                                } else {
                                                    // Handle Status Icon (existing logic)
                                                    const slottedComponents = cellRef.findAll(node => {
                                                        if (node.type === 'INSTANCE') {
                                                            const mainComp = (node as InstanceNode).mainComponent;
                                                            return mainComp !== null && mainComp.id === swapComponentId;
                                                        }
                                                        return false;
                                                    }) as InstanceNode[];
                                                    
                                                    if (slottedComponents.length > 0) {
                                                        const slottedComponent = slottedComponents[0];
                                                        console.log(`  📦 Found Status Icon component: ${slottedComponent.name}`);
                                                        console.log(`  📋 Available properties:`, Object.keys(slottedComponent.componentProperties));
                                                        console.log(`  📝 Properties to apply (from UI):`, slotPropsConfig);
                                                        
                                                        try {
                                                            const slotProps = mapPropertyNames(slotPropsConfig, slottedComponent.componentProperties);
                                                            console.log(`  🔍 Mapped properties (after mapPropertyNames):`, slotProps);
                                                            
                                                            // Debug status property specifically
                                                            if (slotPropsConfig.Status) {
                                                                console.log(`🔍 [STATUS DEBUG] Original Status property: "${slotPropsConfig.Status}"`);
                                                                console.log(`🔍 [STATUS DEBUG] Mapped Status property: "${slotProps.Status || 'NOT MAPPED'}"`);
                                                            }
                                                            
                                                            if (Object.keys(slotProps).length > 0) {
                                                                console.log(`  🚀 Applying properties to Status Icon...`);
                                                                slottedComponent.setProperties(slotProps);
                                                                console.log(`  ✅ Applied properties to Status Icon`);
                                                            } else {
                                                                console.warn(`  ⚠️ No properties matched for Status Icon`);
                                                            }
                                                        } catch (slotError) {
                                                            console.warn(`  ⚠️ Could not set Status Icon properties:`, slotError);
                                                        }
                                                    } else {
                                                        console.warn(`  ⚠️ Status Icon component not found in cell ${cellKeyForLog}`);
                                                    }
                                                }
                                            } catch (error) {
                                                console.error(`  ❌ Error configuring slot component:`, error);
                                            }
                                        }, 100); // 100ms delay to allow Figma to complete the swap
                                    }
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

            // Add toolbar if requested (placed at the top before header)
            console.log('🔍 Toolbar generation check:', {
                includeToolbar,
                toolbarExists: !!toolbar,
                toolbarName: toolbar?.name || 'null',
                condition: includeToolbar && toolbar
            });
            
            if (includeToolbar && toolbar) {
                console.log('[Backend] Adding toolbar to table...');
                try {
                    const toolbarClone = toolbar.createInstance();
                    toolbarClone.name = "Toolbar";
                    toolbarClone.layoutSizingHorizontal = 'FIXED';
                    toolbarClone.resize(totalTableWidth, toolbarClone.height);
                    
                    // Insert toolbar at the beginning (top) of the table
                    tableFrame.insertChild(0, toolbarClone);
                    console.log('✅ Successfully added toolbar to table at the top');
                } catch (error) {
                    console.error('Error creating toolbar:', error);
                    figma.notify('⚠️ Toolbar component is no longer available');
                }
            } else if (includeToolbar && !toolbar) {
                console.log('⚠️ Toolbar requested but not found in scan result');
            } else if (!includeToolbar) {
                console.log('ℹ️ Toolbar not requested (includeToolbar = false)');
            } else {
                console.log('ℹ️ Toolbar condition not met - includeToolbar:', includeToolbar, 'toolbar:', !!toolbar);
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
                includeToolbar: msg.includeToolbar,
                includeSelectable: msg.includeSelectable,
                includeExpandable: msg.includeExpandable,
                cellProperties: msg.cellProps,
                bodyCellComponentId: lastScanResult?.bodyCell?.id || null, // Store body cell component ID for plugin reopen
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
                        toolbarId: lastScanResult.toolbar?.id || null,
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
    } else if (msg.type === 'update-table-metadata') {
        console.log('[Backend] Received update-table-metadata request');
        
        const tableId = msg.tableId;
        const cellProperties = msg.cellProperties;
        
        if (!tableId || !cellProperties) {
            figma.notify('Missing table ID or cell properties for metadata update');
            return;
        }
        
        const tableNode = figma.getNodeById(tableId);
        if (!tableNode || (tableNode.type !== 'FRAME' && tableNode.type !== 'COMPONENT')) {
            figma.notify('Table not found for metadata update');
            return;
        }
        
        try {
            // Get existing table settings
            const tableSettingsData = tableNode.getPluginData('tableSettings');
            let tableSettings: any = {};
            
            if (tableSettingsData) {
                try {
                    tableSettings = JSON.parse(tableSettingsData);
                } catch (e) {
                    console.warn('[Backend] Could not parse existing table settings:', e);
                }
            }
            
            // Update cell properties with the new data
            tableSettings.cellProperties = cellProperties;
            
            // Save updated settings back to table
            tableNode.setPluginData('tableSettings', JSON.stringify(tableSettings));
            
            console.log('[Backend] Successfully updated table metadata with new cell properties');
            console.log('[Backend] Updated cell properties keys:', Object.keys(cellProperties));
            
            // Send confirmation back to UI
            figma.ui.postMessage({
                type: 'metadata-updated',
                success: true,
                message: 'Table metadata updated successfully'
            });
            
        } catch (error) {
            console.error('[Backend] Error updating table metadata:', error);
            figma.notify('❌ Error updating table metadata');
            
            figma.ui.postMessage({
                type: 'metadata-updated',
                success: false,
                message: (error as Error).message
            });
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
                        let headerCell = null, bodyCell = null, footer = null, toolbarComponent = null, selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
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
                        
                        // Restore toolbar
                        if (storedScan.toolbarId) {
                            try {
                                const toolbarNode = figma.getNodeById(storedScan.toolbarId);
                                if (toolbarNode && toolbarNode.type === 'COMPONENT') {
                                    toolbarComponent = toolbarNode;
                                    console.log('✅ Restored toolbar from stored ID:', toolbarComponent.name);
                                }
                            } catch (error) {
                                console.log('Could not restore toolbar from ID:', error);
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
                                toolbar: toolbarComponent,
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
                                        toolbar: toolbarComponent,
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
                let headerCell = null, bodyCell = null, footer = null, toolbarComponent = null, selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
                
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
                            } else if (!toolbarComponent && compName.includes('toolbar')) {
                                toolbarComponent = instance.mainComponent;
                                console.log(`✅ Found toolbar from table instance: ${toolbarComponent.name}`);
                            } else if (compName.includes('toolbar')) {
                                console.log(`🔍 Found toolbar component but already have one: ${instance.mainComponent.name}`);
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
                    toolbar: toolbarComponent as ComponentNode | null,
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
                includeToolbar: msg.includeToolbar,
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
            const { headerCell, bodyCell, footer, toolbar, numCols } = lastScanResult!;
            
            const includeHeader = msg.includeHeader !== false;
            const includeFooter = msg.includeFooter === true;
            const includeToolbar = msg.includeToolbar === true;
            const includeSelectable = msg.includeSelectable === true;
            const includeExpandable = msg.includeExpandable === true;
            const { rows, cols } = msg;
            let { cellProps } = msg;
            
            // Get the original table settings to preserve header/footer/toolbar state
            let originalIncludeHeader = includeHeader;
            let originalIncludeFooter = includeFooter;
            let originalIncludeToolbar = includeToolbar;
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
            if (includeToolbar && !lastScanResult?.toolbar) {
                console.warn('⚠️ Toolbar component not found, toolbar will be skipped');
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
                    headerRowFrame.counterAxisAlignItems = "CENTER"; // Vertically center header content
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
                        rowFrame.counterAxisAlignItems = "CENTER"; // Vertically center cell content
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
                            // Check if slot is enabled using the dedicated slot variable (same logic as create table)
                            const isSlotEnabled = cellData.slot === true;
                            console.log(`🔄 [UPDATE] Cell ${key} slot state: ${isSlotEnabled} (from cellData.slot)`);
                            
                            // Filter out slot properties if slot is disabled
                            let propertiesToApply = { ...cellData.properties };
                            if (!isSlotEnabled) {
                                // Remove slot-related properties if slot is disabled
                                const slotProps = Object.keys(propertiesToApply).filter(prop => 
                                    prop.toLowerCase().includes('slot') || 
                                    prop.toLowerCase().includes('swap')
                                );
                                slotProps.forEach(prop => {
                                    delete propertiesToApply[prop];
                                    console.log(`🔄 [UPDATE] Removed slot property ${prop} for cell ${key} (slot disabled)`);
                                });
                            }
                            
                            const validProps = mapPropertyNames(propertiesToApply, cell.componentProperties);
                            cell.setProperties(validProps);
                            
                            // Apply smart slot configuration if present (UPDATE)
                            const swapSlotKey = Object.keys(validProps).find(key => key.toLowerCase().includes('swap') && key.toLowerCase().includes('slot'));
                            const hasSwapSlot = Object.keys(validProps).some(key => key.toLowerCase().includes('swap') && key.toLowerCase().includes('slot'));
                            const swapComponentId = swapSlotKey ? validProps[swapSlotKey] : null;
                            
                            if (cellData.slotComponentProps && isSlotEnabled && hasSwapSlot && swapComponentId) {
                                // Store the cell reference and config for the async operation
                                const cellRef = cell;
                                const slotPropsConfig = cellData.slotComponentProps;
                                const cellKeyForLog = key;
                                
                                // Check if this is a nested slot case (Slot group with Avatar + Text)
                                const isNestedSlot = slotPropsConfig.nestedSlots === true;
                                
                                // Use setTimeout to allow Figma to complete the component swap
                                setTimeout(async () => {
                                    try {
                                        console.log(`🔧 [UPDATE] Configuring slot component for cell ${cellKeyForLog}:`, slotPropsConfig);
                                        
                                        if (isNestedSlot) {
                                            // Handle Slot Group with Avatar + Text OR Edit + Delete
                                            let avatarComponent: ComponentNode | null = null;
                                            let textComponent: ComponentNode | null = null;
                                            let editComponent: ComponentNode | null = null;
                                            let deleteComponent: ComponentNode | null = null;
                                            
                                            if (slotPropsConfig.userName) {
                                                // User name case: Avatar + Text
                                                console.log(`👥 [UPDATE] Processing nested slots for user: ${slotPropsConfig.userName}`);
                                                
                                                // Import the Avatar and Text components
                                                avatarComponent = await figma.importComponentByKeyAsync(slotPropsConfig.avatarKey);
                                                textComponent = await figma.importComponentByKeyAsync(slotPropsConfig.textKey);
                                                console.log(`  ✅ [UPDATE] Imported Avatar and Text components`);
                                            } else if (slotPropsConfig.editKey && slotPropsConfig.deleteKey) {
                                                // Edit/Delete case: Edit + Delete icons
                                                console.log(`⚡ [UPDATE] Processing nested slots for edit/delete actions`);
                                                
                                                // Import the Edit and Delete components
                                                editComponent = await figma.importComponentByKeyAsync(slotPropsConfig.editKey);
                                                deleteComponent = await figma.importComponentByKeyAsync(slotPropsConfig.deleteKey);
                                                console.log(`  ✅ [UPDATE] Imported Edit and Delete components`);
                                            }
                                            
                                            // Find the Slot group component that was just swapped in
                                            const slotGroupComponents = cellRef.findAll(node => {
                                                if (node.type === 'INSTANCE') {
                                                    const mainComp = (node as InstanceNode).mainComponent;
                                                    return mainComp !== null && mainComp.id === swapComponentId;
                                                }
                                                return false;
                                            }) as InstanceNode[];
                                            
                                            if (slotGroupComponents.length > 0) {
                                                const slotGroup = slotGroupComponents[0];
                                                console.log(`  🔍 [UPDATE] Found Slot Group component: ${slotGroup.name}`);
                                                
                                                // Find the actual slot instances (children of the Slot Group)
                                                const slotInstances = slotGroup.children.filter(node => 
                                                    node.type === 'INSTANCE' && 
                                                    node.name.toLowerCase().includes('slot')
                                                ) as InstanceNode[];
                                                
                                                console.log(`  🔍 [UPDATE] Found ${slotInstances.length} slot instances`);
                                                
                                                if (slotInstances.length >= 2) {
                                                    const avatarSlot = slotInstances[0];
                                                    const textSlot = slotInstances[1];
                                                    
                                                    if (avatarComponent && textComponent && slotPropsConfig.userName) {
                                                        // Swap Avatar into first slot
                                                        console.log(`  📏 [UPDATE] Before swap - avatarSlot layoutSizingHorizontal:`, avatarSlot.layoutSizingHorizontal);
                                                        console.log(`  📏 [UPDATE] Before swap - slotGroup layoutMode:`, slotGroup.layoutMode);
                                                        
                                                        avatarSlot.swapComponent(avatarComponent);
                                                        console.log(`  ✅ [UPDATE] Swapped Avatar component`);
                                                        
                                                        // Configure Avatar properties
                                                        setTimeout(() => {
                                                            const avatarInstances = cellRef.findAll(node => {
                                                                if (node.type === 'INSTANCE') {
                                                                    const mainComp = (node as InstanceNode).mainComponent;
                                                                    return mainComp !== null && mainComp.id === avatarComponent!.id;
                                                                }
                                                                return false;
                                                            }) as InstanceNode[];
                                                            
                                                            if (avatarInstances.length > 0) {
                                                                const avatarInstance = avatarInstances[0];
                                                                // Extract proper initials from the name
                                                                const initials = extractInitials(slotPropsConfig.userName);
                                                                console.log(`  🔧 [UPDATE] Configuring Avatar: Type=Initials, Size=Medium, Initial text=${initials}`);
                                                                
                                                                // Set Avatar properties
                                                                const avatarProps: any = {
                                                                    'Type': 'Initials',
                                                                    'Size': 'Medium'
                                                                };
                                                                
                                                                // Find the Initial text property
                                                                const initialTextProp = Object.keys(avatarInstance.componentProperties).find(prop => 
                                                                    prop.toLowerCase().includes('initial') && prop.toLowerCase().includes('text')
                                                                );
                                                                if (initialTextProp) {
                                                                    avatarProps[initialTextProp] = initials;
                                                                }
                                                                
                                                                avatarInstance.setProperties(avatarProps);
                                                                console.log(`  ✅ [UPDATE] Avatar configured successfully`);
                                                                
                                                                // Set Avatar layout properties to HUG
                                                                avatarInstance.layoutSizingHorizontal = 'HUG';
                                                                avatarInstance.layoutSizingVertical = 'HUG';
                                                                avatarInstance.layoutGrow = 0;
                                                                console.log(`  📏 [UPDATE] Set Avatar layout to HUG`);
                                                                
                                                                // Set fixed dimensions for Avatar (36px x 36px)
                                                                avatarInstance.resize(36, 36);
                                                                console.log(`  📏 [UPDATE] Set Avatar dimensions: 36px x 36px`);
                                                            }
                                                        }, 100);
                                                        
                                                        // Swap Text into second slot
                                                        textSlot.swapComponent(textComponent);
                                                        console.log(`  ✅ [UPDATE] Swapped Text component`);
                                                        
                                                        // Configure Text properties
                                                        setTimeout(() => {
                                                            const textInstances = cellRef.findAll(node => {
                                                                if (node.type === 'INSTANCE') {
                                                                    const mainComp = (node as InstanceNode).mainComponent;
                                                                    return mainComp !== null && mainComp.id === textComponent!.id;
                                                                }
                                                                return false;
                                                            }) as InstanceNode[];
                                                            
                                                            if (textInstances.length > 0) {
                                                                const textInstance = textInstances[0];
                                                                console.log(`  📝 [UPDATE] Found Text instance, setting text to: "${slotPropsConfig.userName}"`);
                                                                
                                                                // Find the text property
                                                                const textPropKey = Object.keys(textInstance.componentProperties).find(prop => 
                                                                    prop.toLowerCase().includes('text') || prop.toLowerCase().includes('label')
                                                                );
                                                                
                                                                if (textPropKey) {
                                                                    textInstance.setProperties({
                                                                        [textPropKey]: slotPropsConfig.userName
                                                                    });
                                                                    console.log(`  ✅ [UPDATE] Set text property: ${textPropKey} = "${slotPropsConfig.userName}"`);
                                                                } else {
                                                                    console.warn(`  ⚠️ [UPDATE] No text property found in Text instance`);
                                                                }
                                                            } else {
                                                                console.warn(`  ⚠️ [UPDATE] Text instance not found after swap`);
                                                            }
                                                        }, 100); // Additional delay for nested component swap
                                                    } else if (editComponent && deleteComponent) {
                                                        // Swap Edit icon into first slot
                                                        avatarSlot.swapComponent(editComponent);
                                                        console.log(`  ✅ [UPDATE] Swapped Edit icon`);
                                                        
                                                        // Swap Delete icon into second slot
                                                        textSlot.swapComponent(deleteComponent);
                                                        console.log(`  ✅ [UPDATE] Swapped Delete icon`);
                                                    }
                                                } else {
                                                    console.warn(`  ⚠️ [UPDATE] Expected 2 slot instances, found ${slotInstances.length}`);
                                                }
                                            } else {
                                                console.warn(`  ⚠️ [UPDATE] Slot Group component not found after swap`);
                                            }
                                        } else if (slotPropsConfig.suggestedComponent === 'tag') {
                                            // Handle Tag component configuration
                                            console.log(`  🏷️ [UPDATE] Configuring Tag component...`);
                                            
                                            // Get the cell value for tag text from slotPropsConfig
                                            const cellValue = slotPropsConfig.tagText || cellData.text || '';
                                            console.log(`  📝 [UPDATE] Cell value for tags: "${cellValue}"`);
                                            
                                            // Split by multiple separators: comma, slash, pipe, semicolon, ampersand
                                            const tagValues = cellValue
                                                .split(/[,\/\|;&]/)  // Split by comma, slash, pipe, semicolon, or ampersand
                                                .map((v: string) => v.trim())
                                                .filter((v: string) => v.length > 0);
                                            console.log(`  🏷️ [UPDATE] Tag values after split:`, tagValues);
                                            
                                            // Available colors for distinct tags
                                            const availableColors = [
                                                'Blue', 'Cyan', 'Teal', 'Green', 'Purple', 
                                                'Magenta', 'Red', 'Gray', 'Cool gray', 'Warm gray'
                                            ];
                                            
                                            // Find the Tag set component that was just swapped in
                                            const tagSetComponents = cellRef.findAll(node => {
                                                if (node.type === 'INSTANCE') {
                                                    const mainComp = (node as InstanceNode).mainComponent;
                                                    return mainComp !== null && mainComp.id === swapComponentId;
                                                }
                                                return false;
                                            }) as InstanceNode[];
                                            
                                            if (tagSetComponents.length > 0) {
                                                const tagSet = tagSetComponents[0];
                                                console.log(`  🏷️ [UPDATE] Found Tag set component: ${tagSet.name}`);
                                                
                                                // Disable Tag overflow property if it exists
                                                const tagSetProps = tagSet.componentProperties || {};
                                                const overflowProp = Object.keys(tagSetProps).find(prop => 
                                                    prop.toLowerCase().includes('overflow') || 
                                                    prop.toLowerCase().includes('tag overflow')
                                                );
                                                if (overflowProp) {
                                                    const overflowProps: any = {};
                                                    overflowProps[overflowProp] = false;
                                                    tagSet.setProperties(overflowProps);
                                                    console.log(`  🚫 [UPDATE] Disabled Tag overflow property: ${overflowProp}`);
                                                }
                                                
                                                // Find all Tag - Read-only instances within the Tag set
                                                const tagInstances = tagSet.findAll(node => 
                                                    node.name === 'Tag - Read-only'
                                                ) as InstanceNode[];
                                                
                                                console.log(`  🏷️ [UPDATE] Found ${tagInstances.length} tag instances`);
                                                
                                                // Configure each tag instance
                                                tagInstances.forEach((tagInstance, index) => {
                                                    if (index < tagValues.length) {
                                                        const tagValue = tagValues[index];
                                                        const color = availableColors[index % availableColors.length];
                                                        
                                                        console.log(`  🏷️ [UPDATE] Configuring tag ${index + 1}: "${tagValue}" with color "${color}"`);
                                                        
                                                        // Set tag properties
                                                        const tagProps: any = {};
                                                        
                                                        // Find text property
                                                        const textProp = Object.keys(tagInstance.componentProperties).find(prop => 
                                                            prop.toLowerCase().includes('text') || prop.toLowerCase().includes('label')
                                                        );
                                                        if (textProp) {
                                                            tagProps[textProp] = tagValue;
                                                        }
                                                        
                                                        // Find color property
                                                        const colorProp = Object.keys(tagInstance.componentProperties).find(prop => 
                                                            prop.toLowerCase().includes('color')
                                                        );
                                                        if (colorProp) {
                                                            tagProps[colorProp] = color;
                                                        }
                                                        
                                                        tagInstance.setProperties(tagProps);
                                                        tagInstance.visible = true;
                                                    } else {
                                                        // Hide unused tags
                                                        tagInstance.visible = false;
                                                    }
                                                });
                                                
                                                console.log(`  ✅ [UPDATE] Tag configuration complete!`);
                                            } else {
                                                console.warn(`  ⚠️ [UPDATE] Tag set component not found after swap`);
                                            }
                                        } else {
                                            // Handle other slot components (Status Icon, etc.)
                                            console.log(`  🔧 [UPDATE] Configuring ${slotPropsConfig.suggestedComponent} component...`);
                                            
                                            // Find the slotted component that was just swapped in
                                            const slottedComponents = cellRef.findAll(node => {
                                                if (node.type === 'INSTANCE') {
                                                    const mainComp = (node as InstanceNode).mainComponent;
                                                    return mainComp !== null && mainComp.id === swapComponentId;
                                                }
                                                return false;
                                            }) as InstanceNode[];
                                            
                                            if (slottedComponents.length > 0) {
                                                const slottedComponent = slottedComponents[0];
                                                console.log(`  🔧 [UPDATE] Found slotted component: ${slottedComponent.name}`);
                                                console.log(`  🔧 [UPDATE] Available properties:`, Object.keys(slottedComponent.componentProperties));
                                                console.log(`  🔧 [UPDATE] slotPropsConfig received:`, slotPropsConfig);
                                                
                                                // Use mapPropertyNames to handle property name variations
                                                try {
                                                    const mappedProps = mapPropertyNames(slotPropsConfig, slottedComponent.componentProperties);
                                                    console.log(`  🔍 [UPDATE] Mapped properties:`, mappedProps);
                                                    
                                                    if (Object.keys(mappedProps).length > 0) {
                                                        slottedComponent.setProperties(mappedProps);
                                                        console.log(`  ✅ [UPDATE] Applied slot component properties successfully`);
                                                    } else {
                                                        console.warn(`  ⚠️ [UPDATE] No properties mapped for slot component`);
                                                    }
                                                } catch (error) {
                                                    console.error(`  ❌ [UPDATE] Error applying slot properties:`, error);
                                                }
                                                
                                                console.log(`  ✅ [UPDATE] Slot component configuration complete!`);
                                            } else {
                                                console.warn(`  ⚠️ [UPDATE] Slotted component not found after swap`);
                                            }
                                        }
                                    } catch (error) {
                                        console.error(`❌ [UPDATE] Error configuring slot component for cell ${cellKeyForLog}:`, error);
                                    }
                                }, 100); // Allow time for component swap to complete
                            }
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

            // Add toolbar if requested (placed at the top before header)
            if (includeToolbar && toolbar) {
                console.log('[Backend] Adding toolbar to updated table...');
                try {
                    const toolbarClone = toolbar.createInstance();
                    toolbarClone.name = "Toolbar";
                    toolbarClone.layoutSizingHorizontal = 'FIXED';
                    toolbarClone.resize(totalTableWidth, toolbarClone.height);
                    
                    // Insert toolbar at the beginning (top) of the table
                    tableFrame.insertChild(0, toolbarClone);
                    console.log('✅ Successfully added toolbar to updated table at the top');
                } catch (error) {
                    console.error('Error creating toolbar in update:', error);
                    figma.notify('⚠️ Toolbar component is no longer available');
                }
            } else if (includeToolbar && !toolbar) {
                console.log('⚠️ Toolbar requested but not found in scan result (update)');
            } else if (!includeToolbar) {
                console.log('ℹ️ Toolbar not requested in update (includeToolbar = false)');
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
                includeToolbar: originalIncludeToolbar,
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
    } else if (msg.type === 'test-swap-component') {
        console.log('🔄 Testing Overflow component swap...');
        (async () => {
            try {
                // Get the selected nodes
                const selection = figma.currentPage.selection;
                if (selection.length === 0) {
                    figma.notify('⚠️ Please select a cell instance to swap with Overflow');
                    return;
                }
                
                const selectedNode = selection[0];
                if (selectedNode.type !== 'INSTANCE') {
                    figma.notify('⚠️ Please select an instance (not a component) to test swap');
                    return;
                }
                
                console.log(`📋 Selected instance: ${selectedNode.name} (ID: ${selectedNode.id})`);
                
                // Try to import the Overflow component using the hardcoded key
                const overflowKey = '608ab39a49081fa1a460e71af33de459e6154ee7';
                console.log(`🔄 Attempting to import Overflow component with key: ${overflowKey}`);
                
                try {
                    const overflowComponent = await figma.importComponentByKeyAsync(overflowKey);
                    
                    if (!overflowComponent) {
                        figma.notify('❌ Could not import Overflow component - is Carbon library enabled?');
                        return;
                    }
                    
                    console.log(`✅ Successfully imported Overflow component: ${overflowComponent.name}`);
                    console.log(`   Component ID: ${overflowComponent.id}`);
                    console.log(`   Component Key: ${overflowComponent.key}`);
                    
                    // Perform the swap
                    (selectedNode as InstanceNode).swapComponent(overflowComponent);
                    figma.notify(`✅ Successfully swapped to Overflow component!`);
                    
                    console.log(`✅ Swap complete! Instance now uses Overflow component.`);
                    
                } catch (importError) {
                    console.error('❌ Error importing Overflow component:', importError);
                    figma.notify(`❌ Could not import Overflow: ${importError instanceof Error ? importError.message : 'Unknown error'}`);
                    console.log('💡 Make sure the Carbon Design System library is enabled in this file');
                }
                
            } catch (error) {
                console.error('❌ Error testing component swap:', error);
                figma.notify(`❌ Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
            }
        })();
    } else if (msg.type === 'analyze-smart-slots') {
        console.log('🧠 Analyzing table data for smart slot suggestions...');
        (async () => {
            try {
                const { gridData, headers, autoApply } = msg;
                
                if (!gridData || !headers) {
                    figma.notify('⚠️ No table data provided for analysis');
                    return;
                }
                
                const suggestions = analyzeTableDataForSmartSlots(gridData, headers);
                
                // If autoApply is true, automatically apply the suggestions with component imports
                if (autoApply && suggestions.length > 0) {
                    console.log(`🤖 Auto-applying ${suggestions.length} smart slot suggestions...`);
                    
                    // Import components for each suggestion
                    const smartSlotComponents: Record<string, string> = {
                        overflow: '608ab39a49081fa1a460e71af33de459e6154ee7',
                        statusIcon: 'd94da074d575aa0ae337fc054f4f19bc311049b8',
                        slotGroup: '761fd2270cc978c88b9cce64926b4c8245831bdc',
                        avatar: 'd80f0d175851756c4601e87e6e0abeb5539b24e8',
                        text: 'e73c62eb16dcb7f54df1384f176fc7a3c0f64df5',
                        tag: 'c95a2fb5332d75515d2ed90ac487bc864fcb09b5',
                        edit: 'a4ba4c4aa1f2b0f0a5206341aafbb7d7eafa47e6',
                        delete: '84a7c6755b83b8e88ca803851c280d1e06255b93',
                    };
                    
                    const suggestionsWithComponents = await Promise.all(suggestions.map(async (suggestion) => {
                        const componentKey = smartSlotComponents[suggestion.suggestedComponent as keyof typeof smartSlotComponents];
                        
                        if (componentKey) {
                            try {
                                console.log(`📦 Importing ${suggestion.suggestedComponent} (key: ${componentKey})...`);
                                const component = await figma.importComponentByKeyAsync(componentKey);
                                
                                return {
                                    ...suggestion,
                                    componentId: component.id,
                                    componentName: component.name
                                };
                            } catch (error) {
                                console.warn(`⚠️ Could not import ${suggestion.suggestedComponent}:`, error);
                                return suggestion;
                            }
                        }
                        
                        return suggestion;
                    }));
                    
                    figma.ui.postMessage({
                        type: 'auto-apply-smart-slots',
                        suggestions: suggestionsWithComponents
                    });
                    
                    const appliedCount = suggestionsWithComponents.filter(s => s.componentId).length;
                    if (appliedCount > 0) {
                        figma.notify(`✨ Auto-applied ${appliedCount} smart component${appliedCount > 1 ? 's' : ''}`);
                    } else {
                        figma.notify('⚠️ Could not import components for auto-apply');
                    }
                } else {
                    // Just send suggestions for manual review
                    figma.ui.postMessage({
                        type: 'smart-slot-suggestions',
                        suggestions
                    });
                }
            } catch (error) {
                console.error('❌ Error analyzing smart slots:', error);
                figma.notify(`❌ Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
            }
        })();
    } else if (msg.type === 'discover-library-components') {
        console.log('🔍 Discovering components from all available libraries...');
        try {
            // This function will help users discover components from libraries
            // that are added to their file but not yet used on any page
            
            console.log('📋 Instructions for discovering library components:');
            console.log('1. Open the Carbon Design System library file in Figma');
            console.log('2. Navigate to the component you want (e.g., Overflow Menu)');
            console.log('3. Select the component and run: figma.currentPage.selection[0].key');
            console.log('4. Copy the returned key');
            console.log('5. Use "Collect Carbon Keys" button to store it in this plugin');
            
            figma.notify('📋 Check console for instructions on discovering library components');
            
            // Try to provide some guidance by checking what's available
            const allComponents = figma.root.findAll(node => node.type === 'COMPONENT') as ComponentNode[];
            const allComponentSets = figma.root.findAll(node => node.type === 'COMPONENT_SET') as ComponentSetNode[];
            
            console.log(`📊 Current file has ${allComponents.length} components and ${allComponentSets.length} component sets`);
            console.log('💡 To access library components not yet used, you need their keys from the library file');
            
        } catch (error) {
            console.error('❌ Error discovering library components:', error);
            figma.notify(`❌ Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    } else if (msg.type === 'find-overflow-component') {
        console.log('🔍 Finding Overflow component automatically...');
        try {
            // First, check if we already found the Overflow component
            let overflowInfo = null;
            try {
                const storedOverflow = figma.currentPage.getPluginData('foundOverflowComponent');
                if (storedOverflow) {
                    overflowInfo = JSON.parse(storedOverflow);
                    console.log(`✅ Found previously discovered Overflow component: ${overflowInfo.name}`);
                    figma.notify(`✅ Overflow component already found: ${overflowInfo.name}`);
                    return;
                }
            } catch (e) {
                console.log('No previously found Overflow component');
            }
            
            // Try to find the Overflow component in existing component sets
            const componentSets = figma.root.findAll(node => node.type === 'COMPONENT_SET') as ComponentSetNode[];
            console.log(`📦 Searching ${componentSets.length} component sets for Overflow...`);
            let overflowFound = false;
            
            for (const componentSet of componentSets) {
                const setName = componentSet.name.toLowerCase();
                console.log(`🔍 Checking: "${componentSet.name}"`);
                
                if (setName.includes('overflow') || setName.includes('menu') || setName.includes('action')) {
                    console.log(`🎯 Found potential Overflow component set: ${componentSet.name}`);
                    
                    // First, check if the component set already has accessible variants
                    const existingVariants = componentSet.children.filter(child => child.type === 'COMPONENT') as ComponentNode[];
                    if (existingVariants.length > 0) {
                        console.log(`✅ Component set already accessible with ${existingVariants.length} variants`);
                        for (const variant of existingVariants) {
                            const variantName = variant.name.toLowerCase();
                            if (variantName.includes('overflow')) {
                                console.log(`✅ Found Overflow component: ${variant.name}`);
                                
                                // Store this for future use
                                const overflowInfo = {
                                    name: variant.name,
                                    key: componentSet.key,
                                    id: variant.id,
                                    description: variant.description || '',
                                    width: variant.width,
                                    height: variant.height,
                                    timestamp: Date.now()
                                };
                                
                                // Save to plugin data
                                figma.currentPage.setPluginData('foundOverflowComponent', JSON.stringify(overflowInfo));
                                
                                figma.notify(`✅ Found Overflow component: ${variant.name}`);
                                overflowFound = true;
                                break;
                            }
                        }
                    } else {
                        // Try to import the component set to access its variants
                        console.log(`🔄 Attempting to import: ${componentSet.name}`);
                        try {
                            const importedSet = await figma.importComponentSetByKeyAsync(componentSet.key);
                            if (importedSet) {
                                const variants = importedSet.children.filter(child => child.type === 'COMPONENT') as ComponentNode[];
                                console.log(`✅ Successfully imported with ${variants.length} variants`);
                                for (const variant of variants) {
                                    const variantName = variant.name.toLowerCase();
                                    if (variantName.includes('overflow')) {
                                        console.log(`✅ Found Overflow component: ${variant.name}`);
                                        
                                        // Store this for future use
                                        const overflowInfo = {
                                            name: variant.name,
                                            key: componentSet.key,
                                            id: variant.id,
                                            description: variant.description || '',
                                            width: variant.width,
                                            height: variant.height,
                                            timestamp: Date.now()
                                        };
                                        
                                        // Save to plugin data
                                        figma.currentPage.setPluginData('foundOverflowComponent', JSON.stringify(overflowInfo));
                                        
                                        figma.notify(`✅ Found Overflow component: ${variant.name}`);
                                        overflowFound = true;
                                        break;
                                    }
                                }
                            }
                        } catch (importError) {
                            console.log(`⚠️ Could not import "${componentSet.name}": ${importError instanceof Error ? importError.message : 'Unknown error'}`);
                            console.log(`ℹ️ This component set may not be accessible in the current file`);
                        }
                    }
                }
                
                if (overflowFound) break;
            }
            
            if (!overflowFound) {
                console.log('ℹ️ No Overflow component found in component sets');
                figma.notify('⚠️ Overflow component not found in component sets. Try using "Collect Carbon Keys" on an Overflow component first.');
            }
            
        } catch (error) {
            console.error('❌ Error finding Overflow component:', error);
            figma.notify(`❌ Error finding Overflow: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    } else if (msg.type === 'collect-carbon-keys') {
        console.log('🔑 Collecting Carbon component keys...');
        try {
            const selection = figma.currentPage.selection;
            if (selection.length === 0) {
                figma.notify('Please select a Carbon component to collect its key');
                return;
            }
            
            const selectedNode = selection[0];
            if (selectedNode.type !== 'COMPONENT' && selectedNode.type !== 'INSTANCE') {
                figma.notify('Please select a Carbon component or instance to collect its key');
                return;
            }
            
            const component = selectedNode.type === 'INSTANCE' ? selectedNode.mainComponent : selectedNode;
            if (!component) {
                figma.notify('Could not get component information');
                return;
            }
            
            // Store the component key for later use
            const componentInfo = {
                name: component.name,
                key: component.key,
                id: component.id,
                description: component.description || '',
                width: component.width,
                height: component.height,
                timestamp: Date.now()
            };
            
            // Get existing collected keys from plugin data (store on current page)
            let collectedKeys = [];
            try {
                const existingKeys = figma.currentPage.getPluginData('collectedCarbonKeys');
                if (existingKeys) {
                    collectedKeys = JSON.parse(existingKeys);
                }
            } catch (e) {
                console.log('No existing collected keys found');
            }
            
            // Add new key if not already collected
            if (!collectedKeys.find((k: any) => k.key === component.key)) {
                collectedKeys.push(componentInfo);
                figma.currentPage.setPluginData('collectedCarbonKeys', JSON.stringify(collectedKeys));
                console.log(`✅ Collected Carbon component key: ${component.name} (${component.key})`);
                figma.notify(`✅ Collected: ${component.name}`);
            } else {
                console.log(`ℹ️ Key already collected: ${component.name}`);
                figma.notify(`ℹ️ Already collected: ${component.name}`);
            }
            
            // Send all collected keys to UI
            figma.ui.postMessage({
                type: 'carbon-keys-collected',
                keys: collectedKeys,
                newKey: componentInfo
            });
            
        } catch (error) {
            console.error('❌ Error collecting Carbon key:', error);
            figma.notify(`❌ Error collecting key: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    } else if (msg.type === 'get-component-keys') {
        console.log('🔑 Getting component keys from selection...');
        try {
            const selection = figma.currentPage.selection;
            if (selection.length === 0) {
                figma.notify('⚠️ Please select a component to get its key');
                return;
            }
            
            const selectedNode = selection[0];
            console.log(`📋 Selected node type: ${selectedNode.type}`);
            console.log(`📋 Selected node name: ${selectedNode.name}`);
            console.log(`📋 Selected node id: ${selectedNode.id}`);
            
            if (selectedNode.type === 'COMPONENT') {
                const component = selectedNode as ComponentNode;
                const componentInfo = {
                    name: component.name,
                    key: component.key,
                    id: component.id,
                    description: component.description || '',
                    width: component.width,
                    height: component.height
                };
                
                figma.ui.postMessage({
                    type: 'component-key-found',
                    component: componentInfo
                });
                
                console.log(`✅ COMPONENT Key: ${component.key}`);
                console.log(`   Name: ${component.name}`);
                console.log(`💡 Add this to main.ts line 1101:`);
                console.log(`   '${component.key}', // ${component.name}`);
                figma.notify(`✅ Component key: ${component.key}`);
                
            } else if (selectedNode.type === 'INSTANCE') {
                const instance = selectedNode as InstanceNode;
                const component = instance.mainComponent;
                
                if (!component) {
                    console.log('❌ Instance has no main component - it may be detached');
                    figma.notify('⚠️ This instance has no main component - select the original component from the library');
                    return;
                }
                
                const componentInfo = {
                    name: component.name,
                    key: component.key,
                    id: component.id,
                    description: component.description || '',
                    width: component.width,
                    height: component.height
                };
                
                figma.ui.postMessage({
                    type: 'component-key-found',
                    component: componentInfo
                });
                
                console.log(`✅ COMPONENT Key (from instance): ${component.key}`);
                console.log(`   Name: ${component.name}`);
                console.log(`   Instance ID: ${instance.id} (this is NOT the key)`);
                console.log(`💡 Add this to main.ts line 1101:`);
                console.log(`   '${component.key}', // ${component.name}`);
                figma.notify(`✅ Component key: ${component.key} (from main component)`);
                
            } else if (selectedNode.type === 'COMPONENT_SET') {
                const componentSet = selectedNode as ComponentSetNode;
                console.log(`✅ COMPONENT SET Key: ${componentSet.key}`);
                console.log(`   Name: ${componentSet.name}`);
                console.log(`   Variants: ${componentSet.children.length}`);
                console.log(`💡 Add this to main.ts line 1101:`);
                console.log(`   '${componentSet.key}', // ${componentSet.name}`);
                
                // Log all variant keys
                console.log('📋 All variant keys:');
                componentSet.children.forEach((child, index) => {
                    if (child.type === 'COMPONENT') {
                        const variant = child as ComponentNode;
                        console.log(`   ${index + 1}. ${variant.name} - ${variant.key}`);
                    }
                });
                
                figma.notify(`✅ Component Set key: ${componentSet.key}`);
                
            } else {
                console.log(`❌ Selected node is type: ${selectedNode.type} (not a component)`);
                figma.notify(`⚠️ Please select a COMPONENT or INSTANCE (you selected: ${selectedNode.type})`);
                return;
            }
            
        } catch (error) {
            console.error('❌ Error getting component key:', error);
            figma.notify(`❌ Error getting component key: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    } else if (msg.type === 'discover-components') {
        console.log('🔍 Received discover-components message');
        (async () => {
            try {
                const availableComponents = await discoverAvailableComponents();
                
                // Calculate breakdown
                const pageComponents = figma.currentPage.findAll(node => node.type === 'COMPONENT') as ComponentNode[];
                const pageComponentIds = new Set(pageComponents.map(c => c.id));
                const libraryComponents = availableComponents.filter(c => !pageComponentIds.has(c.id));
                
                // Send the discovered components to UI
                figma.ui.postMessage({
                    type: 'components-discovered',
                    components: availableComponents.map(comp => ({
                        id: comp.id,
                        name: comp.name,
                        description: comp.description || '',
                        width: comp.width,
                        height: comp.height,
                        isFromLibrary: !pageComponentIds.has(comp.id)
                    })),
                    pageComponents: pageComponents.length,
                    libraryComponents: libraryComponents.length
                });
                
                console.log(`✅ Sent ${availableComponents.length} components to UI (${libraryComponents.length} from libraries)`);
            } catch (error) {
                console.error('❌ Error in discover-components:', error);
                figma.ui.postMessage({
                    type: 'components-discovered',
                    components: [],
                    error: error instanceof Error ? error.message : 'Unknown error'
                });
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
            bodyCell = await findBodyCellComponent();
        
        if (!bodyCell) {
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
            console.log(`[Backend] component-info msg.tableId:`, msg.tableId);
            console.log(`[Backend] component-info msg keys:`, Object.keys(msg));
            
            // Check if this is a generated table and load its cell properties
            let savedCellProperties = null;
            let actualTableData = null;
            
            // Try to get table ID from msg.tableId or current selection
            let tableId = msg.tableId;
            if (!tableId) {
                const selection = figma.currentPage.selection;
                if (selection.length === 1) {
                    tableId = selection[0].id;
                    console.log(`[Backend] No tableId in msg, using current selection: ${tableId}`);
                }
            }
            
            let actualRows: number | undefined;
            let actualCols: number | undefined;
            
            if (tableId) {
                const tableNode = figma.getNodeById(tableId);
                if (tableNode) {
                    // Check if this is a generated table
                    const isGeneratedTable = tableNode.getPluginData('isGeneratedTable') === 'true';
                    console.log(`[Backend] Table ${tableId} is generated table: ${isGeneratedTable}`);
                    
                    if (isGeneratedTable) {
                        // Use saved properties instead of extracting from table
                        console.log(`[Backend] Using saved properties for generated table`);
                        actualTableData = null; // Force use of saved properties
                    }
                    
                    const tableSettingsData = tableNode.getPluginData('tableSettings');
                    if (tableSettingsData) {
                        try {
                            const tableSettings = JSON.parse(tableSettingsData);
                            savedCellProperties = tableSettings.cellProperties;
                            actualRows = tableSettings.rows;
                            actualCols = tableSettings.columns;
                            console.log(`[Backend] Loaded ${Object.keys(savedCellProperties || {}).length} cell properties from table metadata`);
                            console.log(`[Backend] Actual table dimensions: ${actualRows} rows × ${actualCols} columns`);
                        } catch (e) {
                            console.warn('[Backend] Could not parse table settings:', e);
                        }
                    }
                }
            } else {
                console.log(`[Backend] No tableId available for component-info`);
            }
            
            console.log(`[Backend] ========================================`);
            console.log(`[Backend] Sending component-info to UI`);
            console.log(`[Backend] ========================================`);
            
            figma.ui.postMessage({
                type: 'component-info',
                component: {
                    id: bodyCell.id,
                    name: bodyCell.name,
                    width: bodyCell.width,
                    properties: properties,
                    availableProperties: availableProperties,
                    propertyTypes: propertyTypes
                },
                savedCellProperties: savedCellProperties, // Send the saved cell properties to UI
                actualTableData: actualTableData, // Send the actual current table data
                actualRows: actualRows, // Send actual row count from metadata
                actualCols: actualCols  // Send actual column count from metadata
            });
        }
        })();
    }
};

import { createComponentInstanceWithProps, ComponentProperty } from './utils/component';
import { getFakerValue } from './utils/faker';
import { cleanupExternalComponents } from './utils/cleanup';
import { findBodyCellComponent } from './utils/componentSearch';

// Call cleanup when the plugin starts
cleanupExternalComponents();

 