const faker = require('faker');

figma.showUI(__html__, { width: 360, height: 700 });

let selectedComponent: ComponentNode | null = null;
const cellInstanceMap = new Map<string, InstanceNode>();

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
                }
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

// Helper function to find border color variable
function findBorderColorVariable(): VariableAlias | null {
    try {
        // Try to get local variables
        const localVariables = figma.variables.getLocalVariables();
        console.log('🔍 Available local variables:', localVariables.map(v => v.name));
        
        // If no local variables, try to get all variables
        if (localVariables.length === 0) {
            console.log('🔍 No local variables found, trying to get all variables...');
            try {
                // Try to get variables from the current page or document
                console.log('🔍 Trying alternative methods to find variables...');
                // Note: getLocalVariablesByMode() doesn't exist, so we'll try other approaches
                
                // Also try to get variables from the current selection
                const selection = figma.currentPage.selection;
                if (selection.length > 0) {
                    console.log('🔍 Checking variables from current selection...');
                    // Try to get variables from the selected node's parent or document
                    const parent = selection[0].parent;
                    if (parent) {
                        console.log('🔍 Parent node:', parent.name, parent.type);
                    }
                }
            } catch (error) {
                console.log('🔍 Error getting all variables:', error);
            }
        }
        
        // Look for common border variable names
        const borderVariableNames = ['border-subtle-02', 'border-subtle-01', 'border-subtle', 'border-01', 'border-subtle-1'];
        let borderVariable = null;
        
        for (const variableName of borderVariableNames) {
            borderVariable = localVariables.find(v => v.name === variableName);
            if (borderVariable) {
                console.log(`🎨 Found border variable: ${variableName}`);
                break;
            }
        }
        
        if (borderVariable) {
            console.log('✅ Using border color variable for empty fills');
            return {
                id: borderVariable.id,
                type: 'VARIABLE_ALIAS'
            };
        } else {
            console.log('⚠️ No border color variable found');
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
            
            // --- Scan for select/expand cell components and divider ---
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
                    
                    // Check for divider components
                    const isDivider = node.name === 'Divider' || 
                                    node.name.includes('divider') || 
                                    node.name.toLowerCase().includes('divider') ||
                                    node.name.includes('line') ||
                                    node.name.includes('border') ||
                                    node.name.includes('separator') ||
                                    node.name.includes('hr') ||
                                    node.name.includes('horizontal') ||
                                    node.name.includes('rule');
                    
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
            
            // If divider component is not found, try to create a simple one
            if (!dividerComponent) {
                console.log('🔄 Divider component not found, attempting to create simple divider...');
                try {
                    dividerComponent = createSimpleDivider();
                    console.log('✅ Created simple divider component');
                    // Move the divider component to a hidden location to avoid cluttering the document
                    dividerComponent.x = -10000;
                    dividerComponent.y = -10000;
                } catch (error) {
                    console.error('❌ Error creating divider component:', error);
                }
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

// Function to create a simple divider component
function createSimpleDivider(): ComponentNode {
    // Create a simple horizontal line divider
    const dividerComponent = figma.createComponent();
    dividerComponent.name = "Simple Divider";
    dividerComponent.resize(800, 1); // Default width, 1px height to match scanned table
    
    // Create a rectangle for the divider line
    const dividerLine = figma.createRectangle();
    dividerLine.name = "Divider Line";
    dividerLine.resize(800, 1); // Match the height
    dividerLine.x = 0;
    dividerLine.y = 0;
    
    // Try to apply border-subtle-02 color variable
    const borderVariable = findBorderColorVariable();
    if (borderVariable) {
        try {
            const colorVariableFill: Paint = {
                type: 'SOLID',
                color: { r: 0, g: 0, b: 0 }, // Default color (will be overridden by variable)
                boundVariables: {
                    color: borderVariable
                }
            };
            dividerLine.fills = [colorVariableFill];
            console.log('✅ Applied border color variable to simple divider line');
        } catch (error) {
            console.error('❌ Error applying border color variable to simple divider:', error);
            // Fallback to a subtle gray color
            dividerLine.fills = [{
                type: 'SOLID',
                color: { r: 0.9, g: 0.9, b: 0.9 } // Light gray
            }];
            console.log('✅ Applied fallback gray color to simple divider line');
        }
    } else {
        console.log('⚠️ No border color variable found, using fallback gray color for simple divider');
        // Apply a subtle gray color as fallback
        dividerLine.fills = [{
            type: 'SOLID',
            color: { r: 0.9, g: 0.9, b: 0.9 } // Light gray
        }];
        console.log('✅ Applied fallback gray color to simple divider line');
    }
    dividerLine.strokes = []; // No stroke
    
    // Add the line to the component
    dividerComponent.appendChild(dividerLine);
    
    // Set auto-layout properties for responsive behavior
    dividerComponent.layoutMode = "HORIZONTAL";
    dividerComponent.primaryAxisSizingMode = "AUTO"; // Auto sizing
    dividerComponent.counterAxisSizingMode = "FIXED"; // Keep height fixed
    dividerComponent.itemSpacing = 0;
    dividerComponent.paddingLeft = 0;
    dividerComponent.paddingRight = 0;
    dividerComponent.paddingTop = 0;
    dividerComponent.paddingBottom = 0;
    
    // Set the divider line to fill the container width
    dividerLine.layoutGrow = 1; // Make the line grow to fill available space
    
    // Clear default fill to make it theme-compatible
    dividerComponent.fills = [];
    
    console.log('✅ Created simple divider with border color variable and responsive layout');
    
    return dividerComponent;
}


figma.ui.onmessage = async (msg: any) => {
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
    
    } else if (msg.type === "load-watsonx-settings") {
        try {
            const endpoint = await figma.clientStorage.getAsync('watsonx.endpoint');
            const apiKey = await figma.clientStorage.getAsync('watsonx.apiKey');
            const masked = apiKey ? `${String(apiKey).slice(0,4)}••••${String(apiKey).slice(-4)}` : '';
            figma.ui.postMessage({ type: 'watsonx-settings', endpoint, apiKeyMasked: masked });
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
    
    } else if (msg.type === "clear-cell-instances") {
        cellInstanceMap.forEach(instance => instance.remove());

    } else if (msg.type === 'scan-table') {
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
                
                // Update header row properties to enable selectable/expandable functionality
                console.log('🔄 Updating header row properties to add select/expand cells...');
                const headerUpdateSuccess = await updateBodyRowProperties(headerRow);
                if (headerUpdateSuccess) {
                    console.log('✅ Header row properties updated successfully');
                    // Wait a bit more for Figma to fully update the table structure
                    await new Promise(resolve => setTimeout(resolve, 200));
                } else {
                    console.warn('⚠️ Failed to update header row properties');
                }
                
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
                    
                    // Update body row properties to enable selectable/expandable functionality
                    console.log('🔄 Updating body row properties to add select/expand cells...');
                    const updateSuccess = await updateBodyRowProperties(bodyRowInstance);
                    if (updateSuccess) {
                        console.log('✅ Body row properties updated successfully');
                        // Wait a bit more for Figma to fully update the table structure
                        await new Promise(resolve => setTimeout(resolve, 300));
                    } else {
                        console.warn('⚠️ Failed to update body row properties');
                    }
                    
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
                    // Check for divider components with various common names
                    const isDivider = node.name === 'Divider' || 
                                    node.name.includes('divider') || 
                                    node.name.toLowerCase().includes('divider') ||
                                    node.name.includes('line') ||
                                    node.name.includes('border') ||
                                    node.name.includes('separator') ||
                                    node.name.includes('hr') ||
                                    node.name.includes('horizontal') ||
                                    node.name.includes('rule');
                    
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
            
            // If divider component is not found, try to create a simple one
            if (!dividerComponent) {
                console.log('🔄 Divider component not found, attempting to create simple divider...');
                try {
                    dividerComponent = createSimpleDivider();
                    console.log('✅ Created simple divider component');
                    // Move the divider component to a hidden location to avoid cluttering the document
                    dividerComponent.x = -10000;
                    dividerComponent.y = -10000;
                } catch (error) {
                    console.error('❌ Error creating divider component:', error);
                }
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
                    console.log('🔍 Temporary divider instance properties:', {
                        fills: tempDividerInstance.fills,
                        strokes: tempDividerInstance.strokes,
                        strokeWeight: tempDividerInstance.strokeWeight
                    });
                    
                    // Store instance properties
                    dividerProperties.instanceFills = tempDividerInstance.fills;
                    dividerProperties.instanceStrokes = tempDividerInstance.strokes;
                    dividerProperties.instanceStrokeWeight = tempDividerInstance.strokeWeight;
                    dividerProperties.instanceHeight = tempDividerInstance.height; // Capture instance height
                    
                    // Check for color variables in instance fills
                    if (tempDividerInstance.fills && Array.isArray(tempDividerInstance.fills) && tempDividerInstance.fills.length > 0) {
                        for (const fill of tempDividerInstance.fills) {
                            if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                                console.log('🎨 Found color variable in instance:', fill.boundVariables.color);
                                dividerProperties.instanceColorVariable = fill.boundVariables.color;
                            }
                        }
                    }
                    
                    // Check instance children
                    if ('children' in tempDividerInstance) {
                        console.log('🔍 Temporary divider instance children:', tempDividerInstance.children.length);
                        for (const child of tempDividerInstance.children) {
                            console.log('🔍 Temporary instance child:', child.name, child.type, {
                                fills: 'fills' in child ? child.fills : 'N/A',
                                strokes: 'strokes' in child ? child.strokes : 'N/A',
                                strokeWeight: 'strokeWeight' in child ? child.strokeWeight : 'N/A'
                            });
                            
                            if ('fills' in child && child.fills && Array.isArray(child.fills) && child.fills.length > 0) {
                                dividerProperties.instanceChildFills = child.fills;
                                console.log('✅ Found instance child fills:', child.fills);
                                
                                // Check for color variables in fills
                                for (const fill of child.fills) {
                                    if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                                        console.log('🎨 Found color variable in instance child:', fill.boundVariables.color);
                                        dividerProperties.instanceChildColorVariable = fill.boundVariables.color;
                                    }
                                }
                            }
                            if ('strokes' in child && child.strokes && Array.isArray(child.strokes) && child.strokes.length > 0) {
                                dividerProperties.instanceChildStrokes = child.strokes;
                                console.log('✅ Found instance child strokes:', child.strokes);
                            }
                            if ('strokeWeight' in child && child.strokeWeight !== undefined) {
                                dividerProperties.instanceChildStrokeWeight = child.strokeWeight;
                                console.log('✅ Found instance child stroke weight:', child.strokeWeight);
                            }
                            if ('height' in child && child.height !== undefined) {
                                dividerProperties.instanceChildHeight = child.height;
                                console.log('✅ Found instance child height:', child.height);
                            }
                            break;
                        }
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
        const cellProps = msg.cellProps || {};

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
                    columnWidths[c] = 96;
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
                        hCell.setProperties(validProps);
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
                        
                        // Add divider after the data row (except for the last row)
                        if (r < rows - 1) {
                            if (lastScanResult.dividerComponent) {
                                try {
                                    const divider = lastScanResult.dividerComponent.createInstance();
                                    
                                    // Apply divider properties from scanned table
                                    if (lastScanResult && (lastScanResult as any).dividerProperties) {
                                        const dividerProps = (lastScanResult as any).dividerProperties;
                                        console.log('🔄 Applying divider properties:', dividerProps);
                                        
                                        try {
                                                                                // Apply fills if present in the original divider
                                    if (dividerProps.fills && Array.isArray(dividerProps.fills) && dividerProps.fills.length > 0) {
                                        console.log('🎨 Applying divider fills:', dividerProps.fills);
                                        divider.fills = dividerProps.fills;
                                    } else {
                                        // If fills are empty, try to apply border-subtle-01 color variable
                                        console.log('🎨 Divider fills are empty, checking for border color variable');
                                        const borderVariable = findBorderColorVariable();
                                        if (borderVariable) {
                                            try {
                                                const colorVariableFill: Paint = {
                                                type: 'SOLID',
                                                color: { r: 0, g: 0, b: 0 }, // Default color (will be overridden by variable)
                                                boundVariables: {
                                                        color: borderVariable
                                                    }
                                                };
                                                divider.fills = [colorVariableFill];
                                                console.log('✅ Applied border-subtle-01 color variable to divider');
                                            } catch (error) {
                                                console.error('❌ Error applying border color variable to divider:', error);
                                            }
                                        } else {
                                            console.log('⚠️ No border color variable found, applying fallback gray color');
                                            // Apply a subtle gray color as fallback
                                            divider.fills = [{
                                                type: 'SOLID',
                                                color: { r: 0.9, g: 0.9, b: 0.9 } // Light gray
                                            }];
                                            console.log('✅ Applied fallback gray color to divider');
                                        }
                                    }
                                    
                                    // Apply color variable if present
                                            if (dividerProps.colorVariable) {
                                                console.log('🎨 Applying color variable to divider:', dividerProps.colorVariable);
                                                try {
                                                    // Create a fill with the color variable
                                                    const colorVariableFill: Paint = {
                                                        type: 'SOLID',
                                                        color: { r: 0, g: 0, b: 0 }, // Default color (will be overridden by variable)
                                                        boundVariables: {
                                                            color: dividerProps.colorVariable
                                                        }
                                                    };
                                                    divider.fills = [colorVariableFill];
                                                } catch (error) {
                                                    console.error('❌ Error applying color variable to divider:', error);
                                                }
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
                                console.log('⚠️ No divider component available for row', r + 1);
                            }
                        } else {
                            console.log('ℹ️ Skipping divider for last row', r + 1);
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
            const { rows, cols, cellProps, includeHeader, includeFooter, includeSelectable, includeExpandable } = msg;
            
            // Get the original table settings to preserve header/footer state
            let originalIncludeHeader = includeHeader;
            let originalIncludeFooter = includeFooter;
            let originalIncludeSelectable = includeSelectable;
            let originalIncludeExpandable = includeExpandable;
            
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
                    columnWidths[c] = 96;
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
                                hCell.setProperties(validProps);
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
                
                // Add divider after the data row (except for the last row)
                if (r < rows - 1) {
                    if (lastScanResult.dividerComponent) {
                        try {
                            const divider = lastScanResult.dividerComponent.createInstance();
                            
                            // Apply divider properties from scanned table
                            if (lastScanResult && (lastScanResult as any).dividerProperties) {
                                const dividerProps = (lastScanResult as any).dividerProperties;
                                console.log('🔄 Applying divider properties (update):', dividerProps);
                                
                                try {
                                    // Apply fills if present in the original divider
                                    if (dividerProps.fills && Array.isArray(dividerProps.fills) && dividerProps.fills.length > 0) {
                                        console.log('🎨 Applying divider fills (update):', dividerProps.fills);
                                        divider.fills = dividerProps.fills;
                                    }
                                    
                                    // Apply strokes if present in the original divider
                                    if (dividerProps.strokes && Array.isArray(dividerProps.strokes) && dividerProps.strokes.length > 0) {
                                        console.log('🎨 Applying divider strokes (update):', dividerProps.strokes);
                                        divider.strokes = dividerProps.strokes;
                                    }
                                    
                                    // Apply stroke weight if present
                                    if (dividerProps.strokeWeight !== undefined && dividerProps.strokeWeight > 0) {
                                        console.log('🎨 Applying divider stroke weight (update):', dividerProps.strokeWeight);
                                        divider.strokeWeight = dividerProps.strokeWeight;
                                    }
                                    
                                    // Apply corner radius if present
                                    if (dividerProps.cornerRadius !== undefined && dividerProps.cornerRadius > 0) {
                                        console.log('🎨 Applying divider corner radius (update):', dividerProps.cornerRadius);
                                        divider.cornerRadius = dividerProps.cornerRadius;
                                    }
                                    
                                    // Apply effects if present
                                    if (dividerProps.effects && Array.isArray(dividerProps.effects) && dividerProps.effects.length > 0) {
                                        console.log('🎨 Applying divider effects (update):', dividerProps.effects);
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
                    } else {
                        console.log('⚠️ No divider component available for row', r + 1, '(update)');
                    }
                } else {
                    console.log('ℹ️ Skipping divider for last row', r + 1, '(update)');
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
                cellProperties: msg.cellProps,
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

function mapPropertyNames(cellData: any, componentProps: { [key: string]: ComponentProperty }): { [key: string]: any } {
    const validProps: { [key: string]: any } = {};
    
    console.log('🔍 mapPropertyNames input:', { cellData, componentProps });
    console.log('🔍 Available component properties:', Object.keys(componentProps));
    
    // Helper to find matching property by base name (without ID)
    const findMatchingProp = (sourceProp: string, targetProps: { [key: string]: ComponentProperty }) => {
        const sourceBase = sourceProp.split('#')[0];
        
        // First try exact match
        let match = Object.keys(targetProps).find(p => p === sourceProp);
        if (match) return match;
        
        // Then try base name match (without ID)
        match = Object.keys(targetProps).find(p => p.split('#')[0] === sourceBase);
        if (match) return match;
        
        // For Slot properties, try more flexible matching
        if (sourceBase.toLowerCase() === 'slot') {
            match = Object.keys(targetProps).find(p => 
                p.toLowerCase().includes('slot') || 
                p.toLowerCase().includes('icon') ||
                p.toLowerCase().includes('content')
            );
            if (match) return match;
        }
        
        return null;
    };

    // Map each property from cellData to the correct component property
    for (const [sourceKey, sourceValue] of Object.entries(cellData)) {
        // Find matching property in component
        const targetKey = findMatchingProp(sourceKey, componentProps);
        
        if (targetKey) {
            const propDef = componentProps[targetKey];
            
            // Get the actual value (handle case where sourceValue is an object with value property)
            const actualValue = (sourceValue && typeof sourceValue === 'object' && 'value' in sourceValue) 
                ? sourceValue.value 
                : sourceValue;

            // Handle different property types correctly
            if (propDef.type === 'BOOLEAN') {
                validProps[targetKey] = !!actualValue;
            } else if (propDef.type === 'TEXT') {
                validProps[targetKey] = String(actualValue);
            } else if (propDef.type === 'VARIANT') {
                validProps[targetKey] = actualValue;
            } else if (propDef.type === 'INSTANCE_SWAP') {
                // For INSTANCE_SWAP, use the user's selected value if provided, otherwise keep current
                if (actualValue !== undefined && actualValue !== null) {
                    validProps[targetKey] = actualValue;
                } else {
                validProps[targetKey] = propDef.value;
                }
            }

            console.log(`🔍 Mapped property: ${sourceKey} -> ${targetKey} = ${actualValue} (type: ${propDef.type})`);
            if (propDef.type === 'INSTANCE_SWAP') {
                console.log(`🔍 INSTANCE_SWAP property details:`, {
                    sourceKey,
                    targetKey,
                    actualValue,
                    componentValue: propDef.value,
                    finalValue: validProps[targetKey]
                });
            }
        } else {
            console.warn(`⚠️ No matching property found for: ${sourceKey}`);
            // Special logging for Slot properties
            if (sourceKey.toLowerCase().includes('slot')) {
                console.log(`🔍 Slot property search details:`);
                console.log(`  - Source property: ${sourceKey}`);
                console.log(`  - Available properties:`, Object.keys(componentProps));
                console.log(`  - Properties containing 'slot':`, Object.keys(componentProps).filter(p => p.toLowerCase().includes('slot')));
            }
        }
    }

    // Preserve any existing INSTANCE_SWAP values
    for (const [key, prop] of Object.entries(componentProps)) {
        if (prop.type === 'INSTANCE_SWAP' && !(key in validProps)) {
            validProps[key] = prop.value;
        }
    }

    console.log('🔍 mapPropertyNames output:', validProps);
    return validProps;
}

// Utility: Check if a variant combination is valid for a component set
function isValidVariantCombination(componentSet: ComponentSetNode, properties: { [key: string]: any }): boolean {
    console.log('Checking variant combination:', properties);
    console.log('Component set has', componentSet.children.length, 'variants');
    
    // Each child is a variant (ComponentNode) with variantProperties
    const isValid = componentSet.children.some(variant => {
        const variantNode = variant as ComponentNode;
        console.log('Checking variant:', variantNode.name, 'with properties:', variantNode.variantProperties);
        
        // Only compare keys that are in properties
        const matches = Object.entries(properties).every(([key, value]) => {
            const variantValue = variantNode.variantProperties && variantNode.variantProperties[key];
            const isMatch = variantValue == value;
            console.log(`  ${key}: expected=${value}, actual=${variantValue}, match=${isMatch}`);
            // Figma may use string values for variantProperties
            return isMatch;
        });
        
        console.log('  Variant match result:', matches);
        return matches;
    });
    
    console.log('Final validation result:', isValid);
    return isValid;
}

// Function to update body row properties to enable selectable/expandable functionality
async function updateBodyRowProperties(bodyRowInstance: InstanceNode) {
    try {
        console.log('🔄 Updating body row properties to enable selectable/expandable functionality...');
        
        // Define the properties to set
        const propertiesToSet: { [key: string]: string } = {
            'Expandable': 'True',
            'Select type': 'Checkbox', 
            'Selectable': 'True',
            'Selection': 'Checked'
        };
        
        // Validate each property exists and get the correct variant value
        const validProperties: { [key: string]: string } = {};
        
        for (const [propName, targetValue] of Object.entries(propertiesToSet)) {
            if (bodyRowInstance.componentProperties[propName]) {
                const prop = bodyRowInstance.componentProperties[propName];
                if (prop.type === 'VARIANT') {
                    // Get the main component to access variant values
                    const mainComponent = bodyRowInstance.mainComponent;
                    if (mainComponent && mainComponent.parent && mainComponent.parent.type === "COMPONENT_SET") {
                        const componentSet = mainComponent.parent as ComponentSetNode;
                        const variantProps = componentSet.variantGroupProperties;
                        
                        if (variantProps && variantProps[propName] && 'values' in variantProps[propName]) {
                            const possibleValues = (variantProps[propName] as any).values as string[];
                            
                            // Find a case-insensitive match for the target value
                            const found = possibleValues.find(v => 
                                v.trim().toLowerCase() === targetValue.toLowerCase()
                            );
                            
                            if (found) {
                                validProperties[propName] = found;
                                console.log(`✅ Found variant value for ${propName}: ${found}`);
                            } else {
                                console.warn(`⚠️ Target value "${targetValue}" not found for property "${propName}". Available values:`, possibleValues);
                            }
                        }
                    }
                } else {
                    console.warn(`⚠️ Property "${propName}" is not a VARIANT type, it's ${prop.type}`);
                }
            } else {
                console.warn(`⚠️ Property "${propName}" not found in body row component`);
            }
        }
        
        // Check if the combination is valid before setting
        let mainComponentSet = null;
        if (bodyRowInstance.mainComponent && bodyRowInstance.mainComponent.parent && bodyRowInstance.mainComponent.parent.type === "COMPONENT_SET") {
            mainComponentSet = bodyRowInstance.mainComponent.parent as ComponentSetNode;
        }
        
        if (mainComponentSet && isValidVariantCombination(mainComponentSet, validProperties)) {
            try {
                bodyRowInstance.setProperties(validProperties);
                console.log('✅ Successfully updated body row properties:', validProperties);
                
                // Wait for Figma to update the instance tree
                await Promise.resolve();
                // Add additional delay to ensure the instance tree is updated
                await new Promise(resolve => setTimeout(resolve, 200));
                
                return true;
            } catch (error) {
                console.error('❌ Error setting body row properties:', error);
                return false;
            }
        } else {
            console.warn('⚠️ Invalid variant combination for body row:', validProperties);
            return false;
        }
    } catch (error) {
        console.error('❌ Error updating body row properties:', error);
        return false;
    }
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

 