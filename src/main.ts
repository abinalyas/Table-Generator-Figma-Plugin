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
}

let lastScanResult: ScanResult | undefined;
let isCreatingTable = false;

figma.on("selectionchange", async () => {
    const selection = figma.currentPage.selection;

    if (selection.length === 1 && selection[0].type === "FRAME" && selection[0].getPluginData('isGeneratedTable') === 'true') {
        const tableFrame = selection[0] as FrameNode;
        const settings = tableFrame.getPluginData('tableSettings');
        if (settings) {
            try {
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

    if (selection.length === 1 && selection[0].type === "INSTANCE") {
        const instance = selection[0] as InstanceNode;
        try {
            selectedComponent = await instance.getMainComponentAsync();
            
            const propertyValues: { [key: string]: any } = {};
            const propertyTypes: { [key: string]: "TEXT" | "BOOLEAN" | "INSTANCE_SWAP" | "VARIANT" } = {};
            const availableProperties = Object.keys(instance.componentProperties);

            for (const propName of availableProperties) {
                const prop = instance.componentProperties[propName];
                propertyValues[propName] = prop.value;
                propertyTypes[propName] = prop.type;
            }
        
            figma.ui.postMessage({
                type: "component-selected",
                componentName: selection[0].name || "Table Cell",
                componentId: selection[0].id || null,
                componentWidth: Math.round(selection[0].width),
                properties: propertyValues,
                availableProperties: availableProperties,
                propertyTypes: propertyTypes,
                isValidComponent: true
            });
        }
        catch (error) {
            console.error('❌ Error getting main component:', error);
            selectedComponent = null;
            figma.ui.postMessage({
                type: "selection-cleared",
                isValidComponent: false
            });
        }
    } else {
        selectedComponent = null;
        figma.ui.postMessage({
            type: "selection-cleared",
            isValidComponent: false
        });
    }
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

async function createCheckboxComponent() {
    // Check if component already exists
    const existingComponents = figma.root.findAll(node => 
        node.type === "COMPONENT_SET" && node.name === "Data table select cell item"
    );
    
    if (existingComponents.length > 0) {
        const existingComponent = existingComponents[0] as ComponentSetNode;
        const firstChild = existingComponent.children[0];
        
        // Check if this is the old structure (20x20 instead of 52x64)
        if (firstChild && firstChild.width === 20 && firstChild.height === 20) {
            
            existingComponent.remove();
        } else if (firstChild && firstChild.width === 52 && firstChild.height === 64) {
            
            return existingComponent;
        }
    }


    // Create the main wrapper component
    const wrapperComponent = figma.createComponent();
    wrapperComponent.name = "Data table select cell item";
    wrapperComponent.resize(52, 64); // 52px width, 64px height
    
    // Set up auto-layout for the wrapper with specified padding
    wrapperComponent.layoutMode = "HORIZONTAL";
    wrapperComponent.primaryAxisAlignItems = "CENTER";
    wrapperComponent.counterAxisAlignItems = "CENTER";
    wrapperComponent.itemSpacing = 0;
    wrapperComponent.paddingLeft = 16; // 16px left padding
    wrapperComponent.paddingRight = 16; // 16px right padding
    wrapperComponent.paddingTop = 14; // 14px top padding
    wrapperComponent.paddingBottom = 30; // 30px bottom padding

    // Create the checkbox background (outer box) directly in the wrapper
    const checkboxBox = figma.createRectangle();
    checkboxBox.name = "Checkbox Box";
    checkboxBox.resize(20, 20);
    checkboxBox.x = 0;
    checkboxBox.y = 0;
    checkboxBox.cornerRadius = 2; // Slightly rounded corners
    checkboxBox.strokes = [{ type: 'SOLID', color: { r: 0.6, g: 0.6, b: 0.6 } }];
    checkboxBox.strokeWeight = 1;
    checkboxBox.fills = [{ type: 'SOLID', color: { r: 1, g: 1, b: 1 } }]; // White background

    // Create the checkmark directly in the wrapper
    const checkmark = figma.createVector();
    checkmark.name = "Checkmark";
    checkmark.resize(12, 12);
    checkmark.x = 4; // Center horizontally: (20-12)/2 = 4
    checkmark.y = 4; // Center vertically: (20-12)/2 = 4
    
    // Create a simple checkmark using vector network
    const vectorNetwork: VectorNetwork = {
        vertices: [
            { x: 2.5, y: 6.5, strokeCap: "ROUND", strokeJoin: "ROUND", cornerRadius: 0 },
            { x: 5.5, y: 9.5, strokeCap: "ROUND", strokeJoin: "ROUND", cornerRadius: 0 },
            { x: 9.5, y: 2.5, strokeCap: "ROUND", strokeJoin: "ROUND", cornerRadius: 0 }
        ],
        segments: [
            { start: 0, end: 1, tangentStart: { x: 0, y: 0 }, tangentEnd: { x: 0, y: 0 } },
            { start: 1, end: 2, tangentStart: { x: 0, y: 0 }, tangentEnd: { x: 0, y: 0 } }
        ],
        regions: []
    };
    
    checkmark.vectorNetwork = vectorNetwork;
    checkmark.strokeWeight = 2;
    checkmark.strokes = [{ type: 'SOLID', color: { r: 0.2, g: 0.2, b: 0.2 } }]; // Dark gray checkmark
    checkmark.fills = []; // No fill, just stroke
    
    // Add elements directly to the wrapper component
    wrapperComponent.appendChild(checkboxBox);
    wrapperComponent.appendChild(checkmark);
    
    // Ensure checkmark is positioned correctly relative to the checkbox box
    // The checkmark should be positioned at the center of the checkbox box
    // Since the wrapper has padding, we need to account for that
    checkmark.x = 16 + 4; // Left padding (16) + center offset (4)
    checkmark.y = 14 + 4; // Top padding (14) + center offset (4)

    // Create variants by cloning the wrapper component
    const uncheckedVariant = wrapperComponent.clone();
    uncheckedVariant.name = "Unchecked";
    const uncheckedCheckmark = uncheckedVariant.findOne(n => n.name === "Checkmark") as VectorNode;
    if (uncheckedCheckmark) uncheckedCheckmark.visible = false; // Hide checkmark for unchecked state
    
    const checkedVariant = wrapperComponent.clone();
    checkedVariant.name = "Checked";
    const checkedCheckmark = checkedVariant.findOne(n => n.name === "Checkmark") as VectorNode;
    if (checkedCheckmark) checkedCheckmark.visible = true; // Show checkmark for checked state
    
    const disabledUncheckedVariant = wrapperComponent.clone();
    disabledUncheckedVariant.name = "Disabled Unchecked";
    const disabledUncheckedCheckmark = disabledUncheckedVariant.findOne(n => n.name === "Checkmark") as VectorNode;
    if (disabledUncheckedCheckmark) disabledUncheckedCheckmark.visible = false;
    const disabledUncheckedBox = disabledUncheckedVariant.findOne(n => n.name === "Checkbox Box") as RectangleNode;
    if (disabledUncheckedBox) disabledUncheckedBox.opacity = 0.5; // Dim the checkbox for disabled state
    
    const disabledCheckedVariant = wrapperComponent.clone();
    disabledCheckedVariant.name = "Disabled Checked";
    const disabledCheckedCheckmark = disabledCheckedVariant.findOne(n => n.name === "Checkmark") as VectorNode;
    if (disabledCheckedCheckmark) disabledCheckedCheckmark.visible = true;
    const disabledCheckedBox = disabledCheckedVariant.findOne(n => n.name === "Checkbox Box") as RectangleNode;
    if (disabledCheckedBox) disabledCheckedBox.opacity = 0.5; // Dim the checkbox for disabled state

    // Create component set by combining variants
    const componentSet = figma.combineAsVariants([
        uncheckedVariant, 
        checkedVariant, 
        disabledUncheckedVariant, 
        disabledCheckedVariant
    ], figma.currentPage);
    componentSet.name = "Data table select cell item";

    // Add component property to the component set
    componentSet.addComponentProperty("State", "VARIANT", "Unchecked");

    return componentSet;
}


figma.ui.onmessage = async (msg: any) => {
    if (msg.type === "generate-fake-data") {
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
    
    } else if (msg.type === "request-selection-state") {
        figma.ui.postMessage({
            type: "component-selected",
            componentName: (selectedComponent?.name) || "Table Cell",
            componentId: (selectedComponent?.id) || null,
            componentWidth: selectedComponent ? Math.round(selectedComponent.width) : undefined,
            properties: {},
            availableProperties: [],
            propertyTypes: {},
            isValidComponent: selectedComponent !== null
        });
    
    } else if (msg.type === "clear-cell-instances") {
        cellInstanceMap.forEach(instance => instance.remove());

    } else if (msg.type === 'scan-table') {
        const selection = figma.currentPage.selection;
        if (selection.length !== 1 || !['FRAME', 'COMPONENT', 'INSTANCE'].includes(selection[0].type)) {
            figma.ui.postMessage({
                type: 'scan-table-result',
                success: false,
                message: 'Please select a single Frame (table component) in Figma.'
            });
            return;
        }

        const tableFrame = selection[0];
        
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
                                bodyCell = bodyCellInstance;
                            }
                        }
                    }
                }
            }
            // Extract footer template
            if (footerBar && footerBar.type === 'INSTANCE') {
                footer = footerBar;
            }
            
            // --- Scan for select/expand cell components AFTER updating body row properties ---
            if ('findAll' in tableFrame && typeof tableFrame.findAll === 'function') {
                allInstances = tableFrame.findAll(n => n.type === 'INSTANCE' || n.type === 'COMPONENT') as (InstanceNode | ComponentNode)[];
                for (const node of allInstances) {
                    if (node.name === 'Data table expand cell item') {
                      expandCellFound = true;
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        expandCellComponent = node.mainComponent;
                      }
                    }
                    if (node.name === 'Data table select cell item') {
                      selectCellFound = true;
                      if (node.type === 'INSTANCE' && node.mainComponent) {
                        selectCellComponent = node.mainComponent;
                      }
                    }
                    allInstanceNames.push(node.name);
                }
            }
            
            // If select cell component is still not found, try to create it
            if (!selectCellComponent) {
                console.log('🔄 Select cell component not found, attempting to create checkbox component...');
                try {
                    const checkboxComponentSet = await createCheckboxComponent();
                    if (checkboxComponentSet && checkboxComponentSet.children.length > 0) {
                        selectCellComponent = checkboxComponentSet.children[0] as ComponentNode;
                        console.log('✅ Created checkbox component for selectable cells');
                    }
                } catch (error) {
                    console.error('❌ Error creating checkbox component:', error);
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
            lastScanResult = {
                headerCell: headerCell ? headerCell.mainComponent : null,
                headerRowComponent,
                bodyCell: bodyCell ? bodyCell.mainComponent : null,
                bodyRowComponent,
                footer: footer ? footer.mainComponent : null,
                numCols,
                selectCellComponent,
                expandCellComponent
            };
       
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
                    tableFrame.appendChild(bodyWrapperFrame);

                    for (let r = 0; r < rows; r++) {
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
                        bodyWrapperFrame.appendChild(rowFrame);
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
                columns: msg.columns,
                rows: msg.rows,
                includeHeader: msg.includeHeader,
                includeFooter: msg.includeFooter,
                includeSelectable: msg.includeSelectable,
                includeExpandable: msg.includeExpandable,
                cellProperties: msg.cellProperties,
            };
            tableFrame.setPluginData('tableSettings', JSON.stringify(tableSettings));

            figma.currentPage.appendChild(tableFrame);
            figma.viewport.scrollAndZoomIntoView([tableFrame]);
            figma.notify('Table created successfully!');
            figma.ui.postMessage({ type: 'table-created', success: true });

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
        try {
            if (!lastScanResult) {
                figma.notify('❌ No scan result available. Please scan a table first.');
                isCreatingTable = false;
                return;
            }

            const tableFrame = figma.getNodeById(msg.tableId) as FrameNode;
            if (!tableFrame || tableFrame.type !== 'FRAME') {
                figma.notify("❌ Table to update not found.");
                return;
            }

            // Clear existing content
            tableFrame.children.slice().forEach(child => child.remove());

            // Re-use the creation logic, but target the existing frame
            const { headerCell, bodyCell, footer, numCols } = lastScanResult!;
            const { rows, cols, cellProps, includeHeader, includeFooter, includeSelectable, includeExpandable } = msg;

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
                    tableFrame.appendChild(bodyWrapperFrame);

                    for (let r = 0; r < rows; r++) {
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
                        bodyWrapperFrame.appendChild(rowFrame);
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

            // Update plugin data
            const tableSettings = {
                columns: msg.columns,
                rows: msg.rows,
                includeHeader: msg.includeHeader,
                includeFooter: msg.includeFooter,
                includeSelectable: msg.includeSelectable,
                includeExpandable: msg.includeExpandable,
                cellProperties: msg.cellProperties,
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
    
    // Helper to find matching property by base name (without ID)
    const findMatchingProp = (sourceProp: string, targetProps: { [key: string]: ComponentProperty }) => {
        const sourceBase = sourceProp.split('#')[0];
        return Object.keys(targetProps).find(p => p.split('#')[0] === sourceBase);
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
                // For INSTANCE_SWAP, always use the component's current value
                validProps[targetKey] = propDef.value;
            }

            console.log(`🔍 Mapped property: ${sourceKey} -> ${targetKey} = ${actualValue} (type: ${propDef.type})`);
        } else {
            console.warn(`⚠️ No matching property found for: ${sourceKey}`);
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
 