// Function to search for body cell component across all pages
export async function findBodyCellComponent(): Promise<ComponentNode | null> {
    console.log(`[Backend] lastScanResult not available, searching for body cell component...`);
    
    let bodyCell: ComponentNode | null = null;
    
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
        console.log(`[Backend] No Data table found or extraction failed, trying broader search on current page only...`);
    
        // Search for components that might be the body cell component
        // Only search in the current page to avoid loading all pages
        const allComponents: ComponentNode[] = [];
        
        console.log(`[Backend] Searching for components on current page only...`);
        
        // Search only in the current page to avoid loading all pages
        const currentPageComponents = figma.currentPage.findAll(node => 
            node.type === "COMPONENT"
        ) as ComponentNode[];
        console.log(`[Backend] Current page has ${currentPageComponents.length} components:`, currentPageComponents.map(c => c.name));
        allComponents.push(...currentPageComponents);
        
        // Remove duplicates based on ID
        const uniqueComponents = allComponents.filter((comp, index, self) => 
            index === self.findIndex(c => c.id === comp.id)
        );
        
        console.log(`[Backend] Found ${uniqueComponents.length} total components on current page:`, uniqueComponents.map(c => c.name));
        
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
            }
        }
    }
    
    return bodyCell;
}

// This function needs to be declared here since it's used by findBodyCellComponent
// but it's part of the main scanning logic that we shouldn't move
declare function performTableScan(tableFrame: SceneNode): Promise<any>;