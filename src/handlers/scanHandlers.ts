import { performTableScan, ScanResult } from '../scanner';

let lastScanResult: ScanResult | undefined;
let lastScanTime = 0;
const SCAN_DEBOUNCE_MS = 500;

export async function handleTableScan(msg: any) {
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
    
    // Use the imported performTableScan function
    const scanResult = await performTableScan(tableFrame);
    
    if (scanResult) {
        lastScanResult = scanResult;
        
        // Get properties for the bodyCell template to send to UI
        let bodyCellComponent = null;
        if (scanResult.bodyCell && scanResult.bodyCell.type === 'COMPONENT') {
            try {
                const tempInstance = scanResult.bodyCell.createInstance();
                const propertyValues: { [key: string]: any } = {};
                const propertyTypes: { [key: string]: any } = {};
                const availableProperties = Object.keys(tempInstance.componentProperties);
                for (const propName of availableProperties) {
                    const prop = tempInstance.componentProperties[propName];
                    propertyValues[propName] = prop.value;
                    propertyTypes[propName] = prop.type;
                }
                const instanceWidth = tempInstance.width;
                tempInstance.remove();
                
                bodyCellComponent = {
                    id: scanResult.bodyCell.id,
                    name: scanResult.bodyCell.name,
                    width: instanceWidth,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes,
                };
            } catch (error) {
                console.error('Error accessing bodyCell mainComponent:', error);
            }
        }
        
        // Get properties for the headerCell template to send to UI
        let headerCellComponent = null;
        if (scanResult.headerCell && scanResult.headerCell.type === 'COMPONENT') {
            try {
                const tempInstance = scanResult.headerCell.createInstance();
                const propertyValues: { [key: string]: any } = {};
                const propertyTypes: { [key: string]: any } = {};
                const availableProperties = Object.keys(tempInstance.componentProperties);
                for (const propName of availableProperties) {
                    const prop = tempInstance.componentProperties[propName];
                    propertyValues[propName] = prop.value;
                    propertyTypes[propName] = prop.type;
                }
                const instanceWidth = tempInstance.width;
                tempInstance.remove();
                
                headerCellComponent = {
                    id: scanResult.headerCell.id,
                    name: scanResult.headerCell.name,
                    width: instanceWidth,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes,
                };
            } catch (error) {
                console.error('Error accessing headerCell mainComponent:', error);
            }
        }
        
        // Get properties for the footer template to send to UI
        let footerComponent = null;
        if (scanResult.footer && scanResult.footer.type === 'COMPONENT') {
            try {
                const tempInstance = scanResult.footer.createInstance();
                const propertyValues: { [key: string]: any } = {};
                const propertyTypes: { [key: string]: any } = {};
                const availableProperties = Object.keys(tempInstance.componentProperties);
                for (const propName of availableProperties) {
                    const prop = tempInstance.componentProperties[propName];
                    propertyValues[propName] = prop.value;
                    propertyTypes[propName] = prop.type;
                }
                const instanceWidth = tempInstance.width;
                tempInstance.remove();
                
                footerComponent = {
                    id: scanResult.footer.id,
                    name: scanResult.footer.name,
                    width: instanceWidth,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes,
                };
            } catch (error) {
                console.error('Error accessing footer mainComponent:', error);
            }
        }
        
        // Send summary to UI
        figma.ui.postMessage({
            type: 'scan-table-result',
            success: true,
            message: `Tap on the cells below to personalize your table.`,
            details: {
                footer: !!scanResult.footer,
                numCols: scanResult.numCols,
                bodyCellComponent,
                headerCellComponent,
                footerComponent,
                expandCellComponent: scanResult.expandCellComponent,
                selectCellComponent: scanResult.selectCellComponent,
                bodyRowComponent: scanResult.bodyRowComponent
            }
        });
    } else {
        figma.ui.postMessage({
            type: 'scan-table-result',
            success: false,
            message: 'Could not scan the selected table. Please ensure it contains valid table components.'
        });
    }
}

export function handleRequestComponentInfo(msg: any) {
    const { tableId } = msg;
    const tableNode = figma.getNodeById(tableId);
    
    if (tableNode && lastScanResult && lastScanResult.bodyCell) {
        try {
            const tempInstance = lastScanResult.bodyCell.createInstance();
            const propertyValues: { [key: string]: any } = {};
            const propertyTypes: { [key: string]: any } = {};
            const availableProperties = Object.keys(tempInstance.componentProperties);
            
            for (const propName of availableProperties) {
                const prop = tempInstance.componentProperties[propName];
                propertyValues[propName] = prop.value;
                propertyTypes[propName] = prop.type;
            }
            
            const width = tempInstance.width;
            tempInstance.remove();
            
            figma.ui.postMessage({
                type: 'component-info',
                component: {
                    id: lastScanResult.bodyCell.id,
                    name: lastScanResult.bodyCell.name,
                    width: width,
                    properties: propertyValues,
                    availableProperties,
                    propertyTypes
                }
            });
        } catch (error) {
            console.error('Error getting component info:', error);
            figma.ui.postMessage({ type: 'component-info', component: null });
        }
    } else {
        figma.ui.postMessage({ type: 'component-info', component: null });
    }
}

export { lastScanResult };
export function getLastScanResult() {
    return lastScanResult;
}