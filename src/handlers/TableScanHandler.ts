// Handler for table scanning operations
import { Logger } from '../utils/logger';

export class TableScanHandler {
    private static lastScanTime = 0;
    private static readonly SCAN_DEBOUNCE_MS = 500;

    static async handleScanTable(msg: any): Promise<void> {
        const now = Date.now();
        if (now - this.lastScanTime < this.SCAN_DEBOUNCE_MS) {
            Logger.debug(`Debouncing scan - ${now - this.lastScanTime}ms since last scan`);
            return;
        }
        this.lastScanTime = now;

        const selection = figma.currentPage.selection;
        Logger.debug(`Selection check - count: ${selection.length}, types: ${selection.map(s => s.type).join(', ')}, names: ${selection.map(s => s.name).join(', ')}`);

        if (selection.length !== 1 || !['FRAME', 'COMPONENT', 'INSTANCE'].includes(selection[0].type)) {
            Logger.debug(`Invalid selection - count: ${selection.length}, type: ${selection[0]?.type}`);
            figma.ui.postMessage({
                type: 'scan-table-result',
                success: false,
                message: 'Please select a single table component (Frame, Component or Instance) in Figma.'
            });
            return;
        }

        Logger.progress(`Valid selection - type: ${selection[0].type}, name: ${selection[0].name}`);
        const tableFrame = selection[0];

        // Table property analysis and type setting
        await this.analyzeAndSetTableProperties(tableFrame);

        // Perform the actual table scan
        const scanResult = await this.performTableScan(tableFrame);
        
        if (scanResult) {
            Logger.success('Table scan completed successfully');
            figma.ui.postMessage({
                type: 'scan-table-result',
                success: true,
                details: {
                    bodyCellComponent: scanResult.bodyCell,
                    headerCellComponent: scanResult.headerCell,
                    footerComponent: scanResult.footer,
                    numCols: scanResult.numCols || 5,
                    selectCellComponent: scanResult.selectCellComponent,
                    expandCellComponent: scanResult.expandCellComponent,
                    dividerComponent: scanResult.dividerComponent
                }
            });
        } else {
            Logger.error('Table scan failed');
            figma.ui.postMessage({
                type: 'scan-table-result',
                success: false,
                message: 'Failed to scan table. Please try selecting a different table component.'
            });
        }
    }

    private static async analyzeAndSetTableProperties(tableFrame: any): Promise<void> {
        // Analyze table properties and set Type to 'Expandable + Selectable'
        Logger.debug(`TABLE ANALYSIS: ${tableFrame.type} - ${tableFrame.name}`);
        
        if (tableFrame.type === 'INSTANCE' && 'componentProperties' in tableFrame) {
            const componentProps = Object.keys(tableFrame.componentProperties || {});
            Logger.debug(`Table has ${componentProps.length} properties:`, componentProps);
            
            if (componentProps.length > 0) {
                const propValues: { [key: string]: any } = {};
                for (const prop of componentProps) {
                    propValues[prop] = tableFrame.componentProperties[prop].value;
                }
                Logger.debug('Table property values:', propValues);
                
                // Set table Type property to 'Expandable + Selectable' to enable all necessary components
                Logger.progress('Setting table Type property to enable expandable + selectable features');
                
                // Look for Type property (common names: Type, Variant, Style, etc.)
                const typeProperties = componentProps.filter(prop => 
                    prop.toLowerCase().includes('type') || 
                    prop.toLowerCase().includes('variant') ||
                    prop.toLowerCase().includes('style')
                );
                
                Logger.debug('Found potential type properties:', typeProperties);

                for (const typeProp of typeProperties) {
                    try {
                        Logger.debug(`Setting ${typeProp} to 'Expandable + Selectable'`);
                        tableFrame.setProperties({ [typeProp]: 'Expandable + Selectable' });
                        Logger.success(`Successfully set ${typeProp} to 'Expandable + Selectable'`);
                        break; // Stop on first success
                    } catch (error) {
                        Logger.debug(`Failed to set ${typeProp}:`, (error as Error).message);
                        
                        // Try alternative values
                        const alternativeValues = [
                            'Expandable+Selectable',
                            'Selectable + Expandable', 
                            'Selectable+Expandable',
                            'Both',
                            'All',
                            'Full'
                        ];
                        
                        for (const altValue of alternativeValues) {
                            try {
                                Logger.debug(`Trying alternative value ${typeProp}: '${altValue}'`);
                                tableFrame.setProperties({ [typeProp]: altValue });
                                Logger.success(`Successfully set ${typeProp} to '${altValue}'`);
                                break;
                            } catch (altError) {
                                Logger.debug(`Alternative value '${altValue}' failed:`, (altError as Error).message);
                            }
                        }
                    }
                }
            }
        }
    }

    private static async performTableScan(tableFrame: any): Promise<any> {
        Logger.info('Performing table scan...');
        
        try {
            // For now, return a basic scan result with the expected structure
            // TODO: Implement full table scanning logic from the original main.ts
            
            // Try to find basic components in the table
            let headerCell = null;
            let bodyCell = null;
            let footer = null;
            let numCols = 5;
            
            // Enhanced scanning - look for common table components with flexible matching
            if ('children' in tableFrame) {
                // Look for body cells with more flexible matching
                const instances = tableFrame.findAll((node: any) => 
                    node.type === 'INSTANCE' && 
                    (node.name?.toLowerCase().includes('cell') || 
                     node.name?.toLowerCase().includes('row') ||
                     node.name?.toLowerCase().includes('item'))
                );
                
                Logger.debug(`Found ${instances.length} potential cell instances`);
                
                if (instances.length > 0) {
                    // Find the best body cell candidate
                    let bestBodyCell = null;
                    
                    for (const instance of instances) {
                        if (instance.mainComponent) {
                            const compName = instance.mainComponent.name.toLowerCase();
                            const instanceName = instance.name.toLowerCase();
                            
                            // Prefer components that are clearly body cells
                            const isGoodBodyCell = (compName.includes('cell') || compName.includes('row')) && 
                                                 !compName.includes('header') && 
                                                 !compName.includes('footer') && 
                                                 !compName.includes('select') && 
                                                 !compName.includes('expand') &&
                                                 !instanceName.includes('header') &&
                                                 !instanceName.includes('footer');
                            
                            if (isGoodBodyCell) {
                                bestBodyCell = instance.mainComponent;
                                Logger.debug('Found ideal body cell component:', bestBodyCell.name);
                                break;
                            } else if (!bestBodyCell) {
                                // Fallback to any cell-like component
                                bestBodyCell = instance.mainComponent;
                                Logger.debug('Found fallback body cell component:', bestBodyCell.name);
                            }
                        }
                    }
                    
                    if (bestBodyCell) {
                        bodyCell = bestBodyCell;
                        Logger.debug('Using body cell component:', bodyCell.name);
                    }
                }
                
                // Try to count columns with flexible matching
                const colInstances = tableFrame.findAll((node: any) => 
                    node.type === 'INSTANCE' && 
                    (node.name?.toLowerCase().includes('col') ||
                     node.name?.toLowerCase().includes('cell') ||
                     node.name?.toLowerCase().includes('row')) &&
                    node.visible
                );
                
                if (colInstances.length > 0) {
                    numCols = Math.min(colInstances.length, 10); // Cap at 10 columns
                    Logger.debug(`Counted ${numCols} columns from instances`);
                }
            }
            
            return {
                headerCell,
                bodyCell,
                footer,
                numCols,
                selectCellComponent: null,
                expandCellComponent: null,
                dividerComponent: null
            };
        } catch (error) {
            Logger.error('Error during table scan:', error);
            return null;
        }
    }
}