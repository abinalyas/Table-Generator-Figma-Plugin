// Table Scan Service - Handles all table scanning operations
import type { ScanResult } from '../types';

export class TableScanService {
    /**
     * Perform comprehensive table scan on a given component
     */
    static async performTableScan(tableFrame: SceneNode): Promise<ScanResult | null> {
        console.log('🔄 Performing table scan on:', tableFrame.name);
        
        try {
            // Check if this is a generated table first
            const isGeneratedTable = 'getPluginData' in tableFrame && tableFrame.getPluginData('isGeneratedTable') === 'true';
            console.log('🔍 Is generated table:', isGeneratedTable);
            
            if (isGeneratedTable) {
                console.log('🔄 Scanning generated table with fallback method...');
                return await this.scanGeneratedTable(tableFrame as FrameNode | ComponentNode);
            }
            
            // Continue with regular Carbon Design System scanning
            return await this.scanCarbonTable(tableFrame);
        } catch (error) {
            console.error('Error during table scan:', error);
            return null;
        }
    }

    /**
     * Scan Carbon Design System tables
     */
    private static async scanCarbonTable(tableFrame: SceneNode): Promise<ScanResult | null> {
        console.log('🔄 Scanning Carbon Design System table...');
        
        // Implementation for Carbon table scanning
        // This would contain the existing Carbon scanning logic
        return null; // Placeholder
    }

    /**
     * Scan generated tables with their specific structure
     */
    private static async scanGeneratedTable(tableFrame: FrameNode | ComponentNode): Promise<ScanResult | null> {
        console.log('🔄 Scanning generated table:', tableFrame.name);
        
        try {
            // Try to restore from stored plugin data first
            const scanData = tableFrame.getPluginData('tableGeneratorScan');
            if (scanData) {
                try {
                    const storedScan = JSON.parse(scanData);
                    console.log('🔄 Found stored scan data, attempting to restore components...');
                    return await this.restoreFullScanResult(storedScan);
                } catch (error) {
                    console.log('Could not restore scan result from plugin data:', error);
                }
            }
            
            // Fallback to dynamic scanning
            return await this.scanGeneratedTableDynamically(tableFrame);
        } catch (error) {
            console.error('Error scanning generated table:', error);
            return null;
        }
    }

    /**
     * Restore scan result from stored data
     */
    private static async restoreFullScanResult(storedScan: any): Promise<ScanResult | null> {
        console.log('🔄 Restoring scan result from stored data...');
        
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
        
        if (bodyCell) {
            console.log('✅ Successfully restored scan result from stored data');
            return {
                headerCell,
                headerRowComponent: null,
                bodyCell,
                bodyRowComponent: null,
                footer,
                selectCellComponent,
                expandCellComponent,
                dividerComponent,
                numCols: storedScan.numCols || 5
            };
        }
        
        console.log('❌ Could not restore essential components from stored data');
        return null;
    }

    /**
     * Dynamically scan generated table when stored data is unavailable
     */
    private static async scanGeneratedTableDynamically(tableFrame: FrameNode | ComponentNode): Promise<ScanResult | null> {
        console.log('🔄 Performing dynamic scan of generated table...');
        
        // Implementation for dynamic scanning
        // This would contain the existing dynamic scanning logic
        return null; // Placeholder
    }
}
