// Message Router - Central message handling for the plugin
import { TableScanHandler } from './TableScanHandler';
import { TableCreationHandler } from './TableCreationHandler';
import { ComponentHandler } from './ComponentHandler';
import { DataGenerationHandler } from './DataGenerationHandler';
import { SettingsHandler } from './SettingsHandler';
import { handleAITableGeneration } from './aiHandlers';

export class MessageRouter {
    private static lastScanTime = 0;
    private static readonly SCAN_DEBOUNCE_MS = 500;

    // Plugin size constants
    private static readonly pluginMinWidth = 320;
    private static readonly pluginMaxWidth = 800;
    private static readonly pluginMinHeight = 200;
    private static readonly pluginMaxHeight = 600;
    private static readonly pluginDefaultWidth = 400;
    private static readonly pluginDefaultHeight = 500;

    static async handleMessage(msg: any): Promise<void> {
        console.log(`[MessageRouter] Handling message type: ${msg.type}`);

        try {
            switch (msg.type) {
                // UI Management
                case "resize":
                case "resize-ui":
                    await this.handleResize(msg);
                    break;

                // Table Operations
                case "scan-selected-table":
                    await this.handleScanSelectedTable(msg);
                    break;
                case "scan-table":
                    await TableScanHandler.handleScanTable(msg);
                    break;
                case "create-table-from-scan":
                case "create-table-from-ai":
                    await TableCreationHandler.handleCreateTable(msg);
                    break;
                case "update-table":
                    await TableCreationHandler.handleUpdateTable(msg);
                    break;

                // Component Operations
                case "load-instance-by-id":
                case "request-component-info":
                    await ComponentHandler.handleComponentRequest(msg);
                    break;
                case "update-cell-properties":
                    await ComponentHandler.handleUpdateCellProperties(msg);
                    break;
                case "clear-cell-instances":
                    await ComponentHandler.handleClearCellInstances(msg);
                    break;

                // Data Generation
                case "generate-fake-data":
                    await DataGenerationHandler.handleGenerateFakeData(msg);
                    break;
                case "generate-watsonx-data":
                    await DataGenerationHandler.handleGenerateWatsonxData(msg);
                    break;
                case "generate-table-with-ai":
                    await handleAITableGeneration(msg);
                    break;

                // Settings
                case "save-watsonx-api-key":
                    await SettingsHandler.handleSaveWatsonxApiKey(msg);
                    break;
                case "load-watsonx-settings":
                    await SettingsHandler.handleLoadWatsonxSettings(msg);
                    break;
                case "request-selection-state":
                    await SettingsHandler.handleRequestSelectionState(msg);
                    break;

                default:
                    console.warn(`[MessageRouter] Unknown message type: ${msg.type}`);
            }
        } catch (error) {
            console.error(`[MessageRouter] Error handling message ${msg.type}:`, error);
        }
    }

    private static async handleResize(msg: any): Promise<void> {
        const { size, width, height } = msg;
        let w: number, h: number;

        if (size) {
            w = size.w;
            h = size.h;
        } else {
            w = width || this.pluginDefaultWidth;
            h = height || this.pluginDefaultHeight;
        }

        // Apply constraints
        h = Math.max(this.pluginMinHeight, Math.min(this.pluginMaxHeight, h));
        w = Math.max(this.pluginMinWidth, Math.min(this.pluginMaxWidth, w));

        // Resize the plugin
        figma.ui.resize(w, h);

        // Save the size for next time
        figma.clientStorage.setAsync('pluginSize', { w, h }).catch(err => {
            console.log('Failed to save plugin size:', err);
        });

        console.log(`Plugin resized to ${w}×${h}`);
    }

    private static async handleScanSelectedTable(msg: any): Promise<void> {
        console.log(`[MessageRouter] Scanning selected table: ${msg.tableId}`);
        
        const tableNode = figma.getNodeById(msg.tableId);
        console.log(`[MessageRouter] Node found: ${tableNode ? tableNode.type : 'null'}, name: ${tableNode?.name}`);
        
        if (tableNode && (tableNode.type === "FRAME" || tableNode.type === "COMPONENT" || tableNode.type === "INSTANCE")) {
            figma.currentPage.selection = [tableNode];
            console.log('[MessageRouter] Triggering selectionchange for table scan');
        } else {
            console.warn(`[MessageRouter] Invalid node type for scanning: ${tableNode?.type}`);
        }
    }
}