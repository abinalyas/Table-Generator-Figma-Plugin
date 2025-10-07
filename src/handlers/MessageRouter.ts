// Main message router that delegates to specific handlers
import { Logger } from '../utils/logger';
import { TableScanHandler } from './TableScanHandler';
import { TableCreationHandler } from './TableCreationHandler';
import { DataGenerationHandler } from './DataGenerationHandler';
import { ComponentHandler } from './ComponentHandler';
import { SettingsHandler } from './SettingsHandler';

export class MessageRouter {
    static async routeMessage(msg: any): Promise<void> {
        try {
            Logger.debug('Routing message type:', msg.type);

            switch (msg.type) {
                // Settings and UI
                case 'resize':
                    SettingsHandler.handleResize(msg);
                    break;

                case 'save-watsonx-api-key':
                    await SettingsHandler.handleSaveWatsonxApiKey(msg);
                    break;

                case 'load-watsonx-settings':
                    await SettingsHandler.handleLoadWatsonxSettings(msg);
                    break;

                case 'request-selection-state':
                    await SettingsHandler.handleRequestSelectionState(msg);
                    break;

                // Data Generation
                case 'generate-fake-data':
                    await DataGenerationHandler.handleGenerateFakeData(msg);
                    break;

                case 'generate-watsonx-data':
                    await DataGenerationHandler.handleGenerateWatsonxData(msg);
                    break;

                case 'generate-table-with-ai':
                    await DataGenerationHandler.handleGenerateTableWithAI(msg);
                    break;

                // Component Operations
                case 'load-instance-by-id':
                    await ComponentHandler.handleLoadInstanceById(msg);
                    break;

                case 'update-cell-properties':
                    await ComponentHandler.handleUpdateCellProperties(msg);
                    break;

                case 'request-component-info':
                    await ComponentHandler.handleRequestComponentInfo(msg);
                    break;

                case 'clear-cell-instances':
                    await ComponentHandler.handleClearCellInstances(msg);
                    break;

                // Table Operations
                case 'scan-table':
                    await TableScanHandler.handleScanTable(msg);
                    break;

                case 'create-table-from-ai':
                case 'create-table-from-scan':
                    await TableCreationHandler.handleCreateTable(msg);
                    break;

                case 'update-table':
                    await TableCreationHandler.handleUpdateTable(msg);
                    break;

                default:
                    Logger.warning('Unknown message type:', msg.type);
                    break;
            }
        } catch (error) {
            Logger.error('Error routing message:', error);
        }
    }
}