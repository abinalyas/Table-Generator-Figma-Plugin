// Handler for settings operations
import { Logger } from '../utils/logger';

export class SettingsHandler {
  static async handleSaveWatsonxApiKey(msg: any): Promise<void> {
    Logger.info('Saving Watsonx API key...');
    try {
      await figma.clientStorage.setAsync('watsonxApiKey', msg.apiKey);
      figma.ui.postMessage({
        type: 'watsonx-api-key-saved',
        success: true,
        message: 'API key saved successfully'
      });
    } catch (error) {
      Logger.error('Error saving API key:', error);
      figma.ui.postMessage({
        type: 'watsonx-api-key-saved',
        success: false,
        message: 'Failed to save API key'
      });
    }
  }

  static async handleLoadWatsonxSettings(msg: any): Promise<void> {
    Logger.info('Loading Watsonx settings...');
    try {
      const apiKey = await figma.clientStorage.getAsync('watsonxApiKey');
      figma.ui.postMessage({
        type: 'watsonx-settings-loaded',
        success: true,
        apiKey: apiKey || ''
      });
    } catch (error) {
      Logger.error('Error loading settings:', error);
      figma.ui.postMessage({
        type: 'watsonx-settings-loaded',
        success: false,
        message: 'Failed to load settings'
      });
    }
  }

  static async handleRequestSelectionState(msg: any): Promise<void> {
    Logger.info('Handling request selection state...');
    try {
      const selection = figma.currentPage.selection;
      figma.ui.postMessage({
        type: 'selection-state-response',
        success: true,
        selection: selection.map(node => ({
          id: node.id,
          type: node.type,
          name: node.name
        }))
      });
    } catch (error) {
      Logger.error('Error handling selection state request:', error);
      figma.ui.postMessage({
        type: 'selection-state-response',
        success: false,
        message: 'Failed to get selection state'
      });
    }
  }
}
