// Handler for settings and configuration operations
import { Logger } from '../utils/logger';

export class SettingsHandler {
    static async handleSaveWatsonxApiKey(msg: any): Promise<void> {
        try {
            const { apiKey } = msg;
            if (apiKey) {
                await figma.clientStorage.setAsync('watsonx.apiKey', apiKey);
                Logger.success('API key saved successfully');
            }
        } catch (error) {
            Logger.error('Failed to save watsonx API key:', error);
        }
    }

    static async handleLoadWatsonxSettings(msg: any): Promise<void> {
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
            
            Logger.debug('WatsonX settings loaded successfully');
        } catch (error) {
            Logger.error('Error loading WatsonX settings:', error);
            figma.ui.postMessage({ type: 'watsonx-settings', endpoint: '', apiKeyMasked: '' });
        }
    }

    static async handleRequestSelectionState(msg: any): Promise<void> {
        const selection = figma.currentPage.selection;
        
        Logger.debug('Requesting selection state, count:', selection.length);
        
        // The actual selection state logic would be moved here from main.ts
        // This is a placeholder for the extraction
        Logger.info('SettingsHandler: Selection state logic would be implemented here');
    }

    static handleResize(msg: any): void {
        const { size } = msg;
        let { w, h } = size;
        
        // Apply constraints (these should match main.ts constants)
        const pluginMaxHeight = 1000;
        const pluginMinHeight = 400;
        const pluginMaxWidth = 1200;
        const pluginMinWidth = 300;
        
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
            Logger.error('Failed to save plugin size:', err);
        });
        
        Logger.debug(`Resized plugin to ${w}x${h}`);
    }
}