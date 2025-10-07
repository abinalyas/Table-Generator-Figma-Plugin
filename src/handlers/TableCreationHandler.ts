// Handler for table creation operations
import { Logger } from '../utils/logger';

export class TableCreationHandler {
    private static isCreatingTable = false;

    static async handleCreateTable(msg: any): Promise<void> {
        if (this.isCreatingTable) {
            figma.notify("Already creating a table. Please wait.");
            return;
        }

        this.isCreatingTable = true;
        Logger.progress('Starting table creation...');

        try {
            // The actual table creation logic would be moved here from main.ts
            // This is a placeholder for the extraction
            Logger.info('TableCreationHandler: Creation logic would be implemented here');
        } catch (error) {
            Logger.error('Error creating table:', error);
            figma.notify('Error creating table. Please try again.');
        } finally {
            this.isCreatingTable = false;
        }
    }

    static async handleUpdateTable(msg: any): Promise<void> {
        if (this.isCreatingTable) {
            figma.notify("Already updating a table. Please wait.");
            return;
        }

        this.isCreatingTable = true;
        Logger.progress('Starting table update...');

        try {
            // The actual table update logic would be moved here from main.ts
            // This is a placeholder for the extraction
            Logger.info('TableCreationHandler: Update logic would be implemented here');
        } catch (error) {
            Logger.error('Error updating table:', error);
            figma.notify('Error updating table. Please try again.');
        } finally {
            this.isCreatingTable = false;
        }
    }
}