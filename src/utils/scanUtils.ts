// Utility functions for table scanning to reduce code duplication
import { Logger } from './logger';

export class ScanUtils {
    /**
     * Find a component by name in children, handling both instances and components
     */
    static findComponentByName(
        parent: any, 
        targetName: string, 
        options: { exact?: boolean; includes?: boolean } = { exact: true }
    ): any {
        if (!parent || !('children' in parent)) return null;

        for (const child of parent.children) {
            const nameMatch = options.exact 
                ? child.name === targetName
                : child.name?.toLowerCase().includes(targetName.toLowerCase());

            if (nameMatch) {
                if (child.type === 'INSTANCE' && child.mainComponent) {
                    Logger.debug(`Found ${targetName} from instance:`, child.mainComponent.name);
                    return child.mainComponent;
                } else if (child.type === 'COMPONENT') {
                    Logger.debug(`Found ${targetName} component directly:`, child.name);
                    return child;
                }
            }
        }
        return null;
    }

    /**
     * Find a rectangle/divider element in children
     */
    static findDividerElement(parent: any): any {
        if (!parent || !('children' in parent)) return null;

        for (const child of parent.children) {
            if (child.name?.toLowerCase().includes('divider')) {
                if (child.type === 'RECTANGLE' || child.type === 'LINE' || child.type === 'VECTOR') {
                    Logger.debug('Valid divider found with visual element:', child.type);
                    return child;
                }
            }
        }
        return null;
    }

    /**
     * Validate if a component is a valid divider
     */
    static isValidDivider(component: any): boolean {
        if (!component || !('children' in component)) return false;

        for (const child of component.children) {
            if (child.type === 'RECTANGLE' || child.type === 'LINE' || child.type === 'VECTOR') {
                return true;
            } else if (child.type === 'TEXT') {
                Logger.debug('Skipping text-based component:', component.name);
                return false;
            }
        }
        return false;
    }

    /**
     * Extract component from node (handles both instances and components)
     */
    static extractComponent(node: any, componentName: string): any {
        if (node.type === 'INSTANCE' && node.mainComponent) {
            Logger.debug(`Using ${componentName} from instance:`, node.mainComponent.name);
            return node.mainComponent;
        } else if (node.type === 'COMPONENT') {
            Logger.debug(`Using ${componentName} component directly:`, node.name);
            return node;
        }
        return null;
    }

    /**
     * Count visible data columns in a row
     */
    static countDataColumns(row: any): number {
        if (!row || !('children' in row)) return 0;

        return row.children.filter((n: any) => 
            n.type === 'INSTANCE' && 
            n.name?.toLowerCase().includes('col') &&
            !n.name?.toLowerCase().includes('select') &&
            !n.name?.toLowerCase().includes('expand') &&
            n.visible
        ).length;
    }
}