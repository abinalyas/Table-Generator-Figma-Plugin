// Color Utils - Handles color variable operations
export class ColorUtils {
    /**
     * Find border color variable from available variables
     */
    static findBorderColorVariable(): VariableAlias | null {
        try {
            // Get all variables from all libraries (including local and external)
            const allVariableCollections = figma.variables.getLocalVariableCollections();
            const allVariables: Variable[] = [];
            
            // Collect all variables from all collections
            for (const collection of allVariableCollections) {
                const collectionVariables = collection.variableIds.map(id => 
                    figma.variables.getVariableById(id)
                ).filter(v => v !== null) as Variable[];
                allVariables.push(...collectionVariables);
            }
            
            console.log('🔍 Available variables from all libraries:', 
                allVariables.map(v => `${v.name} (ID: ${v.id})`));
            
            // Look for common border variable names, prioritizing border-subtle-01
            const borderVariableNames = [
                'border-subtle-01', 
                'border-subtle-02', 
                'border-subtle', 
                'border-01', 
                'border-subtle-1'
            ];
            
            for (const variableName of borderVariableNames) {
                const borderVariable = allVariables.find(v => v.name === variableName);
                if (borderVariable) {
                    console.log(`🎨 Found border variable from libraries: ${variableName}`);
                    return {
                        id: borderVariable.id,
                        type: 'VARIABLE_ALIAS'
                    };
                }
            }
            
            console.log('⚠️ No border color variable found in any library');
            return null;
        } catch (error) {
            console.error('❌ Error finding border color variable:', error);
            return null;
        }
    }

    /**
     * Extract color variable from footer divider
     */
    static extractFooterDividerColorVariable(footer: ComponentNode): VariableAlias | null {
        try {
            console.log('🔍 Extracting color variable from footer divider:', footer.name);
            
            // Look for divider elements in the footer component
            if ('children' in footer) {
                for (const child of footer.children) {
                    // Check if this child is a divider
                    const isDividerChild = child.name.toLowerCase().includes('divider') ||
                                         child.name.toLowerCase().includes('line') ||
                                         child.name.toLowerCase().includes('border') ||
                                         child.name.toLowerCase().includes('separator') ||
                                         child.type === 'LINE' ||
                                         (child.type === 'RECTANGLE' && child.height <= 2);
                    
                    if (isDividerChild && 'fills' in child && child.fills && Array.isArray(child.fills)) {
                        for (const fill of child.fills) {
                            if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                                console.log('✅ Found color variable in footer divider:', fill.boundVariables.color);
                                return fill.boundVariables.color;
                            }
                        }
                    }
                    
                    // Also check nested children
                    if ('children' in child) {
                        for (const grandChild of child.children) {
                            const isDividerGrandChild = grandChild.name.toLowerCase().includes('divider') ||
                                                      grandChild.name.toLowerCase().includes('line') ||
                                                      grandChild.name.toLowerCase().includes('border') ||
                                                      grandChild.name.toLowerCase().includes('separator') ||
                                                      grandChild.type === 'LINE' ||
                                                      (grandChild.type === 'RECTANGLE' && grandChild.height <= 2);
                            
                            if (isDividerGrandChild && 'fills' in grandChild && grandChild.fills && Array.isArray(grandChild.fills)) {
                                for (const fill of grandChild.fills) {
                                    if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                                        console.log('✅ Found color variable in footer nested divider:', fill.boundVariables.color);
                                        return fill.boundVariables.color;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            console.log('⚠️ No color variable found in footer divider elements');
            return null;
        } catch (error) {
            console.error('❌ Error extracting footer divider color variable:', error);
            return null;
        }
    }
}

