// Property Mapper - Handles property mapping between UI and Figma components
export class PropertyMapper {
    /**
     * Map UI property names to component property names
     */
    static mapPropertyNames(
        uiProperties: { [key: string]: any }, 
        componentProperties: { [key: string]: any }
    ): { [key: string]: any } {
        const mappedProperties: { [key: string]: any } = {};
        
        for (const [uiKey, uiValue] of Object.entries(uiProperties)) {
            const componentKey = this.findMatchingProperty(
                Object.keys(componentProperties), 
                uiKey
            );
            
            if (componentKey) {
                mappedProperties[componentKey] = uiValue;
            } else {
                console.warn(`No matching component property found for UI property: ${uiKey}`);
            }
        }
        
        return mappedProperties;
    }

    /**
     * Find matching property between UI and component properties
     */
    static findMatchingProperty(availableProperties: string[], uiPropName: string): string | null {
        // First try exact match
        const exactMatch = availableProperties.find((prop: string) => prop === uiPropName);
        if (exactMatch) {
            return exactMatch;
        }
        
        // Try match at start of property name
        const startsWithMatch = availableProperties.find((prop: string) => 
            prop.toLowerCase().startsWith(uiPropName.toLowerCase())
        );
        if (startsWithMatch) {
            return startsWithMatch;
        }
        
        // Try contains match
        const containsMatch = availableProperties.find((prop: string) => 
            prop.toLowerCase().includes(uiPropName.toLowerCase())
        );
        if (containsMatch) {
            return containsMatch;
        }
        
        return null;
    }

    /**
     * Clean property name by removing hash suffix
     */
    static cleanPropName(name: string): string {
        return name.split('#')[0].trim();
    }

    /**
     * Validate property combination for component variants
     */
    static isValidVariantCombination(
        componentSet: ComponentSetNode, 
        properties: { [key: string]: any }
    ): boolean {
        try {
            // Check if the component set has variant group properties
            if (!componentSet.variantGroupProperties) {
                return true; // No variant properties to validate
            }

            // Validate each property
            for (const [propName, propValue] of Object.entries(properties)) {
                const variantProp = componentSet.variantGroupProperties[propName];
                if (variantProp && variantProp.values) {
                    if (!variantProp.values.includes(propValue)) {
                        console.warn(`Invalid variant value for ${propName}: ${propValue}`);
                        return false;
                    }
                }
            }

            return true;
        } catch (error) {
            console.error('Error validating variant combination:', error);
            return false;
        }
    }
}
