export interface ComponentProperty {
    type: 'BOOLEAN' | 'TEXT' | 'VARIANT' | 'INSTANCE_SWAP';
    value: any;
}

export function createComponentInstanceWithProps(
  component: ComponentNode,
  cellData: { [key: string]: any },
  componentProps: { [key: string]: ComponentProperty },
  allProps: { [key: string]: any }
): SceneNode {
  console.log('🔧 Creating component instance with props:', { component: component.name, cellData, componentProps });
  
  // Create instance
  const instance = component.createInstance();
  
  // Map properties from cellData to component properties
  const mappedProps = mapPropertyNames(cellData, componentProps);
  console.log('🔧 Mapped properties:', mappedProps);
  
  // Apply the mapped properties to the instance
  for (const [propName, propValue] of Object.entries(mappedProps)) {
    try {
      console.log(`🔧 Setting property: ${propName} = ${propValue}`);
      instance.setProperties({ [propName]: propValue });
    } catch (error) {
      console.error(`❌ Error setting property ${propName}:`, error);
      // Try alternative method for setting properties
      try {
        (instance as any)[propName] = propValue;
        console.log(`🔧 Set property via direct assignment: ${propName} = ${propValue}`);
      } catch (directError) {
        console.error(`❌ Error setting property via direct assignment ${propName}:`, directError);
      }
    }
  }
  
  // Handle any additional properties that weren't mapped, especially slot properties
  for (const [originalKey, originalValue] of Object.entries(cellData)) {
    // Skip already mapped properties
    if (Object.keys(mappedProps).some(mappedKey => mappedKey.split('#')[0] === originalKey.split('#')[0])) {
      continue;
    }
    
    // Handle unmapped properties, especially slot properties
    // Ensure the value is of a valid type before setting
    if (typeof originalValue === 'string' || typeof originalValue === 'boolean') {
      try {
        console.log(`🔧 Setting unmapped property: ${originalKey} = ${originalValue}`);
        instance.setProperties({ [originalKey]: originalValue });
      } catch (error) {
        console.error(`❌ Error setting unmapped property ${originalKey}:`, error);
      }
    } else if (originalValue && typeof originalValue === 'object' && 'value' in originalValue) {
      // Handle case where the value is an object with a value property
      const valueToSet = originalValue.value;
      if (typeof valueToSet === 'string' || typeof valueToSet === 'boolean') {
        try {
          console.log(`🔧 Setting unmapped property from object: ${originalKey} = ${valueToSet}`);
          instance.setProperties({ [originalKey]: valueToSet });
        } catch (error) {
          console.error(`❌ Error setting unmapped property from object ${originalKey}:`, error);
        }
      }
    }
  }
  
  return instance;
}

// Helper function to find matching property
function findMatchingProperty(availableProperties: string[], uiPropName: string): string | null {
    // First try exact match
    const exactMatch = availableProperties.find((prop: string) => prop === uiPropName);
    if (exactMatch) {
        return exactMatch;
    }
    // Try match at start of property name
    const startsWithMatch = availableProperties.find((prop: string) => prop.toLowerCase().startsWith(uiPropName.toLowerCase()));
    if (startsWithMatch) {
        return startsWithMatch;
    }
    // Try contains match
    const containsMatch = availableProperties.find((prop: string) => prop.toLowerCase().includes(uiPropName.toLowerCase()));
    if (containsMatch) {
        return containsMatch;
    }
    
    return null;
}

// Helper function to map property names
function mapPropertyNames(uiProperties: { [key: string]: any }, componentProperties: { [key: string]: any }): { [key: string]: any } {
    const mappedProps: { [key: string]: any } = {};
    
    for (const [uiPropName, uiPropValue] of Object.entries(uiProperties)) {
        const availableProperties = Object.keys(componentProperties);
        const matchedProp = findMatchingProperty(availableProperties, uiPropName);
        
        if (matchedProp) {
            mappedProps[matchedProp] = uiPropValue;
        } else {
            // Try direct mapping as fallback
            if (componentProperties[uiPropName]) {
                mappedProps[uiPropName] = uiPropValue;
            }
        }
    }
    
    return mappedProps;
}