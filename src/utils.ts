export function findMatchingProperty(availableProperties: string[], uiPropName: string): string | null {
    const exactMatch = availableProperties.find((prop: string) => prop === uiPropName);
    if (exactMatch) return exactMatch;
    
    const startsWithMatch = availableProperties.find((prop: string) => prop.toLowerCase().startsWith(uiPropName.toLowerCase()));
    if (startsWithMatch) return startsWithMatch;
    
    const containsMatch = availableProperties.find((prop: string) => prop.toLowerCase().includes(uiPropName.toLowerCase()));
    if (containsMatch) return containsMatch;
    
    return null;
}

export function mapPropertyNames(uiProperties: { [key: string]: any }, componentProperties: { [key: string]: any }): { [key: string]: any } {
    const mappedProps: { [key: string]: any } = {};
    
    for (const [uiPropName, uiPropValue] of Object.entries(uiProperties)) {
        const availableProperties = Object.keys(componentProperties);
        const matchedProp = findMatchingProperty(availableProperties, uiPropName);
        
        if (matchedProp) {
            mappedProps[matchedProp] = uiPropValue;
        } else if (componentProperties[uiPropName]) {
            mappedProps[uiPropName] = uiPropValue;
        }
    }
    
    return mappedProps;
}

export function sortColumnData(cellProps: any, cols: number, rows: number): any {
    const sortedCellProps = {...cellProps};
    
    for (let c = 1; c <= cols; c++) {
        const headerKey = `header-${c}`;
        const headerData = cellProps[headerKey];
        
        if (headerData && headerData.properties) {
            const sortable = headerData.properties['Sortable'];
            const sorted = headerData.properties['Sorted'];
            
            if (sortable === 'True' && (sorted === 'Ascending' || sorted === 'Descending')) {
                const columnData: { rowIndex: number; cellData: any; value: string }[] = [];
                
                for (let r = 0; r < rows; r++) {
                    const cellKey = `${r}-${c - 1}`;
                    const cellData = cellProps[cellKey];
                    
                    let cellValue = '';
                    if (cellData && cellData.properties) {
                        const textProp = Object.keys(cellData.properties).find(
                            prop => prop.toLowerCase().includes('text') && 
                                   !prop.toLowerCase().includes('second') &&
                                   typeof cellData.properties[prop] === 'string'
                        );
                        
                        if (textProp) {
                            cellValue = String(cellData.properties[textProp]);
                        } else if (cellData.properties['Cell text#12234:32']) {
                            cellValue = String(cellData.properties['Cell text#12234:32']);
                        }
                    }
                    
                    columnData.push({
                        rowIndex: r,
                        cellData: cellData,
                        value: cellValue
                    });
                }
                
                columnData.sort((a, b) => {
                    const numA = parseFloat(a.value);
                    const numB = parseFloat(b.value);
                    
                    if (!isNaN(numA) && !isNaN(numB)) {
                        return sorted === 'Ascending' ? numA - numB : numB - numA;
                    } else {
                        return sorted === 'Ascending' 
                            ? a.value.localeCompare(b.value) 
                            : b.value.localeCompare(a.value);
                    }
                });
                
                for (let r = 0; r < rows; r++) {
                    const originalRowIndex = columnData[r].rowIndex;
                    const newCellKey = `${r}-${c - 1}`;
                    sortedCellProps[newCellKey] = columnData[r].cellData;
                }
            }
        }
    }
    
    return sortedCellProps;
}

export function findBorderColorVariable(): VariableAlias | null {
    try {
        const allVariableCollections = figma.variables.getLocalVariableCollections();
        const allVariables: Variable[] = [];
        
        for (const collection of allVariableCollections) {
            const collectionVariables = collection.variableIds.map(id => figma.variables.getVariableById(id)).filter(v => v !== null) as Variable[];
            allVariables.push(...collectionVariables);
        }
        
        const borderVariableNames = ['border-subtle-01', 'border-subtle-02', 'border-subtle', 'border-01', 'border-subtle-1'];
        let borderVariable = null;
        
        for (const variableName of borderVariableNames) {
            borderVariable = allVariables.find(v => v.name === variableName);
            if (borderVariable) break;
        }
        
        if (borderVariable) {
            return {
                id: borderVariable.id,
                type: 'VARIABLE_ALIAS'
            };
        }
        
        return null;
    } catch (error) {
        return null;
    }
}

export function extractFooterDividerColorVariable(footer: ComponentNode): VariableAlias | null {
    try {
        if ('children' in footer) {
            for (const child of footer.children) {
                const isDividerChild = child.name.toLowerCase().includes('divider') ||
                                     child.name.toLowerCase().includes('line') ||
                                     child.name.toLowerCase().includes('border') ||
                                     child.name.toLowerCase().includes('separator') ||
                                     child.type === 'LINE' ||
                                     (child.type === 'RECTANGLE' && child.height <= 2);
                
                if (isDividerChild && 'fills' in child && child.fills && Array.isArray(child.fills)) {
                    for (const fill of child.fills) {
                        if (fill.type === 'SOLID' && fill.boundVariables && fill.boundVariables.color) {
                            return fill.boundVariables.color;
                        }
                    }
                }
                
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
                                    return fill.boundVariables.color;
                                }
                            }
                        }
                    }
                }
            }
        }
        return null;
    } catch (error) {
        return null;
    }
}

export function createStyledDivider(width: number): RectangleNode {
    const divider = figma.createRectangle();
    divider.name = 'Divider';
    divider.resize(width, 1);
    
    const borderColorVariable = findBorderColorVariable();
    if (borderColorVariable) {
        divider.fills = [{
            type: 'SOLID',
            color: { r: 0.9, g: 0.9, b: 0.9 },
            boundVariables: { color: borderColorVariable }
        }];
    } else {
        divider.fills = [{ type: 'SOLID', color: { r: 0.9, g: 0.9, b: 0.9 } }];
    }
    
    return divider;
}

export function applyHeaderStyling(headerFrame: FrameNode): void {
    headerFrame.fills = [{ type: 'SOLID', color: { r: 0.96, g: 0.96, b: 0.96 } }];
}

export function createComponentInstanceWithProps(
    component: ComponentNode,
    cellData: { [key: string]: any },
    componentProps: { [key: string]: any },
    allProps: { [key: string]: any }
): SceneNode {
    console.log('🔧 Creating component instance with props:', { component: component.name, cellData, componentProps });
    
    const instance = component.createInstance();
    
    const mappedProps = mapPropertyNames(cellData, componentProps);
    console.log('🔧 Mapped properties:', mappedProps);
    
    for (const [propName, propValue] of Object.entries(mappedProps)) {
        try {
            console.log(`🔧 Setting property: ${propName} = ${propValue}`);
            instance.setProperties({ [propName]: propValue });
        } catch (error) {
            console.error(`❌ Error setting property ${propName}:`, error);
            try {
                (instance as any)[propName] = propValue;
                console.log(`🔧 Set property via direct assignment: ${propName} = ${propValue}`);
            } catch (directError) {
                console.error(`❌ Error setting property via direct assignment ${propName}:`, directError);
            }
        }
    }
    
    for (const [originalKey, originalValue] of Object.entries(cellData)) {
        if (Object.keys(mappedProps).some(mappedKey => mappedKey.split('#')[0] === originalKey.split('#')[0])) {
            continue;
        }
        
        if (typeof originalValue === 'string' || typeof originalValue === 'boolean') {
            try {
                console.log(`🔧 Setting unmapped property: ${originalKey} = ${originalValue}`);
                instance.setProperties({ [originalKey]: originalValue });
            } catch (error) {
                console.error(`❌ Error setting unmapped property ${originalKey}:`, error);
            }
        } else if (originalValue && typeof originalValue === 'object' && 'value' in originalValue) {
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

export function cleanupExternalComponents() {
    try {
        const simpleDividers = figma.root.findAll(node => 
            node.type === "COMPONENT" && node.name === "Simple Divider"
        );
        
        simpleDividers.forEach(divider => {
            let parent = divider.parent;
            let shouldRemove = true;
            
            while (parent) {
                if (parent.getPluginData('isGeneratedTable') === 'true') {
                    shouldRemove = false;
                    break;
                }
                parent = parent.parent;
            }
            
            if (shouldRemove) {
                divider.remove();
                console.log('🧹 Cleaned up external Simple Divider component');
            }
        });
        
        const checkboxComponents = figma.root.findAll(node => 
            node.type === "COMPONENT_SET" && node.name === "Data table select cell item"
        );
        
        checkboxComponents.forEach(componentSet => {
            let parent = componentSet.parent;
            let shouldRemove = true;
            
            while (parent) {
                if (parent.getPluginData('isGeneratedTable') === 'true') {
                    shouldRemove = false;
                    break;
                }
                parent = parent.parent;
            }
            
            if (shouldRemove) {
                componentSet.remove();
                console.log('🧹 Cleaned up external checkbox component set');
            }
        });
    } catch (error) {
        console.error('❌ Error cleaning up external components:', error);
    }
}