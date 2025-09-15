export interface ScanResult {
    headerCell: ComponentNode | null;
    headerRowComponent: ComponentNode | null;
    bodyCell: ComponentNode | null;
    bodyRowComponent: ComponentNode | null;
    footer: ComponentNode | null;
    numCols: number;
    selectCellComponent: ComponentNode | null;
    expandCellComponent: ComponentNode | null;
    dividerComponent: ComponentNode | null;
}

export async function performTableScan(tableFrame: SceneNode): Promise<ScanResult | null> {
    try {
        const isGeneratedTable = 'getPluginData' in tableFrame && tableFrame.getPluginData('isGeneratedTable') === 'true';
        
        if (isGeneratedTable) {
            return await scanGeneratedTable(tableFrame as FrameNode | ComponentNode);
        }
        
        let expandCellComponent: ComponentNode | null = null;
        let selectCellComponent: ComponentNode | null = null;
        let dividerComponent: ComponentNode | null = null;

        if ('findOne' in tableFrame && typeof tableFrame.findOne === 'function') {
            const headerRow = tableFrame.findOne(n =>
                n.name?.includes('header row') && n.type === 'INSTANCE'
            ) as InstanceNode | undefined;
            
            const bodyGroup = tableFrame.findOne(n => n.name === 'Body' && n.type === 'FRAME');
            const footerBar = tableFrame.findOne(n => n.name === 'Pagination - Table bar' && n.type === 'INSTANCE');
            
            let headerCell: InstanceNode | null = null;
            let bodyCell: InstanceNode | null = null;
            let footer: InstanceNode | null = null;
            let numCols = 0;
            let headerRowComponent: ComponentNode | null = null;
            let bodyRowComponent: ComponentNode | null = null;
            
            if (headerRow && 'children' in headerRow) {
                headerRowComponent = await headerRow.getMainComponentAsync();
                
                const headerCellInstance = headerRow.children.find((n: SceneNode) =>
                    n.type === 'INSTANCE' && n.name?.toLowerCase().includes('col 1')
                ) as InstanceNode | undefined;
                
                if (headerCellInstance) {
                    headerCell = headerCellInstance;
                    numCols = headerRow.children.filter((n: SceneNode) => 
                        n.type === 'INSTANCE' && 
                        n.name?.toLowerCase().includes('col') &&
                        !n.name?.toLowerCase().includes('select') &&
                        !n.name?.toLowerCase().includes('expand') &&
                        n.visible
                    ).length;
                }
            }
            
            if (bodyGroup && 'children' in bodyGroup) {
                const bodyRowInstance = bodyGroup.children.find((n: SceneNode) => 
                    n.type === 'INSTANCE' && n.name === 'Data table body row item'
                ) as InstanceNode | undefined;
                
                if (bodyRowInstance) {
                    bodyRowComponent = await bodyRowInstance.getMainComponentAsync();
                    
                    if ('children' in bodyRowInstance) {
                        const dataTableRow = bodyRowInstance.children.find(child => 
                            child.type === 'FRAME' && child.name === 'Data table row'
                        ) as FrameNode | undefined;
                        
                        if (dataTableRow && 'children' in dataTableRow) {
                            const bodyCellInstance = dataTableRow.children.find((cell: SceneNode) =>
                                cell.type === 'INSTANCE' && cell.name?.toLowerCase().includes('col 1')
                            ) as InstanceNode | undefined;
                            
                            if (bodyCellInstance) {
                                bodyCell = bodyCellInstance;
                            }
                        }
                    }
                }
            }
            
            if (footerBar && footerBar.type === 'INSTANCE') {
                footer = footerBar;
            }
            
            if ('findAll' in tableFrame && typeof tableFrame.findAll === 'function') {
                const allInstances = tableFrame.findAll(n => n.type === 'INSTANCE' || n.type === 'COMPONENT') as (InstanceNode | ComponentNode)[];
                
                for (const node of allInstances) {
                    if (node.name === 'Data table expand cell item' && !expandCellComponent) {
                        expandCellComponent = node.type === 'INSTANCE' ? node.mainComponent : node;
                    }
                    
                    if (node.name === 'Data table select cell item' && !selectCellComponent) {
                        selectCellComponent = node.type === 'INSTANCE' ? node.mainComponent : node;
                    }
                    
                    if (!dividerComponent && !node.name.includes('AI label') && !node.name.includes('Select menu')) {
                        const isDivider = node.name === 'Divider' || 
                                        (node.name.toLowerCase().includes('divider') && !node.name.includes('AI')) ||
                                        node.name.includes('line') ||
                                        node.name.includes('border') ||
                                        node.name.includes('separator');
                        
                        if (isDivider) {
                            const componentToCheck = node.type === 'INSTANCE' ? node.mainComponent : node;
                            
                            if (componentToCheck && 'children' in componentToCheck) {
                                for (const child of componentToCheck.children) {
                                    if (child.type === 'RECTANGLE' || child.type === 'LINE' || child.type === 'VECTOR') {
                                        dividerComponent = node.type === 'INSTANCE' ? node.mainComponent : node;
                                        break;
                                    } else if (child.type === 'TEXT') {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            return {
                headerCell: headerCell ? await headerCell.getMainComponentAsync() : null,
                headerRowComponent: headerRowComponent,
                bodyCell: bodyCell ? await bodyCell.getMainComponentAsync() : null,
                bodyRowComponent: bodyRowComponent,
                footer: footer ? await footer.getMainComponentAsync() : null,
                numCols: numCols || 5,
                selectCellComponent: selectCellComponent,
                expandCellComponent: expandCellComponent,
                dividerComponent: dividerComponent,
            };
        }
        
        return null;
        
    } catch (error) {
        return null;
    }
}

async function scanGeneratedTable(tableFrame: FrameNode | ComponentNode): Promise<ScanResult | null> {
    try {
        const scanData = tableFrame.getPluginData('tableGeneratorScan');
        if (scanData) {
            try {
                const storedScan = JSON.parse(scanData);
                
                let headerCell = null, bodyCell = null, footer = null;
                let selectCellComponent = null, expandCellComponent = null, dividerComponent = null;
                
                if (storedScan.bodyCellId) {
                    const bodyCellNode = figma.getNodeById(storedScan.bodyCellId);
                    if (bodyCellNode && bodyCellNode.type === 'COMPONENT') {
                        bodyCell = bodyCellNode;
                    }
                }
                
                if (storedScan.headerCellId) {
                    const headerCellNode = figma.getNodeById(storedScan.headerCellId);
                    if (headerCellNode && headerCellNode.type === 'COMPONENT') {
                        headerCell = headerCellNode;
                    }
                }
                
                if (storedScan.footerId) {
                    const footerNode = figma.getNodeById(storedScan.footerId);
                    if (footerNode && footerNode.type === 'COMPONENT') {
                        footer = footerNode;
                    }
                }
                
                if (storedScan.selectCellId) {
                    const selectCellNode = figma.getNodeById(storedScan.selectCellId);
                    if (selectCellNode && selectCellNode.type === 'COMPONENT') {
                        selectCellComponent = selectCellNode;
                    }
                }
                
                if (storedScan.expandCellId) {
                    const expandCellNode = figma.getNodeById(storedScan.expandCellId);
                    if (expandCellNode && expandCellNode.type === 'COMPONENT') {
                        expandCellComponent = expandCellNode;
                    }
                }
                
                if (storedScan.dividerId) {
                    const dividerNode = figma.getNodeById(storedScan.dividerId);
                    if (dividerNode && dividerNode.type === 'COMPONENT') {
                        dividerComponent = dividerNode;
                    }
                }
                
                if (bodyCell) {
                    return {
                        headerCell,
                        headerRowComponent: null,
                        bodyCell,
                        bodyRowComponent: null,
                        footer,
                        numCols: storedScan.numCols || 5,
                        selectCellComponent,
                        expandCellComponent,
                        dividerComponent
                    };
                }
            } catch (error) {
                // Continue to fallback scan
            }
        }
        
        return null;
        
    } catch (error) {
        return null;
    }
}