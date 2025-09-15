import { performTableScan, ScanResult } from '../scanner';
import { sortColumnData, mapPropertyNames, createStyledDivider, applyHeaderStyling, extractFooterDividerColorVariable } from '../utils';
import { getLastScanResult } from './scanHandlers';

let isCreatingTable = false;

export async function handleTableCreation(msg: any) {
    if (isCreatingTable) {
        figma.notify("Already creating a table. Please wait.");
        return;
    }
    isCreatingTable = true;
    
    try {
        const lastScanResult = getLastScanResult();
        if (!lastScanResult) {
            figma.notify('❌ No scan result available. Please scan a table first.');
            return;
        }

        const { headerCell, bodyCell, footer, numCols } = lastScanResult;
        figma.ui.postMessage({ type: 'show-loader', message: 'Generating table...' });

        const includeHeader = msg.includeHeader !== false;
        const includeFooter = msg.includeFooter === true;
        const includeSelectable = msg.includeSelectable === true;
        const includeExpandable = msg.includeExpandable === true;
        const rows = msg.rows || 3;
        const cols = msg.cols || numCols || 3;
        let cellProps = msg.cellProps || {};

        // Apply sorting if enabled
        cellProps = sortColumnData(cellProps, cols, rows);

        if (includeSelectable && !lastScanResult.selectCellComponent) {
            figma.notify('⚠️ Selectable cells not available');
        }
        if (includeExpandable && !lastScanResult.expandCellComponent) {
            figma.notify('⚠️ Expandable cells not available');
        }

        const columnWidths: number[] = [];
        let totalTableWidth = 0;

        if (includeExpandable && lastScanResult.expandCellComponent) {
            try {
                const expandCellInstance = lastScanResult.expandCellComponent.createInstance();
                totalTableWidth += expandCellInstance.width;
                expandCellInstance.remove();
            } catch (error) {
                totalTableWidth += 52;
            }
        }
        if (includeSelectable && lastScanResult.selectCellComponent) {
            try {
                const selectCellInstance = lastScanResult.selectCellComponent.createInstance();
                totalTableWidth += selectCellInstance.width;
                selectCellInstance.remove();
            } catch (error) {
                totalTableWidth += 52;
            }
        }

        for (let c = 0; c < cols; c++) {
            let widthFound = false;
            for (let r = 0; r < rows; r++) {
                const key = `${r}-${c}`;
                const cellData = cellProps[key];
                if (cellData && cellData.colWidth) {
                    columnWidths[c] = cellData.colWidth;
                    widthFound = true;
                    break;
                }
            }
            if (!widthFound) {
                columnWidths[c] = 120;
            }
            totalTableWidth += columnWidths[c];
        }

        const tableFrame = figma.createFrame();
        tableFrame.name = "Generated Table";
        tableFrame.layoutMode = "VERTICAL";
        tableFrame.counterAxisSizingMode = "AUTO";
        tableFrame.primaryAxisSizingMode = "AUTO";
        tableFrame.itemSpacing = 0;
        tableFrame.paddingLeft = 0;
        tableFrame.paddingRight = 0;
        tableFrame.paddingTop = 0;
        tableFrame.paddingBottom = 0;
        tableFrame.fills = [];

        if (includeHeader && headerCell) {
            try {
                const headerRowFrame = figma.createFrame();
                headerRowFrame.name = "Header Row";
                headerRowFrame.layoutMode = "HORIZONTAL";
                headerRowFrame.primaryAxisSizingMode = "AUTO";
                headerRowFrame.counterAxisSizingMode = "AUTO";
                headerRowFrame.itemSpacing = 0;
                headerRowFrame.paddingLeft = 0;
                headerRowFrame.paddingRight = 0;
                headerRowFrame.paddingTop = 0;
                headerRowFrame.paddingBottom = 0;
                applyHeaderStyling(headerRowFrame);
                
                if (includeExpandable && lastScanResult.expandCellComponent) {
                    try {
                        const expandCell = lastScanResult.expandCellComponent.createInstance();
                        headerRowFrame.appendChild(expandCell);
                    } catch (error) { console.error('Error creating header expand cell'); }
                }

                if (includeSelectable && lastScanResult.selectCellComponent) {
                    try {
                        const selectCell = lastScanResult.selectCellComponent.createInstance();
                        headerRowFrame.appendChild(selectCell);
                    } catch (error) { console.error('Error creating header select cell'); }
                }

                for (let c = 1; c <= cols; c++) {
                    const hCell = headerCell.createInstance();
                    hCell.layoutSizingHorizontal = 'FIXED';
                    hCell.resize(columnWidths[c - 1], hCell.height);
                    headerRowFrame.appendChild(hCell);

                    const key = `header-${c}`;
                    const cellData = cellProps[key];
                    if (cellData && cellData.properties) {
                        try {
                            const validProps = mapPropertyNames(cellData.properties, hCell.componentProperties);
                            
                            try {
                                hCell.setProperties(validProps);
                            } catch (variantError) {
                                console.warn('Variant combination failed for header cell', key, 'trying individual properties');
                                for (const [propName, propValue] of Object.entries(validProps)) {
                                    try {
                                        hCell.setProperties({ [propName]: propValue });
                                    } catch (individualError) {
                                        console.warn(`Could not set individual property ${propName} for header cell ${key}:`, individualError);
                                    }
                                }
                            }
                        } catch (e) { console.warn('Could not set properties for header cell', key, e); }
                    }
                }
                tableFrame.appendChild(headerRowFrame);
            } catch (error) {
                console.error('Error creating header row:', error);
                figma.notify('⚠️ Header cell component is no longer available');
            }
        }

        if (bodyCell) {
            try {
                const bodyWrapperFrame = figma.createFrame();
                bodyWrapperFrame.name = 'Body';
                bodyWrapperFrame.layoutMode = 'VERTICAL';
                bodyWrapperFrame.primaryAxisSizingMode = 'AUTO';
                bodyWrapperFrame.counterAxisSizingMode = 'AUTO';
                bodyWrapperFrame.itemSpacing = 0;
                bodyWrapperFrame.paddingLeft = 0;
                bodyWrapperFrame.paddingRight = 0;
                bodyWrapperFrame.paddingTop = 0;
                bodyWrapperFrame.paddingBottom = 0;
                bodyWrapperFrame.fills = [];
                tableFrame.appendChild(bodyWrapperFrame);

                for (let r = 0; r < rows; r++) {
                    const bodyRowItemFrame = figma.createFrame();
                    bodyRowItemFrame.name = `Body Row Item ${r + 1}`;
                    bodyRowItemFrame.layoutMode = "VERTICAL";
                    bodyRowItemFrame.primaryAxisSizingMode = "AUTO";
                    bodyRowItemFrame.counterAxisSizingMode = "AUTO";
                    bodyRowItemFrame.itemSpacing = 0;
                    bodyRowItemFrame.paddingLeft = 0;
                    bodyRowItemFrame.paddingRight = 0;
                    bodyRowItemFrame.paddingTop = 0;
                    bodyRowItemFrame.paddingBottom = 0;
                    bodyRowItemFrame.fills = [];
                    
                    const rowFrame = figma.createFrame();
                    rowFrame.name = `Row ${r + 1}`;
                    rowFrame.layoutMode = "HORIZONTAL";
                    rowFrame.primaryAxisSizingMode = "AUTO";
                    rowFrame.counterAxisSizingMode = "AUTO";
                    rowFrame.itemSpacing = 0;
                    rowFrame.paddingLeft = 0;
                    rowFrame.paddingRight = 0;
                    rowFrame.paddingTop = 0;
                    rowFrame.paddingBottom = 0;
                    rowFrame.fills = [];

                    if (includeExpandable && lastScanResult.expandCellComponent) {
                        try {
                            const expandCell = lastScanResult.expandCellComponent.createInstance();
                            rowFrame.appendChild(expandCell);
                        } catch (error) { console.error('Error creating expand cell instance'); }
                    }

                    if (includeSelectable && lastScanResult.selectCellComponent) {
                        try {
                            const selectCell = lastScanResult.selectCellComponent.createInstance();
                            rowFrame.appendChild(selectCell);
                        } catch (error) { console.error('Error creating select cell instance'); }
                    }

                    for (let c = 0; c < cols; c++) {
                        const cell = bodyCell.createInstance();
                        cell.layoutSizingHorizontal = 'FIXED';
                        cell.resize(columnWidths[c], cell.height);
                        rowFrame.appendChild(cell);

                        const key = `${r}-${c}`;
                        const cellData = cellProps[key];
                        if (cellData && cellData.properties) {
                            try {
                                const validProps = mapPropertyNames(cellData.properties, cell.componentProperties);
                                cell.setProperties(validProps);
                            } catch (e) { console.warn('Could not set properties for cell', key, e); }
                        }
                    }
                    
                    bodyRowItemFrame.appendChild(rowFrame);
                    
                    if (lastScanResult.dividerComponent) {
                        try {
                            const divider = lastScanResult.dividerComponent.createInstance();
                            divider.resize(totalTableWidth, divider.height);
                            
                            // Try to apply color variable from footer if available
                            if (lastScanResult.footer) {
                                const colorVariable = extractFooterDividerColorVariable(lastScanResult.footer);
                                if (colorVariable && 'fills' in divider && divider.fills && Array.isArray(divider.fills)) {
                                    const currentFill = divider.fills[0];
                                    if (currentFill && currentFill.type === 'SOLID') {
                                        divider.fills = [{
                                            ...currentFill,
                                            boundVariables: { color: colorVariable }
                                        }];
                                    }
                                }
                            }
                            
                            bodyRowItemFrame.appendChild(divider);
                        } catch (error) { 
                            console.error('Error creating divider instance:', error);
                        }
                    } else {
                        const divider = createStyledDivider(totalTableWidth);
                        bodyRowItemFrame.appendChild(divider);
                    }
                    
                    bodyWrapperFrame.appendChild(bodyRowItemFrame);
                }
            } catch (error) {
                console.error('Error creating body rows:', error);
                figma.notify('❌ Cannot generate table: body cell component is no longer available');
                return;
            }
        }

        if (includeFooter && footer) {
            try {
                const footerClone = footer.createInstance();
                const footerCellData = cellProps['footer'];
                if (footerCellData && footerCellData.properties) {
                    const validProps = mapPropertyNames(footerCellData.properties, footerClone.componentProperties);
                    try {
                        footerClone.setProperties(validProps);
                    } catch (error) { console.error('Error in footer setProperties:', error); }
                }
                if (footerClone) {
                    footerClone.layoutSizingHorizontal = 'FIXED';
                    footerClone.resize(totalTableWidth, footerClone.height);
                    tableFrame.appendChild(footerClone);
                }
            } catch (error) {
                console.error('Error creating footer:', error);
                figma.notify('⚠️ Footer component is no longer available');
            }
        }

        tableFrame.resize(totalTableWidth, tableFrame.height);
        
        // Just use the frame directly as a component-like structure
        tableFrame.setPluginData('isGeneratedTable', 'true');
        const tableSettings = {
            columns: msg.cols,
            rows: msg.rows,
            includeHeader: msg.includeHeader,
            includeFooter: msg.includeFooter,
            includeSelectable: msg.includeSelectable,
            includeExpandable: msg.includeExpandable,
            cellProperties: msg.cellProps,
        };
        tableFrame.setPluginData('tableSettings', JSON.stringify(tableSettings));
        
        figma.currentPage.appendChild(tableFrame);
        figma.viewport.scrollAndZoomIntoView([tableFrame]);
        
        figma.notify('Table created successfully!');
        figma.ui.postMessage({ type: 'table-created', success: true });

    } catch (error) {
        console.error('Error in create-table-from-scan:', error);
        figma.notify('❌ An unexpected error occurred. Please check the console for details.');
    } finally {
        isCreatingTable = false;
    }
}

export async function handleTableUpdate(msg: any) {
    // For now, delegate to table creation handler with update logic
    // This can be expanded later for more sophisticated update handling
    await handleTableCreation(msg);
}

export { isCreatingTable };