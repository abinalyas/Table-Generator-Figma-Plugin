// Sorting utilities for grid data

export function sortColumnData(cellProperties: Map<string, any>, cols: number, rows: number): Map<string, any> {
  let hasSortableColumns = false;

  // Quick check if any columns need sorting
  for (let c = 1; c <= cols; c++) {
    const headerData = cellProperties.get(`header-${c}`);
    if (headerData?.properties?.['Sortable'] === 'True' &&
      (headerData.properties['Sorted'] === 'Ascending' || headerData.properties['Sorted'] === 'Descending')) {
      hasSortableColumns = true;
      break;
    }
  }

  // Return original if no sorting needed
  if (!hasSortableColumns) return cellProperties;

  const sortedCellProperties = new Map(cellProperties);

  for (let c = 1; c <= cols; c++) {
    const headerData = sortedCellProperties.get(`header-${c}`);
    const sortable = headerData?.properties?.['Sortable'];
    const sorted = headerData?.properties?.['Sorted'];

    if (sortable === 'True' && (sorted === 'Ascending' || sorted === 'Descending')) {
      const columnData: { rowIndex: number; value: string }[] = [];

      for (let r = 1; r <= rows; r++) {
        const cellData = sortedCellProperties.get(`${r},${c}`);
        let cellValue = '';
        if (cellData?.properties) {
          if (typeof cellData.properties['Cell text'] === 'string' && cellData.properties['Cell text'].toString().trim() !== '') {
            cellValue = cellData.properties['Cell text'];
          } else {
            const textProp = Object.keys(cellData.properties).find(prop =>
              prop.toLowerCase().includes('text') && !prop.toLowerCase().includes('second') && typeof cellData.properties[prop] === 'string');
            cellValue = (textProp && cellData.properties[textProp]) || cellData.properties['Cell text#12234:32'] || '';
          }
        }
        columnData.push({ rowIndex: r, value: String(cellValue) });
      }

      columnData.sort((a, b) => {
        const numA = parseFloat(a.value);
        const numB = parseFloat(b.value);
        if (!isNaN(numA) && !isNaN(numB)) {
          return sorted === 'Ascending' ? numA - numB : numB - numA;
        }
        return sorted === 'Ascending' ? a.value.localeCompare(b.value) : b.value.localeCompare(a.value);
      });

      // Reorder entire rows based on the sorted order of the selected column
      const newMap = new Map(sortedCellProperties);
      for (let newRow = 1; newRow <= rows; newRow++) {
        const sourceRow = columnData[newRow - 1].rowIndex;
        for (let cc = 1; cc <= cols; cc++) {
          const src = sortedCellProperties.get(`${sourceRow},${cc}`);
          newMap.set(`${newRow},${cc}`, src);
        }
      }
      // Replace with reordered map and stop after first sortable column
      return newMap;
    }
  }

  return sortedCellProperties;
}