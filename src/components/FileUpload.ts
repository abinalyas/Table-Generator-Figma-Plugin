// File upload and processing functionality
import type { State, Elements } from '../types';
import { showMessage, hideMessage, showLoader, hideLoader } from '../utils/helpers';

export interface FileData {
  headers: string[];
  rows: string[][];
}

export class FileUpload {
  private state: State;
  private elements: Elements;
  private stateManager: any = null;

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
  }

  /**
   * Set external dependencies
   */
  setDependencies(dependencies: { stateManager?: any }): void {
    Object.assign(this, dependencies);
  }

  /**
   * Handle file upload event
   */
  async handleFileUpload(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;

    const file = input.files[0];
    
    try {
      showLoader('Parsing file...');
      
      let data: string[][] = [];
      
      if (file.name.endsWith('.csv')) {
        data = await this.parseCSV(file);
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        data = await this.parseExcel(file);
      } else if (file.name.endsWith('.json')) {
        data = await this.parseJSON(file);
      } else {
        throw new Error('Unsupported file format');
      }
      
      hideLoader();
      
      if (data.length === 0) {
        showMessage('File is empty or could not be parsed', 'error');
        return;
      }
      
      this.processFileData(data);
      
    } catch (error) {
      hideLoader();
      console.error('Error parsing file:', error);
      showMessage('Error parsing file: ' + (error as Error).message, 'error');
    }
  }

  /**
   * Parse CSV file
   */
  private parseCSV(file: File): Promise<string[][]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const text = e.target?.result as string;
          const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
          const data = lines.map(line => 
            line.split(',').map(cell => cell.trim())
          );
          resolve(data);
        } catch (error) {
          reject(new Error('Failed to parse CSV file'));
        }
      };
      
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsText(file);
    });
  }

  /**
   * Parse Excel file
   */
  private async parseExcel(file: File): Promise<string[][]> {
    return new Promise((resolve, reject) => {
      // Check if XLSX library is already loaded
      if (typeof (window as any).XLSX !== 'undefined') {
        this.processExcelFile((window as any).XLSX, file, resolve, reject);
        return;
      }

      // Load XLSX library dynamically
      const script = document.createElement('script');
      script.src = 'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js';
      
      script.onload = () => {
        // Give it a moment to initialize
        setTimeout(() => {
          if (typeof (window as any).XLSX !== 'undefined') {
            this.processExcelFile((window as any).XLSX, file, resolve, reject);
          } else {
            reject(new Error('Failed to load Excel parsing library - library not available after loading'));
          }
        }, 100);
      };
      
      script.onerror = () => reject(new Error('Failed to load Excel parsing library from CDN'));
      document.head.appendChild(script);
    });
  }

  /**
   * Process Excel file with XLSX library
   */
  private processExcelFile(XLSX: any, file: File, resolve: (data: string[][]) => void, reject: (error: Error) => void): void {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        resolve(jsonData as string[][]);
      } catch (error) {
        reject(new Error('Failed to parse Excel file'));
      }
    };
    
    reader.onerror = () => reject(new Error('Failed to read Excel file'));
    reader.readAsArrayBuffer(file);
  }

  /**
   * Parse JSON file
   */
  private parseJSON(file: File): Promise<string[][]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const text = e.target?.result as string;
          const jsonData = JSON.parse(text);
          
          if (Array.isArray(jsonData)) {
            if (jsonData.length === 0) {
              resolve([]);
              return;
            }
            
            if (typeof jsonData[0] === 'object' && jsonData[0] !== null) {
              // Array of objects - convert to rows
              const keys = Array.from(new Set(jsonData.flatMap(Object.keys)));
              const headers = [keys];
              jsonData.forEach(item => {
                const row = keys.map(key => item[key] ?? '');
                headers.push(row);
              });
              resolve(headers);
            } else {
              // Array of arrays
              resolve(jsonData);
            }
          } else if (typeof jsonData === 'object' && jsonData !== null) {
            // Object format
            if (Array.isArray(jsonData.columns) && Array.isArray(jsonData.rows)) {
              // Structured format
              const headers = [jsonData.columns];
              jsonData.rows.forEach((row: any) => {
                const rowData = jsonData.columns.map((col: string) => row[col] ?? '');
                headers.push(rowData);
              });
              resolve(headers);
            } else {
              // Key-value object
              const keys = Object.keys(jsonData);
              if (keys.length > 0) {
                const firstKey = keys[0];
                const firstValue = jsonData[firstKey];
                
                if (Array.isArray(firstValue)) {
                  const headers = [keys];
                  for (let i = 0; i < firstValue.length; i++) {
                    const row = keys.map(key => jsonData[key][i] ?? '');
                    headers.push(row);
                  }
                  resolve(headers);
                } else {
                  reject(new Error('Unsupported JSON structure'));
                }
              } else {
                reject(new Error('JSON object is empty'));
              }
            }
          } else {
            reject(new Error('JSON file must contain an array or object'));
          }
        } catch (error) {
          console.error('JSON parsing error:', error);
          reject(new Error('Failed to parse JSON file'));
        }
      };
      
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsText(file);
    });
  }

  /**
   * Process parsed file data
   */
  private processFileData(data: string[][]): void {
    console.log('File data received:', data);
    this.showFilePreview(data);
  }

  /**
   * Show file preview and update grid
   */
  private showFilePreview(data: string[][]): void {
    this.updateMainGridWithData(data);
    
    // Show the inline remove button
    const removeFileBtn = document.getElementById('removeFileBtn');
    if (removeFileBtn) {
      removeFileBtn.style.display = 'flex';
    }
  }

  /**
   * Update main grid with file data
   */
  private updateMainGridWithData(data: string[][]): void {
    const dataRows = data.length;
    const dataCols = data.length > 0 ? data[0].length : 0;
    
    if (dataRows === 0 || dataCols === 0) {
      showMessage('No data found in file', 'error');
      return;
    }
    
    // Update grid size
    const maxRows = Math.max(1, dataRows - 1); // Subtract 1 for header
    const maxCols = Math.max(1, dataCols);
    
    const rowsInput = document.getElementById('scanRowsInput') as HTMLInputElement;
    const colsInput = document.getElementById('scanColsInput') as HTMLInputElement;
    
    if (rowsInput && colsInput) {
      rowsInput.value = String(maxRows);
      colsInput.value = String(maxCols);
      this.stateManager?.setGridDimensions(maxRows, maxCols);
      
      // Emit event to update grid
      const event = new CustomEvent('gridSizeChanged', { 
        detail: { rows: maxRows, cols: maxCols } 
      });
      document.dispatchEvent(event);
    }
    
    // Populate header data
    if (data.length > 0) {
      const headers = data[0];
      const { cols: gridCols } = this.stateManager?.getGridDimensions() || { cols: 0 };
      for (let i = 0; i < headers.length && i < gridCols; i++) {
        const headerKey = `header-${i + 1}`;
        const headerText = String(headers[i] || '');
        
        const headerCellComponent = this.stateManager?.getHeaderCellComponent();
        let textProp = 'Cell text#12234:32'; // default
        
        if (headerCellComponent?.availableProperties && headerCellComponent?.propertyTypes) {
          const foundTextProp = headerCellComponent.availableProperties.find((prop: string) =>
            headerCellComponent.propertyTypes[prop] === 'TEXT'
          );
          if (foundTextProp) {
            textProp = foundTextProp;
          }
        }
        
        this.stateManager?.setCellProperties(headerKey, { [textProp]: headerText });
      }
    }
    
    // Populate body data
    const { rows: gridRows, cols: gridCols2 } = this.stateManager?.getGridDimensions() || { rows: 0, cols: 0 };
    for (let row = 1; row < data.length && row <= gridRows; row++) {
      const rowData = data[row];
      for (let col = 0; col < rowData.length && col < gridCols2; col++) {
        const cellKey = `${row},${col + 1}`;
        const cellText = String(rowData[col] || '');
        
        const selectedComponent = this.stateManager?.getSelectedComponent();
        let textProp = 'Cell text#12234:32'; // default
        
        if (selectedComponent?.availableProperties && selectedComponent?.propertyTypes) {
          const foundTextProp = selectedComponent.availableProperties.find((prop: string) =>
            selectedComponent.propertyTypes[prop] === 'TEXT'
          );
          if (foundTextProp) {
            textProp = foundTextProp;
          }
        }
        
        this.stateManager?.setCellProperties(cellKey, { [textProp]: cellText });
      }
    }
    
    // Emit event to refresh grid
    const event = new CustomEvent('fileDataLoaded');
    document.dispatchEvent(event);
    
    showMessage('File data loaded successfully', 'success');
  }

  /**
   * Reset grid to default state
   */
  resetGridToDefault(): void {
    // Emit event to clear grid
    const event = new CustomEvent('resetToDefault');
    document.dispatchEvent(event);
    
    showMessage('File data removed and grid reset to default state', 'success');
  }
}