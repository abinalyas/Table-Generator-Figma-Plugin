import type { State, Elements, PropertyDefinition } from '../types';
import { RemainingUtils } from './RemainingUtils';

// Property field type definitions
export interface PropertyFieldConfig {
  name: string;
  label: string;
  type: 'TEXT' | 'VARIANT' | 'BOOLEAN';
  options?: string[];
  defaultValue?: string;
  dependsOn?: string;
  showWhen?: string;
}

export interface PropertyRenderOptions {
  availableProps: string[];
  propertyTypes: { [key: string]: any };
  props: any;
}

/**
 * PropertyRenderer component for rendering property fields in the UI
 * Handles all property field creation, updates, and validation
 */
export class PropertyRenderer {
  private state: State;
  private elements: Elements;

  // Property name mappings and configurations
  private readonly labelMap: { [key: string]: string } = {
    'Cell text#12234:32': 'Header Text',
    'Size': 'Size',
    'State': 'State',
    'Sorted': 'Sorted',
    'Sortable': 'Sortable',
    'Total items#12006:49': 'Total items',
    'Current page#12006:39': 'Current page',
    'Total pages#12006:29': 'Total pages',
    'Type': 'Type',
  };

  private readonly optionsMap: { [key: string]: string[] } = {
    'Size': ['Extra large', 'Large', 'Small'],
    'State': ['Enabled', 'Disabled', 'Focus'],
    'Sorted': ['Ascending', 'Descending', 'None'],
    'Sortable': ['True', 'False'],
    'Type': ['Advanced', 'Simple'],
  };

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
  }

  /**
   * Clean property name by removing hash suffix
   */
  cleanPropName(name: string): string {
    return RemainingUtils.cleanPropName(name);
  }

  /**
   * Find matching property from available properties
   */
  findMatchingProperty(availableProperties: string[], uiPropName: string): string | null {
    console.log('🔍 Finding match for:', uiPropName);
    console.log('Available properties:', availableProperties);
    
    // First try exact match
    const exactMatch = availableProperties.find((prop: string) => prop === uiPropName);
    if (exactMatch) {
      console.log('✅ Found exact match:', exactMatch);
      return exactMatch;
    }
    
    // Try match at start of property name
    const startsWithMatch = availableProperties.find((prop: string) => 
      prop.toLowerCase().startsWith(uiPropName.toLowerCase())
    );
    if (startsWithMatch) {
      console.log('✅ Found starts-with match:', startsWithMatch);
      return startsWithMatch;
    }
    
    // Try contains match
    const containsMatch = availableProperties.find((prop: string) => 
      prop.toLowerCase().includes(uiPropName.toLowerCase())
    );
    if (containsMatch) {
      console.log('✅ Found contains match:', containsMatch);
      return containsMatch;
    }
    
    console.log('❌ No match found');
    return null;
  }

  /**
   * Create a property field element based on type
   */
  private createPropertyField(
    propName: string, 
    type: string, 
    label: string, 
    value: any, 
    options?: string[]
  ): HTMLElement | null {
    let field: HTMLElement | null = null;

    if (type === 'VARIANT') {
      field = document.createElement('div');
      field.className = 'property-field';
      
      const select = document.createElement('select');
      select.className = 'styled-input';
      select.id = `dynamic-${propName}`;
      
      // Use mapped options if available, else fallback
      const selectOptions = options || this.optionsMap[propName] || [String(value)];
      for (const opt of selectOptions) {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        if (String(value) === opt) option.selected = true;
        select.appendChild(option);
      }
      
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      field.appendChild(labelEl);
      field.appendChild(select);
      
      // Add event listeners for special fields
      this.attachSelectEventListeners(select, label);
      
    } else if (type === 'BOOLEAN') {
      // Skip creating "Show text" checkbox - always hide it
      if (label.toLowerCase().includes('show text')) {
        return null; // Skip this field entirely
      }

      field = document.createElement('div');
      field.className = 'property-field checkbox';
      
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.id = `dynamic-${propName}`;
      input.checked = !!value;
      
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      field.appendChild(input);
      field.appendChild(labelEl);
      
      // Add event listeners for special fields
      this.attachCheckboxEventListeners(input, label);
      
    } else if (type === 'TEXT') {
      field = document.createElement('div');
      field.className = 'property-field';
      
      const labelEl = document.createElement('label');
      labelEl.textContent = label;
      
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'styled-input';
      input.id = `dynamic-${propName}`;
      input.value = value;
      
      field.appendChild(labelEl);
      field.appendChild(input);
    }

    return field;
  }

  /**
   * Attach event listeners for select elements
   */
  private attachSelectEventListeners(select: HTMLSelectElement, label: string): void {
    // --- Attach event for Show text ---
    if (label.toLowerCase().includes('show text')) {
      select.addEventListener('change', function () {
        const cellTextInput = document.getElementById('dynamic-Cell text#12234:16') as HTMLInputElement;
        if (cellTextInput) {
          const cellTextField = cellTextInput.parentElement as HTMLElement;
          if (cellTextField) {
            cellTextField.style.display = this.value === 'Cell text#12234:16' ? '' : 'none';
          }
        }
      });
    }
    
    // --- Attach event for Second text line ---
    if (label.toLowerCase().includes('second text line')) {
      select.addEventListener('change', function () {
        const secondCellTextInput = document.getElementById('dynamic-Second cell text#105573:16') as HTMLInputElement;
        if (secondCellTextInput) {
          const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
          if (secondCellTextField) {
            secondCellTextField.style.display = this.value === 'Second cell text#105573:16' ? '' : 'none';
          }
        }
      });
    }
  }

  /**
   * Attach event listeners for checkbox elements
   */
  private attachCheckboxEventListeners(input: HTMLInputElement, label: string): void {
    // --- Attach event for Show text ---
    if (label.toLowerCase().includes('show text')) {
      input.addEventListener('change', function () {
        const cellTextInput = document.getElementById('dynamic-Cell text#12234:16') as HTMLInputElement;
        if (cellTextInput) {
          const cellTextField = cellTextInput.parentElement as HTMLElement;
          if (cellTextField) {
            cellTextField.style.display = this.checked ? '' : 'none';
          }
        }
      });
    }
    
    // --- Attach event for Second text line ---
    if (label.toLowerCase().includes('second text line')) {
      input.addEventListener('change', function () {
        const secondCellTextInput = document.getElementById('dynamic-Second cell text#105573:16') as HTMLInputElement;
        if (secondCellTextInput) {
          const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
          if (secondCellTextField) {
            secondCellTextField.style.display = this.checked ? '' : 'none';
          }
        }
      });
    }
  }

  /**
   * Render dynamic property fields from available properties
   */
  renderDynamicPropertyFields(availableProps: string[], propertyTypes: { [key: string]: any }, props: any): void {
    const container = document.getElementById('dynamicPropertyFields');
    if (!container) return;
    
    container.innerHTML = '';
    container.style.border = '';
    container.style.background = '';
    let fieldCount = 0;
    
    console.log('[DEBUG] renderDynamicPropertyFields', { availableProps, propertyTypes, props });
    console.log('[DEBUG] renderDynamicPropertyFields props keys:', Object.keys(props));
    console.log('[DEBUG] renderDynamicPropertyFields props length:', Object.keys(props).length);

    // Check if we have wrong component info (only Size property or empty)
    const isWrongComponent = (availableProps.length === 1 && availableProps[0] === 'Size') ||
                            availableProps.length === 0;
    
    console.log('[DEBUG] renderDynamicPropertyFields isWrongComponent:', isWrongComponent);
    
    let propsToProcess = availableProps;
    let typesToUse = propertyTypes;
    
    // If wrong component info, use actual properties from cell data
    if (isWrongComponent && Object.keys(props).length > 0) {
      console.log('[DEBUG] Wrong component detected in renderDynamicPropertyFields, using actual cell properties:', Object.keys(props));
      propsToProcess = Object.keys(props);
      // Create fallback property types based on property names
      typesToUse = {};
      for (const propName of propsToProcess) {
        if (propName.toLowerCase().includes('text')) {
          typesToUse[propName] = 'TEXT';
        } else if (propName.toLowerCase().includes('slot') || propName.toLowerCase().includes('show')) {
          typesToUse[propName] = 'BOOLEAN';
        } else {
          typesToUse[propName] = 'VARIANT';
        }
      }
      console.log('[DEBUG] Using fallback property types in renderDynamicPropertyFields:', typesToUse);
    } else {
      console.log('[DEBUG] Not using fallback - isWrongComponent:', isWrongComponent, 'props length:', Object.keys(props).length);
    }

    for (const propName of propsToProcess) {
      // Skip Size property for all cell types since it can't be updated
      if (propName === 'Size') continue;

      // Skip "Cell text" property since we now have custom cell text functionality
      // But allow it when using fallback logic (wrong component info)
      if (propName.toLowerCase().includes('cell text') && !propName.toLowerCase().includes('second') && !isWrongComponent) {
        continue; // Skip this field entirely
      }

      const type = typesToUse[propName];
      const label = this.labelMap[propName] || this.cleanPropName(propName);
      const value = props[propName] ?? '';
      console.log(`[DEBUG] PropertyRenderer creating field for ${propName}, current value:`, value);
      
      const field = this.createPropertyField(propName, type, label, value);
      
      if (field) {
        container.appendChild(field);
        fieldCount++;
      }
    }
    
    // Post-processing: ensure proper field visibility
    this.postProcessFieldVisibility(container);
    
    console.log('[DEBUG] renderDynamicPropertyFields created fields:', fieldCount);
  }

  /**
   * Render dynamic property fields from a model definition
   */
  renderDynamicPropertyFieldsFromModel(model: PropertyDefinition[], props: any): void {
    const container = document.getElementById('dynamicPropertyFields');
    if (!container) return;
    
    container.innerHTML = '';
    container.style.border = '';
    container.style.background = '';
    let fieldCount = 0;

    console.log('[DEBUG] renderDynamicPropertyFieldsFromModel called with props:', props);

    for (const fieldDef of model) {
      const { name, label, type, options, defaultValue, dependsOn, showWhen } = fieldDef;

      // Check if field should be shown based on dependencies
      if (dependsOn && showWhen !== undefined) {
        const dependentValue = props[dependsOn];
        console.log(`[DEBUG] Checking dependency for ${name}: dependsOn=${dependsOn}, showWhen=${showWhen}, currentValue=${dependentValue}`);
        if (String(dependentValue) !== String(showWhen)) {
          console.log(`[DEBUG] Skipping ${name} - dependency condition not met`);
          continue; // Skip this field if dependency condition is not met
        }
        console.log(`[DEBUG] Showing ${name} - dependency condition met`);
      }

      // Use default value if no value is set
      const value = props[name] ?? defaultValue ?? '';

      const field = this.createPropertyField(name, type, label, value, options);
      
      if (field) {
        container.appendChild(field);
        fieldCount++;

        // Add change event listener for dependency handling
        const input = field.querySelector('select, input') as HTMLSelectElement | HTMLInputElement;
        if (input && type === 'VARIANT') {
          this.attachModelFieldEventListeners(input as HTMLSelectElement, name, model, props);
        }
      }
    }

    console.log('[DEBUG] renderDynamicPropertyFieldsFromModel created fields:', fieldCount);
  }

  /**
   * Render body cell properties with specific visibility rules
   */
  renderBodyCellProperties(availableProps: string[], propertyTypes: { [key: string]: any }, props: any): void {
    const container = document.getElementById('dynamicPropertyFields');
    if (!container) return;

    // Reset visibility
    container.style.display = '';

    // Check if text-related properties exist in availableProperties
    const hasShowTextProp = availableProps.some(p => {
      const type = propertyTypes[p];
      return type === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot');
    });
    const hasCellTextProp = availableProps.some(p => {
      const type = propertyTypes[p];
      return type === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second');
    });
    const hasSecondTextProp = availableProps.some(p => p.toLowerCase().includes('second'));
    const hasSlotProp = availableProps.some(p => p.toLowerCase().includes('slot'));

    // Hide text-related fields that duplicate dynamic properties
    if (hasShowTextProp && this.elements.showText) {
      const showTextField = this.elements.showText.parentElement;
      if (showTextField) showTextField.style.display = 'none';
    }

    // Always show custom cell text toggle for body cells
    if (this.elements.customCellTextToggle) {
      const customCellTextField = this.elements.customCellTextToggle.parentElement;
      if (customCellTextField) customCellTextField.style.display = '';
      // Ensure the custom cell text input is shown when toggle is checked
      if (this.elements.customCellTextToggle.checked && this.elements.customCellTextContainer) {
        this.elements.customCellTextContainer.style.display = 'block';
        console.log('[DEBUG] Setting customCellTextContainer display to block in renderBodyCellProperties');
      }
    }

    if (hasCellTextProp) {
      const cellTextField = this.elements.customCellTextContainer;
      if (cellTextField) cellTextField.style.display = 'none';
    } else {
      // Show cell text container if no dynamic cell text property exists
      const cellTextField = this.elements.customCellTextContainer;
      if (cellTextField) cellTextField.style.display = 'block';
    }

    if (hasSecondTextProp) {
      const secondTextLineField = this.elements.secondLineContainer;
      if (secondTextLineField) secondTextLineField.style.display = 'none';
      const secondCellTextField = this.elements.secondCellTextContainer;
      if (secondCellTextField) secondCellTextField.style.display = 'none';
    }

    if (hasSlotProp) {
      const slotField = this.elements.slotCheckbox.parentElement;
      if (slotField) slotField.style.display = 'none';
    }

    // Ensure "Generate sample data using AI" and "Column width" are always visible for body cells.
    if (this.elements.generateSampleCheckbox && this.elements.generateSampleCheckbox.parentElement) {
      this.elements.generateSampleCheckbox.parentElement.style.display = '';
    }
    if (this.elements.aiPromptContainer) {
      this.elements.aiPromptContainer.style.display = this.elements.generateSampleCheckbox?.checked ? 'block' : 'none';
    }
    if (this.elements.colWidthContainer) {
      this.elements.colWidthContainer.style.display = 'none';
    }
  }

  /**
   * Update static fields visibility and values for body cells
   */
  updateStaticFieldsVisibilityAndValues(availableProps: string[], propertyTypes: { [key: string]: any }, props: any): void {
    // Find dynamic property keys
    const showTextPropKey = availableProps.find(p => propertyTypes[p] === 'BOOLEAN' && (p.toLowerCase().includes('show') || p.toLowerCase().includes('text')) && !p.toLowerCase().includes('slot'));
    const cellTextPropKey = availableProps.find(p => propertyTypes[p] === 'TEXT' && p.toLowerCase().includes('text') && !p.toLowerCase().includes('second'));
    const secondTextPropKey = availableProps.find(p => p.toLowerCase().includes('second'));
    const slotPropKey = availableProps.find(p => propertyTypes[p] === 'BOOLEAN' && p.toLowerCase().includes('slot'));
    const statePropKey = availableProps.find(p => p.toLowerCase() === 'state' && propertyTypes[p] === 'VARIANT');

    // Hide static 'Second text line' if dynamic one exists
    if (secondTextPropKey && this.elements.secondLineContainer) {
      this.elements.secondLineContainer.style.display = 'none';
    }

    // Show/hide static fields
    if (this.elements.showText) {
      const showTextField = this.elements.showText.parentElement;
      if (showTextField) showTextField.style.display = 'none'; // Always hide show text checkbox
    }
    
    // Set custom cell text toggle to checked by default and show the input
    if (this.elements.customCellTextToggle) {
      this.elements.customCellTextToggle.checked = true;
      // Ensure the input is shown when toggle is checked by default
      if (this.elements.customCellTextContainer) {
        this.elements.customCellTextContainer.style.display = cellTextPropKey ? 'none' : 'block';
        console.log('[DEBUG] Setting customCellTextContainer display to', cellTextPropKey ? 'none' : 'block', 'in updateStaticFieldsVisibilityAndValues');
      }
    }
    
    console.log('[DEBUG] updateStaticFieldsVisibilityAndValues elements presence', {
      customCellText: !!this.elements.customCellText,
      secondCellText: !!this.elements.secondCellText,
      secondLineContainer: !!this.elements.secondLineContainer,
      secondCellTextContainer: !!this.elements.secondCellTextContainer,
      slotCheckbox: !!this.elements.slotCheckbox,
      state: !!this.elements.state
    });
    
    if (this.elements.secondLineContainer && this.elements.secondCellText && this.elements.customCellTextToggle) {
      this.elements.secondLineContainer.style.display = secondTextPropKey ? 'none' : (!!this.elements.secondCellText.value && this.elements.customCellTextToggle?.checked ? '' : 'none');
    }
    if (this.elements.secondCellTextContainer && this.elements.secondCellText && this.elements.customCellTextToggle) {
      this.elements.secondCellTextContainer.style.display = secondTextPropKey ? 'none' : (!!this.elements.secondCellText.value && this.elements.customCellTextToggle?.checked ? 'block' : 'none');
    }
    
    const slotField = this.elements.slotCheckbox?.parentElement;
    if (slotField) slotField.style.display = slotPropKey ? 'none' : '';
    const stateField = this.elements.state?.parentElement;
    if (stateField) stateField.style.display = statePropKey ? 'none' : '';

    // Populate static fields from props or reset to default
    if (this.elements.showText) {
      this.elements.showText.checked = showTextPropKey ? (props[showTextPropKey] || false) : true; // Always default to true since checkbox is hidden
    }
    if (this.elements.customCellTextToggle) {
      this.elements.customCellTextToggle.checked = true; // Always default to true for custom cell text
    }
    
    // Set custom cell text value - show existing cell text if available
    if (this.elements.customCellText) {
      let existingText = '';

      // Try to find existing text from various property keys
      if (cellTextPropKey && props[cellTextPropKey]) {
        existingText = props[cellTextPropKey];
      } else {
        // Look for any text property in the cell properties
        const textProps = Object.keys(props).filter(prop =>
          prop.toLowerCase().includes('text') &&
          !prop.toLowerCase().includes('second') &&
          typeof props[prop] === 'string' &&
          props[prop].trim() !== ''
        );
        if (textProps.length > 0) {
          existingText = props[textProps[0]];
        }
      }

      this.elements.customCellText.value = existingText;
      this.elements.customCellText.placeholder = 'Enter custom cell text';
    }
    
    const secondTextValue = secondTextPropKey ? (props[secondTextPropKey] || '') : '';
    if (this.elements.secondCellText) this.elements.secondCellText.value = secondTextValue;
    if (this.elements.secondTextLine) this.elements.secondTextLine.checked = !!secondTextValue;
    if (this.elements.slotCheckbox) {
      const slotValue = slotPropKey ? (props[slotPropKey] === true || props[slotPropKey] === 'true') : false;
      this.elements.slotCheckbox.checked = slotValue;
    }
    if (this.elements.state) this.elements.state.value = props['State'] || 'Enabled';

    // Always reset AI-related fields
    if (this.elements.generateSampleCheckbox) this.elements.generateSampleCheckbox.checked = false;
    const fakerInput = document.getElementById('fakerMethodInput') as HTMLInputElement | null;
    if (fakerInput) fakerInput.value = '';
    if (this.elements.aiPromptContainer) this.elements.aiPromptContainer.style.display = 'none';
    if (this.elements.customCellText) this.elements.customCellText.disabled = false;
    if (this.elements.customCellTextToggle) this.elements.customCellTextToggle.disabled = false;

    // --- For static property fields ---
    if (this.elements.secondLineContainer && this.elements.secondCellText && this.elements.secondCellTextContainer && this.elements.secondTextLine && !secondTextPropKey) {
      this.elements.secondLineContainer.style.display = !!this.elements.secondCellText.value ? '' : 'none';
      if (!this.elements.secondCellText.value) {
        this.elements.secondTextLine.checked = false;
        this.elements.secondCellTextContainer.style.display = 'none';
      }
    }
  }

  /**
   * Attach event listeners for model-based fields with dependency handling
   */
  private attachModelFieldEventListeners(
    select: HTMLSelectElement, 
    fieldName: string, 
    model: PropertyDefinition[], 
    props: any
  ): void {
    select.addEventListener('change', () => {
      // Update the props object with the new value
      props[fieldName] = select.value;
      console.log(`[DEBUG] Field ${fieldName} changed to: ${select.value}`);

      // Special logic: When Sortable is True and Sorted is None, set State to Hover
      if (fieldName === 'Sorted' && select.value === 'None') {
        const sortableInput = document.getElementById('dynamic-Sortable') as HTMLSelectElement;
        if (sortableInput && sortableInput.value === 'True') {
          const stateInput = document.getElementById('dynamic-State') as HTMLSelectElement;
          if (stateInput) {
            stateInput.value = 'Hover';
            props['State'] = 'Hover';
          }
        }
      }

      // Check if this field is a dependency for other fields
      const hasDependentFields = model.some(field => field.dependsOn === fieldName);
      if (hasDependentFields) {
        console.log(`[DEBUG] Field ${fieldName} has dependent fields, re-rendering form`);

        // Collect all current field values before re-rendering
        const currentValues = { ...props };
        
        // Collect values from all form fields
        const formInputs = document.querySelectorAll('#dynamicPropertyFields input, #dynamicPropertyFields select');
        formInputs.forEach((input) => {
          const element = input as HTMLInputElement | HTMLSelectElement;
          const fieldName = element.id.replace('dynamic-', '');
          if (element.type === 'checkbox') {
            currentValues[fieldName] = (element as HTMLInputElement).checked;
          } else {
            currentValues[fieldName] = element.value;
          }
        });

        // Re-render with updated values
        this.renderDynamicPropertyFieldsFromModel(model, currentValues);
      }
    });
  }

  /**
   * Post-process field visibility after rendering
   */
  private postProcessFieldVisibility(container: HTMLElement): void {
    // Since we're hiding the "Show text" checkbox, always show the Cell text field
    const cellTextInput = container.querySelector('input[type="text"][id^="dynamic-Cell text"]') as HTMLInputElement;
    if (cellTextInput) {
      const cellTextField = cellTextInput.parentElement as HTMLElement;
      if (cellTextField) {
        cellTextField.style.display = ''; // Always show cell text field
      }
    }
    
    // After all fields are rendered, ensure Second cell text is hidden if Second text line is unchecked
    const secondTextLineCheckbox = container.querySelector('input[type="checkbox"][id^="dynamic-Second text line"]') as HTMLInputElement;
    const secondCellTextInput = container.querySelector('input[type="text"][id^="dynamic-Second cell text"]') as HTMLInputElement;
    if (secondTextLineCheckbox && secondCellTextInput) {
      const secondCellTextField = secondCellTextInput.parentElement as HTMLElement;
      if (secondCellTextField) {
        secondCellTextField.style.display = secondTextLineCheckbox.checked ? '' : 'none';
      }
    }
    
    // Since we're hiding the "Show text" checkbox, always show the second text line container
    if (secondTextLineCheckbox && secondTextLineCheckbox.parentElement) {
      const secondLineContainerDyn = secondTextLineCheckbox.parentElement as HTMLElement;
      secondLineContainerDyn.style.display = ''; // Always show second text line container
    }
  }
}