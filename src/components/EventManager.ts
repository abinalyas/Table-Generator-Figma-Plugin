import type { State, Elements } from '../types';
import type { MessageHandler as MessageHandlerClass } from './MessageHandler';

/**
 * EventManager component for centralized event handling
 * Manages all DOM event listeners and user interactions
 */
export class EventManager {
  private state: State;
  private elements: Elements;
  private messageHandler: MessageHandlerClass | null = null;
  private fileUpload: any = null;
  private propertyEditor: any = null;

  // Event listener cleanup tracking
  private eventListeners: Array<{
    element: Element | Document | Window;
    event: string;
    handler: EventListener;
  }> = [];

  constructor(state: State, elements: Elements) {
    this.state = state;
    this.elements = elements;
  }

  /**
   * Set external dependencies
   */
  setDependencies(dependencies: {
    messageHandler?: MessageHandlerClass | null;
    fileUpload?: any;
    propertyEditor?: any;
    showCellTooltip?: (cell: HTMLDivElement) => void;
    hideCellTooltip?: () => void;
    resetTableProperties?: () => void;
    createTable?: () => void;
  }): void {
    Object.assign(this, dependencies);
  }

  /**
   * Add event listener with cleanup tracking
   */
  private addEventListener(
    element: Element | Document | Window,
    event: string,
    handler: EventListener,
    options?: boolean | AddEventListenerOptions
  ): void {
    element.addEventListener(event, handler, options);
    this.eventListeners.push({ element, event, handler });
  }

  /**
   * Setup all event listeners
   */
  setupEventListeners(): void {
    // Property editor events
    this.addEventListener(this.elements.clearSelectionBtn, 'click', () => {
      (this as any).resetTableProperties?.();
    });

    this.addEventListener(this.elements.createTableBtn, 'click', () => {
      (this as any).createTable?.();
    });

    this.addEventListener(this.elements.cancelPropsBtn, 'click', () => {
      this.propertyEditor?.closeEditor();
    });

    this.addEventListener(this.elements.propertyEditorOverlay, 'click', (e) => {
      if (e.target === this.elements.propertyEditorOverlay) {
        this.propertyEditor?.closeEditor();
      }
    });

    this.addEventListener(this.elements.savePropsBtn, 'click', () => {
      this.propertyEditor?.saveProperties();
    });

    // Custom cell text toggle
    this.addEventListener(this.elements.customCellTextToggle, 'change', (e) => {
      const target = e.target as HTMLInputElement;
      const show = target.checked;
      this.elements.customCellTextContainer.style.display = show ? 'block' : 'none';
      
      // Only show second line if custom cell text is enabled
      if (show) {
        this.elements.secondLineContainer.style.display = 'flex';
      } else {
        this.elements.secondLineContainer.style.display = 'none';
        if (this.elements.secondTextLine) {
          this.elements.secondTextLine.checked = false;
          this.elements.secondCellTextContainer.style.display = 'none';
        }
      }
    });

    // Second text line toggle
    this.addEventListener(this.elements.secondTextLine, 'change', (e) => {
      const target = e.target as HTMLInputElement;
      this.elements.secondCellTextContainer.style.display = target.checked ? 'block' : 'none';
    });

    // Apply option clicks
    document.querySelectorAll('.apply-option').forEach(option => {
      this.addEventListener(option, 'click', this.handleApplyOptionClick.bind(this));
    });

    // AI generation checkbox
    this.addEventListener(this.elements.generateSampleCheckbox, 'change', () => {
      const isChecked = this.elements.generateSampleCheckbox.checked;
      this.elements.customCellText.disabled = isChecked;

      if (isChecked) {
        this.elements.aiPromptContainer.style.display = 'block';
        requestAnimationFrame(() => {
          this.elements.aiPromptContainer.classList.add('show');
        });
        // Ensure the correct AI config is shown immediately on first enable
        this.updateAiConfigVisibility();
      } else {
        this.elements.aiPromptContainer.classList.remove('show');
        setTimeout(() => {
          this.elements.aiPromptContainer.style.display = 'none';
        }, 250);
        // Hide both configs when disabling
        this.updateAiConfigVisibility();
      }

      // Disable custom cell text toggle when AI generation is enabled
      this.elements.customCellTextToggle.disabled = isChecked;
    });

    // AI source toggle
    const aiSource = document.getElementById('aiSource') as HTMLSelectElement | null;
    if (aiSource) {
      this.addEventListener(aiSource, 'change', () => {
        this.updateAiConfigVisibility();
      });
    }

    // Cell tooltip events
    this.addEventListener(document, 'mouseover', (e) => {
      const target = e.target as HTMLElement;
      if (target.classList.contains('cell') || target.classList.contains('header-cell') || target.classList.contains('footer-cell')) {
        (this as any).showCellTooltip?.(target as HTMLDivElement);
      }
    });

    this.addEventListener(document, 'mouseout', (e) => {
      const target = e.target as HTMLElement;
      if (target.classList.contains('cell') || target.classList.contains('header-cell') || target.classList.contains('footer-cell')) {
        (this as any).hideCellTooltip?.();
      }
    });

    // Load persisted watsonx settings
    this.messageHandler?.loadWatsonxSettings();

    // Setup API key synchronization
    this.setupApiKeySync();
  }

  /**
   * Handle apply option clicks
   */
  private handleApplyOptionClick(e: Event): void {
    const target = e.target as HTMLElement;
    console.log(`[DEBUG] Apply option clicked: ${target.dataset.apply}`);

    // Remove active class from all options
    document.querySelectorAll('.apply-option').forEach(opt => {
      opt.classList.remove('active');
    });

    // Add active class to clicked option
    target.classList.add('active');

    // Update state
    const newApplyMode = target.dataset.apply as 'cell' | 'row' | 'column';
    this.state.applyMode = newApplyMode;

    // Show column width option only for cell and column apply modes
    if (this.elements.colWidthContainer) {
      if (this.state.applyMode === 'cell' || this.state.applyMode === 'column') {
        this.elements.colWidthContainer.style.display = 'block';
      } else {
        this.elements.colWidthContainer.style.display = 'none';
      }
    }
  }

  /**
   * Update AI config visibility based on selection
   */
  private updateAiConfigVisibility(): void {
    const aiSource = document.getElementById('aiSource') as HTMLSelectElement | null;
    const fakerConfig = document.getElementById('fakerConfig') as HTMLElement | null;
    const watsonxConfig = document.getElementById('watsonxConfig') as HTMLElement | null;
    const isChecked = this.elements.generateSampleCheckbox.checked;

    if (!aiSource || (!fakerConfig && !watsonxConfig)) return;

    if (!isChecked) {
      // Hide both when AI generation is disabled
      if (fakerConfig) { 
        fakerConfig.classList.remove('show'); 
        fakerConfig.style.display = 'none'; 
      }
      if (watsonxConfig) { 
        watsonxConfig.classList.remove('show'); 
        watsonxConfig.style.display = 'none'; 
      }
      return;
    }

    // Show only the selected source config
    const selectedValue = aiSource.value;
    if (fakerConfig) {
      if (selectedValue === 'faker') {
        fakerConfig.style.display = 'block';
        requestAnimationFrame(() => { fakerConfig.classList.add('show'); });
      } else {
        fakerConfig.classList.remove('show');
        setTimeout(() => { fakerConfig.style.display = 'none'; }, 250);
      }
    }
    if (watsonxConfig) {
      if (selectedValue === 'watsonx') {
        watsonxConfig.style.display = 'block';
        requestAnimationFrame(() => { watsonxConfig.classList.add('show'); });
      } else {
        watsonxConfig.classList.remove('show');
        setTimeout(() => { watsonxConfig.style.display = 'none'; }, 250);
      }
    }
  }

  /**
   * Setup API key synchronization between inputs
   */
  setupApiKeySync(): void {
    const spApiKeyInput = document.getElementById('spApiKey') as HTMLInputElement | null;
    const watsonxApiKeyInput = document.getElementById('watsonxApiKey') as HTMLInputElement | null;

    if (!spApiKeyInput || !watsonxApiKeyInput) {
      console.warn('API key inputs not found');
      return;
    }

    // Function to sync API key between both inputs
    const syncApiKey = (sourceInput: HTMLInputElement, targetInput: HTMLInputElement, apiKey: string) => {
      if (apiKey.trim() && targetInput.value.trim() === '') {
        targetInput.value = apiKey;
        // Save the API key
        this.messageHandler?.saveWatsonxApiKey(apiKey);
      }
    };

    // Sync from single prompt to watsonx cell generation
    this.addEventListener(spApiKeyInput, 'input', () => {
      const apiKey = spApiKeyInput.value;
      syncApiKey(spApiKeyInput, watsonxApiKeyInput, apiKey);
    });

    // Sync from watsonx cell generation to single prompt
    this.addEventListener(watsonxApiKeyInput, 'input', () => {
      const apiKey = watsonxApiKeyInput.value;
      syncApiKey(watsonxApiKeyInput, spApiKeyInput, apiKey);
    });

    // Also sync on blur (when user finishes typing)
    this.addEventListener(spApiKeyInput, 'blur', () => {
      const apiKey = spApiKeyInput.value;
      if (apiKey.trim()) {
        watsonxApiKeyInput.value = apiKey;
        this.messageHandler?.saveWatsonxApiKey(apiKey);
      }
    });

    this.addEventListener(watsonxApiKeyInput, 'blur', () => {
      const apiKey = watsonxApiKeyInput.value;
      if (apiKey.trim()) {
        spApiKeyInput.value = apiKey;
        this.messageHandler?.saveWatsonxApiKey(apiKey);
      }
    });
  }

  /**
   * Setup file upload controls
   */
  setupFileUploadControls(): void {
    const removeFileBtn = document.getElementById('removeFileBtn');

    if (removeFileBtn) {
      this.addEventListener(removeFileBtn, 'click', () => {
        // Clear the file input
        const fileInput = document.getElementById('dataFileInput') as HTMLInputElement | null;
        if (fileInput) {
          fileInput.value = '';
        }

        // Hide the remove button
        removeFileBtn.style.display = 'none';

        // Reset grid using FileUpload component
        this.fileUpload?.resetGridToDefault();
      });
    }
  }

  /**
   * Setup resize corner functionality
   */
  setupResizeCorner(): void {
    const resizeCorner = document.getElementById('resizeCorner') as HTMLElement;

    if (!resizeCorner) {
      console.warn('Resize corner element not found');
      return;
    }

    let isResizing = false;

    const resizeWindow = (e: Event) => {
      if (!isResizing) return;
      const pointerEvent = e as PointerEvent;

      const size = {
        w: Math.max(300, Math.floor(pointerEvent.clientX + 5)),
        h: Math.max(400, Math.floor(pointerEvent.clientY + 5))
      };

      // Send resize message to plugin
      this.messageHandler?.resize(size);
    };

    const handlePointerDown = (e: Event) => {
      const pointerEvent = e as PointerEvent;
      isResizing = true;
      resizeCorner.setPointerCapture(pointerEvent.pointerId);

      // Add event listeners for move and up
      this.addEventListener(resizeCorner, 'pointermove', resizeWindow);

      // Prevent default to avoid text selection
      pointerEvent.preventDefault();
    };

    const handlePointerUp = (e: Event) => {
      if (!isResizing) return;
      const pointerEvent = e as PointerEvent;

      isResizing = false;
      resizeCorner.releasePointerCapture(pointerEvent.pointerId);

      // Remove event listeners
      resizeCorner.removeEventListener('pointermove', resizeWindow);
    };

    // Attach event listeners
    this.addEventListener(resizeCorner, 'pointerdown', handlePointerDown);
    this.addEventListener(resizeCorner, 'pointerup', handlePointerUp);

    // Handle pointer leave to stop resizing if pointer goes outside
    this.addEventListener(resizeCorner, 'pointerleave', (e: Event) => {
      if (isResizing) {
        handlePointerUp(e);
      }
    });

    console.log('Resize corner functionality initialized');
  }

  /**
   * Clean up all event listeners
   */
  cleanup(): void {
    this.eventListeners.forEach(({ element, event, handler }) => {
      element.removeEventListener(event, handler);
    });
    this.eventListeners = [];
  }
}