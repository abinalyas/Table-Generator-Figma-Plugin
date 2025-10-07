// UI utility functions
import type { State, Elements } from '../types';

/**
 * Set the application mode
 */
export function setMode(state: State, newMode: State['mode']): void {
  state.mode = newMode;
  document.body.classList.remove('selection-mode', 'edit-mode');
  document.body.classList.add(`${newMode}-mode`);

  state.isDragging = false;
  const gridHighlight = document.getElementById('gridHighlight');
  if (gridHighlight) {
    gridHighlight.style.display = 'none';
  }

  updateModeDependentVisibility(state);

  // Show colWidthInput if default mode is column
  if (state.applyMode === 'column') {
    const colWidthInput = document.getElementById('colWidthInput') as HTMLInputElement;
    if (colWidthInput) {
      colWidthInput.style.display = 'inline-block';
    }
  }
}

/**
 * Update visibility of mode-dependent UI elements
 */
export function updateModeDependentVisibility(state: State): void {
  const actionButtons = document.getElementById('actionButtons');
  if (!actionButtons) return;

  // Only show action buttons if a valid component is selected and table size is confirmed
  if (state.hasComponent && state.sizeConfirmed) {
    actionButtons.style.display = 'flex';
  } else {
    actionButtons.style.display = 'none';
  }
}

/**
 * Update the create button state
 */
export function updateCreateButtonState(state: State, elements: Elements): void {
  elements.createTableBtn.disabled = !state.sizeConfirmed;
  elements.createTableBtn.style.opacity = state.sizeConfirmed ? '1' : '0.5';
}

/**
 * Get or create cell state
 */
export function getCellState(state: State, key: string): any {
  if (!state.cellProperties.has(key)) {
    state.cellProperties.set(key, { properties: {} });
  }
  return state.cellProperties.get(key);
}

/**
 * Clean property name by removing Figma IDs
 */
export function cleanPropName(name: string): string {
  return name.split('#')[0].trim();
}

/**
 * Handle apply option clicks
 */
export function handleApplyOptionClick(state: State, elements: Elements, target: HTMLElement): void {
  console.log(`[DEBUG] Apply option clicked: ${target.dataset.apply}`);

  // Remove active class from all options
  document.querySelectorAll('.apply-option').forEach(opt => {
    opt.classList.remove('active');
  });

  // Add active class to clicked option
  target.classList.add('active');

  // Update state
  const newApplyMode = target.dataset.apply as 'cell' | 'row' | 'column';
  state.applyMode = newApplyMode;

  // Show column width option only for cell and column apply modes
  if (elements.colWidthContainer) {
    if (state.applyMode === 'cell' || state.applyMode === 'column') {
      elements.colWidthContainer.style.display = 'block';
      console.log(`[DEBUG] Showing column width option for ${state.applyMode} mode`);
    } else {
      elements.colWidthContainer.style.display = 'none';
      console.log(`[DEBUG] Hiding column width option for ${state.applyMode} mode`);
    }
  }
}

/**
 * Setup API key synchronization between inputs
 */
export function setupApiKeySync(): void {
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
      parent.postMessage({
        pluginMessage: {
          type: 'save-watsonx-api-key',
          apiKey: apiKey
        }
      }, '*');
    }
  };

  // Sync from single prompt to watsonx cell generation
  spApiKeyInput.addEventListener('input', () => {
    const apiKey = spApiKeyInput.value;
    syncApiKey(spApiKeyInput, watsonxApiKeyInput, apiKey);
  });

  // Sync from watsonx cell generation to single prompt
  watsonxApiKeyInput.addEventListener('input', () => {
    const apiKey = watsonxApiKeyInput.value;
    syncApiKey(watsonxApiKeyInput, spApiKeyInput, apiKey);
  });

  // Also sync on blur (when user finishes typing)
  spApiKeyInput.addEventListener('blur', () => {
    const apiKey = spApiKeyInput.value;
    if (apiKey.trim()) {
      watsonxApiKeyInput.value = apiKey;
      parent.postMessage({
        pluginMessage: {
          type: 'save-watsonx-api-key',
          apiKey: apiKey
        }
      }, '*');
    }
  });

  watsonxApiKeyInput.addEventListener('blur', () => {
    const apiKey = watsonxApiKeyInput.value;
    if (apiKey.trim()) {
      spApiKeyInput.value = apiKey;
      parent.postMessage({
        pluginMessage: {
          type: 'save-watsonx-api-key',
          apiKey: apiKey
        }
      }, '*');
    }
  });
}

/**
 * Setup resize corner functionality
 */
export function setupResizeCorner(): void {
  const resizeCorner = document.getElementById('resizeCorner') as HTMLElement;

  if (!resizeCorner) {
    console.warn('Resize corner element not found');
    return;
  }

  let isResizing = false;

  const resizeWindow = (e: PointerEvent) => {
    if (!isResizing) return;

    const size = {
      w: Math.max(300, Math.floor(e.clientX + 5)),
      h: Math.max(400, Math.floor(e.clientY + 5))
    };

    // Send resize message to plugin
    parent.postMessage({
      pluginMessage: {
        type: 'resize',
        size: size
      }
    }, '*');
  };

  const handlePointerDown = (e: PointerEvent) => {
    isResizing = true;
    resizeCorner.setPointerCapture(e.pointerId);

    // Add event listeners for move and up
    resizeCorner.addEventListener('pointermove', resizeWindow);

    // Prevent default to avoid text selection
    e.preventDefault();
  };

  const handlePointerUp = (e: PointerEvent) => {
    if (!isResizing) return;

    isResizing = false;
    resizeCorner.releasePointerCapture(e.pointerId);

    // Remove event listeners
    resizeCorner.removeEventListener('pointermove', resizeWindow);
  };

  // Attach event listeners
  resizeCorner.addEventListener('pointerdown', handlePointerDown);
  resizeCorner.addEventListener('pointerup', handlePointerUp);

  // Handle pointer leave to stop resizing if pointer goes outside
  resizeCorner.addEventListener('pointerleave', (e: PointerEvent) => {
    if (isResizing) {
      handlePointerUp(e);
    }
  });

  console.log('Resize corner functionality initialized');
}