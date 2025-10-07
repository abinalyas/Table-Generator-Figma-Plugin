// Reset and state management utilities
import type { State, Elements } from '../types';
import { showMessage } from './helpers';

// Track if user has made any changes that can be reset
let hasChangesToReset = false;

/**
 * Update reset button state
 */
export function updateResetButtonState(elements: Elements): void {
  const resetBtn = elements.clearSelectionBtn;
  if (resetBtn) {
    resetBtn.disabled = !hasChangesToReset;
    resetBtn.style.opacity = hasChangesToReset ? '1' : '0.5';
    resetBtn.style.cursor = hasChangesToReset ? 'pointer' : 'not-allowed';
    resetBtn.title = hasChangesToReset ? 'Reset all table properties' : 'No changes to reset';
  }
}

/**
 * Mark that changes have been made
 */
export function markChangesForReset(elements: Elements): void {
  hasChangesToReset = true;
  updateResetButtonState(elements);
}

/**
 * Clear the changes flag
 */
export function clearChangesForReset(elements: Elements): void {
  hasChangesToReset = false;
  updateResetButtonState(elements);
}

/**
 * Reset table properties
 */
export function resetTableProperties(state: State, elements: Elements, gridManager: any): void {
  if (gridManager) {
    gridManager.clearGrid();
  } else {
    // Fallback if gridManager is not available
    state.cellProperties.clear();
    state.selectedCells.clear();
    state.currentEditingCell = null;
    state.sizeConfirmed = false;
  }

  // Clear file upload state
  const useFileUpload = document.getElementById('useFileUpload') as HTMLInputElement | null;
  const fileInput = document.getElementById('fileInput') as HTMLInputElement | null;
  const fileUploadContainer = document.getElementById('fileUploadContainer') as HTMLElement | null;
  const filePreview = document.getElementById('filePreview') as HTMLElement | null;

  if (useFileUpload?.checked) {
    useFileUpload.checked = false;
    if (fileUploadContainer) {
      fileUploadContainer.style.display = 'none';
    }
    if (fileInput) {
      fileInput.value = '';
    }
    if (filePreview) {
      filePreview.style.display = 'none';
    }
  }

  clearChangesForReset(elements);
  showMessage("All cell properties have been reset.", "success");
}