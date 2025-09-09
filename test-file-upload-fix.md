# File Upload Header Display Fix

## Issue Description
When users upload a file using the "Populate data from file" option, the header cells in the grid preview don't show the file data on the first load. The header data only appears after making changes to rows/columns count because that triggers a re-render with proper "edited" state.

## Root Cause
The `renderHeaderFooterGrids()` function was always setting header cell text to `H${c}` (like "H1", "H2") instead of checking for actual data from uploaded files stored in `state.cellProperties`.

## Solution
Updated the `renderHeaderFooterGrids()` function to:

1. **Check for header data**: Look for data in `state.cellProperties` using the key `header-${c}`
2. **Find the correct text property**: Use the same logic as body cells to find the appropriate text property from the component
3. **Apply edited styling**: Add the "edited" class and bold font weight when displaying file data
4. **Fallback gracefully**: Show default `H${c}` text when no data is available

## Code Changes

### Before (renderHeaderFooterGrids function):
```typescript
cell.textContent = `H${c}`;  // Always shows H1, H2, etc.
```

### After (renderHeaderFooterGrids function):
```typescript
// Check for header data from uploaded file or user input
const headerKey = `header-${c}`;
const headerState = state.cellProperties.get(headerKey);
if (headerState && headerState.properties) {
  // Find the actual text property from the component
  let displayText = '';
  if (state.headerCellComponent?.availableProperties && state.headerCellComponent?.propertyTypes) {
    const textProp = state.headerCellComponent.availableProperties.find(p => 
      state.headerCellComponent!.propertyTypes[p] === 'TEXT');
    if (textProp && headerState.properties[textProp]) {
      displayText = headerState.properties[textProp];
    }
  }
  
  // Fallback to hardcoded property name
  if (!displayText && headerState.properties['Cell text#12234:32']) {
    displayText = headerState.properties['Cell text#12234:32'];
  }
  
  if (displayText) {
    cell.textContent = displayText;
    cell.classList.add('edited');
    cell.style.fontWeight = 'bold';
  } else {
    cell.textContent = `H${c}`;
  }
} else {
  cell.textContent = `H${c}`;
}
```

## Testing Steps
1. Upload a CSV/Excel/JSON file using "Populate data from file"
2. Verify header cells immediately show the file's header row data
3. Verify header cells have the "edited" class and bold styling
4. Verify changing rows/columns still works correctly
5. Verify the fix doesn't break existing functionality

## Additional Improvements
- Also updated footer grid rendering for consistency
- Maintained the same property lookup logic used by body cells
- Preserved all existing event handlers and tooltip functionality