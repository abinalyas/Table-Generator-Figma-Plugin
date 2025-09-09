# Remove File Button Layout Fix

## Issue
The remove file button was overlapping with the file upload input box when placed inline, causing layout issues and poor user experience.

## Solution
Moved the remove file button below the file input and added both trash icon and "Remove file" text for better clarity and accessibility.

## Changes Made

### 1. Layout Structure
**Before**: Button inline with file input (causing overlap)
```html
<div class="file-input-container">
  <input type="file" id="dataFileInput" class="styled-input file-input">
  <button id="removeFileBtn">🗑️</button>
</div>
```

**After**: Button below file input (clean separation)
```html
<div class="property-field">
  <label for="dataFileInput">Upload Data File</label>
  <input type="file" id="dataFileInput" class="styled-input">
  <div class="hint">Supported formats: CSV, Excel (.xlsx, .xls), JSON</div>
  <button id="removeFileBtn">
    <svg>...</svg>
    <span>Remove file</span>
  </button>
</div>
```

### 2. Visual Improvements
- **Icon + Text**: Combined trash icon with "Remove file" text
- **Better Spacing**: 8px margin-top for proper separation
- **Proper Sizing**: `width: fit-content` prevents button stretching
- **Gap**: 6px gap between icon and text

### 3. Enhanced Styling
```css
.remove-file-btn {
  background: none;
  border: none;
  cursor: pointer;
  padding: 8px 12px;
  border-radius: 4px;
  color: #c00000;
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 13px;
  margin-top: 8px;
  transition: all 0.2s ease;
  width: fit-content;
}

.remove-file-btn:hover {
  background-color: rgba(192, 0, 0, 0.1);
  color: #a00000;
}

.remove-file-btn svg {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
}

.remove-file-btn span {
  font-weight: 500;
}
```

## Benefits

### 1. No Overlap Issues
- **Clean Layout**: File input has full width without interference
- **Proper Spacing**: Clear visual separation between elements
- **Responsive**: Works well on different screen sizes

### 2. Better Accessibility
- **Clear Labeling**: Both icon and text make purpose obvious
- **Screen Readers**: Text content is accessible to assistive technology
- **Visual Hierarchy**: Logical flow from input to hint to action

### 3. Professional Appearance
- **Consistent Styling**: Matches overall plugin design
- **Hover Effects**: Clear interactive feedback
- **Icon + Text**: Best of both worlds for recognition and clarity

## Layout Flow
```
┌─────────────────────────────────┐
│ Upload Data File                │ ← Label
├─────────────────────────────────┤
│ [Choose File] No file chosen    │ ← File Input (full width)
├─────────────────────────────────┤
│ Supported formats: CSV, Excel...│ ← Hint
├─────────────────────────────────┤
│ 🗑️ Remove file                  │ ← Remove Button (when file loaded)
└─────────────────────────────────┘
```

## User Experience
- **Clear Actions**: Users can easily see and access the remove option
- **No Confusion**: Button appears only when relevant (file is loaded)
- **Intuitive Flow**: Natural progression from upload to remove
- **Professional Feel**: Clean, organized layout

This fix resolves the overlap issue while maintaining a professional and accessible user interface.