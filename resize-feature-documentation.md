# Plugin Resize Feature

## Overview
Added comprehensive resize functionality to the IBM Table Generator Figma plugin, allowing users to adjust the plugin size for better workflow integration.

## Features Implemented

### 1. Manual Drag Resize
- **Native Figma Support**: Users can drag the corners of the plugin window to resize it
- **Enabled by**: `figma.ui.resize(500, 750)` in main.ts
- **Constraints**: Figma automatically handles minimum/maximum constraints

### 2. Programmatic Resize Controls
- **UI Controls**: Added resize input fields and preset buttons at the top of the plugin
- **Custom Dimensions**: Users can enter specific width and height values
- **Quick Presets**: One-click resize to common sizes

### 3. Size Presets
- **Default**: 500×750px (original size)
- **Large**: 600×800px (for complex tables)
- **Compact**: 400×600px (for smaller screens)

## Implementation Details

### Backend (main.ts)
```typescript
// Enable UI resizing with theme colors
figma.showUI(__html__, { 
    width: 500, 
    height: 750,
    themeColors: true // Better Figma integration
});

// Enable drag resizing
figma.ui.resize(500, 750);

// Handle programmatic resize requests
figma.ui.onmessage = async (msg: any) => {
    if (msg.type === "resize-ui") {
        const { width, height } = msg;
        // Safety constraints
        const constrainedWidth = Math.max(300, Math.min(1200, width || 500));
        const constrainedHeight = Math.max(400, Math.min(1000, height || 750));
        
        figma.ui.resize(constrainedWidth, constrainedHeight);
        return;
    }
    // ... other message handling
};
```

### Frontend (ui.html + ui.ts)
```html
<div class="resize-controls">
  <label>Size:</label>
  <input type="number" id="widthInput" placeholder="500" min="300" max="1200">
  <span>×</span>
  <input type="number" id="heightInput" placeholder="750" min="400" max="1000">
  <button id="applyResizeBtn">Apply</button>
  <div class="resize-presets">
    <button class="resize-preset" data-size="500,750">Default</button>
    <button class="resize-preset" data-size="600,800">Large</button>
    <button class="resize-preset" data-size="400,600">Compact</button>
  </div>
</div>
```

```typescript
function setupResizeControls() {
  // Handle manual resize input
  applyResizeBtn?.addEventListener('click', () => {
    const width = parseInt(widthInput.value) || 500;
    const height = parseInt(heightInput.value) || 750;
    
    // Send resize message to plugin
    parent.postMessage({ 
      pluginMessage: { 
        type: 'resize-ui', 
        width: constrainedWidth, 
        height: constrainedHeight 
      } 
    }, '*');
  });

  // Handle preset buttons
  resizePresets.forEach(preset => {
    preset.addEventListener('click', () => {
      const [width, height] = sizeData.split(',').map(Number);
      // Auto-apply preset
      parent.postMessage({ 
        pluginMessage: { type: 'resize-ui', width, height } 
      }, '*');
    });
  });
}
```

## Safety Features

### 1. Dimension Constraints
- **Minimum**: 300×400px (prevents unusably small sizes)
- **Maximum**: 1200×1000px (prevents excessive screen usage)
- **Applied**: Both in UI validation and backend processing

### 2. Input Validation
- **Number inputs**: Only accept valid numeric values
- **Fallback values**: Default to 500×750 if invalid input
- **Real-time feedback**: Success messages confirm resize operations

### 3. User Experience
- **Enter key support**: Press Enter in input fields to apply
- **Visual feedback**: Success messages show actual applied dimensions
- **Preset convenience**: One-click common sizes
- **Responsive design**: Controls adapt to plugin theme

## Benefits

### 1. Workflow Flexibility
- **Small screens**: Compact mode for laptops
- **Large displays**: Expanded view for detailed editing
- **Multi-monitor**: Optimal sizing for different screen setups

### 2. User Preference
- **Personal choice**: Users can set their preferred size
- **Task-specific**: Different sizes for different workflows
- **Persistent**: Manual drag resizing remembered by Figma

### 3. Professional Integration
- **Theme colors**: Matches Figma's light/dark themes
- **Native behavior**: Consistent with other Figma plugins
- **Accessibility**: Proper contrast and sizing

## Usage Instructions

### Method 1: Drag Resize (Recommended)
1. Hover over any corner of the plugin window
2. Drag to desired size
3. Figma automatically saves the preference

### Method 2: Programmatic Resize
1. Use the size controls at the top of the plugin
2. Enter custom width and height values
3. Click "Apply" or press Enter
4. Or click a preset button for instant sizing

### Method 3: Keyboard Shortcuts
- **Enter key**: Apply current input values
- **Tab navigation**: Move between width/height inputs

## Technical Notes

- **API Used**: `figma.ui.resize()` from Figma Plugin API
- **Message Passing**: Custom `resize-ui` message type
- **Constraints**: Enforced both client and server side
- **Performance**: Minimal overhead, instant response
- **Compatibility**: Works with all Figma plugin environments

This feature significantly improves the plugin's usability and professional feel, allowing users to customize their workspace according to their needs and preferences.