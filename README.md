# Table Generator Figma Plugin

A powerful Figma plugin for generating customizable data tables with AI integration and file import capabilities.

## Features

- **AI-Powered Content Generation**: Generate table content using IBM Watson AI
- **File Import Support**: Import data from CSV, Excel, and JSON files
- **Component-Based Architecture**: Modular design for maintainability and performance
- **Real-time Property Editing**: Edit cell properties with live preview
- **Flexible Grid System**: Customizable rows, columns, headers, and footers
- **Performance Optimized**: Lazy loading and efficient component initialization

## Architecture

The plugin uses a modular component architecture for better maintainability:

### Core Components

- **StateManager**: Centralized state management and UI synchronization
- **MessageHandler**: Plugin-to-Figma communication management
- **GridManager**: Grid rendering and cell interaction handling
- **PropertyEditor**: Dynamic property editing interface
- **PropertyRenderer**: UI rendering for different property types
- **EventManager**: Centralized event handling and DOM management
- **TooltipManager**: Cell tooltip display and management
- **FileUpload**: File parsing and data import functionality
- **RemainingUtils**: Utility functions for grid operations

### Performance Optimizations

- Lazy loading of non-critical components
- Optimized component initialization order
- Efficient dependency management
- Reduced bundle size through code splitting

## Development

### Build

```bash
npm run build
```

### Development Mode

```bash
npm run dev
```

## File Structure

```
src/
├── components/          # Modular components
├── types/              # TypeScript type definitions
├── utils/              # Utility functions
├── ui.ts              # Main UI entry point (2,112 lines)
└── main.ts            # Plugin main thread
```
