# Code Walkthrough Guide - IBM Table Generator Figma Plugin

## 📋 Table of Contents
1. [Project Overview](#project-overview)
2. [Architecture](#architecture)
3. [Project Structure](#project-structure)
4. [Entry Points](#entry-points)
5. [Key Components](#key-components)
6. [Data Flow](#data-flow)
7. [Main Workflows](#main-workflows)
8. [Key Features](#key-features)
9. [Server Architecture](#server-architecture)
10. [Analytics System](#analytics-system)

---

## 🎯 Project Overview

**What it does:**
A Figma plugin that generates customizable data tables with:
- AI-powered content generation (IBM Watsonx)
- File import (CSV, JSON, Excel)
- Smart component detection and slot suggestions
- Real-time property editing
- Table scanning from existing Figma tables

**Tech Stack:**
- **Frontend:** TypeScript, HTML/CSS
- **Figma API:** Plugin API (main thread + UI thread)
- **Server:** Node.js/Express (IBM Code Engine)
- **AI:** IBM Watsonx.ai via proxy server
- **Analytics:** Google Analytics 4 (Measurement Protocol)

---

## 🏗️ Architecture

### Two-Thread Architecture (Figma Plugin Standard)

```
┌─────────────────┐         ┌─────────────────┐
│   UI Thread     │ ◄─────► │  Main Thread    │
│   (ui.ts)       │         │   (main.ts)     │
│   - User UI     │         │   - Figma API   │
│   - State Mgmt  │         │   - Node ops    │
│   - Events      │         │   - Component   │
└─────────────────┘         └─────────────────┘
        │                            │
        │                            │
        ▼                            ▼
┌─────────────────────────────────────────┐
│      Proxy Server (Code Engine)        │
│   - /generate (AI text generation)      │
│   - /generateTable (AI table gen)       │
│   - /analytics (GA4 forwarding)         │
│   - /token (IAM token)                  │
└─────────────────────────────────────────┘
```

**Key Concept:** 
- **UI Thread** (`ui.ts`): Runs in iframe, handles user interactions
- **Main Thread** (`main.ts`): Runs in Figma sandbox, has access to Figma API
- **Communication:** Via `postMessage` between threads

---

## 📁 Project Structure

```
Table-Generator-Figma-Plugin/
├── src/
│   ├── main.ts                    # Main thread (Figma API access)
│   ├── ui.ts                      # UI thread (user interface)
│   ├── ui.html                    # HTML template
│   │
│   ├── core/                      # Core state management
│   │   ├── state.ts              # Application state
│   │   ├── selection.ts          # Selection handling
│   │   └── metadata.ts           # Metadata management
│   │
│   ├── handlers/                  # Feature handlers
│   │   ├── MessageRouter.ts      # Routes messages to handlers
│   │   ├── TableCreationHandler.ts
│   │   ├── TableScanHandler.ts
│   │   ├── ComponentHandler.ts
│   │   ├── DataGenerationHandler.ts
│   │   ├── aiHandlers.ts
│   │   └── SettingsHandler.ts
│   │
│   ├── services/
│   │   └── TableScanService.ts   # Table scanning logic
│   │
│   ├── utils/                      # Utility functions
│   │   ├── analytics.ts          # Analytics tracking
│   │   ├── component.ts          # Component utilities
│   │   ├── faker.ts              # Faker.js integration
│   │   ├── helpers.ts            # General helpers
│   │   ├── PropertyMapper.ts    # Property mapping
│   │   └── ... (other utilities)
│   │
│   └── types/
│       └── index.ts              # TypeScript definitions
│
├── proxy-server.js                 # Express server (Code Engine)
├── server-package.json            # Server dependencies
├── Dockerfile                     # Docker build config
├── webpack.config.js              # Build configuration
└── manifest.json                  # Figma plugin manifest
```

---

## 🚀 Entry Points

### 1. **main.ts** (Main Thread - ~6900 lines)
**Purpose:** Figma plugin main thread with Figma API access

**Key Responsibilities:**
- Component scanning and discovery
- Table creation and manipulation
- Figma node operations
- Message handling from UI

**Key Functions:**
```typescript
// Message handler
figma.ui.onmessage = async (msg: any) => {
  // Routes different message types
  if (msg.type === 'create-table-from-scan') { ... }
  if (msg.type === 'scan-table') { ... }
  if (msg.type === 'generate-table-with-ai') { ... }
  // ... etc
}

// Component scanning
async function scanTableForComponents() { ... }

// Table creation
async function createTableFromScan(...) { ... }
```

**Important State:**
- `lastScanResult`: Stores scanned table components
- `isCreatingTable`: Prevents concurrent table creation
- `cellInstanceMap`: Maps cell keys to Figma instances

---

### 2. **ui.ts** (UI Thread - ~6000 lines)
**Purpose:** User interface and state management

**Key Responsibilities:**
- Grid rendering and cell visualization
- Property editor UI
- User input handling
- State synchronization
- Analytics tracking

**Key Functions:**
```typescript
// State management
const state: State = { ... }
function getCellState(key: string) { ... }

// Grid rendering
function renderGrid() { ... }
function updateCellVisuals() { ... }

// Property editing
function openPropertyEditor(key: string) { ... }
function saveCellProperties() { ... }

// Message handling
window.onmessage = (event) => { ... }
```

**Important State:**
- `state.cellProperties`: Map of cell configurations
- `state.selectedCells`: Currently selected cells
- `state.gridRows/gridCols`: Grid dimensions
- `state.selectedComponent`: Current component info

---

## 🔑 Key Components

### 1. **MessageRouter.ts**
Routes messages between UI and main thread to appropriate handlers.

**Flow:**
```
UI sends message → MessageRouter → Handler → Response → UI
```

**Message Types:**
- `create-table-from-scan`
- `scan-table`
- `generate-table-with-ai`
- `update-table`
- `generate-watsonx-data`
- `generate-fake-data`

---

### 2. **TableScanService.ts**
Scans existing Figma tables to extract:
- Component structure
- Cell components
- Header/footer components
- Layout information

**Key Method:**
```typescript
static async scanTable(tableNode: FrameNode): Promise<ScanResult>
```

---

### 3. **ComponentHandler.ts**
Handles Figma component operations:
- Component discovery
- Property mapping
- Component swapping

---

### 4. **DataGenerationHandler.ts**
Manages AI data generation:
- Watsonx API calls
- Faker.js integration
- Data formatting

---

### 5. **analytics.ts** (Utility)
Tracks user events and sends to GA4:
- `init()`: Initialize analytics
- `track()`: Send event
- `trackSessionEnd()`: Track session closure

**Tracks:**
- `plugin_open`
- `table_created`
- `ai_generate_single`
- `ai_generate_watsonx`
- `session_start`
- `session_end`

---

## 🔄 Data Flow

### Table Creation Flow

```
1. User clicks "Create Table"
   ↓
2. UI sends: { type: 'create-table-from-scan', rows, cols, cellProps }
   ↓
3. Main thread receives message
   ↓
4. Main thread:
   - Gets scanned components (lastScanResult)
   - Creates table frame
   - Creates header row (if enabled)
   - Creates body rows with cells
   - Applies cell properties
   - Configures slots (tags, avatars, etc.)
   ↓
5. Main thread sends: { type: 'table-created', success: true }
   ↓
6. UI receives success message
   ↓
7. UI tracks: table_created event
   ↓
8. Table appears in Figma canvas
```

### AI Generation Flow

```
1. User enters prompt in UI
   ↓
2. UI sends: { type: 'generate-table-with-ai', prompt, rows, cols }
   ↓
3. Main thread forwards to proxy server
   ↓
4. Proxy server (/generateTable):
   - Calls Watsonx.ai API
   - Generates headers + rows
   - Returns JSON structure
   ↓
5. Main thread receives AI response
   ↓
6. Main thread sends: { type: 'create-table-from-ai', headers, rows }
   ↓
7. Same flow as table creation (step 4 above)
```

### Table Scanning Flow

```
1. User selects table in Figma
   ↓
2. UI sends: { type: 'scan-table' }
   ↓
3. Main thread:
   - Gets selected node
   - Calls TableScanService.scanTable()
   - Extracts components
   - Stores in lastScanResult
   ↓
4. Main thread sends: { type: 'scan-result', result: {...} }
   ↓
5. UI receives scan result
   ↓
6. UI updates component info
   ↓
7. User can now create tables using scanned structure
```

---

## 🔧 Main Workflows

### Workflow 1: Creating a Table from Scan

**Prerequisites:**
- User has scanned a table (has `lastScanResult`)

**Steps:**
1. User configures grid (rows, columns)
2. User edits cell properties (optional)
3. User clicks "Create Table"
4. Main thread creates table structure
5. Applies cell properties
6. Configures smart slots (if enabled)
7. Table appears in Figma

**Key Code Locations:**
- `main.ts`: `create-table-from-scan` handler (line ~3157)
- `ui.ts`: `saveCellProperties()` (line ~2830)

---

### Workflow 2: AI Table Generation

**Steps:**
1. User enters prompt (e.g., "Create a user management table")
2. User sets rows/columns
3. User clicks "Generate with AI"
4. UI sends prompt to main thread
5. Main thread calls proxy server `/generateTable`
6. Proxy server calls Watsonx.ai
7. Returns headers + rows
8. Main thread creates table from AI data
9. Table appears in Figma

**Key Code Locations:**
- `ui.ts`: Generate button handler (line ~401)
- `main.ts`: `generate-table-with-ai` handler (line ~1847)
- `proxy-server.js`: `/generateTable` endpoint (line ~170)

---

### Workflow 3: Smart Slot Application

**What are Smart Slots?**
Automatically suggests appropriate components (tags, avatars, status icons) based on column content.

**Steps:**
1. User creates table with data
2. System analyzes column content
3. Suggests slot types (tag, status, avatar, etc.)
4. User accepts suggestions (or manual selection)
5. System applies components to cells

**Key Code Locations:**
- `main.ts`: `analyzeTableDataForSmartSlots()` (line ~1364)
- `main.ts`: `analyzeColumnContent()` (line ~1065)
- `main.ts`: Slot configuration logic (line ~5804)

---

## ✨ Key Features

### 1. **Smart Component Detection**
- Scans existing Figma tables
- Extracts component structure
- Reuses components for new tables

**Files:**
- `src/services/TableScanService.ts`
- `src/main.ts` (scan logic)

---

### 2. **AI Content Generation**
- Single prompt: Generate full table
- Per-column: Generate column-specific data
- Uses IBM Watsonx.ai

**Files:**
- `proxy-server.js` (`/generateTable`, `/generate`)
- `src/handlers/aiHandlers.ts`
- `src/main.ts` (AI table creation)

---

### 3. **File Import**
- CSV parsing
- JSON parsing
- Excel parsing (via SheetJS)

**Files:**
- `src/ui.ts` (file upload handlers)

---

### 4. **Smart Slots**
Automatically suggests components based on content:
- **Tags**: For comma-separated values, categories
- **Status Icons**: For status values (Failed, Succeeded, etc.)
- **Avatars**: For user names
- **Links**: For URLs, emails
- **Checkboxes**: For boolean values

**Files:**
- `src/main.ts` (`analyzeColumnContent`, `suggestSlotComponent`)

---

### 5. **Property Mapping**
Maps UI property names to Figma component properties.

**Files:**
- `src/utils/PropertyMapper.ts`
- `src/main.ts` (`mapPropertyNames`)

---

### 6. **Analytics Tracking**
Tracks user behavior and sends to GA4:
- Plugin usage
- Table generation
- AI usage
- User location (server-inferred)
- Repeat usage

**Files:**
- `src/utils/analytics.ts`
- `proxy-server.js` (`/analytics` endpoint)

---

## 🌐 Server Architecture

### Proxy Server (`proxy-server.js`)

**Purpose:** 
- Bypasses CORS restrictions
- Handles IBM Watsonx.ai API calls
- Forwards analytics to GA4
- Manages IAM tokens

**Endpoints:**

1. **POST /token**
   - Gets IAM access token from IBM Cloud
   - Used for Watsonx authentication

2. **POST /generate**
   - Generates text values using Watsonx
   - Used for per-column AI generation
   - Returns array of strings

3. **POST /generateTable**
   - Generates full table (headers + rows)
   - Uses structured prompt
   - Returns JSON: `{ headers: [], rows: [] }`

4. **POST /analytics**
   - Receives events from plugin
   - Forwards to Google Analytics 4
   - Includes user tracking and geographic data

**Deployment:**
- Hosted on IBM Code Engine
- Uses Docker container
- Environment variables: `GA4_MEASUREMENT_ID`, `GA4_API_SECRET`, `WATSON_API_KEY`

---

## 📊 Analytics System

### How It Works

```
Plugin → analytics.ts → /analytics endpoint → GA4
```

### Events Tracked

1. **plugin_open**: Plugin initialized
2. **session_start**: Session begins (with user type, session count)
3. **session_end**: Session ends (with duration)
4. **table_created**: Table generated (with rows, columns, source)
5. **ai_generate_single**: Single prompt AI generation
6. **ai_generate_watsonx**: Per-column AI generation
7. **slot_applied**: Component slot applied
8. **scan_table**: Table scanned

### Key Properties

- `user_id`: Unique user identifier
- `user_type`: 'new' or 'returning'
- `session_count`: Number of sessions
- `prompt_text`: AI prompts used
- `rows`, `columns`: Table dimensions
- `source`: 'scan', 'ai', etc.

**Files:**
- `src/utils/analytics.ts`
- `proxy-server.js` (analytics endpoint)

---

## 🔍 Important Code Patterns

### 1. **Message Passing Pattern**
```typescript
// UI → Main
parent.postMessage({ 
  pluginMessage: { 
    type: 'create-table-from-scan', 
    data: {...} 
  } 
}, '*');

// Main → UI
figma.ui.postMessage({ 
  type: 'table-created', 
  success: true 
});
```

### 2. **State Management Pattern**
```typescript
// Get cell state
const cellState = getCellState('1,1');
cellState.properties = { ... };
state.cellProperties.set('1,1', cellState);
```

### 3. **Component Property Mapping**
```typescript
const validProps = mapPropertyNames(
  propertiesToApply, 
  cell.componentProperties
);
cell.setProperties(validProps);
```

### 4. **Slot Configuration Pattern**
```typescript
// Find slot component
const slotComponents = cellRef.findAll(node => {
  return node.type === 'INSTANCE' && 
         node.mainComponent.id === swapComponentId;
});

// Configure slot
slotComponent.setProperties(slotProps);
```

---

## 🎓 Key Concepts to Understand

### 1. **Figma Plugin Architecture**
- **Main Thread**: Access to Figma API, runs in sandbox
- **UI Thread**: HTML/CSS/JS, runs in iframe
- **Communication**: PostMessage API

### 2. **Component System**
- Figma components are reusable elements
- Properties control component variants
- Instances are copies of components

### 3. **Smart Slots**
- Analyzes column content
- Suggests appropriate components
- Auto-applies or manual selection

### 4. **State Management**
- Centralized state in `ui.ts`
- Cell properties stored in Map
- Synchronized between UI and visual grid

### 5. **Analytics Flow**
- Client-side tracking (analytics.ts)
- Server-side forwarding (proxy-server.js)
- GA4 Measurement Protocol

---

## 📝 Development Tips

### Adding a New Feature

1. **UI Changes**: Edit `ui.ts`
2. **Main Thread Changes**: Edit `main.ts`
3. **Server Changes**: Edit `proxy-server.js`
4. **Message Routing**: Update `MessageRouter.ts`
5. **Types**: Update `types/index.ts`

### Testing

1. **Local Testing**: `npm run watch` (development mode)
2. **Build**: `npm run build`
3. **Server**: `node proxy-server.js` (local)
4. **Deploy**: Push to Code Engine

### Debugging

- **UI Console**: Browser DevTools (Figma plugin UI)
- **Main Console**: Figma → Plugins → Development → Open Console
- **Server Logs**: Code Engine → Instances → View logs

---

## 🚨 Important Notes

1. **Code Size**: 
   - `main.ts` is ~6900 lines (consider splitting)
   - `ui.ts` is ~6000 lines (consider splitting)

2. **Performance**:
   - Large tables can be slow
   - Component scanning is async
   - Property mapping is optimized

3. **Error Handling**:
   - Try-catch blocks in critical paths
   - User-friendly error messages
   - Console logging for debugging

4. **Analytics**:
   - Privacy-compliant (anonymous IDs)
   - Tracks user behavior
   - Helps understand usage patterns

---

## 📚 Additional Resources

- **Figma Plugin API**: https://www.figma.com/plugin-docs/
- **IBM Watsonx**: https://www.ibm.com/products/watsonx-ai
- **GA4 Measurement Protocol**: https://developers.google.com/analytics/devguides/collection/protocol/ga4

---

## 🎯 Quick Reference

**Main Files:**
- `src/main.ts` - Figma API operations
- `src/ui.ts` - User interface
- `proxy-server.js` - Backend server

**Key Handlers:**
- `MessageRouter.ts` - Message routing
- `TableScanService.ts` - Table scanning
- `ComponentHandler.ts` - Component operations

**Utilities:**
- `analytics.ts` - Analytics tracking
- `PropertyMapper.ts` - Property mapping
- `helpers.ts` - General utilities

---

**End of Code Walkthrough Guide**

