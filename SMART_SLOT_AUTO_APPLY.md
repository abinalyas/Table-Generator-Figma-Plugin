# Smart Slot Auto-Apply Feature

## ✨ Overview

The plugin now **automatically detects and applies smart component suggestions** when generating tables with AI, Faker, or file upload data.

---

## 🎯 What It Does

When you generate a table, the plugin will:

1. **Analyze the data** in each column
2. **Detect content types**:
   - 📊 **Status** values (Pending, In-progress, Succeeded, Failed, etc.)
   - 👤 **User** data (emails, names)
   - ⚡ **Actions** (last column with "action" in name)
   - ✅ **Booleans** (true/false, yes/no)
   - 🔗 **URLs** (http/https links)
   - 📅 **Dates** (various formats)
   - 🔢 **Numbers**
3. **Automatically enable slots** for detected columns
4. **Suggest Carbon components**:
   - `statusIcon` for status columns
   - `avatar` for user columns
   - `overflow` for action columns
   - `checkbox` for boolean columns
   - `link` for URL columns

---

## 🚀 How To Use

### **Automatic Mode** (Default)

1. **Generate table data** using:
   - AI generation (Watsonx/OpenAI)
   - Faker data generation
   - File upload (CSV/JSON)

2. **Wait for analysis** (happens automatically after 500ms)

3. **Check notification**: 
   ```
   ✨ Auto-applied 2 smart components
   ```

4. **Check console** for details:
   ```
   🤖 Auto-applying 2 smart slot suggestions...
   🎯 Applying statusIcon to column 1 (Status)
     ✅ Enabled slot for cell 1,1
     ✅ Enabled slot for cell 2,1
     ✅ Enabled slot for cell 3,1
   ```

5. **Create table** - slots are already enabled and ready for component swapping!

### **Manual Mode**

1. **Generate or load data** into the grid
2. **Click "Analyze Smart Slots"** button
3. **Review suggestions** in console
4. **Manually enable slots** if needed

---

## 📊 Detection Logic

### **Status Detection**
Carbon Design System values:
- `Failed`, `Succeeded`, `Pending`, `In-progress`
- `Not started`, `Incomplete`, `Unknown`
- `Normal`, `Informative`

Plus common keywords:
- Active, Inactive, Approved, Rejected
- Completed, Cancelled, Draft, Published
- Archived, Enabled, Disabled
- Success, Error, Warning

### **User Detection**
- Email pattern: `user@domain.com`
- Name pattern: `First Last` (capitalized)

### **Action Detection**
- Column name contains: `action`, `option`, `menu`
- Typically the last column in a table

### **Other Types**
- **Boolean**: true, false, yes, no, 0, 1
- **URL**: starts with `http://` or `https://`
- **Date**: `YYYY-MM-DD`, `MM/DD/YYYY`, etc.
- **Number**: numeric values with optional decimals

---

## 🔧 Technical Details

### **Files Modified**

1. **`src/ui.ts`**:
   - `analyzeSmartSlots(autoApply: boolean)` - now accepts auto-apply flag
   - `updateMainGridWithData()` - triggers auto-analysis after data load
   - `case 'auto-apply-smart-slots'` - new message handler to apply suggestions

2. **`src/main.ts`**:
   - `analyze-smart-slots` handler - checks `autoApply` flag
   - Sends `auto-apply-smart-slots` message instead of `smart-slot-suggestions`

### **Message Flow**

```
┌─────────────────────────────────────────────────────┐
│ 1. User generates table data (AI/Faker/File)       │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│ 2. UI: updateMainGridWithData()                     │
│    → Populates grid with data                       │
│    → setTimeout(() => analyzeSmartSlots(true), 500) │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│ 3. UI → Backend: analyze-smart-slots message        │
│    { gridData, headers, autoApply: true }           │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│ 4. Backend: analyzeTableDataForSmartSlots()         │
│    → Detects content types                          │
│    → Generates suggestions                          │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│ 5. Backend → UI: auto-apply-smart-slots message     │
│    { suggestions: [...] }                           │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│ 6. UI: Apply suggestions                            │
│    → Enable slot property for detected columns      │
│    → Update cell visuals                            │
│    → Show notification                              │
└─────────────────────────────────────────────────────┘
```

---

## 🧪 Testing

### **Test Case 1: Status Column**
```
Column 1: Pending, In-progress, Succeeded, Failed
Expected: statusIcon suggested & auto-applied
```

### **Test Case 2: Email Column**
```
Column 2: john@example.com, jane@example.com
Expected: avatar suggested & auto-applied
```

### **Test Case 3: Action Column**
```
Column name: "Actions" or last column
Expected: overflow suggested & auto-applied
```

### **Test Case 4: Mixed Data**
```
Column 1: Status values → statusIcon
Column 2: Names/Emails → avatar  
Column 3: Regular text → no suggestion
Expected: Only columns 1 & 2 have slots enabled
```

---

## 💡 Future Enhancements

1. **Visual Indicators**: Show badges/icons in grid for detected columns
2. **Component Swapping**: Automatically import and swap components (not just enable slots)
3. **Confidence Threshold**: Only auto-apply suggestions above 80% confidence
4. **User Preferences**: Toggle auto-apply on/off in settings
5. **Custom Rules**: Allow users to define their own detection patterns
6. **Undo**: Quick undo for auto-applied suggestions

---

## ⚠️ Known Limitations

1. **Name Detection**: Currently only detects "First Last" pattern (capitalized)
2. **Manual Component Swap**: Slots are enabled, but you still need to manually swap components
3. **500ms Delay**: Fixed delay may not be optimal for all scenarios
4. **No Undo**: Cannot undo auto-applied suggestions (use reset button)

---

## 📝 Example Console Output

```
🤖 [UI] Auto-analyzing smart slots after data load...
🧠 [UI] Analyzing table data for smart slots...
🧠 [UI] Header 1: "Status" (key: header-1)
🧠 [UI] Header 2: "Owner" (key: header-2)
🧠 [UI] Header 3: "Email" (key: header-3)
🧠 [UI] Row 1 data: ['Pending', 'Jensen', 'test@email.com']
🧠 [UI] Row 2 data: ['In-progress', 'Junior', 'another@email.com']
📊 [UI] Analyzing 3 columns x 2 rows
✅ [UI] Message sent to backend

🧠 Analyzing table data for smart slot suggestions...
🔍 Analyzing table data for smart slot suggestions...
💡 Column "Status" (0): status → suggest statusIcon
   Samples: Pending, In-progress
💡 Column "Email" (2): user → suggest avatar
   Samples: test@email.com, another@email.com
✅ Found 2 smart slot suggestions

🤖 Auto-applying 2 smart slot suggestions...
🤖 [UI] Auto-applying smart slot suggestions: [...]
🎯 [UI] Applying statusIcon to column 1 (Status)
  ✅ Enabled slot for cell 1,1
  ✅ Enabled slot for cell 2,1
🎯 [UI] Applying avatar to column 3 (Email)
  ✅ Enabled slot for cell 1,3
  ✅ Enabled slot for cell 2,3
✨ [UI] Auto-applied 2 smart slot suggestion(s)

Notification: ✨ Auto-applied 2 smart components
```

---

## 🎉 Success Criteria

✅ Detects status values (Carbon + common keywords)  
✅ Detects user data (emails, names)  
✅ Auto-enables slots for detected columns  
✅ Shows notification with count  
✅ Logs detailed info to console  
✅ Works with AI/Faker/File data  
✅ Manual mode still works (button)  

---

**Created**: October 12, 2025  
**Version**: 1.0.0  
**Feature Status**: ✅ Complete and Working

