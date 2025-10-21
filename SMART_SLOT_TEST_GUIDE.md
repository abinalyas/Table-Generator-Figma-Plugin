# Testing Smart Slot Detection Without Watson AI

## 🎯 What We're Testing

The smart slot detection feature that automatically detects:
- **Status columns** → Suggests Status Icon component
- **User columns** → Suggests Avatar component
- **Action columns** → Suggests Overflow component

## ✅ Method 1: Use Faker (No API Required!)

### Step 1: Create a Table with Faker

1. **Open the plugin** in Figma
2. **Create a 5×5 table** (or any size)
3. **Click on cells** and fill them with Faker data:

#### Column 1 - Project Name:
- Select cell → Faker → Search "company" → Use `company.companyName()`

#### Column 2 - Status (THIS IS THE KEY COLUMN!):
- Cell 1: Manually type "Pending"
- Cell 2: Manually type "In-progress"
- Cell 3: Manually type "Succeeded"
- Cell 4: Manually type "Failed"
- Cell 5: Manually type "Completed"

#### Column 3 - Owner:
- Select cell → Faker → Search "name" → Use `name.findName()`

#### Column 4 - Email:
- Select cell → Faker → Search "email" → Use `internet.email()`

#### Column 5 - Actions:
- Manually type "Actions" in header
- Leave cells empty or type "..."

### Step 2: Test Smart Slot Detection

1. **Click "Analyze Smart Slots"** button
2. **Check the console** - you should see:
   ```
   🔍 Analyzing table data for smart slot suggestions...
   💡 Column "Status" (1): status → suggest statusIcon
      Samples: Pending, In-progress, Succeeded
   💡 Column "Email" (2): user → suggest avatar
      Samples: john@example.com, jane@example.com
   💡 Column "Actions" (4): action → suggest overflow
      Samples: ...
   ✅ Found 3 smart slot suggestions
   ```

## ✅ Method 2: Manual Test Data Entry

### Quick Test Table:

| Project Name | Status | Owner | Email | Priority |
|--------------|--------|-------|-------|----------|
| Project Alpha | Pending | John Doe | john@example.com | High |
| Project Beta | In-progress | Jane Smith | jane@example.com | Medium |
| Project Gamma | Succeeded | Bob Johnson | bob@example.com | Low |
| Project Delta | Failed | Alice Brown | alice@example.com | Critical |
| Project Epsilon | Completed | Charlie Wilson | charlie@example.com | High |

### Expected Smart Slot Suggestions:

1. **Column "Status"**:
   - Content Type: `status`
   - Suggested Component: `statusIcon`
   - Reason: Contains keywords like "Pending", "In-progress", "Succeeded", "Failed"

2. **Column "Email"**:
   - Content Type: `user`
   - Suggested Component: `avatar`
   - Reason: Contains email pattern `xxx@xxx.xxx`

3. **Column "Owner"** (if detected):
   - Content Type: `user`
   - Suggested Component: `avatar`
   - Reason: Contains "First Last" name pattern

## ✅ Method 3: File Upload (CSV/Excel)

1. **Create a CSV file** with this content:
   ```csv
   Project,Status,Owner,Email,Actions
   Alpha,Pending,John Doe,john@example.com,Edit
   Beta,In-progress,Jane Smith,jane@example.com,Delete
   Gamma,Succeeded,Bob Wilson,bob@example.com,View
   Delta,Failed,Alice Brown,alice@example.com,Edit
   Epsilon,Completed,Charlie Davis,charlie@example.com,Archive
   ```

2. **Save as `test-data.csv`**

3. **In plugin**:
   - Check "Populate data from file"
   - Upload the CSV
   - Data fills the table automatically

4. **Click "Analyze Smart Slots"**

## 🔍 What to Look For in Console

When you click "Analyze Smart Slots", check for:

### ✅ Success Indicators:
```
🔍 Analyzing table data for smart slot suggestions...
💡 Column "Status" (1): status → suggest statusIcon
   Samples: Pending, In-progress, Succeeded
💡 Column "Email" (3): user → suggest avatar
   Samples: john@example.com, jane@example.com
✅ Found 2 smart slot suggestions
```

### 📊 UI Notification:
You should see a notification:
```
💡 Found 2 smart slot suggestions! Check console for details.
```

## 🎨 Testing Different Content Types

### Status Detection Test:
- **Try these values**: Active, Inactive, Pending, Approved, Rejected, Completed, In-progress, Cancelled, Draft, Published, Failed, Succeeded, Warning, Error

### User Detection Test:
- **Email pattern**: anything@domain.com
- **Name pattern**: First Last (e.g., "John Smith", "Jane Doe")

### Boolean Detection Test:
- **Values**: true, false, yes, no, 0, 1

### URL Detection Test:
- **Values**: http://example.com, https://google.com

### Action Column Test:
- **Column name must contain**: "action", "actions", "options", "menu"

## 🐛 Troubleshooting

**No suggestions found:**
- Make sure at least 50% of cells in a column match the pattern
- Check that status values are spelled correctly
- Try all lowercase (the detection is case-insensitive)

**Console shows no output:**
- Check that you have data in the table
- Make sure you clicked "Analyze Smart Slots" AFTER filling data
- Verify the plugin is using the latest build: `npm run build`

## 💡 Next Steps After Testing

Once smart slot detection is working, the next phase would be:
1. **Auto-apply suggestions** - Automatically enable slot property
2. **Import suggested components** - Use the hardcoded component keys
3. **Set component properties** - Match status values to component variants
4. **Show UI suggestions** - Display suggestions in a panel with accept/reject options

The detection is working now - we just need to wire up the auto-apply logic!


