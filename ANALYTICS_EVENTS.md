# Analytics Events - Additional Tracking Opportunities

## Currently Tracked Events ✅
- `plugin_open` - Plugin initialized with version
- `ai_generate_single` - Single prompt AI generation with prompt text and length
- `ai_generate_watsonx` - Watsonx column generation with mode, column, prompt, length

---

## Recommended Additional Events

### 1. Table Creation & Structure
**Event: `table_created`**
- **Properties:**
  - `rows` (number) - Number of rows in table
  - `columns` (number) - Number of columns in table
  - `source` (string) - How table was created: 'ai_generated', 'csv_import', 'json_import', 'manual', 'scan'
  - `theme` (string) - Carbon theme used: 'light', 'dark', 'white', 'g10', etc.
  - `time_taken_ms` (number) - Time to generate table (for AI-generated)
  - `has_headers` (boolean) - Whether table has header row

**Event: `table_updated`**
- **Properties:**
  - `rows` (number)
  - `columns` (number)
  - `change_type` (string) - 'rows_added', 'rows_removed', 'columns_added', 'columns_removed', 'data_modified'
  - `change_count` (number) - How many rows/columns changed

### 2. Table Scanning
**Event: `scan_table`**
- **Properties:**
  - `rows` (number) - Rows detected
  - `columns` (number) - Columns detected
  - `detected_content_types` (array) - ['text', 'number', 'label', 'status', 'user', 'link', etc.]
  - `smart_slots_applied` (boolean) - Whether smart slots were auto-applied
  - `scan_duration_ms` (number) - Time to scan

### 3. Slot/Component Application
**Event: `slot_applied`**
- **Properties:**
  - `slot_type` (string) - 'tag', 'avatar', 'link', 'status', 'checkbox', 'slotGroup', 'overflow', 'edit', 'delete'
  - `column_index` (number) - Column where slot was applied
  - `row_count` (number) - Number of rows affected
  - `apply_mode` (string) - 'cell', 'row', 'column'
  - `component_id` (string) - Figma component ID if applicable
  - `is_auto_applied` (boolean) - Whether smart slot auto-applied it

**Event: `slot_configured`**
- **Properties:**
  - `slot_type` (string)
  - `column_index` (number)
  - `config_change` (string) - 'color_changed', 'separator_changed', 'text_updated', etc.

### 4. Feature Usage - Specific Components
**Event: `tags_configured`**
- **Properties:**
  - `column_index` (number)
  - `tag_count` (number) - Number of tags in cell/column
  - `separator` (string) - Separator used: ',', '|', '/', '&', etc.
  - `color_count` (number) - Number of distinct colors used

**Event: `link_configured`**
- **Properties:**
  - `column_index` (number)
  - `link_type` (string) - 'url', 'email', 'phone'
  - `link_count` (number) - Number of links in column

**Event: `status_configured`**
- **Properties:**
  - `column_index` (number)
  - `status_type` (string) - 'Failed', 'Succeeded', 'In-progress', etc.
  - `status_count` (number) - Number of distinct statuses

**Event: `checkbox_configured`**
- **Properties:**
  - `column_index` (number)
  - `checked_count` (number) - Number of checked boxes
  - `total_count` (number) - Total checkboxes

### 5. Data Import/Export
**Event: `data_imported`**
- **Properties:**
  - `source_type` (string) - 'csv', 'json', 'file_upload'
  - `rows` (number)
  - `columns` (number)
  - `file_size_kb` (number)
  - `import_success` (boolean)

**Event: `table_exported`**
- **Properties:**
  - `export_format` (string) - 'csv', 'json', 'figma_nodes'
  - `rows` (number)
  - `columns` (number)

### 6. User Interactions & Engagement
**Event: `settings_changed`**
- **Properties:**
  - `setting_name` (string) - 'theme', 'column_width', 'show_text', 'slot_enabled', etc.
  - `old_value` (string/number) - Previous value
  - `new_value` (string/number) - New value

**Event: `plugin_closed`**
- **Properties:**
  - `session_duration_seconds` (number) - How long plugin was open
  - `tables_created` (number) - Tables created in this session
  - `events_triggered` (number) - Total events in session

**Event: `ui_interaction`**
- **Properties:**
  - `element_id` (string) - 'generate_button', 'settings_tab', 'add_row_button', etc.
  - `action_type` (string) - 'click', 'toggle', 'input', 'select'
  - `context` (string) - Additional context about the interaction

### 7. Performance & Errors
**Event: `api_call_duration`**
- **Properties:**
  - `api_endpoint` (string) - '/generate', '/generateTable', '/analytics'
  - `duration_ms` (number) - Response time
  - `success` (boolean) - Whether call succeeded
  - `status_code` (number) - HTTP status code

**Event: `error_raised`**
- **Properties:**
  - `error_type` (string) - 'setProperties', 'import_failed', 'api_error', etc.
  - `error_message` (string) - Sanitized error message (first 200 chars)
  - `surface` (string) - 'ui', 'main', 'server'
  - `context` (string) - Where error occurred: 'table_creation', 'slot_application', etc.

**Event: `table_generation_time`**
- **Properties:**
  - `source` (string) - 'ai', 'import', 'scan'
  - `duration_ms` (number) - Total generation time
  - `rows` (number)
  - `columns` (number)

### 8. Smart Features
**Event: `smart_slots_auto_applied`**
- **Properties:**
  - `columns_analyzed` (number) - Number of columns analyzed
  - `slots_applied` (number) - Number of slots auto-applied
  - `slot_types` (array) - ['tag', 'status', 'avatar', etc.]
  - `accuracy` (number) - User acceptance rate (if tracked)

**Event: `smart_slot_suggestion`**
- **Properties:**
  - `column_index` (number)
  - `suggested_type` (string) - Suggested slot type
  - `confidence` (number) - Confidence score (0-1)
  - `user_accepted` (boolean) - Whether user accepted suggestion

### 9. AI/ML Specific
**Event: `ai_generate_faker`**
- **Properties:**
  - `faker_method` (string) - Faker method used
  - `mode` (string) - 'cell', 'row', 'column'
  - `count` (number) - Number of values generated
  - `column_index` (number) - If column mode

**Event: `ai_prompt_optimized`**
- **Properties:**
  - `original_length` (number)
  - `optimized_length` (number)
  - `prompt_type` (string) - 'table', 'column', 'single'

### 10. Plugin Environment (if accessible)
**Event: `plugin_environment`**
- **Properties:**
  - `figma_version` (string)
  - `plugin_version` (string)
  - `os` (string) - If accessible
  - `browser` (string) - If accessible

---

## Implementation Priority

### High Priority (Core Features)
1. ✅ `plugin_open` - Already implemented
2. ✅ `ai_generate_single` - Already implemented
3. ✅ `ai_generate_watsonx` - Already implemented
4. 🔲 `table_created` - Track table generation
5. 🔲 `slot_applied` - Track slot usage
6. 🔲 `scan_table` - Track scanning usage
7. 🔲 `error_raised` - Track errors for debugging

### Medium Priority (Feature Usage)
8. 🔲 `tags_configured` - Track tag feature usage
9. 🔲 `link_configured` - Track link feature usage
10. 🔲 `status_configured` - Track status feature usage
11. 🔲 `data_imported` - Track import usage
12. 🔲 `settings_changed` - Track settings usage

### Low Priority (Nice to Have)
13. 🔲 `plugin_closed` - Track session duration
14. 🔲 `api_call_duration` - Track performance
15. 🔲 `smart_slots_auto_applied` - Track smart feature usage

---

## Notes
- All events should include `anonId` (already handled by analytics client)
- All timestamps should use `ts` (already handled)
- Sanitize sensitive data (passwords, tokens, etc.)
- Keep event names lowercase with underscores
- Keep property names concise but descriptive

