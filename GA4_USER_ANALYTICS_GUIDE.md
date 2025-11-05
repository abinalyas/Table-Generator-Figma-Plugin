# GA4 User Analytics Guide

## What Data You Can Now Track

### ✅ **User Location (Geographic)**
- **Automatic:** GA4 automatically infers geographic location from the server IP address (Code Engine proxy)
- **Where to see it:**
  - GA4 → Reports → Realtime → Geographic data (map view)
  - GA4 → Reports → User → Demographics → Geographic
  - Shows: Country, Region, City based on server requests

### ✅ **Unique Users**
- **How it works:** Each user gets a persistent anonymous ID stored in localStorage
- **Where to see it:**
  - GA4 → Reports → User → Overview
  - Look for "Total users" and "Active users"
  - Each `anonId` = 1 unique user
  - `user_id` property in events also tracks unique users

### ✅ **Repeat Usage (Returning Users)**
- **Events tracked:**
  - `session_start` - Includes:
    - `is_first_visit` (boolean) - New vs returning
    - `user_type` ('new' or 'returning')
    - `session_count` (number) - How many times user opened plugin
    - `days_since_last_visit` (number) - Days since last use
- **Where to see it:**
  - GA4 → Reports → User → User lifecycle
  - GA4 → Reports → User → Demographics → User attributes
  - Filter by `user_type` parameter to see new vs returning users
  - Look at `session_start` events and filter by `is_first_visit`

### ✅ **AI Prompts Used**
- **Events tracked:**
  - `ai_generate_single` - Full prompt text in:
    - `prompt` (full text)
    - `prompt_text` (same, for better filtering)
    - `prompt_length` (character count)
  - `ai_generate_watsonx` - Column-level prompts with:
    - `prompt` (full text)
    - `prompt_text` (same, for better filtering)
    - `prompt_length` (character count)
    - `column` (column index if column mode)
    - `mode` ('cell', 'row', 'column')
- **Where to see it:**
  - GA4 → Explore → Free form
  - Create a report with:
    - Dimensions: `Event name`, `prompt_text`
    - Metrics: `Event count`
  - Or use GA4 → Reports → Engagement → Events
  - Click on `ai_generate_single` or `ai_generate_watsonx` events
  - View "Event parameters" → `prompt_text` to see all prompts

## Creating Custom Reports in GA4

### 1. **User Retention Report**
1. GA4 → Explore → Free form
2. Dimensions:
   - `Event name` = `session_start`
   - `user_type`
3. Metrics:
   - `Active users`
   - `Event count`
4. Filter: `user_type` = `returning` to see repeat users

### 2. **Prompt Usage Report**
1. GA4 → Explore → Free form
2. Dimensions:
   - `Event name` (filter to `ai_generate_single` or `ai_generate_watsonx`)
   - `prompt_text`
3. Metrics:
   - `Event count`
   - `Total users`
4. Sort by `Event count` to see most popular prompts

### 3. **Geographic User Distribution**
1. GA4 → Reports → User → Demographics → Geographic
2. View by:
   - Country
   - Region
   - City
3. See which locations have most users

### 4. **Session Duration & Engagement**
1. GA4 → Explore → Free form
2. Dimensions:
   - `Event name` = `session_end`
   - `user_type`
3. Metrics:
   - `session_duration_seconds` (average)
   - `Total users`
4. See how long users stay in the plugin

## Events Now Being Tracked

### Core Events
- ✅ `plugin_open` - Plugin initialized
- ✅ `session_start` - Session begins (with user type, session count)
- ✅ `session_end` - Session ends (with duration)

### AI Generation Events
- ✅ `ai_generate_single` - Single prompt AI generation
  - Properties: `prompt`, `prompt_text`, `prompt_length`, `ai_source`
- ✅ `ai_generate_watsonx` - Watsonx column generation
  - Properties: `prompt`, `prompt_text`, `prompt_length`, `mode`, `column`, `count`

### All Events Include
- `user_id` - Unique user identifier
- `session_duration_seconds` - Time since session start
- `user_type` - 'new' or 'returning'

## Quick Queries in GA4

### "How many unique users used the plugin?"
- GA4 → Reports → User → Overview → Total users

### "How many users are returning?"
- GA4 → Explore → Free form
- Filter: `session_start` event, `user_type` = `returning`
- Metric: `Total users`

### "What prompts are users entering most?"
- GA4 → Explore → Free form
- Event: `ai_generate_single` or `ai_generate_watsonx`
- Dimension: `prompt_text`
- Metric: `Event count`
- Sort: Descending

### "Where are users located?"
- GA4 → Reports → User → Demographics → Geographic
- View by Country/Region/City

### "How often do users reopen the plugin?"
- GA4 → Explore → Free form
- Event: `session_start`
- Dimension: `session_count`
- Metric: `Event count`
- See distribution of session counts

## Notes

- **Geographic data:** GA4 infers location from server IP (Code Engine proxy location may not reflect actual user location perfectly)
- **User privacy:** All tracking uses anonymous IDs, no PII collected
- **Session tracking:** Sessions are 30-minute windows (inactive timeout)
- **Prompt data:** Full prompt text is stored (first 1000 chars), visible in GA4 event parameters

