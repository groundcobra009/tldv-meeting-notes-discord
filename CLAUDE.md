# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script (GAS) project that automatically processes tldv meeting minutes from Gmail and logs them into a Google Spreadsheet. The project name "tldvの議事録をまとめる" means "Summarizing tldv Meeting Minutes" in Japanese.

## Development Environment

This is a Google Apps Script project that runs in Google's cloud environment. There is no local build process or npm commands. All development and testing must be done through the Google Apps Script editor.

### Key Configuration
- Runtime: V8 (modern JavaScript)
- Timezone: Asia/Tokyo (JST)
- Exception logging: STACKDRIVER
- Main code file: `コード.gs`

## Architecture

The system follows a simple event-driven architecture:

1. **Email Processing Pipeline**:
   - Gmail → Filter by "tldv" label → Extract data → Write to Spreadsheet → Move to "処理済み" label

2. **Trigger System**:
   - Hourly automatic execution via time-based trigger
   - Manual execution through custom spreadsheet menu

3. **Data Storage**:
   - Configuration (email address) stored in Script Properties
   - Processed data written to active spreadsheet (Column A: date, B: subject, C: body)

## Core Functionality

The main processing logic in `processEmails()` performs:
1. Retrieves emails with "tldv" label
2. Extracts and cleans email content (removes promotional text and URLs)
3. Writes to spreadsheet with proper formatting
4. Moves processed emails to "処理済み" (processed) label

## Testing and Deployment

Since this is a GAS project:
- No unit tests or test commands
- Testing is done manually through the spreadsheet menu options
- Deployment happens automatically when saving in the GAS editor
- Use the spreadsheet's custom menu (スクリプト → メール処理を実行) to test functionality

## Important Implementation Details

- Email body cleaning removes specific patterns:
  - Text starting with "👋 Hi,"
  - Text starting with "P.S."
  - All URLs (http/https links)
- Row height is set to 24 pixels for consistency
- Uses Google's PropertiesService for secure configuration storage
- All user-facing text is in Japanese

## Common Tasks

To modify email processing logic:
1. Edit the `processEmails()` function in `コード.gs`
2. Test using the spreadsheet menu: スクリプト → メール処理を実行

To change email filtering criteria:
1. Modify the label name in `GmailApp.getUserLabelByName('tldv')`
2. Update the processed label in `GmailApp.getUserLabelByName('処理済み')`

To adjust spreadsheet formatting:
1. Modify column assignments in the `processEmails()` function
2. Change row height in `setRowHeight()` function