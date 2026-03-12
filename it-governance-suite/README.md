# IT Governance Suite

This folder isolates the Google Workspace governance work from the Wasp/Katana sync project.

## Components

- `security-tracker-scr/`: Apps Script project for the `2026_Security & GMP Connection Tracker` sheet
- `systems-health-scr/`: Apps Script project for the `2026_Systems-Health-Tracker` sheet
- `sop-correction/`: SOP extraction, design analysis, and rewrite workspace
- `sops/working-docx/`: local `.docx` copies of the current SOP drafts

## Google Sheets

- Security tracker sheet ID: `1WBzed_RMtPC3kii-ybJ3lE4Fo1pAqWBt9OJyCSVR-mQ`
- Systems health tracker sheet ID: `1jnWtdBPzR7DreihCHQASiN7splmRS7HYJTR77GpfI5w`

## Slack Pieces

- `security-tracker-scr/slack/slack-notifications.js`
- `security-tracker-scr/slack/slash-command.js`
- `systems-health-scr/slack/it-alert.js`

## SOP Format Baseline

The current SOP structure is defined in `sop-correction/scripts/rewrite.js`.

- Template order: Title, Metadata, 1. Purpose, 2. Scope, 3. Responsibilities, 4. Definitions, 5. Procedure, 6. References, 7. Attachments, 8. Revision History
- Style baseline: Arial throughout, 26pt centered title, 14pt section headings, 11pt body text, right-aligned metadata, bordered tables, standard blue links
- Current rewrite targets: `IT-014 Automation Maintenance` and `IT-015 IT Security Tracker Maintenance`
- Visual examples used as reference: `LOG-006` and `QA-012`

Use this folder as the home for future tracker, Slack, and SOP maintenance work.
