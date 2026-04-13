# NEO Quiz Automation Tool

## What This Project Is

A web-based automation tool that populates Microsoft Forms quizzes using Playwright browser automation. It was built to support the **LUSA NEO (New Employee Onboarding)** program, which requires creating quiz forms across multiple SAP process tribes for end-user training assessments.

The tool ships with **89 pre-built quiz JSON files** covering all tribes and topics. Users provide a Microsoft Forms template URL, and the automation duplicates template questions, fills in the real quiz content, marks correct answers, and cleans up the templates — all hands-free.

## Architecture

- **Backend**: Node.js + Express server (`server.js`) on port 3000
- **Frontend**: Vanilla HTML/CSS/JS single-page app in `public/`
- **Automation**: Playwright-based browser automation (`automation.js`) using a `QuizAutomation` class that extends EventEmitter for real-time progress streaming via SSE
- **Quiz Data**: JSON files in `quiz_json/` with a manifest (`quiz_manifest.json`) indexing all forms

## Key Files

- `server.js` — Express server, API routes, SSE streaming, processing orchestration
- `automation.js` — `QuizAutomation` class: browser lifecycle, form navigation, question duplication/filling, cleanup
- `public/index.html` — UI shell with dashboard, quiz editor, process view, bulk progress, setup guide modal
- `public/app.js` — Frontend logic: quiz list, editor, progress tracking, SSE handling, auto-refresh polling
- `public/style.css` — All styles
- `quiz_json/quiz_manifest.json` — Index of all 89 quiz forms with metadata (tribe, filename, question count)
- `quiz_json/quiz_*.json` — Individual quiz data files with questions, options, correct answers
- `processing-log.json` — Runtime tracking of which forms have been processed (gitignored)
- `auth-state.json` — Playwright browser auth session (gitignored, contains tokens)

## Who Uses This Tool

- **NEO Training Coordinators**: Non-technical users who need to create Microsoft Forms quizzes from pre-approved question banks. They use the web UI to review questions, paste form URLs, and kick off automation.
- **Tribe Leads / SMEs**: Subject matter experts across SAP process areas (F2S, O2C, P2P, R2R, etc.) who review and approve quiz content before it goes into Forms.
- **IT/Automation Support**: Technical users who clone the repo, set up the tool locally, and may troubleshoot issues.

All end users interact through the localhost web UI at `http://localhost:3000`. The tool runs locally on each user's machine — there is no shared server deployment.

## SAP Process Tribes

The quiz data covers these SAP implementation tribes:
- **F2S** (Forecast to Stock) — demand planning, inventory, inbound delivery, goods reception
- **O2C** (Order to Cash) — sales orders, delivery, billing, customer management
- **P2P** (Procure to Pay) — purchase requisitions, purchase orders, invoice verification
- **R2R** (Record to Report) — general ledger, cost centers, profit centers, financial closing
- **Finance Master Data** — cost centers, profit centers, GL accounts

## How the Automation Works

1. User selects a quiz and provides a Microsoft Forms editor URL (the form must have "Sample Two Option" and "Sample Four Option" template questions)
2. The tool opens a Chromium browser, navigates to the form, and handles Microsoft sign-in if needed
3. For each quiz question:
   - Clicks the appropriate sample question (2-option for True/False, 4-option for multiple choice)
   - Duplicates it
   - Replaces the question text and option texts
   - Marks the correct answer
4. After all questions are filled, deletes the original sample template questions
5. Reports progress in real-time via SSE to the web UI

## Development Notes

- The frontend auto-polls quiz data every 10 seconds so status changes appear without manual reload
- OS-aware keyboard shortcuts: Meta (macOS) vs Control (Windows/Linux) for select-all and duplicate
- Auth state persists across sessions via `auth-state.json`
- Screenshots are saved to `screenshots/` for debugging failed runs
- The `processing-log.json` tracks which forms have been processed to prevent accidental re-runs
