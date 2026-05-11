# Report Automation Tool — Stakeholder Overview

## What This Tool Does

The Report Automation Tool is an AI-powered system that automatically transforms project documents into professional weekly status reports. It takes raw project data (from SOW documents or task spreadsheets) and generates two outputs:

1. **Google Sheets Task Tracker** — A structured spreadsheet with all tasks, dates, and status
2. **Google Slides Presentation** — A polished presentation with executive summary, activity lists, and a Gantt chart timeline

---

## How It Works (Simple Flow)

```
Input (PDF/DOCX/Excel) 
        ↓
  Agent 1: Extract Tasks & Metadata
        ↓
  Agent 2: Generate Task Tracker (Sheets)
        ↓
  Agent 3: Generate Slides with Gantt Chart
        ↓
Output: Google Sheets + Google Slides
```

---

## The Two Agents

### Agent 1: SOW Report Agent (Phase 1)

**Purpose:** Extracts project information from Scope of Work (SOW) documents.

**Input:** PDF, DOCX, or Excel file containing project scope

**Output:** Google Sheets task tracker with:
- Project metadata (name, client, dates)
- Task list (phase, module, task description, assigned to, dates, status)
- Color-coded status (Green=Completed, Orange=In Progress, Red=To Do, Gray=Delayed)
- Dropdown menus for easy status updates

**Key Features:**
- Reads multiple file formats (PDF, DOCX, Excel)
- AI extracts tasks and organizes them by phase
- Creates professional spreadsheet with formatting
- Automatically publishes to Google Drive

---

### Agent 2: Report Generator Agent (Phase 2)

**Purpose:** Transforms the task tracker into a presentation.

**Input:** Google Sheets task tracker (from Agent 1) or direct Excel file

**Output:** Google Slides presentation with:
- Executive Summary (AI-generated 60-word narrative)
- Key Activities Completed (bullet points with category headers)
- Key Activities In Progress
- Key Activities Upcoming
- Gantt Chart/Timeline visualization
- Project status summary (counts of completed/in-progress/upcoming tasks)

**Key Features:**
- Auto-generates Gantt chart from timeline data
- Smart formatting (bold category headers, consistent fonts)
- Embeds timeline image into slides
- Professional template-based output

---

## Key Capabilities

| Capability | Description |
|------------|-------------|
| **AI Extraction** | Uses Gemini AI to understand document structure and extract tasks |
| **Gantt Charts** | Automatically generates visual timeline from task dates |
| **Multi-Format Support** | Handles PDF, DOCX, Excel, TXT inputs |
| **Google Integration** | Publishes directly to Google Sheets and Slides |
| **Smart Formatting** | Applies consistent styling, bold headers, proper fonts |
| **Status Tracking** | Color-coded task status with dropdown selection |

---

## Technical Details (For Technical Team)

- **Framework:** Google Agent Development Kit (ADK)
- **AI Model:** Gemini 2.5 Flash
- **APIs Used:** Google Drive, Google Sheets, Google Slides
- **Languages:** Python
- **Chart Rendering:** Matplotlib (for Gantt charts)

---

## Setup Requirements

1. **Google Cloud Project** with Drive, Sheets, and Slides APIs enabled
2. **OAuth2 Credentials** (client secret + token) for user authentication
3. **API Key** for Gemini AI
4. **Template Slides** — A master Google Slides template with placeholders

---

## Example Output

### Google Sheets Output
| Phase | Module | Task | Assigned | Start | End | Status |
|-------|--------|------|----------|-------|-----|--------|
| Phase 1 | Documentation | Create SRS | John | 01/01 | 15/01 | Completed |
| Phase 1 | Development | Build API | Jane | 16/01 | 28/01 | In Progress |
| Phase 2 | Testing | Unit Tests | Mike | 01/02 | 10/02 | To Do |

### Google Slides Output
- Slide 1: Executive Summary (60-word narrative)
- Slide 2: Activities (Completed/In Progress/Upcoming) with bold category headers
- Slide 3: Timeline with Gantt chart image

---

## Configuration

Key environment variables:
- `GOOGLE_API_KEY` — Gemini AI key
- `GOOGLE_OAUTH_CLIENT_SECRET_JSON` — OAuth client secret
- `GOOGLE_OAUTH_TOKEN_JSON` — Cached OAuth token
- `GOOGLE_SLIDES_TEMPLATE_ID` — ID of presentation template

---

## Benefits for Stakeholders

1. **Saves Time** — No manual copying of tasks to spreadsheets or slides
2. **Consistent Format** — Same professional output every week
3. **Real-Time Updates** — Task tracker updates reflect immediately in slides
4. **Visual Clarity** — Gantt charts make project progress easy to understand
5. **AI-Powered** — Reduces human effort in summarizing progress

---

## Support

For technical issues or questions, refer to:
- Main README: `/Users/shubhamjain/Documents/Report Automation Tool/README.md`
- SOW Agent docs: `/Users/shubhamjain/Documents/Report Automation Tool/sow_report_agent/README.md`
- Report Generator docs: `/Users/shubhamjain/Documents/Report Automation Tool/Report_generator/README.md`