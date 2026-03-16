# UiPath RPA Project Analyzer

A fully offline, command-line utility that scans **UiPath Studio RPA projects** — including all `.xaml` workflow files and `project.json` — and generates a detailed report covering:

- Every **activity** used, by category and file
- Every **application** automated (Windows desktop & web)
- All **absolute URLs** and **browser page titles** targeted
- All **Windows process names** and **window titles** targeted
- A **Reusability Matrix** scoring each workflow and activity category
- **Shared library candidate** recommendations

Output formats: **Excel (multi-sheet)**, **JSON**, **CSV**, and **console**.

---

## Table of Contents

- [Quick Start](#quick-start)
- [Prerequisites](#prerequisites)
  - [Offline / Airgapped Installation](#offline--airgapped-installation)
- [Usage](#usage)
  - [Single Project](#single-project)
  - [Full Repository Scan](#full-repository-scan)
  - [All Output Flags](#all-output-flags)
- [What Gets Extracted](#what-gets-extracted)
  - [Activities](#activities)
  - [Web Browser Targets](#web-browser-targets)
  - [Windows Desktop Targets](#windows-desktop-targets)
- [Reusability Matrix](#reusability-matrix)
  - [How Scoring Works](#how-scoring-works)
  - [Score Labels](#score-labels)
  - [Shared Library Candidates](#shared-library-candidates)
- [Output Files](#output-files)
- [Airgapped & Data Privacy](#airgapped--data-privacy)
- [Project Structure Supported](#project-structure-supported)
- [Known Limitations](#known-limitations)

---

## Quick Start

```bash
# Clone / copy the script to your machine
# Install dependencies (one-time)
pip3 install openpyxl pandas tabulate

# Analyse a single project — console output
python3 uipath_analyzer.py ./MyProject

# Analyse an entire repo — generate all output files
python3 uipath_analyzer.py ./MyRepo --repo --output all --report-dir ./reports
```

---

## Prerequisites

**Python 3.8 or higher** is required.

| Package | Version | Purpose |
|---------|---------|---------|
| `openpyxl` | ≥ 3.0 | Excel report generation (`.xlsx`) |
| `pandas` | ≥ 1.3 | CSV export and tabular data handling |
| `tabulate` | ≥ 0.8 | Console table formatting |

> **Built-in Python libraries used** (no install needed):
> `xml.etree.ElementTree`, `json`, `pathlib`, `re`, `argparse`, `collections`, `datetime`, `os`, `sys`

### Install all at once

```bash
pip3 install openpyxl pandas tabulate
```

### Verify your environment

```bash
python3 --version          # Should be 3.8+
python3 -c "import openpyxl, pandas, tabulate; print('All dependencies OK')"
```

---

### Offline / Airgapped Installation

If your target machine has **no internet access**, pre-bundle all dependencies on a connected machine and carry them across on USB or a network share.

**Step 1 — On an internet-connected machine:**

```bash
# Download all required wheels (including transitive dependencies) into a folder
pip3 download openpyxl pandas tabulate -d ./uipath_analyzer_packages
```

**Step 2 — Transfer to the airgapped machine:**

Copy both of these across (USB / share / internal repo):
```
uipath_analyzer.py
uipath_analyzer_packages/   ← the folder from Step 1
```

**Step 3 — On the airgapped machine:**

```bash
# Install from local folder, no internet required
pip3 install --no-index --find-links=./uipath_analyzer_packages openpyxl pandas tabulate

# Verify
python3 -c "import openpyxl, pandas, tabulate; print('Ready')"
```

**Step 4 — Run normally:**

```bash
python3 uipath_analyzer.py ./MyRepo --repo --output all
```

> ✅ Once dependencies are installed, **no internet connection is ever needed** to run the tool.

---

## Usage

### Single Project

Analyse one UiPath project folder (must contain a `project.json`):

```bash
# Console report only
python3 uipath_analyzer.py ./MyProject

# Console + all file exports
python3 uipath_analyzer.py ./MyProject --output all --report-dir ./reports

# Excel only
python3 uipath_analyzer.py ./MyProject --output excel --report-dir ./reports

# JSON only
python3 uipath_analyzer.py ./MyProject --output json

# Suppress console when exporting files
python3 uipath_analyzer.py ./MyProject --output all --quiet
```

### Full Repository Scan

Use `--repo` to auto-discover **all UiPath projects** anywhere under the root folder, regardless of nesting depth. Supports flat, nested, and mixed structures.

```bash
python3 uipath_analyzer.py ./MyRepo --repo --output all --report-dir ./reports
```

The tool will find every subfolder that contains a `project.json` — you don't need to specify project paths individually.

**Example structures it handles:**

```
# Flat
MyRepo/
  ProjectA/project.json
  ProjectB/project.json

# Nested
MyRepo/
  Finance/
    PayrollBot/project.json
    ExpenseClaims/project.json
  HR/
    Onboarding/project.json

# Mixed — any combination of the above
```

### All Output Flags

| Flag | Description |
|------|-------------|
| `--repo` | Scan for multiple projects under the root folder |
| `--output console` | Console report only (default) |
| `--output excel` | Generate `.xlsx` report |
| `--output json` | Generate `.json` report |
| `--output csv` | Generate `.csv` files |
| `--output all` | Generate Excel + JSON + CSV |
| `--report-dir <path>` | Folder to write output files (default: current dir) |
| `--quiet` | Suppress console output when exporting files |

---

## What Gets Extracted

### Activities

Every XML element in every `.xaml` file is classified against a catalog of **80+ known UiPath activity tags**, grouped into categories:

| Category | Examples |
|----------|---------|
| Excel | Read Range, Write Range, Excel Application Scope |
| Web Browser | Open Browser, Navigate To, Upload File |
| UI Automation | Click, Type Into, Attach Window, Send Hotkey |
| Outlook / Email | Send Outlook Mail, Get Mail Messages |
| PDF | Read PDF Text, Read PDF With OCR |
| Orchestrator | Get Transaction Item, Set Transaction Status, Get Credential |
| System | Log Message, Invoke Workflow File, Invoke PowerShell |
| Database | Execute Query, Connect, Insert DataTable |
| Document Understanding | Validation Station, Classify Document |
| SAP | SAP-specific selectors and GUI interactions |
| Control Flow | For Each, Try Catch, If, Throw |

Unknown tags are flagged separately for review.

---

### Web Browser Targets

The tool deep-parses every `Selector` attribute and `Url` attribute across all activities to extract:

| Field | Source | Example |
|-------|--------|---------|
| **URL** | `Url=` attribute on `OpenBrowser`/`Navigate` | `https://portal.acme.com/login` |
| **URL** | `url=` inside selector XML | `https://acme.workday.com/d/inst/...` |
| **Page Title** | `title=` inside `<html ...>` selector node | `ACME Invoice Portal - Dashboard` |
| **Browser** | `app=` in selector / `BrowserType=` attribute | `Google Chrome` / `Microsoft Edge` |
| **Activity** | `DisplayName` of the activity containing it | `Click Login Button` |
| **XAML File** | Relative path of the workflow file | `Main/Main.xaml` |

UiPath stores selectors as HTML-entity-encoded XML strings like:

```xml
Selector="&lt;html app='chrome.exe' title='ACME Portal' /&gt;&lt;webctrl tag='BUTTON' /&gt;"
```

The tool decodes and fully parses these — including nested `<html>`, `<webctrl>`, and `<wnd>` nodes — rather than just doing a text search.

---

### Windows Desktop Targets

Extracted from `AttachWindow`, `AttachBrowser`, and any activity whose `Selector` contains a `<wnd ...>` node:

| Field | Source | Example |
|-------|--------|---------|
| **Process (EXE)** | `app=` in `<wnd>` selector node | `saplogon.exe` |
| **Friendly App Name** | Mapped from known EXE names | `SAP Logon / SAP GUI` |
| **Window Title** | `title=` in `<wnd>` selector node | `Park Vendor Invoice: Company Code` |
| **Activity** | `DisplayName` of the containing activity | `Attach SAP ECC` |
| **XAML File** | Relative path of the workflow file | `Main/Main.xaml` |

Known EXE → friendly name mappings include: `saplogon.exe`, `excel.exe`, `winword.exe`, `outlook.exe`, `acrobat.exe`, `powershell.exe`, `cmd.exe`, and more. Custom EXE names are preserved as-is.

---

## Reusability Matrix

The Reusability Matrix is the key feature for **repository-level analysis**. It answers: *"Which workflows and activity patterns appear across multiple projects and should be extracted into shared, reusable components?"*

It operates at two levels:

1. **Workflow-level** — scores each individual `.xaml` file
2. **Activity Category-level** — scores each category (Excel, Orchestrator, etc.)

### How Scoring Works

Each workflow is scored 0–100 based on four rule-based signals. No AI or LLM is involved — all scoring is deterministic arithmetic.

#### Workflow Score Breakdown

| Signal | Max Points | Description |
|--------|-----------|-------------|
| **Cross-project presence** | +20 pts | How many projects contain this workflow as a percentage of total projects |
| **Invocation frequency** | +15 pts | How many times other workflows call this one via `InvokeWorkflowFile` |
| **Name pattern match** | +20 pts | Workflow filename matches known reusable patterns (e.g. `SendEmail`, `ErrorHandler`, `Login`, `GetCredential`) |
| **High-value categories** | +3 pts each | Contains categories like Orchestrator, System, or Email which are inherently cross-project |

**Base score starts at 40.** Maximum is capped at 100.

#### Activity Category Score Breakdown

| Signal | Weight |
|--------|--------|
| **Category base score** | 70% of known base (Orchestrator=92, System=88, Email=85, Excel=82...) |
| **Cross-project spread** | 30% — how many projects use this category as a ratio of total |

### Score Labels

| Score | Label | Meaning |
|-------|-------|---------|
| 80 – 100 | 🟢 High | Strong candidate for a shared reusable library |
| 60 – 79 | 🟡 Medium | Good candidate — parameterise and document before sharing |
| 40 – 59 | 🟠 Low-Medium | Project-specific for now — monitor as more projects are added |
| 0 – 39 | 🔴 Low | Highly context/selector-dependent — keep project-local |

### Shared Library Candidates

Any workflow or activity category scoring **≥ 60** and appearing in **2 or more projects** is flagged as a **Shared Library Candidate**, with a specific recommended action:

| Score Range | Recommended Action |
|------------|-------------------|
| ≥ 80, 2+ projects | Extract to shared `.xaml`, define standard In/Out arguments, publish as NuGet package to Orchestrator |
| ≥ 60, 2+ projects | Parameterise hardcoded values, document the interface, share within the team |
| ≥ 60, 1 project | Abstract hardcoded values to arguments now, ready for future reuse |
| < 60 | Keep project-specific; review if selector-dependent logic can be generalised |

---

## Output Files

When using `--output all`, the following files are generated (all named with project/repo name + timestamp):

| File | Contents |
|------|---------|
| `*_UiPath_Report_*.xlsx` | Full Excel workbook — 9 sheets (see below) |
| `*_UiPath_Report_*.json` | Complete structured JSON — all data including targets |
| `*_Activities_*.csv` | All activities across all projects |
| `*_WorkflowMatrix_*.csv` | Workflow-level reusability scores |
| `*_ActivityMatrix_*.csv` | Activity category-level reusability scores |
| `*_WebTargets_*.csv` | All web URLs and page titles extracted |
| `*_WindowTargets_*.csv` | All Windows process names and window titles |

### Excel Workbook — 9 Sheets

| Sheet | Contents |
|-------|---------|
| 📊 Repo Summary | Project list with stats — XAML count, activity count, app count |
| 🔁 Workflow Matrix | Reusability scores, descriptions, recommendations per workflow |
| 📦 Activity Matrix | Reusability scores per activity category |
| ⭐ Shared Library Candidates | Prioritised list of components to extract |
| 📋 All Activities | Every activity across every project |
| 🖥️ Applications Map | Which apps are automated and in which projects |
| 📦 Dependencies | NuGet package versions cross-referenced by project |
| 🌐 Web Targets | URLs and page titles with clickable hyperlinks |
| 🖥️ Window Targets | Windows process names and window titles |

---

## Airgapped & Data Privacy

**This tool is fully airgapped compatible.** It contains zero network calls of any kind.

| What it does | Network? |
|---|---|
| Parse XAML / XML files | ❌ No — Python stdlib `xml.etree.ElementTree` |
| Classify activities | ❌ No — static dictionary lookups |
| Parse selectors | ❌ No — local string and XML operations |
| Score reusability | ❌ No — deterministic arithmetic only |
| Generate Excel / CSV / JSON | ❌ No — `openpyxl` and `pandas` write to local disk |
| Activity descriptions | ❌ No — hardcoded static text in `WORKFLOW_PATTERNS` dict |

**No data ever leaves the machine.** The tool does not call any LLM, AI API, telemetry service, or external endpoint at any point — not during analysis, not during export, not during startup.

The "AI-generated descriptions" label used in earlier documentation was inaccurate. All descriptions and recommendations are produced by **static rule-based lookups and arithmetic scoring** embedded directly in the script. Reviewing the source confirms this — there are no `requests`, `urllib`, `socket`, `http`, or subprocess calls of any kind.

This makes the tool safe for use in:
- 🏦 Banking and financial services environments
- 🏥 Healthcare and HIPAA-regulated systems
- 🏛️ Government and classified networks
- 🏭 Industrial/OT networks with strict isolation requirements
- Any environment where **source code / workflow IP must not leave the network**

---

## Project Structure Supported

The tool supports any folder structure as long as each UiPath project root contains a `project.json`:

```
MyRepo/
├── ProjectA/
│   ├── project.json        ← detected
│   ├── Main/
│   │   └── Main.xaml
│   └── Workflows/
│       └── SendEmail.xaml
├── Finance/
│   ├── PayrollBot/
│   │   ├── project.json    ← detected
│   │   └── Main/
│   │       └── Main.xaml
│   └── ExpenseClaims/
│       ├── project.json    ← detected
│       └── Main/
│           └── Main.xaml
└── HR/
    └── Onboarding/
        ├── project.json    ← detected
        └── Main/
            └── Main.xaml
```

For a single project, the root itself must contain `project.json`:

```
MyProject/
├── project.json            ← required
├── Main/
│   └── Main.xaml
└── Workflows/
    ├── SendEmail.xaml
    └── ErrorHandler.xaml
```

---

## Known Limitations

| Limitation | Detail |
|-----------|--------|
| **Dynamic selectors** | Selectors built at runtime using variables (e.g. `"<wnd title='" + varTitle + "'>"`) cannot be statically parsed. The variable placeholder will appear as-is or be skipped. |
| **Encrypted workflows** | `.xaml` files encrypted via Orchestrator Studio Web cannot be read. |
| **Custom activity packages** | Activities from in-house NuGet packages not in the built-in catalog will appear as `unknown`. Add them to `ACTIVITY_CATALOG` in the script. |
| **VB.NET / C# code** | `InvokeCode` blocks are not parsed for logic — only the activity itself is recorded. |
| **Hardcoded credential values** | If credentials are hardcoded in XAML (not via `GetCredential`), the tool does not flag this as a security issue. |
| **Cross-machine selectors** | Window titles with machine-specific values (e.g. `title='Server01 - SAP'`) will extract the literal title string. |
