# AI-Powered Excel Automation

An intelligent Excel automation tool with **three modes of operation**: AI Assistant (LLM), AI Offline (rule-based), and Manual Builder (form-based).

## Quick Start

1. **Install dependencies**: `pip install -r requirements.txt`
2. **Run the app**: `streamlit run app.py`
3. **Choose your mode** and start building actions!

## Three Modes of Operation

| Mode | Description | Best For |
|------|-------------|----------|
| **🤖 AI (LLM)** | AI-powered parsing with local LLM | Complex, multi-step requests |
| **🤖 AI (Offline)** | Rule-based pattern matching | Simple requests, no LLM available |
| **🔧 Manual Builder** | Form-based action builder with dropdowns | Precise control, learning the system |

## Project Structure

```
Excel_Automation/
├── app.py                  # Main Streamlit application
├── llm_client.py           # LLM integration + rule-based fallback
├── rule_based_parser.py    # Pattern-based action extraction
├── manual_builder.py       # Form-based action builder
├── excel_executor.py       # Excel action execution engine
├── config.py               # Application configuration
├── requirements.txt        # Python dependencies
└── Automation.xlsm         # Original VBA macro file
```

## Available Actions (19 total)

| Category | Actions |
|----------|---------|
| **File Operations** | `OPEN`, `CLOSE`, `SAVE`, `SAVEAS` |
| **Cell Operations** | `CHANGE`, `CHECK`, `COPY`, `PASTE`, `PASTEVALUE`, `PASTELINK` |
| **Sheet Operations** | `REFRESH`, `DELETE`, `CLEAR` |
| **File System** | `COPYFILE`, `COPYFOLDER`, `CREATFOLDER`, `RENAME` |
| **Macros** | `MACRO`, `AUTOCAL` |

## Manual Builder Mode

In Manual Builder mode, you:
1. **Select action type** from dropdown (e.g., OPEN, CHANGE, COPY)
2. **Fill in required fields** (file paths, sheet names, cell refs, values)
3. **Preview** the action line before adding
4. **Add to list** and reorder as needed
5. **Run All** to execute

Each action has a form with:
- **Required fields** (marked with *) - must be filled to enable "Add Action"
- **Optional fields** - can be left blank
- **Preview button** - shows the formatted action line before adding

## How It Works

```
User Input
    ↓
┌───────────────────────────────┐
│     Mode Selector             │
│  ┌─────────┐  ┌────────────┐ │
│  │ AI Mode │  │ Manual     │ │
│  │         │  │ Builder    │ │
│  └────┬────┘  └─────┬──────┘ │
│       │              │        │
│  ┌────┴────┐    ┌────┴──────┐ │
│  │ LLM or  │    │ Form-based│ │
│  │ Rules   │    │ Builder   │ │
│  └────┬────┘    └────┬──────┘ │
└───────┼──────────────┼────────┘
        ↓              ↓
    Action Lines (Structured)
            ↓
    Excel Executor (xlwings)
            ↓
    Excel Automation (Results)
```
