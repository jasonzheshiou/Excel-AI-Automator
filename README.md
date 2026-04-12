# AI-Powered Excel Automation

An intelligent Excel automation tool that uses local LLMs (like LM Studio) to convert natural language descriptions into executable Excel automation steps.

## Features

- **Natural Language Input**: Describe your Excel task in plain English
- **Dual Mode Operation**: 
  - **🤖 AI Assistant**: AI-powered parsing with local LLMs or rule-based fallback
  - **🔧 Manual Builder**: Form-based action builder - pick action types, fill in fields
- **Step-by-Step Execution**: Review and edit each action before running
- **Flexible Actions**: Supports 19 Excel operations (open, copy, paste, save, macros, file operations, etc.)
- **Visual Feedback**: Real-time status indicators for each action

## Prerequisites

1. **Python 3.10+**
2. **Microsoft Excel** (installed on your system)
3. **Local LLM Server** (optional, for AI-powered parsing):
   - [LM Studio](https://lmstudio.ai/) (recommended)
   - [Ollama](https://ollama.com/)
   - Any OpenAI-compatible API server
   
**Note**: The application works without LLM using built-in rule-based parsing!

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Start your local LLM server:
   - **LM Studio**: Load a model and start the server (default: `http://localhost:1234/v1`)
   - Note the server URL and model name for configuration

## Usage

1. **Start the Streamlit app**:
```bash
streamlit run app.py
```

2. **Choose Your Mode** (segmented control at top):

### 🤖 AI Assistant Mode
- Describe your task in natural language
- Click "Generate Actions" to create the steps
  - **LLM Mode**: AI intelligently parses complex requests (requires LM Studio)
  - **Offline Mode**: Rule-based pattern matching (no LLM needed, auto-fallback)
- Review & Edit actions, then Execute

### 🔧 Manual Builder Mode
- **Pick an action type** from the dropdown (OPEN, CHANGE, COPY, etc.)
- **Fill in the required fields** (marked with *)
  - File paths, sheet names, cell references, values, etc.
- **Click 'Add Action'** to add to the list
- **Reorder** (⬆️) or **Delete** (🗑️) actions as needed
- **Run All** to execute the sequence

3. **Execute**:
   - Click "Run All" to execute all actions sequentially
   - Or click the ▶️ button next to individual actions (AI mode)

## Available Actions

| Action | Description | Example |
|--------|-------------|---------|
| `OPEN` | Open an Excel file | `OPEN \| file_path=C:\data\report.xlsx` |
| `CLOSE` | Close an Excel file | `CLOSE \| file_path=C:\data\report.xlsx` |
| `SAVE` | Save the current file | `SAVE \| file_path=C:\data\report.xlsx` |
| `SAVEAS` | Save as a new file | `SAVEAS \| file_path=C:\data\r.xlsx \| save_as=C:\data\final.xlsx` |
| `REFRESH` | Refresh/recalculate a sheet | `REFRESH \| file_path=C:\data\r.xlsx \| sheet=Data` |
| `CHANGE` | Change a cell value | `CHANGE \| file_path=C:\data\r.xlsx \| sheet=Summary \| cell=B5 \| value=100` |
| `CHECK` | Verify a cell value | `CHECK \| file_path=C:\data\r.xlsx \| sheet=S \| cell=A1 \| expected=Total \| condition==` |
| `COPY` | Copy a range of cells | `COPY \| file_path=C:\data\r.xlsx \| sheet=Data \| range=A1:D100` |
| `PASTE` | Paste clipboard | `PASTE \| file_path=C:\data\r.xlsx \| sheet=R \| target_cell=A1` |
| `PASTEVALUE` | Paste values only | `PASTEVALUE \| file_path=C:\data\r.xlsx \| sheet=R \| target_cell=A1` |
| `PASTELINK` | Paste as linked data | `PASTELINK \| file_path=C:\data\r.xlsx \| sheet=R \| target_cell=A1` |
| `DELETE` | Delete cells/rows/columns | `DELETE \| file_path=C:\data\r.xlsx \| sheet=D \| range=A1:A100` |
| `CLEAR` | Clear cell contents | `CLEAR \| file_path=C:\data\r.xlsx \| sheet=D \| range=A1:D100` |
| `COPYFILE` | Copy a file | `COPYFILE \| source=C:\data\r.xlsx \| destination=C:\backup\r.xlsx` |
| `COPYFOLDER` | Copy a folder | `COPYFOLDER \| source=C:\data \| destination=C:\backup` |
| `CREATFOLDER` | Create a folder | `CREATFOLDER \| path=C:\data\new_folder` |
| `RENAME` | Rename a file | `RENAME \| source=C:\data\old.xlsx \| destination=C:\data\new.xlsx` |
| `MACRO` | Run an Excel macro | `MACRO \| file_path=C:\data\r.xlsm \| macro=FormatReport` |

## Project Structure

```
Excel_Automation/
├── app.py                  # Main Streamlit application
├── llm_client.py           # LLM integration (OpenAI-compatible API)
├── excel_executor.py       # Excel action execution engine
├── config.py               # Application configuration
├── requirements.txt        # Python dependencies
├── .env.example            # Example environment variables
├── Automation.xlsm         # Original VBA macro file
└── QWEN.md                 # Project documentation
```

## Configuration

Create a `.env` file in the project root (copy from `.env.example`):

```env
LLM_BASE_URL=http://localhost:1234/v1
LLM_MODEL=local-model
LLM_TEMPERATURE=0.1
LLM_MAX_TOKENS=2000
DEFAULT_EXCEL_FILE=Automation.xlsm
```

## Troubleshooting

### LLM Connection Failed
- Ensure LM Studio (or your LLM server) is running
- Verify the base URL and model name
- Check if the model is fully loaded

### Excel Automation Issues
- Make sure Excel is installed on your system
- Grant necessary permissions when prompted
- Check file paths are correct and accessible

### Actions Fail to Execute
- Verify file paths use double backslashes (`\\`) in Windows
- Ensure the target Excel files exist
- Check sheet names match exactly

## Architecture

```
User Input (Natural Language)
        ↓
    ┌─────────────────────┐
    │   LLM Client        │
    │  ┌───────────────┐  │
    │  │ LLM (if avail)│  │ ← AI-powered parsing
    │  └───────────────┘  │
    │         ↓ fallback  │
    │  ┌───────────────┐  │
    │  │Rule-Based Parser│ ← Pattern matching (always works)
    │  └───────────────┘  │
    └─────────────────────┘
        ↓
    Action Lines (Structured)
        ↓
    Excel Executor (xlwings)
        ↓
    Excel Automation (Results)
```

## How Offline Mode Works

The offline mode uses pattern matching to extract actions from natural language:

| Pattern | Extracted Action |
|---------|------------------|
| "Open C:\file.xlsx" | `OPEN \| file_path=C:\file.xlsx` |
| "Copy cells A1 to D50 from Data sheet" | `COPY \| sheet=Data \| range=A1:D50` |
| "Paste values to Summary at A1" | `PASTEVALUE \| sheet=Summary \| target_cell=A1` |
| "Change B3 to 'value' in Dashboard" | `CHANGE \| cell=B3 \| value=value \| sheet=Dashboard` |

For complex requests, **LLM mode** is recommended as it handles ambiguity better.

## License

This project is for educational and internal use.
