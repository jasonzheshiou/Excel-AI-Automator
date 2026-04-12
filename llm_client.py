"""LLM Integration module for converting user requests to Excel actions."""
import json
from openai import OpenAI
from config import Config
from typing import List, Dict, Optional
from rule_based_parser import RuleBasedParser

# System prompt that instructs the LLM how to generate action lines
SYSTEM_PROMPT = """You are an AI assistant that converts natural language Excel task descriptions into structured action lines.

## Available Actions:
1. **OPEN** - Open an Excel file
   - Format: `OPEN | file_path=<path_to_excel_file>`
   - Example: `OPEN | file_path=C:\\data\\report.xlsx`

2. **CLOSE** - Close an Excel file
   - Format: `CLOSE | file_path=<path_to_excel_file>`
   - Example: `CLOSE | file_path=C:\\data\\report.xlsx`

3. **SAVE** - Save the current Excel file
   - Format: `SAVE | file_path=<path_to_excel_file>`
   - Example: `SAVE | file_path=C:\\data\\report.xlsx`

4. **SAVEAS** - Save Excel file with a new name/location
   - Format: `SAVEAS | file_path=<current_file> | save_as=<new_path>`
   - Example: `SAVEAS | file_path=C:\\data\\report.xlsx | save_as=C:\\data\\report_final.xlsx`

5. **REFRESH** - Refresh/recalculate a worksheet
   - Format: `REFRESH | file_path=<file> | sheet=<sheet_name>`
   - Example: `REFRESH | file_path=C:\\data\\report.xlsx | sheet=Data`

6. **CHANGE** - Change a cell value
   - Format: `CHANGE | file_path=<file> | sheet=<sheet> | cell=<cell_ref> | value=<new_value>`
   - Example: `CHANGE | file_path=C:\\data\\report.xlsx | sheet=Summary | cell=B5 | value=100`

7. **CHECK** - Verify a cell value matches expected
   - Format: `CHECK | file_path=<file> | sheet=<sheet> | cell=<cell_ref> | expected=<value> | condition=<=|<>`
   - Example: `CHECK | file_path=C:\\data\\report.xlsx | sheet=Summary | cell=A1 | expected=Total | condition==`

8. **COPY** - Copy a range of cells
   - Format: `COPY | file_path=<file> | sheet=<sheet> | range=<start_cell>:<end_cell>`
   - Example: `COPY | file_path=C:\\data\\report.xlsx | sheet=Data | range=A1:D100`

9. **PASTE** - Paste clipboard contents
   - Format: `PASTE | file_path=<file> | sheet=<sheet> | target_cell=<cell_ref>`
   - Example: `PASTE | file_path=C:\\data\\report.xlsx | sheet=Results | target_cell=A1`

10. **PASTEVALUE** - Paste values only
    - Format: `PASTEVALUE | file_path=<file> | sheet=<sheet> | target_cell=<cell_ref>`
    - Example: `PASTEVALUE | file_path=C:\\data\\report.xlsx | sheet=Results | target_cell=A1`

11. **PASTELINK** - Paste as linked data
    - Format: `PASTELINK | file_path=<file> | sheet=<sheet> | target_cell=<cell_ref>`
    - Example: `PASTELINK | file_path=C:\\data\\report.xlsx | sheet=Results | target_cell=A1`

12. **DELETE** - Delete cells/rows/columns
    - Format: `DELETE | file_path=<file> | sheet=<sheet> | range=<start_cell>:<end_cell>`
    - Example: `DELETE | file_path=C:\\data\\report.xlsx | sheet=Data | range=A1:A100`

13. **CLEAR** - Clear cell contents
    - Format: `CLEAR | file_path=<file> | sheet=<sheet> | range=<start_cell>:<end_cell>`
    - Example: `CLEAR | file_path=C:\\data\\report.xlsx | sheet=Data | range=A1:D100`

14. **COPYFILE** - Copy a file
    - Format: `COPYFILE | source=<file_path> | destination=<dest_path>`
    - Example: `COPYFILE | source=C:\\data\\report.xlsx | destination=C:\\backup\\report_backup.xlsx`

15. **COPYFOLDER** - Copy a folder
    - Format: `COPYFOLDER | source=<folder_path> | destination=<dest_path>`
    - Example: `COPYFOLDER | source=C:\\data | destination=C:\\backup`

16. **CREATFOLDER** - Create a folder
    - Format: `CREATFOLDER | path=<folder_path>`
    - Example: `CREATFOLDER | path=C:\\data\\new_folder`

17. **RENAME** - Rename a file
    - Format: `RENAME | source=<old_path> | destination=<new_path>`
    - Example: `RENAME | source=C:\\data\\old_name.xlsx | destination=C:\\data\\new_name.xlsx`

18. **MACRO** - Run an Excel macro
    - Format: `MACRO | file_path=<file> | macro=<macro_name>`
    - Example: `MACRO | file_path=C:\\data\\report.xlsm | macro=FormatReport`

## Rules:
- Output ONLY the action lines, one per line
- Each action line follows the format: `ACTION | param1=value1 | param2=value2 | ...`
- Use double backslashes (\\\\) in Windows paths
- Do NOT include explanations, just the actions
- If a task requires multiple steps, output each step as a separate line
- For file paths, use the exact path provided by the user
- Sheet names should match exactly as they appear in Excel

## Examples:

User: "Open the sales report at C:\\Reports\\sales.xlsx, then copy cells A1 to D50 from the Data sheet and paste them as values into the Summary sheet starting at A1"
Output:
OPEN | file_path=C:\\Reports\\sales.xlsx
COPY | file_path=C:\\Reports\\sales.xlsx | sheet=Data | range=A1:D50
PASTEVALUE | file_path=C:\\Reports\\sales.xlsx | sheet=Summary | target_cell=A1

User: "Change cell B3 to 'Q4 Results' in the Dashboard sheet of C:\\Reports\\monthly.xlsx, then save the file"
Output:
CHANGE | file_path=C:\\Reports\\monthly.xlsx | sheet=Dashboard | cell=B3 | value=Q4 Results
SAVE | file_path=C:\\Reports\\monthly.xlsx
"""


class LLMClient:
    """Client for interacting with local LLM (LM Studio or similar)."""
    
    def __init__(self):
        self.client = None
        self.model = Config.LLM_MODEL
        self.use_llm = True
        self.rule_parser = RuleBasedParser()
        
        # Try to initialize OpenAI client, but don't fail if unreachable
        try:
            self.client = OpenAI(
                base_url=Config.LLM_BASE_URL,
                api_key="not-needed"
            )
        except Exception:
            self.use_llm = False
    
    def generate_actions(self, user_request: str) -> List[str]:
        """Convert user natural language request to action lines.
        
        Tries LLM first, falls back to rule-based parser if unavailable.
        """
        # Try LLM if available
        if self.use_llm and self.client:
            try:
                return self._generate_with_llm(user_request)
            except Exception:
                self.use_llm = False
        
        # Fallback to rule-based parser
        return self._generate_with_rules(user_request)
    
    def _generate_with_llm(self, user_request: str) -> List[str]:
        """Generate actions using LLM."""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_request}
            ],
            temperature=Config.LLM_TEMPERATURE,
            max_tokens=Config.LLM_MAX_TOKENS
        )
        
        actions_text = response.choices[0].message.content.strip()
        actions = [line.strip() for line in actions_text.split('\n') if line.strip()]
        return actions
    
    def _generate_with_rules(self, user_request: str) -> List[str]:
        """Generate actions using rule-based parsing."""
        return self.rule_parser.parse(user_request)
    
    def test_connection(self) -> bool:
        """Test if the LLM server is reachable."""
        try:
            if not self.client:
                return False
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": "Say 'connected'"}],
                max_tokens=10
            )
            return response.choices[0].message.content is not None
        except Exception:
            return False
    
    def is_llm_available(self) -> bool:
        """Check if LLM is available."""
        return self.use_llm and self.client is not None
