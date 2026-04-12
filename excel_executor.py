"""Excel Action Executor - Performs individual actions on Excel files."""
import os
import shutil
import xlwings as xw
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from enum import Enum


class ActionStatus(Enum):
    PENDING = "pending"
    SUCCESS = "success"
    ERROR = "error"
    SKIPPED = "skipped"


@dataclass
class ActionResult:
    """Result of executing an action."""
    action_line: str
    status: ActionStatus
    message: str = ""


class ActionParser:
    """Parse action lines into structured format."""
    
    @staticmethod
    def parse(action_line: str) -> Tuple[str, Dict[str, str]]:
        """Parse an action line into action type and parameters.
        
        Returns:
            Tuple of (action_type, params_dict)
        """
        parts = [p.strip() for p in action_line.split('|')]
        action_type = parts[0].strip()
        
        params = {}
        for part in parts[1:]:
            if '=' in part:
                key, value = part.split('=', 1)
                params[key.strip()] = value.strip()
        
        return action_type, params


class ExcelExecutor:
    """Executes Excel actions using xlwings."""
    
    def __init__(self):
        self.app: Optional[xw.App] = None
        self.workbooks: Dict[str, xw.Book] = {}
    
    def _get_app(self) -> xw.App:
        """Get or create Excel application instance."""
        if self.app is None:
            self.app = xw.App(visible=True)
            self.app.display_alerts = False
            self.app.screen_updating = True
        return self.app
    
    def _get_workbook(self, file_path: str) -> xw.Book:
        """Get or open a workbook."""
        # Normalize path
        file_path = os.path.abspath(file_path)
        
        if file_path in self.workbooks:
            return self.workbooks[file_path]
        
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        app = self._get_app()
        wb = app.books.open(file_path)
        self.workbooks[file_path] = wb
        return wb
    
    def _get_sheet(self, wb: xw.Book, sheet_name: str) -> xw.Sheet:
        """Get a worksheet by name."""
        for sheet in wb.sheets:
            if sheet.name.lower() == sheet_name.lower():
                return sheet
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook")
    
    def execute_action(self, action_line: str) -> ActionResult:
        """Execute a single action line."""
        try:
            action_type, params = ActionParser.parse(action_line)
            
            if action_type == "OPEN":
                return self._action_open(params)
            elif action_type == "CLOSE":
                return self._action_close(params)
            elif action_type == "SAVE":
                return self._action_save(params)
            elif action_type == "SAVEAS":
                return self._action_saveas(params)
            elif action_type == "REFRESH":
                return self._action_refresh(params)
            elif action_type == "CHANGE":
                return self._action_change(params)
            elif action_type == "CHECK":
                return self._action_check(params)
            elif action_type == "COPY":
                return self._action_copy(params)
            elif action_type == "PASTE":
                return self._action_paste(params)
            elif action_type == "PASTEVALUE":
                return self._action_pastevalue(params)
            elif action_type == "PASTELINK":
                return self._action_pastelink(params)
            elif action_type == "DELETE":
                return self._action_delete(params)
            elif action_type == "CLEAR":
                return self._action_clear(params)
            elif action_type == "COPYFILE":
                return self._action_copyfile(params)
            elif action_type == "COPYFOLDER":
                return self._action_copyfolder(params)
            elif action_type == "CREATFOLDER":
                return self._action_createfolder(params)
            elif action_type == "RENAME":
                return self._action_rename(params)
            elif action_type == "MACRO":
                return self._action_macro(params)
            else:
                return ActionResult(
                    action_line=action_line,
                    status=ActionStatus.ERROR,
                    message=f"Unknown action type: {action_type}"
                )
        except Exception as e:
            return ActionResult(
                action_line=action_line,
                status=ActionStatus.ERROR,
                message=str(e)
            )
    
    def execute_actions(self, action_lines: List[str]) -> List[ActionResult]:
        """Execute a sequence of actions."""
        results = []
        for action_line in action_lines:
            result = self.execute_action(action_line)
            results.append(result)
            if result.status == ActionStatus.ERROR:
                break  # Stop on error
        return results
    
    # Action implementations
    
    def _action_open(self, params: Dict[str, str]) -> ActionResult:
        file_path = params.get("file_path", "")
        if not file_path:
            return ActionResult("", ActionStatus.ERROR, "Missing file_path parameter")
        
        file_path = os.path.abspath(file_path)
        if not os.path.exists(file_path):
            return ActionResult("", ActionStatus.ERROR, f"File not found: {file_path}")
        
        try:
            self._get_workbook(file_path)
            return ActionResult(
                action_line=f"OPEN | file_path={file_path}",
                status=ActionStatus.SUCCESS,
                message=f"Opened: {os.path.basename(file_path)}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Failed to open: {str(e)}")
    
    def _action_close(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        if file_path in self.workbooks:
            wb = self.workbooks[file_path]
            wb.save()
            wb.close()
            del self.workbooks[file_path]
            return ActionResult(
                action_line=f"CLOSE | file_path={file_path}",
                status=ActionStatus.SUCCESS,
                message=f"Closed: {os.path.basename(file_path)}"
            )
        return ActionResult(
            action_line=f"CLOSE | file_path={file_path}",
            status=ActionStatus.SKIPPED,
            message="File was not open"
        )
    
    def _action_save(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        if file_path in self.workbooks:
            self.workbooks[file_path].save()
            return ActionResult(
                action_line=f"SAVE | file_path={file_path}",
                status=ActionStatus.SUCCESS,
                message=f"Saved: {os.path.basename(file_path)}"
            )
        # If not open, just check file exists
        if os.path.exists(file_path):
            return ActionResult(
                action_line=f"SAVE | file_path={file_path}",
                status=ActionStatus.SUCCESS,
                message=f"File exists: {os.path.basename(file_path)}"
            )
        return ActionResult("", ActionStatus.ERROR, "File not found")
    
    def _action_saveas(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        save_as = os.path.abspath(params.get("save_as", ""))
        
        if file_path in self.workbooks:
            self.workbooks[file_path].save(save_as)
            return ActionResult(
                action_line=f"SAVEAS | file_path={file_path}",
                status=ActionStatus.SUCCESS,
                message=f"Saved as: {save_as}"
            )
        return ActionResult("", ActionStatus.ERROR, "Source file not open")
    
    def _action_refresh(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        
        try:
            wb = self._get_workbook(file_path)
            if sheet_name:
                sheet = self._get_sheet(wb, sheet_name)
                sheet.calculate()
            else:
                wb.sheets[0].calculate()
            return ActionResult(
                action_line=f"REFRESH | file_path={file_path}",
                status=ActionStatus.SUCCESS,
                message="Sheet refreshed successfully"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Refresh failed: {str(e)}")
    
    def _action_change(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        cell = params.get("cell", "")
        value = params.get("value", "")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            sheet.range(cell).value = value
            return ActionResult(
                action_line=f"CHANGE | cell={cell}",
                status=ActionStatus.SUCCESS,
                message=f"Changed {cell} to '{value}'"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Change failed: {str(e)}")
    
    def _action_check(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        cell = params.get("cell", "")
        expected = params.get("expected", "")
        condition = params.get("condition", "==")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            actual_value = sheet.range(cell).value
            
            # Convert expected to appropriate type
            if expected.isdigit():
                expected = int(expected)
            elif expected.replace('.', '', 1).isdigit():
                expected = float(expected)
            
            # Check condition
            passed = False
            if condition == "==":
                passed = actual_value == expected
            elif condition == "<>":
                passed = actual_value != expected
            
            status = ActionStatus.SUCCESS if passed else ActionStatus.ERROR
            return ActionResult(
                action_line=f"CHECK | cell={cell}",
                status=status,
                message=f"Check {'passed' if passed else 'failed'}: {cell} = {actual_value}, expected {condition} {expected}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Check failed: {str(e)}")
    
    def _action_copy(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        cell_range = params.get("range", "")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            sheet.range(cell_range).copy()
            return ActionResult(
                action_line=f"COPY | range={cell_range}",
                status=ActionStatus.SUCCESS,
                message=f"Copied range {cell_range}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Copy failed: {str(e)}")
    
    def _action_paste(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        target_cell = params.get("target_cell", "")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            sheet.range(target_cell).sheet.api.Paste()
            return ActionResult(
                action_line=f"PASTE | target={target_cell}",
                status=ActionStatus.SUCCESS,
                message=f"Pasted to {target_cell}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Paste failed: {str(e)}")
    
    def _action_pastevalue(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        target_cell = params.get("target_cell", "")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            sheet.range(target_cell).api.PasteSpecial(Paste=-4163)  # xlPasteValues
            return ActionResult(
                action_line=f"PASTEVALUE | target={target_cell}",
                status=ActionStatus.SUCCESS,
                message=f"Pasted values to {target_cell}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Paste values failed: {str(e)}")
    
    def _action_pastelink(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        target_cell = params.get("target_cell", "")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            sheet.api.Paste(Link=True)
            return ActionResult(
                action_line=f"PASTELINK | target={target_cell}",
                status=ActionStatus.SUCCESS,
                message=f"Pasted link to {target_cell}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Paste link failed: {str(e)}")
    
    def _action_delete(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        cell_range = params.get("range", "")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            sheet.range(cell_range).api.Delete()
            return ActionResult(
                action_line=f"DELETE | range={cell_range}",
                status=ActionStatus.SUCCESS,
                message=f"Deleted range {cell_range}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Delete failed: {str(e)}")
    
    def _action_clear(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        sheet_name = params.get("sheet", "")
        cell_range = params.get("range", "")
        
        try:
            wb = self._get_workbook(file_path)
            sheet = self._get_sheet(wb, sheet_name)
            sheet.range(cell_range).clear_contents()
            return ActionResult(
                action_line=f"CLEAR | range={cell_range}",
                status=ActionStatus.SUCCESS,
                message=f"Cleared contents of {cell_range}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Clear failed: {str(e)}")
    
    # File system actions
    
    def _action_copyfile(self, params: Dict[str, str]) -> ActionResult:
        source = os.path.abspath(params.get("source", ""))
        destination = os.path.abspath(params.get("destination", ""))
        
        try:
            shutil.copy2(source, destination)
            return ActionResult(
                action_line=f"COPYFILE | source={source}",
                status=ActionStatus.SUCCESS,
                message=f"Copied file to {destination}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Copy file failed: {str(e)}")
    
    def _action_copyfolder(self, params: Dict[str, str]) -> ActionResult:
        source = os.path.abspath(params.get("source", ""))
        destination = os.path.abspath(params.get("destination", ""))
        
        try:
            shutil.copytree(source, destination)
            return ActionResult(
                action_line=f"COPYFOLDER | source={source}",
                status=ActionStatus.SUCCESS,
                message=f"Copied folder to {destination}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Copy folder failed: {str(e)}")
    
    def _action_createfolder(self, params: Dict[str, str]) -> ActionResult:
        path = os.path.abspath(params.get("path", ""))
        
        try:
            os.makedirs(path, exist_ok=True)
            return ActionResult(
                action_line=f"CREATFOLDER | path={path}",
                status=ActionStatus.SUCCESS,
                message=f"Created folder: {path}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Create folder failed: {str(e)}")
    
    def _action_rename(self, params: Dict[str, str]) -> ActionResult:
        source = os.path.abspath(params.get("source", ""))
        destination = os.path.abspath(params.get("destination", ""))
        
        try:
            os.rename(source, destination)
            return ActionResult(
                action_line=f"RENAME | source={source}",
                status=ActionStatus.SUCCESS,
                message=f"Renamed to {destination}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Rename failed: {str(e)}")
    
    def _action_macro(self, params: Dict[str, str]) -> ActionResult:
        file_path = os.path.abspath(params.get("file_path", ""))
        macro_name = params.get("macro", "")
        
        try:
            wb = self._get_workbook(file_path)
            wb.app.macro(f"'{wb.name}'!{macro_name}")()
            return ActionResult(
                action_line=f"MACRO | macro={macro_name}",
                status=ActionStatus.SUCCESS,
                message=f"Ran macro: {macro_name}"
            )
        except Exception as e:
            return ActionResult("", ActionStatus.ERROR, f"Macro failed: {str(e)}")
    
    def cleanup(self):
        """Close all workbooks and Excel application."""
        if self.app:
            for file_path, wb in list(self.workbooks.items()):
                try:
                    wb.save()
                    wb.close()
                except:
                    pass
            self.app.quit()
            self.app = None
            self.workbooks.clear()
