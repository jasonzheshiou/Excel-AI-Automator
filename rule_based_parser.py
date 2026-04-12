"""Rule-based Action Parser - Converts natural language to actions without LLM."""
import re
from typing import List, Tuple, Optional
from dataclasses import dataclass


@dataclass
class ParsedAction:
    """A parsed action with type and parameters."""
    action_type: str
    params: dict
    confidence: float  # 0.0 to 1.0
    raw_text: str


class RuleBasedParser:
    """Parses natural language into Excel actions using pattern matching."""
    
    # Common file path patterns
    FILE_PATH_PATTERN = re.compile(r'([A-Z]:\\[^\s,]+(?:\.xlsx?|\.xlsm))', re.IGNORECASE)
    
    # Sheet reference patterns - more precise matching
    SHEET_PATTERNS = [
        re.compile(r'(?:in|from|to)\s+(?:the\s+)?["\']?([A-Z_a-z]\w*)["\']?\s+sheet', re.IGNORECASE),
        re.compile(r'sheet\s+(?:named\s+)?["\']?([A-Z_a-z]\w*)["\']?', re.IGNORECASE),
    ]
    
    # Cell reference patterns: "A1", "B3", "A1:D50", "A1 to D50"
    CELL_REF_PATTERN = re.compile(r'\b([A-Z]{1,3}\d{1,7})\b', re.IGNORECASE)
    
    # Range patterns: "A1 to D50", "A1:D50", "cells A1 through D50"
    RANGE_PATTERN = re.compile(r'(?:cells?\s+)?\b([A-Z]{1,3}\d{1,7})\s+(?:to|through|:|-)\s+([A-Z]{1,3}\d{1,7})\b', re.IGNORECASE)
    
    # Value patterns: "to 'value'", "to 100", "=value"
    VALUE_PATTERN = re.compile(r'(?:to|as|=|value\s+(?:of\s+)?)\s*["\']?([^"\'.\n]+?)["\']?(?:\s|$|,)', re.IGNORECASE)
    
    def parse(self, text: str) -> List[str]:
        """Parse natural language text into action lines.
        
        Returns:
            List of action line strings in the format: ACTION | param=value | ...
        """
        actions = []
        
        # Split into sentences/processable chunks
        sentences = self._split_sentences(text)
        
        for sentence in sentences:
            sentence = sentence.strip()
            if not sentence:
                continue
            
            # Try to match known patterns
            parsed_actions = self._parse_sentence(sentence)
            for action in parsed_actions:
                action_line = self._format_action(action)
                actions.append(action_line)
        
        # If no specific actions matched, try generic extraction
        if not actions:
            actions = self._try_generic_extraction(text)
        
        return actions
    
    def _split_sentences(self, text: str) -> List[str]:
        """Split text into sentences or logical chunks."""
        # Split on common delimiters
        parts = re.split(r'[.,;]\s+|then\s+', text, flags=re.IGNORECASE)
        return [p.strip() for p in parts if p.strip()]
    
    def _extract_sheet(self, text: str) -> Optional[str]:
        """Extract sheet name from text."""
        for pattern in self.SHEET_PATTERNS:
            match = pattern.search(text)
            if match:
                sheet_name = match.group(1).strip()
                # Filter out common words that aren't sheet names
                if sheet_name.lower() not in ['the', 'a', 'an', 'to', 'from', 'in', 'is', 'it']:
                    return sheet_name
        return None
    
    def _parse_sentence(self, sentence: str) -> List[ParsedAction]:
        """Parse a single sentence into actions."""
        actions = []
        lower = sentence.lower()
        
        # OPEN pattern: "open C:\file.xlsx", "open the file at..."
        if any(word in lower for word in ['open', 'load', 'start with']):
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            if paths:
                actions.append(ParsedAction(
                    action_type="OPEN",
                    params={"file_path": paths[0]},
                    confidence=0.9,
                    raw_text=sentence
                ))
        
        # SAVE pattern: "save the file", "save C:\file.xlsx"
        if any(word in lower for word in ['save', 'store']) and 'save as' not in lower and 'saveas' not in lower:
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            if paths:
                actions.append(ParsedAction(
                    action_type="SAVE",
                    params={"file_path": paths[0]},
                    confidence=0.9,
                    raw_text=sentence
                ))
            else:
                actions.append(ParsedAction(
                    action_type="SAVE",
                    params={"file_path": "[current_file]"},
                    confidence=0.7,
                    raw_text=sentence
                ))
        
        # SAVE AS pattern: "save as C:\new.xlsx", "save it as..."
        if 'save as' in lower or 'save it as' in lower:
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            if len(paths) >= 2:
                actions.append(ParsedAction(
                    action_type="SAVEAS",
                    params={"file_path": paths[0], "save_as": paths[1]},
                    confidence=0.9,
                    raw_text=sentence
                ))
            elif len(paths) == 1:
                actions.append(ParsedAction(
                    action_type="SAVEAS",
                    params={"file_path": "[current_file]", "save_as": paths[0]},
                    confidence=0.8,
                    raw_text=sentence
                ))
        
        # CLOSE pattern: "close the file", "close C:\file.xlsx"
        if 'close' in lower:
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            if paths:
                actions.append(ParsedAction(
                    action_type="CLOSE",
                    params={"file_path": paths[0]},
                    confidence=0.9,
                    raw_text=sentence
                ))
        
        # COPY pattern: "copy cells A1 to D50", "copy A1:D100"
        if 'copy' in lower and 'paste' not in lower and 'copyfile' not in lower and 'copy folder' not in lower and 'copy file' not in lower:
            range_match = self.RANGE_PATTERN.search(sentence)
            cell_matches = self.CELL_REF_PATTERN.findall(sentence)
            sheet_name = self._extract_sheet(sentence)
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            
            if range_match:
                start_cell = range_match.group(1).upper()
                end_cell = range_match.group(2).upper()
                params = {
                    "range": f"{start_cell}:{end_cell}",
                }
                if sheet_name:
                    params["sheet"] = sheet_name
                if paths:
                    params["file_path"] = paths[0]
                
                actions.append(ParsedAction(
                    action_type="COPY",
                    params=params,
                    confidence=0.85,
                    raw_text=sentence
                ))
        
        # PASTE patterns
        if 'paste' in lower:
            cell_matches = self.CELL_REF_PATTERN.findall(sentence)
            sheet_name = self._extract_sheet(sentence)
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            
            params = {}
            if sheet_name:
                params["sheet"] = sheet_name
            if paths:
                params["file_path"] = paths[0]
            
            # Find target cell - look for "at", "starting at", "to", "into"
            target_patterns = [
                re.compile(r'(?:at|starting\s+at|to|into)\s+([A-Z]{1,3}\d{1,7})', re.IGNORECASE),
            ]
            target_cell = None
            for pat in target_patterns:
                m = pat.search(sentence)
                if m:
                    target_cell = m.group(1).upper()
                    break
            
            if not target_cell and cell_matches:
                target_cell = cell_matches[-1].upper()
            
            if target_cell:
                params["target_cell"] = target_cell
            
            # Determine paste type
            if 'value' in lower and 'paste' in lower:
                action_type = "PASTEVALUE"
            elif 'link' in lower:
                action_type = "PASTELINK"
            else:
                action_type = "PASTE"
            
            if params:
                actions.append(ParsedAction(
                    action_type=action_type,
                    params=params,
                    confidence=0.8,
                    raw_text=sentence
                ))
        
        # CHANGE pattern: "change B3 to 'value'", "set cell A1 to 100"
        if any(word in lower for word in ['change', 'set', 'update', 'modify']):
            cell_matches = self.CELL_REF_PATTERN.findall(sentence)
            sheet_name = self._extract_sheet(sentence)
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            
            # Try to extract the value
            value_match = re.search(r'(?:to|as|=)\s*["\']?([^"\'.\n]+)', sentence, re.IGNORECASE)
            
            if cell_matches:
                params = {
                    "cell": cell_matches[0].upper(),
                }
                if value_match:
                    params["value"] = value_match.group(1).strip()
                if sheet_name:
                    params["sheet"] = sheet_name
                if paths:
                    params["file_path"] = paths[0]
                
                actions.append(ParsedAction(
                    action_type="CHANGE",
                    params=params,
                    confidence=0.8,
                    raw_text=sentence
                ))
        
        # CHECK/VERIFY pattern: "check if A1 equals 'value'", "verify B3 is 100"
        if any(word in lower for word in ['check', 'verify', 'confirm', 'validate']):
            cell_matches = self.CELL_REF_PATTERN.findall(sentence)
            sheet_name = self._extract_sheet(sentence)
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            
            # Try to extract expected value
            expected_patterns = re.findall(r'(?:equals?|is|should\s+be|=|value\s+(?:of\s+)?)\s*["\']?([^"\'.\n]+)', sentence, re.IGNORECASE)
            
            if cell_matches:
                params = {
                    "cell": cell_matches[0].upper(),
                    "condition": "=="
                }
                if expected_patterns:
                    params["expected"] = expected_patterns[0].strip()
                if sheet_name:
                    params["sheet"] = sheet_name
                if paths:
                    params["file_path"] = paths[0]
                
                actions.append(ParsedAction(
                    action_type="CHECK",
                    params=params,
                    confidence=0.75,
                    raw_text=sentence
                ))
        
        # DELETE pattern
        if any(word in lower for word in ['delete', 'remove']):
            range_match = self.RANGE_PATTERN.search(sentence)
            cell_matches = self.CELL_REF_PATTERN.findall(sentence)
            sheet_name = self._extract_sheet(sentence)
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            
            params = {}
            if range_match:
                params["range"] = f"{range_match.group(1).upper()}:{range_match.group(2).upper()}"
            elif cell_matches:
                params["range"] = f"{cell_matches[0].upper()}:{cell_matches[-1].upper()}"
            if sheet_name:
                params["sheet"] = sheet_name
            if paths:
                params["file_path"] = paths[0]
            
            if params:
                actions.append(ParsedAction(
                    action_type="DELETE",
                    params=params,
                    confidence=0.8,
                    raw_text=sentence
                ))
        
        # CLEAR pattern
        if any(word in lower for word in ['clear', 'empty']):
            range_match = self.RANGE_PATTERN.search(sentence)
            cell_matches = self.CELL_REF_PATTERN.findall(sentence)
            sheet_name = self._extract_sheet(sentence)
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            
            params = {}
            if range_match:
                params["range"] = f"{range_match.group(1).upper()}:{range_match.group(2).upper()}"
            elif cell_matches:
                params["range"] = f"{cell_matches[0].upper()}:{cell_matches[-1].upper()}"
            if sheet_name:
                params["sheet"] = sheet_name
            if paths:
                params["file_path"] = paths[0]
            
            if params:
                actions.append(ParsedAction(
                    action_type="CLEAR",
                    params=params,
                    confidence=0.8,
                    raw_text=sentence
                ))
        
        # REFRESH pattern
        if any(word in lower for word in ['refresh', 'recalculate', 'update.*data']):
            sheet_name = self._extract_sheet(sentence)
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            
            params = {}
            if sheet_name:
                params["sheet"] = sheet_name
            if paths:
                params["file_path"] = paths[0]
            
            if params:
                actions.append(ParsedAction(
                    action_type="REFRESH",
                    params=params,
                    confidence=0.8,
                    raw_text=sentence
                ))
        
        # COPYFILE pattern
        if 'copy file' in lower or 'copyfile' in lower:
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            if len(paths) >= 2:
                actions.append(ParsedAction(
                    action_type="COPYFILE",
                    params={"source": paths[0], "destination": paths[1]},
                    confidence=0.9,
                    raw_text=sentence
                ))
        
        # CREATFOLDER pattern
        if any(word in lower for word in ['create folder', 'make directory', 'mkdir', 'create directory']):
            # Try to extract folder path - could be Windows path
            folder_patterns = [
                re.compile(r'(?:folder|directory|path)\s+(?:at\s+)?([A-Z]:\\[^\s,]+)', re.IGNORECASE),
                re.compile(r'(?:create|make)\s+(?:a\s+)?(?:folder|directory)\s+([A-Z]:\\[^\s,]+)', re.IGNORECASE),
            ]
            folder_path = None
            for pat in folder_patterns:
                m = pat.search(sentence)
                if m:
                    folder_path = m.group(1)
                    break
            
            if folder_path:
                actions.append(ParsedAction(
                    action_type="CREATFOLDER",
                    params={"path": folder_path},
                    confidence=0.85,
                    raw_text=sentence
                ))
        
        # MACRO pattern
        if any(word in lower for word in ['run macro', 'execute macro', 'macro']):
            paths = self.FILE_PATH_PATTERN.findall(sentence)
            # Try to extract macro name
            macro_match = re.search(r'macro\s+(?:named\s+)?["\']?(\w+)["\']?', sentence, re.IGNORECASE)
            
            params = {}
            if macro_match:
                params["macro"] = macro_match.group(1)
            if paths:
                params["file_path"] = paths[0]
            
            if params:
                actions.append(ParsedAction(
                    action_type="MACRO",
                    params=params,
                    confidence=0.8,
                    raw_text=sentence
                ))
        
        return actions
    
    def _try_generic_extraction(self, text: str) -> List[str]:
        """Fallback: Try to extract actions using generic patterns."""
        actions = []
        paths = self.FILE_PATH_PATTERN.findall(text)
        
        # If we found file paths, at least suggest OPEN
        if paths:
            for path in paths:
                actions.append(f"OPEN | file_path={path}")
        
        # Look for cell references and suggest CHANGE
        cell_matches = self.CELL_REF_PATTERN.findall(text)
        if len(cell_matches) >= 1:
            value_match = re.search(r'(?:to|as|=)\s*["\']?([^"\'.\n]+)', text, re.IGNORECASE)
            if value_match:
                actions.append(
                    f"CHANGE | cell={cell_matches[0].upper()} | value={value_match.group(1).strip()}"
                )
        
        return actions
    
    def _format_action(self, action: ParsedAction) -> str:
        """Format a ParsedAction into the action line string."""
        parts = [action.action_type]
        for key, value in action.params.items():
            parts.append(f"{key}={value}")
        return " | ".join(parts)
