"""Configuration management for the application."""
import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    """Application configuration."""
    
    # LLM Settings
    LLM_BASE_URL = os.getenv("LLM_BASE_URL", "http://localhost:1234/v1")
    LLM_MODEL = os.getenv("LLM_MODEL", "local-model")
    LLM_TEMPERATURE = float(os.getenv("LLM_TEMPERATURE", "0.1"))
    LLM_MAX_TOKENS = int(os.getenv("LLM_MAX_TOKENS", "2000"))
    
    # Excel Settings
    DEFAULT_EXCEL_FILE = os.getenv("DEFAULT_EXCEL_FILE", "Automation.xlsm")
    
    # Action definitions
    VALID_ACTIONS = [
        "OPEN", "CLOSE", "SAVE", "SAVEAS", "REFRESH",
        "CHANGE", "CHECK", "COPY", "PASTE", "PASTEVALUE", "PASTELINK",
        "DELETE", "CLEAR",
        "COPYFILE", "COPYFOLDER", "CREATFOLDER", "RENAME", "SetFileAttr",
        "MACRO", "AUTOCAL"
    ]
