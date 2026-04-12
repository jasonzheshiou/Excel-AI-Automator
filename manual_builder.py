"""Manual Action Builder - UI components for building actions without AI."""
import streamlit as st
from typing import Dict, List, Optional


# Action type definitions with their required parameters
ACTION_DEFINITIONS = {
    "OPEN": {
        "description": "Open an Excel file",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True, 
                         "placeholder": "C:\\data\\report.xlsx"},
        }
    },
    "CLOSE": {
        "description": "Close an Excel file",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
        }
    },
    "SAVE": {
        "description": "Save the current Excel file",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
        }
    },
    "SAVEAS": {
        "description": "Save Excel file with a new name/location",
        "params": {
            "file_path": {"label": "Current File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "save_as": {"label": "Save As (New Path)", "type": "text", "required": True,
                       "placeholder": "C:\\data\\report_final.xlsx"},
        }
    },
    "REFRESH": {
        "description": "Refresh/recalculate a worksheet",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": False,
                     "placeholder": "Data"},
        }
    },
    "CHANGE": {
        "description": "Change a cell value",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Summary"},
            "cell": {"label": "Cell Reference", "type": "text", "required": True,
                    "placeholder": "B5"},
            "value": {"label": "New Value", "type": "text", "required": True,
                     "placeholder": "100"},
        }
    },
    "CHECK": {
        "description": "Verify a cell value matches expected",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Summary"},
            "cell": {"label": "Cell Reference", "type": "text", "required": True,
                    "placeholder": "A1"},
            "expected": {"label": "Expected Value", "type": "text", "required": True,
                        "placeholder": "Total"},
            "condition": {"label": "Condition", "type": "select", "required": True,
                         "options": ["==", "<>"]},
        }
    },
    "COPY": {
        "description": "Copy a range of cells",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Data"},
            "range": {"label": "Cell Range", "type": "text", "required": True,
                     "placeholder": "A1:D100"},
        }
    },
    "PASTE": {
        "description": "Paste clipboard contents",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Results"},
            "target_cell": {"label": "Target Cell", "type": "text", "required": True,
                           "placeholder": "A1"},
        }
    },
    "PASTEVALUE": {
        "description": "Paste values only",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Results"},
            "target_cell": {"label": "Target Cell", "type": "text", "required": True,
                           "placeholder": "A1"},
        }
    },
    "PASTELINK": {
        "description": "Paste as linked data",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Results"},
            "target_cell": {"label": "Target Cell", "type": "text", "required": True,
                           "placeholder": "A1"},
        }
    },
    "DELETE": {
        "description": "Delete cells/rows/columns",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Data"},
            "range": {"label": "Cell Range", "type": "text", "required": True,
                     "placeholder": "A1:A100"},
        }
    },
    "CLEAR": {
        "description": "Clear cell contents",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsx"},
            "sheet": {"label": "Sheet Name", "type": "text", "required": True,
                     "placeholder": "Data"},
            "range": {"label": "Cell Range", "type": "text", "required": True,
                     "placeholder": "A1:D100"},
        }
    },
    "COPYFILE": {
        "description": "Copy a file",
        "params": {
            "source": {"label": "Source File", "type": "text", "required": True,
                      "placeholder": "C:\\data\\report.xlsx"},
            "destination": {"label": "Destination", "type": "text", "required": True,
                           "placeholder": "C:\\backup\\report_backup.xlsx"},
        }
    },
    "COPYFOLDER": {
        "description": "Copy a folder",
        "params": {
            "source": {"label": "Source Folder", "type": "text", "required": True,
                      "placeholder": "C:\\data"},
            "destination": {"label": "Destination", "type": "text", "required": True,
                           "placeholder": "C:\\backup"},
        }
    },
    "CREATFOLDER": {
        "description": "Create a folder",
        "params": {
            "path": {"label": "Folder Path", "type": "text", "required": True,
                    "placeholder": "C:\\data\\new_folder"},
        }
    },
    "RENAME": {
        "description": "Rename a file",
        "params": {
            "source": {"label": "Current Path", "type": "text", "required": True,
                      "placeholder": "C:\\data\\old_name.xlsx"},
            "destination": {"label": "New Path", "type": "text", "required": True,
                           "placeholder": "C:\\data\\new_name.xlsx"},
        }
    },
    "MACRO": {
        "description": "Run an Excel macro",
        "params": {
            "file_path": {"label": "File Path", "type": "text", "required": True,
                         "placeholder": "C:\\data\\report.xlsm"},
            "macro": {"label": "Macro Name", "type": "text", "required": True,
                     "placeholder": "FormatReport"},
        }
    },
    "AUTOCAL": {
        "description": "Set calculation to automatic",
        "params": {}
    },
}


def render_action_form(index: int = 0) -> Optional[str]:
    """Render a form to build a single action.
    
    Returns:
        Formatted action line string, or None if not submitted
    """
    # Action type selector
    action_types = list(ACTION_DEFINITIONS.keys())
    selected_action = st.selectbox(
        "Action Type",
        options=action_types,
        format_func=lambda x: f"{x} - {ACTION_DEFINITIONS[x]['description']}",
        key=f"action_type_{index}"
    )
    
    if not selected_action:
        return None
    
    action_def = ACTION_DEFINITIONS[selected_action]
    
    # Show description
    st.caption(action_def["description"])
    
    # Build parameter inputs
    params = {}
    cols = st.columns(2)
    col_idx = 0
    
    for param_name, param_def in action_def["params"].items():
        col = cols[col_idx % 2]
        col_idx += 1
        
        with col:
            required_mark = " *" if param_def.get("required", False) else ""
            label = f"{param_def['label']}{required_mark}"
            
            if param_def["type"] == "text":
                value = st.text_input(
                    label,
                    placeholder=param_def.get("placeholder", ""),
                    key=f"param_{index}_{param_name}",
                    help="Required" if param_def.get("required") else "Optional"
                )
            elif param_def["type"] == "select":
                value = st.selectbox(
                    label,
                    options=param_def.get("options", []),
                    key=f"param_{index}_{param_name}"
                )
            
            if value:
                params[param_name] = value
    
    # Validation and submit
    required_params = [k for k, v in action_def["params"].items() if v.get("required", False)]
    missing_required = [p for p in required_params if p not in params]
    
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 2])
    
    with col1:
        submit = st.button("➕ Add Action", use_container_width=True, disabled=len(missing_required) > 0)
    
    with col2:
        preview = st.button("👁️ Preview", use_container_width=True)
    
    if missing_required:
        st.error(f"Missing required fields: {', '.join(missing_required)}")
    
    # Build action line
    action_line = None
    if params or not action_def["params"]:
        parts = [selected_action]
        for key, value in params.items():
            parts.append(f"{key}={value}")
        action_line = " | ".join(parts)
    
    if preview and action_line:
        st.success("Preview:")
        st.code(action_line)
    
    if submit and action_line:
        return action_line
    
    return None


def render_quick_actions() -> Optional[str]:
    """Render quick action buttons for common operations."""
    st.subheader("⚡ Quick Actions")
    
    quick_actions = [
        ("OPEN", "OPEN | file_path="),
        ("SAVE", "SAVE | file_path="),
        ("CHANGE", "CHANGE | file_path= | sheet= | cell= | value="),
        ("COPY", "COPY | file_path= | sheet= | range="),
        ("PASTEVALUE", "PASTEVALUE | file_path= | sheet= | target_cell="),
    ]
    
    cols = st.columns(len(quick_actions))
    for i, (label, template) in enumerate(quick_actions):
        with cols[i]:
            if st.button(label, key=f"quick_{label}", use_container_width=True):
                return template
    
    return None


def render_action_list_editor(actions: List[str]) -> List[str]:
    """Render an editable list of actions.
    
    Returns:
        Updated list of actions
    """
    if not actions:
        st.info("No actions added yet. Use the form above to build actions.")
        return actions
    
    st.subheader(f"📋 Action List ({len(actions)} actions)")
    
    updated_actions = []
    to_delete = []
    
    for i, action in enumerate(actions):
        with st.container():
            col1, col2, col3, col4 = st.columns([0.3, 5, 0.4, 0.4])
            
            with col1:
                st.markdown(f"**{i + 1}**")
            
            with col2:
                edited = st.text_input(
                    f"Action {i + 1}",
                    value=action,
                    key=f"edit_action_{i}",
                    label_visibility="collapsed"
                )
                updated_actions.append(edited)
            
            with col3:
                if st.button("⬆️", key=f"up_{i}", help="Move up"):
                    if i > 0:
                        actions[i], actions[i-1] = actions[i-1], actions[i]
                        st.rerun()
            
            with col4:
                if st.button("🗑️", key=f"del_{i}", help="Delete"):
                    to_delete.append(i)
    
    # Delete marked actions
    if to_delete:
        for i in sorted(to_delete, reverse=True):
            actions.pop(i)
        st.rerun()
    
    return updated_actions
