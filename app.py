"""AI-Powered Excel Automation Streamlit Application."""
import streamlit as st
import os
from typing import List

from llm_client import LLMClient
from excel_executor import ExcelExecutor, ActionResult, ActionStatus
from manual_builder import render_action_form, render_action_list_editor
from config import Config


def render_action_line(action_text: str, index: int, results: dict, editing: bool = False):
    """Render a single action line with controls."""
    result = results.get(index)
    
    col1, col2, col3, col4 = st.columns([0.5, 5, 0.5, 0.5])
    
    with col1:
        status_icon = "⏳"
        if result:
            if result.status == ActionStatus.SUCCESS:
                status_icon = "✅"
            elif result.status == ActionStatus.ERROR:
                status_icon = "❌"
            elif result.status == ActionStatus.SKIPPED:
                status_icon = "⏭️"
        st.markdown(f"**{status_icon}**")
    
    with col2:
        if editing:
            # Allow editing the action
            edited = st.text_input(
                f"Action {index + 1}",
                value=action_text,
                key=f"action_input_{index}",
                label_visibility="collapsed"
            )
            return edited
        else:
            st.text(action_text)
            return action_text
    
    with col3:
        if st.button("▶️", key=f"run_{index}", help="Run this action only"):
            return True  # Signal to run
    return False


def main():
    st.set_page_config(
        page_title="AI Excel Automation",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Excel Automation")
    st.caption("Choose your mode: AI-powered or Manual Builder")
    
    # Initialize session state
    if "actions" not in st.session_state:
        st.session_state.actions = []
    if "results" not in st.session_state:
        st.session_state.results = {}
    if "llm_connected" not in st.session_state:
        st.session_state.llm_connected = False
    if "executor" not in st.session_state:
        st.session_state.executor = ExcelExecutor()
    if "mode" not in st.session_state:
        st.session_state.mode = "ai"  # "ai" or "manual"
    
    # Mode Selector Tabs
    mode_tab = st.segmented_control(
        "Operation Mode",
        options=["ai", "manual"],
        default=st.session_state.mode,
        format_func=lambda x: "🤖 AI Assistant" if x == "ai" else "🔧 Manual Builder",
        key="mode_selector"
    )
    
    if mode_tab != st.session_state.mode:
        st.session_state.mode = mode_tab
        st.rerun()
    
    # Route to appropriate mode
    if st.session_state.mode == "ai":
        render_ai_mode()
    else:
        render_manual_mode()


def render_ai_mode():
    """Render the AI-assisted mode UI."""
    
    # Sidebar - Configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # LLM Connection
        st.subheader("🤖 LLM Settings")
        llm_base_url = st.text_input("LLM Base URL", value=Config.LLM_BASE_URL, 
                                      help="Leave empty to use rule-based parser (offline mode)")
        llm_model = st.text_input("Model Name", value=Config.LLM_MODEL)
        
        # Update config if URL changed
        if llm_base_url != Config.LLM_BASE_URL:
            Config.LLM_BASE_URL = llm_base_url
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔌 Test LLM", use_container_width=True):
                with st.spinner("Testing..."):
                    try:
                        test_client = LLMClient()
                        if test_client.test_connection():
                            st.success("✅ LLM Connected")
                            st.session_state.llm_connected = True
                        else:
                            st.warning("⚠️ LLM unavailable")
                            st.session_state.llm_connected = False
                    except Exception as e:
                        st.warning(f"⚠️ {str(e)[:50]}")
                        st.session_state.llm_connected = False
        with col2:
            if st.button("🔄 Refresh", use_container_width=True):
                # Force re-initialization
                st.session_state.llm_client = LLMClient()
                st.rerun()
        
        # Status indicator
        llm_client = st.session_state.get("llm_client", LLMClient())
        if llm_client.is_llm_available():
            st.success("🟢 LLM Mode (AI-powered)")
            st.caption("Using local LLM for intelligent parsing")
        else:
            st.info("🔵 Offline Mode (Rule-based)")
            st.caption("Using pattern matching for parsing")
        
        st.divider()
        
        # Actions Management
        st.subheader("📋 Actions")
        if st.session_state.actions:
            st.write(f"{len(st.session_state.actions)} actions loaded")
            if st.button("🗑️ Clear All", use_container_width=True, key="clear_all_ai"):
                st.session_state.actions = []
                st.session_state.results = {}
                st.rerun()
        else:
            st.info("No actions loaded")
        
        st.divider()
        
        # Help
        with st.expander("❓ How to use"):
            st.markdown("""
            1. **Describe your task**: Enter a natural language description
            2. **Generate actions**: Click 'Generate Actions' to create the steps
               - **LLM Mode**: AI intelligently parses your request (requires LM Studio)
               - **Offline Mode**: Rule-based pattern matching (no LLM needed)
            3. **Review & Edit**: Review the generated actions, edit if needed
            4. **Execute**: Run all actions or individual ones
            
            **Offline Mode Examples:**
            - "Open C:\\Reports\\sales.xlsx"
            - "Copy cells A1 to D50 from Data sheet"
            - "Paste values to Summary sheet at A1"
            - "Change B3 to 'Q4 Results' in Dashboard sheet"
            """)
        
        with st.expander("📖 Available Actions"):
            st.markdown("""
            - **OPEN** - Open Excel file
            - **CLOSE** - Close Excel file
            - **SAVE** - Save file
            - **SAVEAS** - Save as new file
            - **REFRESH** - Refresh sheet
            - **CHANGE** - Change cell value
            - **CHECK** - Verify cell value
            - **COPY** - Copy range
            - **PASTE** - Paste clipboard
            - **PASTEVALUE** - Paste values only
            - **PASTELINK** - Paste as link
            - **DELETE** - Delete cells
            - **CLEAR** - Clear contents
            - **COPYFILE** - Copy file
            - **COPYFOLDER** - Copy folder
            - **CREATFOLDER** - Create folder
            - **RENAME** - Rename file
            - **MACRO** - Run macro
            """)
    
    # Main content - AI Mode
    st.subheader("📝 Describe Your Task")
    user_input = st.text_area(
        "What would you like to automate?",
        placeholder="Example: Open C:\\Reports\\sales.xlsx, copy cells A1 to D50 from the Data sheet, and paste them as values into the Summary sheet starting at A1",
        height=100
    )
    
    col1, col2 = st.columns([1, 3])
    with col1:
        generate_btn = st.button("🤖 Generate Actions", use_container_width=True)
    
    if generate_btn and user_input:
        with st.spinner("🤖 Generating actions..."):
            try:
                # Initialize LLM client (will auto-fallback to rule-based if LLM unavailable)
                llm_client = LLMClient()
                st.session_state.llm_client = llm_client
                
                actions = llm_client.generate_actions(user_input)
                st.session_state.actions = actions
                st.session_state.results = {}
                
                if llm_client.is_llm_available():
                    st.success(f"✅ Generated {len(actions)} actions (LLM-powered)")
                else:
                    st.info(f"✅ Generated {len(actions)} actions (Offline mode)")
                
                if actions:
                    st.rerun()
                else:
                    st.warning("⚠️ No actions generated. Try rephrasing your request or check the action format.")
            except Exception as e:
                st.error(f"❌ Failed to generate actions: {str(e)}")
    elif generate_btn and not user_input:
        st.warning("Please describe your task first")
    
    # Display Actions
    if st.session_state.actions:
        st.divider()
        st.subheader("📋 Generated Actions")
        
        # Edit mode toggle
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            edit_mode = st.checkbox("✏️ Edit Mode")
        with col2:
            if st.button("▶️ Run All", use_container_width=True, key="run_all_ai"):
                with st.spinner("Executing actions..."):
                    executor = ExcelExecutor()
                    results = executor.execute_actions(st.session_state.actions)
                    for i, result in enumerate(results):
                        st.session_state.results[i] = result
                    executor.cleanup()
                    st.rerun()
        
        st.divider()
        
        # Render each action
        updated_actions = []
        run_single = -1
        
        for i, action in enumerate(st.session_state.actions):
            cols = st.columns([0.3, 5.5, 0.6, 0.6])
            
            # Status icon
            result = st.session_state.results.get(i)
            status_icon = "⏳"
            if result:
                if result.status == ActionStatus.SUCCESS:
                    status_icon = "✅"
                elif result.status == ActionStatus.ERROR:
                    status_icon = "❌"
                elif result.status == ActionStatus.SKIPPED:
                    status_icon = "⏭️"
            
            with cols[0]:
                st.markdown(f"**{i + 1}** {status_icon}")
            
            # Action text (editable if edit mode)
            with cols[1]:
                if edit_mode:
                    edited = st.text_input(
                        f"Action {i + 1}",
                        value=action,
                        key=f"edit_{i}",
                        label_visibility="collapsed"
                    )
                    updated_actions.append(edited)
                else:
                    st.code(action, language=None)
            
            # Run single button
            with cols[2]:
                if st.button("▶️", key=f"run_{i}", help="Run this action"):
                    run_single = i
            
            # Delete button
            with cols[3]:
                if st.button("🗑️", key=f"del_{i}", help="Delete this action"):
                    st.session_state.actions.pop(i)
                    if i in st.session_state.results:
                        del st.session_state.results[i]
                    st.rerun()
            
            # Show result message if exists
            if result and result.message:
                with cols[1]:
                    msg_color = "green" if result.status == ActionStatus.SUCCESS else "red"
                    st.caption(f":{msg_color}[{result.message}]")
        
        # Update actions if in edit mode
        if edit_mode and updated_actions:
            st.session_state.actions = updated_actions
    
    # Execution Results
    if st.session_state.results:
        st.divider()
        st.subheader("📊 Execution Results")
        
        success_count = sum(1 for r in st.session_state.results.values() if r.status == ActionStatus.SUCCESS)
        error_count = sum(1 for r in st.session_state.results.values() if r.status == ActionStatus.ERROR)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Actions", len(st.session_state.results))
        with col2:
            st.metric("✅ Successful", success_count)
        with col3:
            st.metric("❌ Failed", error_count)
        
        if error_count == 0:
            st.success("🎉 All actions executed successfully!")
        else:
            st.error("⚠️ Some actions failed. Check the results above for details.")


def render_manual_mode():
    """Render the manual action builder mode UI."""
    
    # Sidebar for manual mode
    with st.sidebar:
        st.header("🔧 Manual Builder")
        
        st.subheader("📋 Actions")
        if st.session_state.actions:
            st.write(f"{len(st.session_state.actions)} actions loaded")
            if st.button("🗑️ Clear All", use_container_width=True, key="clear_all_manual"):
                st.session_state.actions = []
                st.session_state.results = {}
                st.rerun()
        else:
            st.info("No actions loaded")
        
        st.divider()
        
        # Run actions
        st.subheader("▶️ Execute")
        if st.session_state.actions:
            if st.button("▶️ Run All Actions", use_container_width=True, key="run_all_manual"):
                with st.spinner("Executing actions..."):
                    executor = ExcelExecutor()
                    results = executor.execute_actions(st.session_state.actions)
                    for i, result in enumerate(results):
                        st.session_state.results[i] = result
                    executor.cleanup()
                    st.rerun()
        else:
            st.info("Add actions first")
        
        st.divider()
        
        with st.expander("❓ How to use"):
            st.markdown("""
            1. **Pick an action type** from the dropdown
            2. **Fill in the required fields** (marked with *)
            3. **Click 'Add Action'** to add to the list
            4. **Reorder or delete** actions as needed
            5. **Run All** to execute the sequence
            """)
    
    # Main content - Manual Builder Mode
    st.subheader("🔧 Build Actions")
    
    # Action form
    new_action = render_action_form()
    
    if new_action:
        st.session_state.actions.append(new_action)
        st.success(f"✅ Added: {new_action}")
        st.rerun()
    
    st.divider()
    
    # Action list editor
    if st.session_state.actions:
        st.session_state.actions = render_action_list_editor(st.session_state.actions)
    
    # Execution Results
    if st.session_state.results:
        st.divider()
        st.subheader("📊 Execution Results")
        
        success_count = sum(1 for r in st.session_state.results.values() if r.status == ActionStatus.SUCCESS)
        error_count = sum(1 for r in st.session_state.results.values() if r.status == ActionStatus.ERROR)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Actions", len(st.session_state.results))
        with col2:
            st.metric("✅ Successful", success_count)
        with col3:
            st.metric("❌ Failed", error_count)
        
        if error_count == 0:
            st.success("🎉 All actions executed successfully!")
        else:
            st.error("⚠️ Some actions failed. Check the results above for details.")


if __name__ == "__main__":
    main()
