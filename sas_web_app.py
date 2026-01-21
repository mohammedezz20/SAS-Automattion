import streamlit as st
import threading
import queue
import time
import os
import shutil
import zipfile
import io
import tempfile
from datetime import datetime
import pandas as pd

# Import the class
from sas_automation import SASFormAutomator

st.set_page_config(page_title="SAS Automation Pro",
                   layout="wide", page_icon="🤖")

# Title and description
st.title("🤖 SAS Form Automation Pro")
st.markdown("**Automated form filling for SAS certification submissions**")

# Sidebar for browser selection
with st.sidebar:
    st.header("⚙️ Settings")
    browser_choice = st.selectbox(
        "Browser Selection",
        options=['auto', 'chrome', 'firefox', 'edge'],
        index=0,
        help="Auto mode will try Chrome → Edge → Firefox in order"
    )

    st.markdown("---")
    st.markdown("""
    ### 📋 Instructions
    1. Upload Excel file with required columns
    2. Preview the data
    3. Click "Start Automation"
    4. Monitor progress in real-time
    5. Download results when complete
    """)

    st.markdown("---")
    st.markdown("**Browser Info:**")
    if browser_choice == 'auto':
        st.info("🔄 Will auto-detect available browser")
    elif browser_choice == 'chrome':
        st.info("🌐 Using Google Chrome")
    elif browser_choice == 'firefox':
        st.info("🦊 Using Mozilla Firefox")
    elif browser_choice == 'edge':
        st.info("🔷 Using Microsoft Edge")

# Initialize session state
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'results' not in st.session_state:
    st.session_state.results = []
if 'running' not in st.session_state:
    st.session_state.running = False
if 'stop_flag' not in st.session_state:
    st.session_state.stop_flag = False
if 'browser_used' not in st.session_state:
    st.session_state.browser_used = None

# File upload
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Preview data
    df = pd.read_excel(uploaded_file)

    # Show file info
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.metric("Total Students", len(df))
    with col_info2:
        st.metric("Columns", len(df.columns))
    with col_info3:
        st.metric("File Size", f"{uploaded_file.size / 1024:.1f} KB")

    # Data preview
    with st.expander("📊 Preview Data (First 10 rows)", expanded=True):
        st.dataframe(df.head(10), use_container_width=True)

    # Validate required columns
    required_cols = ['First Name', 'Last Name', 'Email', 'Certificate Link']
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
        st.stop()

    # Save file temporarily
    temp_dir = tempfile.mkdtemp()
    excel_path = os.path.join(temp_dir, "data.xlsx")
    with open(excel_path, "wb") as f:
        f.write(uploaded_file.getvalue())

    # Control buttons
    col1, col2, col3 = st.columns([2, 2, 1])

    with col1:
        start_button = st.button(
            "▶️ Start Automation",
            type="primary",
            disabled=st.session_state.running,
            use_container_width=True
        )

        if start_button:
            st.session_state.running = True
            st.session_state.stop_flag = False
            st.session_state.logs = ["🚀 Starting automation..."]
            st.session_state.results = []
            st.session_state.browser_used = None

            # Create queues for thread communication
            log_queue = queue.Queue()
            result_queue = queue.Queue()

            def run_automation():
                try:
                    automator = SASFormAutomator(
                        "", excel_path, browser_choice=browser_choice)

                    # Store which browser was used
                    result_queue.put(('browser', automator.browser_name))

                    students = automator.read_excel()

                    for i, student in enumerate(students, 1):
                        # Check if user pressed Stop
                        try:
                            if not result_queue.empty():
                                stop_signal = result_queue.get_nowait()
                                if stop_signal == "STOP":
                                    log_queue.put(
                                        "⏸️ Automation stopped by user!")
                                    return
                        except:
                            pass

                        log_queue.put(
                            f"[{time.strftime('%H:%M:%S')}] 🔄 Processing {i}/{len(students)}: {student['firstName']} {student['lastName']}")

                        result = automator.fill_form(student)
                        result_queue.put(('result', result))

                        status_emoji = "✅" if result['status'] == "Success" else "❌"
                        log_queue.put(
                            f"{status_emoji} {result['status']}: {student['email']}")

                    log_queue.put("🎉 All students processed successfully!")
                    result_queue.put(('done', None))
                except Exception as e:
                    log_queue.put(f"❌ Critical error: {str(e)}")
                    result_queue.put(('error', str(e)))
                finally:
                    try:
                        automator.close_driver()
                    except:
                        pass

            # Start thread
            thread = threading.Thread(target=run_automation, daemon=True)
            thread.start()

            # Store queues in session state
            st.session_state.log_queue = log_queue
            st.session_state.result_queue = result_queue
            st.session_state.temp_dir = temp_dir
            st.rerun()

    with col2:
        stop_button = st.button(
            "⏹️ Stop Automation",
            type="secondary",
            disabled=not st.session_state.running,
            use_container_width=True
        )

        if stop_button:
            st.session_state.stop_flag = True
            if hasattr(st.session_state, 'result_queue'):
                st.session_state.result_queue.put("STOP")
            st.warning("⏸️ Stopping automation...")
            time.sleep(1)
            st.session_state.running = False
            st.rerun()

    with col3:
        if st.session_state.running:
            st.markdown("### 🔄")
            st.markdown("**Running**")

    # Update logs and results from queues
    if hasattr(st.session_state, 'log_queue'):
        while not st.session_state.log_queue.empty():
            try:
                msg = st.session_state.log_queue.get_nowait()
                st.session_state.logs.append(msg)
            except:
                break

    if hasattr(st.session_state, 'result_queue'):
        while not st.session_state.result_queue.empty():
            try:
                item = st.session_state.result_queue.get_nowait()
                if isinstance(item, tuple):
                    msg_type, data = item
                    if msg_type == 'result':
                        st.session_state.results.append(data)
                    elif msg_type == 'done':
                        st.session_state.running = False
                    elif msg_type == 'error':
                        st.session_state.running = False
                    elif msg_type == 'browser':
                        st.session_state.browser_used = data
            except:
                break

    # Show browser being used
    if st.session_state.browser_used:
        st.info(f"🌐 Using browser: **{st.session_state.browser_used}**")

    # Display logs
    if st.session_state.logs:
        st.markdown("---")
        st.subheader("📋 Live Logs")
        logs_container = st.container()
        with logs_container:
            # Show last 50 logs
            display_logs = st.session_state.logs[-50:]
            st.code("\n".join(display_logs), language=None)

    # Auto-refresh while running
    if st.session_state.running:
        time.sleep(2)
        st.rerun()

    # After completion
    if not st.session_state.running and st.session_state.results:
        st.markdown("---")
        st.success(
            f"✅ Automation completed! Processed {len(st.session_state.results)} students")

        # Statistics
        success_count = sum(
            1 for r in st.session_state.results if r['status'] == "Success")
        failed_count = len(st.session_state.results) - success_count

        col_stats1, col_stats2, col_stats3, col_stats4 = st.columns(4)
        with col_stats1:
            st.metric("Total Processed", len(st.session_state.results))
        with col_stats2:
            st.metric("✅ Success", success_count)
        with col_stats3:
            st.metric("❌ Failed", failed_count)
        with col_stats4:
            success_rate = (success_count / len(st.session_state.results)
                            * 100) if st.session_state.results else 0
            st.metric("Success Rate", f"{success_rate:.1f}%")

        # Results table
        st.subheader("📊 Detailed Results")
        results_df = pd.DataFrame(st.session_state.results)

        # Add color coding
        def color_status(val):
            color = 'background-color: #d4edda' if val == 'Success' else 'background-color: #f8d7da'
            return color

        styled_df = results_df.style.map(color_status, subset=['status'])
        st.dataframe(styled_df, use_container_width=True)

        # Create ZIP file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            # Add logs
            zf.writestr("logs.txt", "\n".join(st.session_state.logs))

            # Add results CSV
            zf.writestr("results.csv", results_df.to_csv(
                index=False, encoding='utf-8-sig'))

            # Add screenshots if any
            if hasattr(st.session_state, 'temp_dir') and os.path.exists(st.session_state.temp_dir):
                screenshot_count = 0
                for file in os.listdir(st.session_state.temp_dir):
                    if file.endswith(".png"):
                        file_path = os.path.join(
                            st.session_state.temp_dir, file)
                        zf.write(file_path, f"screenshots/{file}")
                        screenshot_count += 1

                if screenshot_count > 0:
                    st.info(
                        f"📸 {screenshot_count} screenshots included in download")

        zip_buffer.seek(0)

        # Download button
        col_dl1, col_dl2 = st.columns([3, 1])
        with col_dl1:
            st.download_button(
                label="📥 Download All Results (ZIP)",
                data=zip_buffer,
                file_name=f"SAS_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True,
                type="primary"
            )

        with col_dl2:
            if st.button("🔄 New Session", use_container_width=True):
                if hasattr(st.session_state, 'temp_dir') and os.path.exists(st.session_state.temp_dir):
                    try:
                        shutil.rmtree(st.session_state.temp_dir)
                    except:
                        pass
                for key in ['logs', 'results', 'running', 'stop_flag', 'log_queue', 'result_queue', 'temp_dir', 'browser_used']:
                    if key in st.session_state:
                        del st.session_state[key]
                st.success("✅ Session cleared! Upload a new file to start")
                st.rerun()

else:
    # Welcome screen
    st.info("👆 Upload an Excel file to get started")

    # Instructions
    with st.expander("📖 How to Use This Tool", expanded=True):
        st.markdown("""
        ### Required Excel Columns:
        
        | Column Name | Description | Required |
        |------------|-------------|----------|
        | **First Name** | Student's first name | ✅ Yes |
        | **Last Name** | Student's last name | ✅ Yes |
        | **Email** | Student's email address | ✅ Yes |
        | **Certificate Link** | SAS form URL | ✅ Yes |
        | **Certificate Name** | Name of certification | ⚪ Optional |
        | **Badge Opt-In** | Yes/No (defaults to Yes if empty) | ⚪ Optional |
        
        ### Features:
        - ✅ **Multi-browser support**: Works with Chrome, Firefox, or Edge
        - ✅ **Auto-detection**: Automatically finds available browser
        - ✅ **Real-time logs**: Monitor progress as it happens
        - ✅ **Error screenshots**: Captures errors for debugging
        - ✅ **Bulk download**: Get all results, logs, and screenshots in one ZIP file
        - ✅ **Resume capability**: Can stop and restart anytime
        
        ### Browser Priority (Auto Mode):
        1. **Chrome** (recommended)
        2. **Edge** 
        3. **Firefox**
        """)

    # System requirements
    with st.expander("💻 System Requirements"):
        st.markdown("""
        ### Required Software:
        - Python 3.8 or higher
        - One of: Chrome, Firefox, or Edge browser
        - Excel file with proper formatting
        
        ### Required Python Packages:
        ```bash
        pip install streamlit selenium openpyxl pandas webdriver-manager
        ```
        """)
