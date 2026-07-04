import streamlit as st
import threading
import queue
import time
import os
import shutil
import zipfile
import io
import re
import tempfile
from datetime import date, datetime
import pandas as pd

# Import the class
from sas_automation import SASFormAutomator
from sheet_link_mapper import (
    get_sheet_names,
    read_all_sheets,
    detect_columns,
    default_form_url_for_sheet,
)
from certificate_generator import (
    generate_all_certificates,
    format_certificate_date,
    find_libreoffice,
)

_DEFAULT_CERT_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "SAS Certificate Template.docx",
)


def _resolve_cert_template_bytes(uploaded_file):
    """Uploaded template wins; otherwise use bundled SAS Certificate Template.docx."""
    if uploaded_file is not None:
        return uploaded_file.getvalue(), uploaded_file.name
    if os.path.isfile(_DEFAULT_CERT_TEMPLATE_PATH):
        with open(_DEFAULT_CERT_TEMPLATE_PATH, "rb") as fh:
            return fh.read(), os.path.basename(_DEFAULT_CERT_TEMPLATE_PATH)
    return None, None

st.set_page_config(page_title="SAS Automation Pro",
                   layout="wide", page_icon="🤖")

# Title and description
st.title("🤖 SAS Form Automation Pro")
st.markdown("**Automated SAS form filling & certificate generation**")

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
    st.subheader("⚡ Performance")
    use_parallel = st.checkbox(
        "Enable Parallel Processing",
        value=False,
        help="Process multiple forms simultaneously (faster but uses more resources)"
    )
    
    if use_parallel:
        num_workers = st.slider(
            "Number of Parallel Browsers",
            min_value=2,
            max_value=10,
            value=3,
            help="More browsers = faster but uses more CPU/RAM. Recommended: 3-4"
        )
        st.info(f"⚡ Will use {num_workers} browsers in parallel")
    else:
        num_workers = 1
        st.info("🐌 Sequential processing (one browser at a time)")
    
    st.markdown("---")
    use_headless = st.checkbox(
        "Headless Mode (Faster)",
        value=False,
        help="Run browsers in background without GUI. Faster and uses less resources, but you won't see the browser windows."
    )
    if use_headless:
        st.info("🚀 Headless mode enabled - browsers will run in background")

    st.markdown("---")
    st.markdown("""
    ### 📋 Instructions
    1. Upload Excel file with required columns
    2. Click **Preview Final Data**
    3. **Automation tab:** run SAS form filling
    4. **Certificates tab:** generate PDF certificates
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
if 'sheet_links' not in st.session_state:
    st.session_state.sheet_links = {}
if 'preview_df' not in st.session_state:
    st.session_state.preview_df = None
if 'preview_read_warnings' not in st.session_state:
    st.session_state.preview_read_warnings = []
if 'certificates_zip' not in st.session_state:
    st.session_state.certificates_zip = None

# File upload (stable key keeps selection across reruns; only .xlsx / openpyxl)
uploaded_file = st.file_uploader(
    "📂 Upload Excel File (.xlsx)",
    type=["xlsx"],
    key="excel_workbook_upload",
    help="استخدم ملف Excel بصيغة .xlsx فقط. Use .xlsx format only.",
)
st.caption(
    "إذا الاختيار مش بيظهر بعد الرفع، جرّب تحديث الصفحة أو متصفح مختلف. "
    "If the file does not stick after choosing it, refresh or try another browser."
)

if uploaded_file:
    upload_id = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("upload_id") != upload_id:
        st.session_state.upload_id = upload_id
        st.session_state.upload_temp_dir = tempfile.mkdtemp()
        st.session_state.excel_path = os.path.join(
            st.session_state.upload_temp_dir, "data.xlsx"
        )
        for k in list(st.session_state.keys()):
            if isinstance(k, str) and k.startswith("sheet_link_input_"):
                del st.session_state[k]
        st.session_state.sheet_links = {}
        st.session_state.preview_df = None
        st.session_state.preview_read_warnings = []
        st.session_state.certificates_zip = None

    excel_path = st.session_state.excel_path
    temp_dir = st.session_state.upload_temp_dir

    # Always flush latest bytes to disk (fixes missing/stale file after reruns)
    try:
        with open(excel_path, "wb") as f:
            f.write(uploaded_file.getvalue())
    except OSError as e:
        st.error(f"❌ Could not save uploaded file: {e}")
        st.stop()

    try:
        sheet_names = get_sheet_names(excel_path)
        for n in sheet_names:
            st.session_state.sheet_links.setdefault(n, default_form_url_for_sheet(n))

        df_preview = pd.read_excel(
            excel_path, sheet_name=sheet_names[0], engine="openpyxl"
        )
    except Exception as e:
        st.error(
            "❌ تعذّر قراءة ملف Excel. تأكد أنه .xlsx وليس .xls، وأن الملف غير تالف أو محمي بكلمة مرور.\n\n"
            f"Could not read the workbook: {e}"
        )
        st.stop()

    # Show file info
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.metric("Sheets", len(sheet_names))
    with col_info2:
        st.metric("Columns (first sheet)", len(df_preview.columns))
    with col_info3:
        st.metric("File Size", f"{uploaded_file.size / 1024:.1f} KB")

    st.caption(f"Preview shows first sheet: **{sheet_names[0]}** ({len(df_preview)} rows).")

    # Data preview
    with st.expander("📊 Preview Data (First 10 rows)", expanded=True):
        st.dataframe(df_preview.head(10), use_container_width=True)

    col_map = detect_columns(
        pd.read_excel(excel_path, sheet_name=sheet_names[0], engine="openpyxl")
    )
    has_name = "english_name" in col_map or "first_name" in col_map
    has_email = "email" in col_map
    if not has_name or not has_email:
        st.warning(
            f"⚠️ Could not detect name/email columns in '{sheet_names[0]}'. "
            f"Detected keys: {list(col_map.keys())}. Will still show link inputs below."
        )

    # Same column order/labels as SASFormAutomator.read_excel / Excel template
    _AUTOMATION_PREVIEW_COLS = [
        "First Name",
        "Last Name",
        "Email",
        "Student ID",
        "Certificate Name",
        "Certificate Link",
    ]

    if st.button("🔍 Preview Final Data", key="btn_preview_final_data"):
        configured = {
            n: (st.session_state.sheet_links.get(n) or "").strip()
            for n in sheet_names
            if (st.session_state.sheet_links.get(n) or "").strip()
        }
        if not configured:
            st.warning("Add at least one non-empty form URL for a sheet to preview.")
        else:
            try:
                rows, read_warnings = read_all_sheets(excel_path, configured)
                st.session_state.preview_read_warnings = list(read_warnings)
                if rows:
                    raw_df = pd.DataFrame(rows)
                    st.session_state.preview_df = pd.DataFrame(
                        {
                            "First Name": raw_df["firstName"],
                            "Last Name": raw_df["lastName"],
                            "Email": raw_df["email"],
                            "Student ID": raw_df["studentId"],
                            "Certificate Name": raw_df["certificationName"],
                            "Certificate Link": raw_df["certificationLink"],
                        }
                    )[_AUTOMATION_PREVIEW_COLS]
                else:
                    st.session_state.preview_df = pd.DataFrame(
                        columns=_AUTOMATION_PREVIEW_COLS
                    )
                    first_sheet = next(iter(configured.keys()))
                    detected_cols = pd.read_excel(
                        excel_path, sheet_name=first_sheet, engine="openpyxl"
                    ).columns.tolist()
                    st.error(
                        f"Preview loaded **0 rows**. Columns detected in sheet "
                        f"**{first_sheet!r}**: `{detected_cols}`. "
                        f"Expected headers like **English Name**, **Personal Email** / **Academic Email**. "
                        f"See terminal/logs for `col_map` debug output."
                    )
                    if read_warnings:
                        st.warning("Reader messages:\n\n" + "\n\n".join(read_warnings))
            except Exception as e:
                st.error(f"Could not build preview: {e}")

    # Drain automation queues on every rerun (works across tabs)
    if hasattr(st.session_state, 'log_queue'):
        while not st.session_state.log_queue.empty():
            try:
                msg = st.session_state.log_queue.get_nowait()
                st.session_state.logs.append(msg)
            except Exception:
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
            except Exception:
                break

    tab_automation, tab_certificates = st.tabs(["🤖 Automation", "🎓 Certificates"])

    with tab_automation:
        st.subheader("🔗 Form URL per sheet")
        st.caption(
            "Default SAS form URLs are pre-filled for known sheets. Edit any URL before running. "
            "Sheets with an empty URL are skipped."
        )
        for i, name in enumerate(sheet_names):
            key = f"sheet_link_input_{i}"
            st.session_state.sheet_links[name] = st.text_input(
                f"Form URL — `{name}`",
                value=st.session_state.sheet_links.get(name, ""),
                key=key,
            )

        with st.expander("📋 Preview Final Data Before Start", expanded=True):
            if st.session_state.preview_df is None:
                st.caption(
                    "Click **🔍 Preview Final Data** above to see merged rows "
                    "as they will be sent to the forms."
                )
            else:
                pdf = st.session_state.preview_df
                total = len(pdf)
                grouped = None
                if total > 0:
                    grouped = {
                        str(name): grp[_AUTOMATION_PREVIEW_COLS].copy()
                        for name, grp in pdf.groupby("Certificate Name", sort=True)
                    }

                if total == 0:
                    st.metric("Total rows", 0)
                    st.caption("No rows — check that linked sheets have the required columns.")
                    if st.session_state.get("preview_read_warnings"):
                        for w in st.session_state.preview_read_warnings:
                            st.caption(w)

                missing_n = 0
                if total > 0 and "Email" in pdf.columns:
                    missing_n = int((pdf["Email"] == "noemail@example.com").sum())
                if missing_n:
                    st.markdown(
                        f"⚠️ **Missing email:** {missing_n} row(s) use `noemail@example.com`"
                    )
                elif total > 0:
                    st.markdown("No rows flagged with placeholder email.")

                preview_xlsx_buf = io.BytesIO()
                with pd.ExcelWriter(preview_xlsx_buf, engine="openpyxl") as writer:
                    if grouped:
                        used_excel = set()

                        def _writer_tab_name(raw_name: str) -> str:
                            base = re.sub(
                                r"[\[\]:*?/\\]", "_", str(raw_name).strip()
                            )[:31] or "Sheet"
                            wname = base
                            n = 1
                            while wname in used_excel:
                                suffix = f"_{n}"
                                wname = (base[: max(1, 31 - len(suffix))] + suffix)[:31]
                                n += 1
                            used_excel.add(wname)
                            return wname

                        for sheet_name, sheet_df in grouped.items():
                            sheet_df.to_excel(
                                writer,
                                sheet_name=_writer_tab_name(sheet_name),
                                index=False,
                            )
                    else:
                        pd.DataFrame(columns=_AUTOMATION_PREVIEW_COLS).to_excel(
                            writer, sheet_name="Preview", index=False
                        )
                preview_xlsx_buf.seek(0)
                st.download_button(
                    label="📥 Download Preview as Excel",
                    data=preview_xlsx_buf.getvalue(),
                    file_name=f"SAS_Preview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_preview_xlsx",
                )

                if grouped:
                    per = pdf["Certificate Name"].value_counts().sort_index()
                    summary_cols = st.columns(1 + len(per))
                    with summary_cols[0]:
                        st.metric("Total", total)
                    for idx, (sn, cnt) in enumerate(per.items(), start=1):
                        with summary_cols[idx]:
                            st.metric(str(sn), int(cnt))

                    sheet_tab_labels = list(grouped.keys())
                    sheet_tabs = st.tabs(sheet_tab_labels)

                    def _yellow_missing(row):
                        if row["Email"] == "noemail@example.com":
                            return ["background-color: #fff3cd"] * len(row)
                        return [""] * len(row)

                    for sheet_tab, sheet_name in zip(sheet_tabs, sheet_tab_labels):
                        with sheet_tab:
                            sheet_df = grouped[sheet_name]
                            st.metric("Students", len(sheet_df))
                            styled = sheet_df.style.apply(_yellow_missing, axis=1)
                            st.dataframe(
                                styled, use_container_width=True, height=400
                            )

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
                sheet_link_snapshot = {
                    n: (st.session_state.sheet_links.get(n) or "").strip()
                    for n in sheet_names
                    if (st.session_state.sheet_links.get(n) or "").strip()
                }
                if not sheet_link_snapshot:
                    st.error("❌ Add at least one form URL for a sheet before starting.")
                    st.stop()

                st.session_state.running = True
                st.session_state.stop_flag = False
                st.session_state.logs = ["🚀 Starting automation..."]
                st.session_state.results = []
                st.session_state.browser_used = None
                st.session_state.temp_dir = temp_dir

                # Create queues for thread communication
                log_queue = queue.Queue()
                result_queue = queue.Queue()

                excel_for_run = excel_path
                snapshot_for_run = dict(sheet_link_snapshot)

                def run_automation():
                    try:
                        # Create checkpoint directory in temp folder
                        checkpoint_dir = os.path.join(temp_dir, "checkpoints")
                        automator = SASFormAutomator(
                            "", excel_for_run,
                            browser_choice=browser_choice,
                            checkpoint_dir=checkpoint_dir,
                            restart_browser_interval=100,  # Restart browser every 100 forms
                            headless=use_headless
                        )

                        # Store which browser was used
                        result_queue.put(('browser', automator.browser_name))

                        all_students, sheet_read_warnings = read_all_sheets(
                            excel_for_run, snapshot_for_run
                        )
                        for w in sheet_read_warnings:
                            log_queue.put(w)
                        students = all_students
                        total_students = len(students)
                    
                        # Stop flag for parallel processing
                        stop_flag = threading.Event()
                    
                        def log_callback(message):
                            log_queue.put(f"[{time.strftime('%H:%M:%S')}] {message}")
                    
                        def result_callback(result):
                            result_queue.put(('result', result))
                    
                        if use_parallel and num_workers >= 1:
                            # Parallel processing
                            log_queue.put(f"📊 Total students to process: {total_students}")
                            log_queue.put(f"⚡ Parallel mode: {num_workers} browser(s) working simultaneously")
                            log_queue.put(f"💾 Checkpoint will be saved every 50 forms")
                            log_queue.put(f"⏱️ Estimated time: ~{total_students * 8 // num_workers // 60} minutes (vs ~{total_students * 10 // 60} minutes sequential)")
                        
                            # Store stop flag
                            st.session_state.stop_flag_event = stop_flag
                        
                            results = automator.process_students_parallel(
                                students,
                                num_workers=num_workers,
                                log_callback=log_callback,
                                result_callback=result_callback,
                                stop_flag=stop_flag,
                                headless=use_headless
                            )
                        
                            # Put all results
                            for result in results:
                                result_queue.put(('result', result))
                        else:
                            # Sequential processing (original method)
                            log_queue.put(f"📊 Total students to process: {total_students}")
                            log_queue.put(f"💾 Checkpoint will be saved every 50 forms")
                            log_queue.put(f"🔄 Browser will restart every 100 forms to prevent crashes")
                            log_queue.put(f"⏱️ Estimated time: ~{total_students * 10 // 60} minutes")

                            for i, student in enumerate(students, 1):
                                # Check if user pressed Stop
                                try:
                                    if not result_queue.empty():
                                        stop_signal = result_queue.get_nowait()
                                        if stop_signal == "STOP":
                                            log_queue.put(
                                                "⏸️ Automation stopped by user!")
                                            # Save checkpoint before stopping
                                            automator.save_checkpoint(i-1, total_students)
                                            log_queue.put(f"💾 Progress saved: {i-1}/{total_students} processed")
                                            return
                                except:
                                    pass

                                log_queue.put(
                                    f"[{time.strftime('%H:%M:%S')}] 🔄 Processing {i}/{total_students}: "
                                    f"[{student.get('certificationName', '')}] {student['firstName']} {student['lastName']}"
                                )

                                # Restart browser if needed (for large datasets)
                                automator.restart_browser_if_needed()

                                result = automator.fill_form(student)
                                result_queue.put(('result', result))
                            
                                # Add result to automator's results list for CSV saving
                                automator.results.append(result)

                                status_emoji = "✅" if result['status'] == "Success" else "❌"
                                log_queue.put(
                                    f"{status_emoji} {result['status']}: {student['email']}")

                                # Save checkpoint every 50 forms (for large datasets)
                                if i % 50 == 0:
                                    automator.save_checkpoint(i, total_students)
                                    log_queue.put(f"💾 Checkpoint saved: {i}/{total_students} processed")

                            # Final save - clear any existing results file first to avoid duplicates
                            if os.path.exists(automator.results_file):
                                os.remove(automator.results_file)
                            automator.save_checkpoint(total_students, total_students)
                            automator.save_results()
                    
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
                if hasattr(st.session_state, 'stop_flag_event'):
                    st.session_state.stop_flag_event.set()
                st.warning("⏸️ Stopping automation...")
                time.sleep(1)
                st.session_state.running = False
                st.rerun()

        with col3:
            if st.session_state.running:
                st.markdown("### 🔄")
                st.markdown("**Running**")

        # Show browser being used
        if st.session_state.browser_used:
            st.info(f"🌐 Using browser: **{st.session_state.browser_used}**")

        # Display logs
        if st.session_state.logs:
            st.markdown("---")
            st.subheader("📋 Live Logs")
            logs_container = st.container()
            with logs_container:
                # Show last 100 logs (increased for large datasets)
                display_logs = st.session_state.logs[-100:]
                st.code("\n".join(display_logs), language=None)
            
                # For large datasets, show memory optimization info
                if len(st.session_state.logs) > 500:
                    st.info("💡 **Memory Optimization**: Logs are limited to last 100 entries. Full logs are saved in checkpoint files.")

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

            if "certificationName" in results_df.columns:
                st.subheader("📑 Results by Sheet")
                by_sheet = (
                    results_df.groupby("certificationName", dropna=False)
                    .agg(
                        total=("status", "count"),
                        success=("status", lambda s: int((s == "Success").sum())),
                    )
                    .reset_index()
                )
                by_sheet["failed"] = by_sheet["total"] - by_sheet["success"]
                st.dataframe(by_sheet, use_container_width=True)

            # Create ZIP file
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                # Add logs
                zf.writestr("logs.txt", "\n".join(st.session_state.logs))

                # Add results CSV
                zf.writestr("results.csv", results_df.to_csv(
                    index=False, encoding='utf-8-sig'))
            
                # Add checkpoint files if they exist
                if hasattr(st.session_state, 'temp_dir') and os.path.exists(st.session_state.temp_dir):
                    checkpoint_dir = os.path.join(st.session_state.temp_dir, "checkpoints")
                    if os.path.exists(checkpoint_dir):
                        # Add checkpoint JSON
                        checkpoint_file = os.path.join(checkpoint_dir, "progress.json")
                        if os.path.exists(checkpoint_file):
                            zf.write(checkpoint_file, "checkpoints/progress.json")
                    
                        # Add incremental results CSV
                        results_file = os.path.join(checkpoint_dir, "results.csv")
                        if os.path.exists(results_file):
                            zf.write(results_file, "checkpoints/results.csv")

                        # Add screenshots if any
                        screenshot_count = 0
                        for file in os.listdir(checkpoint_dir):
                            if file.endswith(".png"):
                                file_path = os.path.join(checkpoint_dir, file)
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
                    for path in {
                        st.session_state.get("temp_dir"),
                        st.session_state.get("upload_temp_dir"),
                    }:
                        if path and os.path.exists(path):
                            try:
                                shutil.rmtree(path)
                            except Exception:
                                pass
                    for key in list(st.session_state.keys()):
                        if isinstance(key, str) and key.startswith("sheet_link_input_"):
                            del st.session_state[key]
                    for key in [
                        'logs', 'results', 'running', 'stop_flag', 'log_queue',
                        'result_queue', 'temp_dir', 'browser_used', 'sheet_links',
                        'upload_id', 'excel_path', 'upload_temp_dir', 'stop_flag_event',
                        'preview_df', 'preview_read_warnings', 'certificates_zip',
                    ]:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.success("✅ Session cleared! Upload a new file to start")
                    st.rerun()

    with tab_certificates:
        if (
            st.session_state.preview_df is None
            or len(st.session_state.preview_df) == 0
        ):
            st.info(
                "Click **🔍 Preview Final Data** above first — "
                "certificates are built from the preview rows."
            )
        else:
            pdf = st.session_state.preview_df
            cert_total = len(pdf)
            with_id = int(pdf["Student ID"].astype(str).str.strip().ne("").sum())
            st.metric("Rows ready for certificates", with_id)
            missing_id_n = cert_total - with_id
            if missing_id_n:
                st.warning(
                    f"⚠️ **Missing Student ID:** {missing_id_n} row(s) will be "
                    "excluded from the ZIP."
                )

            st.caption(
                "Default template: **SAS Certificate Template.docx** "
                "(placeholders `{Name}`, `{Certificate Name}`, `{Date}`). "
                "Upload a different `.docx` below to override. "
                "Requires LibreOffice for PDF export."
            )

            default_template_ok = os.path.isfile(_DEFAULT_CERT_TEMPLATE_PATH)
            if default_template_ok:
                st.info(f"📄 Using default template: `{os.path.basename(_DEFAULT_CERT_TEMPLATE_PATH)}`")
            else:
                st.warning(
                    "Default template file not found in the app folder. "
                    "Upload a `.docx` template to generate certificates."
                )

            cert_template = st.file_uploader(
                "📄 Custom template (.docx) — optional",
                type=["docx"],
                key="cert_template_upload",
                help="Leave empty to use SAS Certificate Template.docx",
            )

            if cert_template is not None:
                st.caption(f"Override active: **{cert_template.name}**")

            template_bytes, template_name = _resolve_cert_template_bytes(cert_template)

            cert_col1, cert_col2 = st.columns(2)
            with cert_col1:
                cert_date = st.date_input(
                    "Certificate date",
                    value=date.today(),
                    key="cert_date_input",
                )
            with cert_col2:
                date_format = st.selectbox(
                    "Date format",
                    options=["us", "british", "iso"],
                    index=0,
                    format_func=lambda x: {
                        "us": "US (July 4, 2026)",
                        "british": "British (4 July 2026)",
                        "iso": "ISO (2026-07-04)",
                    }[x],
                    key="cert_date_format",
                )

            lo_path = find_libreoffice()
            if not lo_path:
                st.warning(
                    "LibreOffice not found. Install LibreOffice "
                    "(Mac: /Applications/LibreOffice.app, Windows: Program Files) "
                    "before generating PDFs."
                )

            generate_certs = st.button(
                "🎓 Generate Certificates",
                key="btn_generate_certificates",
                disabled=template_bytes is None or not lo_path or with_id == 0,
            )

            if generate_certs and template_bytes is not None:
                old_zip = st.session_state.certificates_zip
                if old_zip and os.path.isfile(old_zip):
                    try:
                        os.remove(old_zip)
                    except OSError:
                        pass
                st.session_state.certificates_zip = None

                date_str = format_certificate_date(cert_date, date_format)
                cert_out_dir = os.path.join(temp_dir, "certificates")
                os.makedirs(cert_out_dir, exist_ok=True)

                with st.status(
                    "Generating certificates…", expanded=True
                ) as cert_status:
                    progress_bar = st.progress(0.0)

                    def _cert_progress(current, total, message):
                        frac = (current / total) if total else 0.0
                        progress_bar.progress(min(frac, 1.0))
                        cert_status.update(label=message)

                    try:
                        zip_path, cert_warnings = generate_all_certificates(
                            st.session_state.preview_df,
                            template_bytes,
                            date_str,
                            cert_out_dir,
                            progress_callback=_cert_progress,
                        )
                        st.session_state.certificates_zip = zip_path
                        cert_status.update(
                            label=f"Done — {os.path.basename(zip_path)}",
                            state="complete",
                        )
                        for w in cert_warnings:
                            st.warning(w)
                        st.success(
                            f"Generated certificates ZIP: **{os.path.basename(zip_path)}** "
                            f"(template: {template_name})"
                        )
                    except Exception as e:
                        cert_status.update(
                            label="Certificate generation failed",
                            state="error",
                        )
                        st.error(f"Could not generate certificates: {e}")

            if (
                st.session_state.certificates_zip
                and os.path.isfile(st.session_state.certificates_zip)
            ):
                zip_path = st.session_state.certificates_zip
                with open(zip_path, "rb") as zf:
                    zip_bytes = zf.read()
                st.download_button(
                    label="📥 Download Certificates (ZIP)",
                    data=zip_bytes,
                    file_name=os.path.basename(zip_path),
                    mime="application/zip",
                    key="download_certificates_zip",
                )

    # Auto-refresh while automation runs (works even on Certificates tab)
    if st.session_state.running:
        time.sleep(3)
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
        | **Form URL** | Pasted per sheet in the app (not an Excel column) | ✅ Yes |
        | **Certificate Name** | Optional column per sheet | ⚪ Optional |
        | **Badge Opt-In** | Yes/No (defaults to Yes if empty) | ⚪ Optional |
        
        ### Features:
        - ✅ **Multi-browser support**: Works with Chrome, Firefox, or Edge
        - ✅ **Auto-detection**: Automatically finds available browser
        - ✅ **Real-time logs**: Monitor progress as it happens
        - ✅ **Error screenshots**: Captures errors for debugging
        - ✅ **Bulk download**: Get all results, logs, and screenshots in one ZIP file
        - ✅ **Resume capability**: Can stop and restart anytime
        - ✅ **Multi-sheet workbooks**: One form URL per sheet; rows use that sheet’s URL
        
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
