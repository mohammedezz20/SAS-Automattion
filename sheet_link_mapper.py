# -*- coding: utf-8 -*-
"""Multi-sheet Excel helpers: sheet discovery and row loading with per-sheet form URLs."""

import pandas as pd


def _build_column_map(df):
    """Map logical fields to actual column names (same rules as SASFormAutomator.read_excel)."""
    col_map = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if "first name" in col_lower or col_lower == "firstname":
            col_map["firstName"] = col
        elif "last name" in col_lower or col_lower == "lastname":
            col_map["lastName"] = col
        elif "email" in col_lower:
            col_map["email"] = col
        elif "certificate name" in col_lower or "cert name" in col_lower:
            col_map["certificationName"] = col
        elif (
            "certificate link" in col_lower
            or "cert link" in col_lower
            or "link" in col_lower
        ):
            col_map["certificationLink"] = col
        elif "badge opt" in col_lower or "badgeopt" in col_lower:
            col_map["badgeOptIn"] = col
    return col_map


def dataframe_has_student_columns(df):
    """True if df has mappable First Name, Last Name, and Email columns."""
    m = _build_column_map(df)
    return bool(m.get("firstName") and m.get("lastName") and m.get("email"))


def detect_columns(df):
    """
    Auto-detect name/email columns for UI hints and read_sheet_students.
    Keys may include: english_name, first_name, last_name, email, badge.
    Prefers Personal Email, then any other email column (e.g. Academic Email).
    """
    col_map = {}
    columns = list(df.columns)

    for col in columns:
        col_lower = str(col).strip().lower()
        if "english" in col_lower and "name" in col_lower:
            col_map["english_name"] = col
            break
        if col_lower in ("english name", "englishname", "name_en", "name en"):
            col_map["english_name"] = col
            break

    if "english_name" not in col_map:
        for col in columns:
            col_lower = str(col).strip().lower()
            if "first" in col_lower and "name" in col_lower:
                col_map["first_name"] = col
            if "last" in col_lower and "name" in col_lower:
                col_map["last_name"] = col

    for col in columns:
        col_lower = str(col).strip().lower()
        if "personal" in col_lower and "email" in col_lower:
            col_map["email"] = col
            break
    if "email" not in col_map:
        for col in columns:
            col_lower = str(col).strip().lower()
            if "email" in col_lower:
                col_map["email"] = col
                break

    for col in columns:
        col_lower = str(col).strip().lower()
        if "badge" in col_lower or ("opt" in col_lower and "badge" in col_lower):
            col_map["badge"] = col
            break

    print(f"col_map: {col_map}")
    return col_map


def split_english_name(full_name: str):
    """First token = first name, remainder = last name."""
    full_name = str(full_name).strip()
    parts = full_name.split()
    if not parts:
        return ("Unknown", "User")
    first = parts[0]
    last = " ".join(parts[1:]) if len(parts) > 1 else "."
    return (first, last)


def read_sheet_students(excel_file_path: str, sheet_name: str, form_link: str):
    """
    Read one sheet into student dicts for SASFormAutomator.
    Supports English Name + Personal/Academic Email layout (and First/Last + Email).
    Returns (students, warning_or_none). warning is set when no rows are produced
    or required columns cannot be detected.
    """
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine="openpyxl")
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    print(f"Sheet '{sheet_name}' columns: {list(df.columns)}")
    col_map = detect_columns(df)

    has_name = "english_name" in col_map or "first_name" in col_map
    has_email = "email" in col_map

    if not col_map or (not has_name and not has_email):
        msg = (
            f"Sheet '{sheet_name}': cannot detect name or email columns. "
            f"col_map keys={list(col_map.keys())}, columns={list(df.columns)}"
        )
        print(f"WARNING: {msg}")
        return [], msg

    if not has_name:
        msg = (
            f"Sheet '{sheet_name}': no English Name or First Name column detected. "
            f"col_map keys={list(col_map.keys())}, columns={list(df.columns)}"
        )
        print(f"WARNING: {msg}")
        return [], msg

    if not has_email:
        msg = (
            f"Sheet '{sheet_name}': no Email column detected. "
            f"col_map keys={list(col_map.keys())}, columns={list(df.columns)}"
        )
        print(f"WARNING: {msg}")
        return [], msg

    students = []
    for row in df.to_dict("records"):
        if pd.Series(row).isna().all():
            continue

        if "english_name" in col_map:
            raw_name = str(row.get(col_map["english_name"], "")).strip()
            if not raw_name or raw_name.lower() in ("nan", "none", ""):
                continue
            first_name, last_name = split_english_name(raw_name)
        else:
            first_name = str(row.get(col_map["first_name"], "Unknown")).strip()
            last_col = col_map.get("last_name")
            last_name = (
                str(row.get(last_col, "User")).strip()
                if last_col
                else "User"
            )

        email_raw = row.get(col_map["email"], "")
        email = (
            str(email_raw).strip()
            if email_raw is not None and not (isinstance(email_raw, float) and pd.isna(email_raw))
            else ""
        )
        if not email or email.lower() in ("nan", "none", ""):
            email = "noemail@example.com"
        if "@" not in email:
            email = "noemail@example.com"

        badge_raw = row.get(col_map.get("badge"), None)
        if badge_raw is None or str(badge_raw).strip().lower() in ("nan", "none", ""):
            badge_final = "yes"
        else:
            badge_clean = str(badge_raw).strip().lower()
            badge_final = "yes" if badge_clean in ("yes", "y", "1", "true", "ok") else "no"

        students.append(
            {
                "firstName": first_name,
                "lastName": last_name,
                "email": email,
                "certificationName": sheet_name,
                "certificationLink": form_link,
                "badgeOptIn": badge_final,
            }
        )

    if not students:
        msg = (
            f"Sheet '{sheet_name}': 0 student rows after parsing (empty names or all rows blank?). "
            f"columns={list(df.columns)}, col_map keys={list(col_map.keys())}"
        )
        print(f"WARNING: {msg}")
        return [], msg

    return students, None


def get_sheet_names(excel_path):
    """Return list of sheet names in the workbook."""
    xl = pd.ExcelFile(excel_path, engine="openpyxl")
    return xl.sheet_names


def read_all_sheets(excel_path, sheet_link_snapshot):
    """
    Read students from each sheet that has a form URL in sheet_link_snapshot.

    Returns:
        (all_students: list[dict], warnings: list[str])
    """
    all_students = []
    warnings = []

    for sheet_name, form_url_raw in sheet_link_snapshot.items():
        form_url = (form_url_raw or "").strip()
        if not form_url:
            continue

        students, sheet_warn = read_sheet_students(excel_path, sheet_name, form_url)

        if sheet_warn:
            warnings.append(sheet_warn)

        if not students:
            df_dbg = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            cols = [str(c).strip() for c in df_dbg.columns.tolist()]
            line = (
                f"WARNING: Sheet '{sheet_name}' returned 0 students. "
                f"Columns found: {cols}"
            )
            print(line)
            warnings.append(line)
        else:
            all_students.extend(students)

    return all_students, warnings
