# -*- coding: utf-8 -*-
"""Generate personalized PDF certificates from a Word template and package into a ZIP."""

from __future__ import annotations

import io
import os
import platform
import re
import shutil
import subprocess
import tempfile
import uuid
import zipfile
from datetime import date, datetime
from typing import Callable, Optional
from xml.etree import ElementTree as ET

import pandas as pd

from sheet_link_mapper import _build_display_name

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W_TAG_P = f"{{{W_NS}}}p"
W_TAG_T = f"{{{W_NS}}}t"
MC_TAG_FALLBACK = f"{{{MC_NS}}}Fallback"

_ILLEGAL_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')
_PLACEHOLDER_PATTERN = re.compile(
    r"\{Certificate Name\}|\{Name\}|\{Date\}"
)

ProgressCallback = Optional[Callable[[int, int, str], None]]


def format_certificate_date(d: date, fmt: str) -> str:
    """Format a date for certificate placeholders."""
    if fmt == "british":
        return d.strftime("%d %B %Y")
    if fmt == "iso":
        return d.strftime("%Y-%m-%d")
    return d.strftime("%B %d, %Y")


def sanitize_filename_part(value: str) -> str:
    """Remove illegal path characters and collapse whitespace."""
    text = _ILLEGAL_FILENAME_CHARS.sub("", str(value))
    return re.sub(r"\s+", " ", text).strip()


def find_libreoffice() -> Optional[str]:
    """Locate LibreOffice soffice binary on Mac, Windows, or PATH."""
    system = platform.system()

    if system == "Darwin":
        candidates = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            shutil.which("soffice"),
            shutil.which("libreoffice"),
        ]
    elif system == "Windows":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            shutil.which("soffice"),
            shutil.which("libreoffice"),
        ]
    else:
        candidates = [
            shutil.which("soffice"),
            shutil.which("libreoffice"),
            "/usr/bin/libreoffice",
            "/usr/local/bin/libreoffice",
        ]

    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def _apply_placeholders(text: str, name: str, cert_name: str, date_str: str) -> str:
    """Replace placeholders; longest match first via alternation regex."""

    def _repl(match: re.Match) -> str:
        token = match.group(0)
        if token == "{Certificate Name}":
            return cert_name
        if token == "{Name}":
            return name
        if token == "{Date}":
            return date_str
        return token

    return _PLACEHOLDER_PATTERN.sub(_repl, text)


def _build_parent_map(root: ET.Element) -> dict:
    parent_map = {}
    for parent in root.iter():
        for child in parent:
            parent_map[child] = parent
    return parent_map


def _is_inside_mc_fallback(element: ET.Element, parent_map: dict) -> bool:
    """True if element lives under mc:Fallback (legacy duplicate of mc:Choice)."""
    node = element
    while node in parent_map:
        node = parent_map[node]
        if node.tag == MC_TAG_FALLBACK:
            return True
    return False


def _collapse_duplicate_placeholders(text: str) -> str:
    """Word sometimes stores '{Name}{Name}' as split runs in one paragraph."""
    for token in ("{Certificate Name}", "{Name}", "{Date}"):
        doubled = token + token
        while doubled in text:
            text = text.replace(doubled, token)
    return text


def _replace_placeholders_in_xml(xml_bytes: bytes, name: str, cert_name: str, date_str: str) -> bytes:
    """Merge w:t runs per paragraph and replace placeholders in document XML."""
    root = ET.fromstring(xml_bytes)
    parent_map = _build_parent_map(root)
    changed = False

    for paragraph in root.iter(W_TAG_P):
        if _is_inside_mc_fallback(paragraph, parent_map):
            continue

        t_nodes = list(paragraph.iter(W_TAG_T))
        if not t_nodes:
            continue
        full_text = _collapse_duplicate_placeholders(
            "".join(node.text or "" for node in t_nodes)
        )
        new_text = _apply_placeholders(full_text, name, cert_name, date_str)
        if new_text == full_text:
            continue
        t_nodes[0].text = new_text
        for node in t_nodes[1:]:
            node.text = ""
        changed = True

    if not changed:
        return xml_bytes
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _replace_placeholders_in_docx(
    template_docx_bytes: bytes,
    name: str,
    cert_name: str,
    date_str: str,
) -> bytes:
    """Fill placeholders in all Word XML parts that may contain body text."""
    xml_parts = (
        "word/document.xml",
    )
    # Include headers/footers when present (discovered while reading the zip).
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, "template.docx")
        with open(docx_path, "wb") as fh:
            fh.write(template_docx_bytes)

        with zipfile.ZipFile(docx_path, "r") as zin:
            part_names = zin.namelist()
            header_footer = [
                n
                for n in part_names
                if n.startswith("word/header") or n.startswith("word/footer")
            ]
            targets = list(xml_parts) + header_footer

            out_buf = io.BytesIO()
            with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in part_names:
                    data = zin.read(item)
                    if item in targets:
                        data = _replace_placeholders_in_xml(data, name, cert_name, date_str)
                    zout.writestr(item, data)
            return out_buf.getvalue()


def _convert_docx_batch_to_pdf(docx_paths: list[str], pdf_dir: str, lo_binary: str) -> None:
    """Convert multiple DOCX files in one LibreOffice invocation."""
    if not docx_paths:
        return

    profile_dir = os.path.join(tempfile.gettempdir(), f"lo_{uuid.uuid4().hex}")
    os.makedirs(profile_dir, exist_ok=True)
    profile_uri = f"file://{profile_dir}"
    if platform.system() == "Windows":
        profile_uri = "file:///" + profile_dir.replace("\\", "/")

    try:
        cmd = [
            lo_binary,
            "--headless",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to",
            "pdf",
            "--outdir",
            pdf_dir,
            *docx_paths,
        ]
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=max(120, 30 * len(docx_paths)),
        )
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(
                f"LibreOffice PDF conversion failed (exit {result.returncode}): {detail}"
            )
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def _build_pdf_filename(
    student_id: str,
    display_name: str,
    cert_name: str,
    seen: set[str],
) -> str:
    """Build a unique sanitized PDF filename inside a certificate folder."""
    base = sanitize_filename_part(
        f"{student_id} {display_name} {cert_name}"
    )
    filename = f"{base}.pdf"
    if filename not in seen:
        seen.add(filename)
        return filename

    n = 1
    while True:
        candidate = f"{base} ({n}).pdf"
        if candidate not in seen:
            seen.add(candidate)
            return candidate
        n += 1


def _display_name_from_row(row: pd.Series) -> str:
    first = str(row.get("First Name", "")).strip()
    last = str(row.get("Last Name", "")).strip()
    return _build_display_name(first, last)


def render_certificate(
    template_docx_path: str,
    name: str,
    cert_name: str,
    date_str: str,
    output_pdf_path: str,
    lo_binary: Optional[str] = None,
) -> None:
    """Fill template placeholders and write a single PDF."""
    lo = lo_binary or find_libreoffice()
    if not lo:
        raise RuntimeError(
            "LibreOffice not found. Install LibreOffice (Mac .app, Windows Program Files, or PATH)."
        )

    with open(template_docx_path, "rb") as fh:
        template_bytes = fh.read()

    filled_docx = _replace_placeholders_in_docx(template_bytes, name, cert_name, date_str)

    out_dir = os.path.dirname(output_pdf_path) or "."
    os.makedirs(out_dir, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, "cert.docx")
        with open(docx_path, "wb") as fh:
            fh.write(filled_docx)

        pdf_dir = os.path.join(tmp, "pdf")
        os.makedirs(pdf_dir, exist_ok=True)
        _convert_docx_batch_to_pdf([docx_path], pdf_dir, lo)

        generated = os.path.join(pdf_dir, "cert.pdf")
        if not os.path.isfile(generated):
            raise RuntimeError("LibreOffice did not produce the expected PDF output.")
        shutil.move(generated, output_pdf_path)


def generate_all_certificates(
    preview_df: pd.DataFrame,
    template_docx_bytes: bytes,
    date_str: str,
    output_dir: str,
    progress_callback: ProgressCallback = None,
) -> tuple[str, list[str]]:
    """
    Generate PDF certificates for all preview rows with a Student ID.

    Returns:
        (zip_path, warnings) — ZIP written to disk under output_dir.
    """
    warnings: list[str] = []
    lo_binary = find_libreoffice()
    if not lo_binary:
        raise RuntimeError(
            "LibreOffice not found. Install LibreOffice (Mac .app, Windows Program Files, or PATH)."
        )

    required_cols = {"First Name", "Last Name", "Student ID", "Certificate Name"}
    missing = required_cols - set(preview_df.columns)
    if missing:
        raise ValueError(f"preview_df missing required columns: {sorted(missing)}")

    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = os.path.join(output_dir, f"SAS_Certificates_{timestamp}.zip")

    eligible_rows = []
    skipped_missing_id = 0
    for _, row in preview_df.iterrows():
        student_id = str(row.get("Student ID", "")).strip()
        if not student_id:
            skipped_missing_id += 1
            continue
        eligible_rows.append(row)

    if skipped_missing_id:
        warnings.append(
            f"Skipped {skipped_missing_id} row(s) with missing Student ID."
        )

    if not eligible_rows:
        raise ValueError("No rows with a Student ID to generate certificates for.")

    total = len(eligible_rows)

    with tempfile.TemporaryDirectory() as work_dir:
        docx_dir = os.path.join(work_dir, "docx")
        pdf_dir = os.path.join(work_dir, "pdf")
        os.makedirs(docx_dir, exist_ok=True)
        os.makedirs(pdf_dir, exist_ok=True)

        # Map docx stem -> metadata for zip layout after batch conversion.
        job_meta: list[dict] = []
        seen_names: dict[str, set[str]] = {}

        for idx, row in enumerate(eligible_rows, start=1):
            cert_name = str(row["Certificate Name"]).strip()
            display_name = _display_name_from_row(row)
            student_id = str(row["Student ID"]).strip()
            folder_key = sanitize_filename_part(cert_name) or "Certificates"
            seen_names.setdefault(folder_key, set())
            pdf_name = _build_pdf_filename(
                student_id, display_name, cert_name, seen_names[folder_key]
            )
            stem = f"{idx:05d}_{uuid.uuid4().hex[:8]}"
            docx_path = os.path.join(docx_dir, f"{stem}.docx")

            filled = _replace_placeholders_in_docx(
                template_docx_bytes,
                display_name,
                cert_name,
                date_str,
            )
            with open(docx_path, "wb") as fh:
                fh.write(filled)

            job_meta.append(
                {
                    "stem": stem,
                    "folder": folder_key,
                    "pdf_name": pdf_name,
                    "label": f"{student_id} {display_name}",
                }
            )

            if progress_callback:
                progress_callback(
                    idx,
                    total,
                    f"Prepared {idx}/{total}: {student_id} {display_name}",
                )

        docx_paths = [
            os.path.join(docx_dir, f"{job['stem']}.docx") for job in job_meta
        ]
        if progress_callback:
            progress_callback(total, total, "Converting DOCX files to PDF via LibreOffice…")
        _convert_docx_batch_to_pdf(docx_paths, pdf_dir, lo_binary)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for job in job_meta:
                pdf_src = os.path.join(pdf_dir, f"{job['stem']}.pdf")
                if not os.path.isfile(pdf_src):
                    warnings.append(
                        f"Missing PDF for {job['label']} — LibreOffice may have failed on this file."
                    )
                    continue
                arcname = f"{job['folder']}/{job['pdf_name']}"
                zf.write(pdf_src, arcname=arcname)

    if progress_callback:
        progress_callback(total, total, f"ZIP saved: {os.path.basename(zip_path)}")

    return zip_path, warnings
