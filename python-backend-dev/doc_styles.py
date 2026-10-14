"""
doc_styles.py — Shared DOCX formatting pipeline.

Contains all document styling functions that are reusable across both
Design Document and Final Config Document generators:
- LibreOffice-based conversions (PPTX→PDF, TOC update)
- Pandoc MD→DOCX conversion
- Enterprise table formatting (borders, widths, banding)
- Typography enforcement (fonts, captions, page breaks)
- Image centering and resizing
- Header/footer application (parameterized for different document types)
- Heading number stripping for Word multilevel lists
"""

import os
import re
import shutil
import subprocess
from pathlib import Path
from typing import Optional

from config import SCRIPT_DIR, _get_soffice, _make_user_profile


# -----------------------------
# Conversions (LibreOffice-only, cross-platform)
# -----------------------------
def convert_pptx_to_pdf(pptx_path: str, pdf_path: str) -> None:
    """Convert PPTX to PDF using LibreOffice headless."""
    abs_pptx = os.path.abspath(pptx_path)
    abs_pdf = os.path.abspath(pdf_path)
    out_dir = os.path.dirname(abs_pdf)
    soffice = _get_soffice()
    user_profile = _make_user_profile()

    try:
        subprocess.run(
            [soffice, "--headless", "--nologo", "--nofirststartwizard",
             "--norestore",
             f"-env:UserInstallation=file:///{user_profile.replace(os.sep, '/')}",
             "--convert-to", "pdf", "--outdir", out_dir, abs_pptx],
            check=True, timeout=180,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        expected_out = os.path.join(out_dir, os.path.splitext(os.path.basename(abs_pptx))[0] + ".pdf")
        if expected_out != abs_pdf and os.path.exists(expected_out):
            os.rename(expected_out, abs_pdf)
        if not os.path.exists(abs_pdf):
            raise RuntimeError("LibreOffice conversion succeeded but output PDF not found.")
    except subprocess.TimeoutExpired:
        raise RuntimeError("PPTX to PDF conversion timed out (180s).")
    except Exception as e:
        raise RuntimeError(f"PPTX to PDF conversion failed: {e}")
    finally:
        shutil.rmtree(user_profile, ignore_errors=True)


def _update_toc_with_libreoffice(docx_path: str) -> None:
    """Fallback TOC update: round-trip DOCX through LibreOffice to force field recalculation."""
    abs_path = os.path.abspath(docx_path)
    out_dir = os.path.dirname(abs_path)
    base_name = os.path.splitext(os.path.basename(abs_path))[0]
    soffice = _get_soffice()

    # Create a backup
    backup_path = abs_path + ".bak"
    shutil.copy2(abs_path, backup_path)

    try:
        # Convert DOCX -> DOCX through LibreOffice (forces field update)
        new_profile = _make_user_profile()
        try:
            subprocess.run(
                [soffice, "--headless", "--norestore",
                 f"-env:UserInstallation=file:///{new_profile.replace(os.sep, '/')}",
                 "--infilter=Microsoft Word 2007-2019 XML",
                 "--convert-to", "docx:Microsoft Word 2007-2019 XML",
                 "--outdir", out_dir, abs_path],
                check=True, timeout=120,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
        finally:
            shutil.rmtree(new_profile, ignore_errors=True)

        # LibreOffice may produce the output with the same name, check if it exists
        if os.path.exists(abs_path):
            # Remove backup since conversion succeeded
            os.remove(backup_path)
        else:
            # Restore from backup
            os.rename(backup_path, abs_path)
    except Exception as e:
        # Restore from backup on failure
        if os.path.exists(backup_path):
            if os.path.exists(abs_path):
                os.remove(abs_path)
            os.rename(backup_path, abs_path)
        print(f"[TOC Roundtrip] Warning: {e}")
        raise e


# =============================================================================
# Markdown -> DOCX (Pandoc + python-docx post-processing)
# =============================================================================
# This section is structured as a pipeline of focused sub-functions for
# maintainability and testability.

def _pandoc_md_to_raw_docx(md_path: str, docx_path: str) -> None:
    """Phase 1: Use Pandoc to convert Markdown to a raw DOCX with the Enfrasys template."""
    import pypandoc

    template_file = str(SCRIPT_DIR / "enfrasys_template.docx")
    if not os.path.exists(template_file):
        raise FileNotFoundError(f"Deployment Error: enfrasys_template.docx is missing at {template_file}")

    task_dir_str = str(Path(md_path).parent.resolve())
    extra_args = [
        f"--reference-doc={template_file}",
        f"--resource-path={task_dir_str}"
    ]

    pypandoc.convert_file(
        str(md_path),
        "docx",
        outputfile=str(docx_path),
        extra_args=extra_args
    )


def _apply_header_footer(doc, document_title, client_logo_path,
                          document_type_label="Planning & Design Document",
                          title_separator=" - "):
    """Apply header logos and dynamic footer to all sections.
    
    Args:
        doc: python-docx Document object
        document_title: The project/client title for the footer
        client_logo_path: Path to client logo image (or None)
        document_type_label: Label for footer (e.g., "Planning & Design Document",
                             "Final Configuration Document"). Defaults to Design Document.
    """
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    enfrasys_logo = SCRIPT_DIR / "enfrasys-logo.jpg"

    # --- Header ---
    doc.sections[0].different_first_page_header_footer = True
    header = doc.sections[0].header

    htable = header.add_table(1, 2, Inches(6.0))
    for p in header.paragraphs:
        if not p._element.getparent() == htable._element and p.text.strip() == "":
            p_elem_h = p._element
            p_elem_h.getparent().remove(p_elem_h)

    tblPr_h = htable._tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "none")
        tblBorders.append(border)
    tblPr_h.append(tblBorders)

    cell_left = htable.cell(0, 0)
    cell_right = htable.cell(0, 1)

    if client_logo_path and os.path.exists(client_logo_path):
        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
        try:
            p_left.add_run().add_picture(str(client_logo_path), height=Inches(0.6))
        except Exception as e:
            print(f"[build_docx] Failed to insert client logo in header: {e}")

    cell_right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if os.path.exists(enfrasys_logo):
        p_right = cell_right.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        try:
            p_right.add_run().add_picture(str(enfrasys_logo), width=Inches(1.5))
        except Exception as e:
            print(f"[build_docx] Failed to insert Enfrasys logo in header: {e}")

    # --- Footer ---
    clean_title = document_title.replace('\n', '').replace('\r', '').strip()

    for section in doc.sections:
        footer = section.footer
        for p in footer.paragraphs:
            p_element = p._element
            p_element.getparent().remove(p_element)

        para = footer.add_paragraph()
        para.paragraph_format.tab_stops.clear_all()
        para.paragraph_format.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

        run_left = para.add_run(f"{clean_title}{title_separator}{document_type_label}\t")
        run_left.font.name = 'Segoe UI'
        run_left.font.size = Pt(9)

        run_right = para.add_run()
        run_right.font.name = 'Segoe UI'
        run_right.font.size = Pt(9)

        fldChar_begin = OxmlElement('w:fldChar')
        fldChar_begin.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"
        fldChar_end = OxmlElement('w:fldChar')
        fldChar_end.set(qn('w:fldCharType'), 'end')

        run_right._r.append(fldChar_begin)
        run_right._r.append(instrText)
        run_right._r.append(fldChar_end)


def _fix_tables(doc):
    """Enterprise-grade table fix: explicit XML borders, fixed layout, header banding.
    Includes dynamic page width calculation and the LibreOffice <w:tblGrid> skeleton fix.
    """
    from docx.shared import Pt, RGBColor
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    ENFRASYS_BLUE = "4472C4"
    ENFRASYS_BLUE_LIGHT = "D6E4F0"
    
    # ---------------------------------------------------------
    # DYNAMIC WIDTH CALCULATION (Fixes the "Weird Spaces")
    # ---------------------------------------------------------
    # Get the actual page dimensions from the first section
    section = doc.sections[0]
    # Dimensions are in EMU. 1 inch = 1440 twips = 914400 EMU. 
    # Therefore, 1 twip = 635 EMU.
    printable_width_emu = section.page_width - section.left_margin - section.right_margin
    PAGE_WIDTH_TWIPS = int(printable_width_emu / 635)

    for table in doc.tables:
        tblPr = table._tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            table._tbl.insert(0, tblPr)

        # --- Detect if this is a signature/front-matter table ---
        is_signature_table = False
        if len(table.rows) > 0:
            for cell in table.rows[0].cells:
                cell_text = cell.text.lower()
                if "prepared by:" in cell_text or "reviewed by:" in cell_text or "verified by:" in cell_text:
                    is_signature_table = True
                    break

        # --- Detect if this is a horizontal column table (section 3.0 style) ---
        # These tables have labels in the LEFT column instead of a top header row.
        # They are identified by having 2 columns and known left-column labels.
        is_horizontal_table = False
        num_cols = len(table.rows[0].cells) if len(table.rows) > 0 else 0
        horizontal_header_keywords = [
            "field", "exchange version", "policy setting",
            "problem description", "total user", "shared",
            "total mailbox", "smtp relay", "mail flow",
            "public folders", "exchange rules", "outlook client"
        ]
        if num_cols == 2 and not is_signature_table and len(table.rows) > 0:
            first_cell_text = table.rows[0].cells[0].text.strip().lower()
            if any(kw in first_cell_text for kw in horizontal_header_keywords):
                is_horizontal_table = True

        # --- 1. Remove any existing borders and set explicit ones ---
        existing_borders = tblPr.find(qn('w:tblBorders'))
        if existing_borders is not None:
            tblPr.remove(existing_borders)

        tblBorders = OxmlElement('w:tblBorders')
        # Use black borders for ALL tables for consistent visibility
        border_color = "000000"
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '8')       # ← Increased from 4 to 8 (1pt visible)
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), border_color)
            tblBorders.append(border)
        tblPr.append(tblBorders)

        # --- 2. Set table width and align it perfectly to the margin ---
        existing_tblW = tblPr.find(qn('w:tblW'))
        if existing_tblW is not None:
            tblPr.remove(existing_tblW)
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:type'), 'dxa')
        tblW.set(qn('w:w'), str(PAGE_WIDTH_TWIPS))
        tblPr.append(tblW)

        # Force Left Alignment to prevent weird side spacing
        existing_jc = tblPr.find(qn('w:jc'))
        if existing_jc is not None:
            tblPr.remove(existing_jc)
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'left')
        tblPr.append(jc)

        # Strip any rogue indentation
        existing_ind = tblPr.find(qn('w:tblInd'))
        if existing_ind is not None:
            tblPr.remove(existing_ind)
        tblInd = OxmlElement('w:tblInd')
        tblInd.set(qn('w:w'), '0')
        tblInd.set(qn('w:type'), 'dxa')
        tblPr.append(tblInd)

        # --- 3. Set fixed table layout ---
        existing_layout = tblPr.find(qn('w:tblLayout'))
        if existing_layout is not None:
            tblPr.remove(existing_layout)
        tblLayout = OxmlElement('w:tblLayout')
        tblLayout.set(qn('w:type'), 'fixed')
        tblPr.append(tblLayout)

        # --- 4. Distribute column widths and build <w:tblGrid> ---
        # num_cols was already computed above during horizontal table detection
        
        if num_cols > 0:
            # Default: equal distribution
            ratios = [1.0 / num_cols] * num_cols
            
            # Identify table by headers to apply organic proportional widths
            if len(table.rows) > 0 and not is_signature_table:
                headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
                
                if num_cols == 3 and 'id' in headers[0] and 'description' in headers:
                    ratios = [0.10, 0.65, 0.25]  # Design Considerations
                elif num_cols == 3 and 'no' in headers[0] and 'decision' in headers[-1]:
                    ratios = [0.10, 0.45, 0.45]  # Design Decisions
                elif num_cols == 4 and 'item' in headers[0] and 'activities' in headers[1]:
                    ratios = [0.15, 0.45, 0.20, 0.20]  # Task Rollout
                elif num_cols == 2 and 'capability' in headers[0]:
                    ratios = [0.30, 0.70]  # Capabilities
                elif num_cols == 2 and 'component' in headers[0]:
                    ratios = [0.25, 0.75]  # VM Specs
                elif num_cols == 2 and next(iter(headers), "") == "setting":
                    ratios = [0.35, 0.65]  # Final Config Setting | Value
                elif num_cols >= 8 and 'priority' in headers[1]:
                    ratios = [0.15, 0.10, 0.15, 0.15, 0.15, 0.10, 0.10, 0.10] # NSG Rules

            # Horizontal column tables get a 35/65 split
            if is_horizontal_table:
                ratios = [0.35, 0.65]
            
            # Constrain Signature Tables to not span the full page width
            target_page_twips = PAGE_WIDTH_TWIPS
            if is_signature_table:
                target_page_twips = int(PAGE_WIDTH_TWIPS * 0.75) # 75% width centered look
            
            col_widths = [int(target_page_twips * r) for r in ratios]

            # Build the new structural grid for LibreOffice
            existing_tblGrid = table._tbl.find(qn('w:tblGrid'))
            if existing_tblGrid is not None:
                table._tbl.remove(existing_tblGrid)

            tblGrid = OxmlElement('w:tblGrid')
            for w in col_widths:
                gridCol = OxmlElement('w:gridCol')
                gridCol.set(qn('w:w'), str(w))
                tblGrid.append(gridCol)

            # tblGrid MUST be inserted strictly before the first <w:tr>
            first_tr_idx = 0
            for idx, child in enumerate(table._tbl):
                if child.tag == qn('w:tr'):
                    first_tr_idx = idx
                    break
            
            if first_tr_idx > 0:
                table._tbl.insert(first_tr_idx, tblGrid)
            else:
                table._tbl.append(tblGrid)

            # Apply the exact explicit width to every single cell
            for row in table.rows:
                for idx, cell in enumerate(row.cells):
                    # Safety: skip cells beyond expected column count
                    if idx >= len(col_widths):
                        continue
                    
                    tcPr = cell._tc.get_or_add_tcPr()
                    
                    existing_tcW = tcPr.find(qn('w:tcW'))
                    if existing_tcW is not None:
                        tcPr.remove(existing_tcW)
                        
                    tcW = OxmlElement('w:tcW')
                    tcW.set(qn('w:type'), 'dxa')
                    tcW.set(qn('w:w'), str(col_widths[idx]))
                    tcPr.append(tcW)

        # --- 5. Apply table styling ---
        if not is_signature_table and len(table.rows) > 0:

            if is_horizontal_table:
                # --- HORIZONTAL COLUMN TABLE STYLING (Section 3.0) ---
                # The first row is a dummy "Field | Detail" header required by Markdown.
                # Remove it so the table starts with actual data rows.
                if len(table.rows) > 1:
                    header_row_text = table.rows[0].cells[0].text.strip().lower()
                    if header_row_text == "field":
                        table._tbl.remove(table.rows[0]._tr)

                # Style LEFT column (col 0) with blue bg + white bold text for ALL rows
                # Style RIGHT column (col 1) with alternating bands
                for row_idx, row in enumerate(table.rows):
                    # Left cell: blue background, white bold text
                    left_cell = row.cells[0]
                    left_tcPr = left_cell._tc.get_or_add_tcPr()
                    existing_shd = left_tcPr.find(qn('w:shd'))
                    if existing_shd is not None:
                        left_tcPr.remove(existing_shd)
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:val'), 'clear')
                    shading.set(qn('w:color'), 'auto')
                    shading.set(qn('w:fill'), ENFRASYS_BLUE)
                    left_tcPr.append(shading)
                    for p in left_cell.paragraphs:
                        for run in p.runs:
                            run.font.color.rgb = RGBColor(255, 255, 255)
                            run.bold = True
                            run.font.name = 'Segoe UI'
                            run.font.size = Pt(10)

                    # Right cell: alternating light blue band
                    right_cell = row.cells[1]
                    right_tcPr = right_cell._tc.get_or_add_tcPr()
                    existing_shd_r = right_tcPr.find(qn('w:shd'))
                    if existing_shd_r is not None:
                        right_tcPr.remove(existing_shd_r)
                    if row_idx % 2 == 0:
                        shading_r = OxmlElement('w:shd')
                        shading_r.set(qn('w:val'), 'clear')
                        shading_r.set(qn('w:color'), 'auto')
                        shading_r.set(qn('w:fill'), ENFRASYS_BLUE_LIGHT)
                        right_tcPr.append(shading_r)

                # Do NOT add <w:tblHeader> for horizontal tables (no repeating header)

            else:
                # --- NORMAL TABLE STYLING (top row = header) ---
                for cell in table.rows[0].cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    existing_shd = tcPr.find(qn('w:shd'))
                    if existing_shd is not None:
                        tcPr.remove(existing_shd)
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:val'), 'clear')
                    shading.set(qn('w:color'), 'auto')
                    shading.set(qn('w:fill'), ENFRASYS_BLUE)
                    tcPr.append(shading)

                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.color.rgb = RGBColor(255, 255, 255)
                            run.bold = True
                            run.font.name = 'Segoe UI'
                            run.font.size = Pt(10)

                for row_idx, row in enumerate(table.rows):
                    if row_idx == 0:
                        continue
                    if row_idx % 2 == 0:
                        for cell in row.cells:
                            tcPr = cell._tc.get_or_add_tcPr()
                            existing_shd = tcPr.find(qn('w:shd'))
                            if existing_shd is not None:
                                tcPr.remove(existing_shd)
                            shading = OxmlElement('w:shd')
                            shading.set(qn('w:val'), 'clear')
                            shading.set(qn('w:color'), 'auto')
                            shading.set(qn('w:fill'), ENFRASYS_BLUE_LIGHT)
                            tcPr.append(shading)

                first_tr = table.rows[0]._tr
                trPr = first_tr.find(qn('w:trPr'))
                if trPr is None:
                    trPr = OxmlElement('w:trPr')
                    first_tr.insert(0, trPr)
                tblHeader = OxmlElement('w:tblHeader')
                trPr.append(tblHeader)

        # --- 6. Cell formatting: vertical alignment, padding, fonts ---
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                tcPr = cell._tc.get_or_add_tcPr()
                tcMar = OxmlElement('w:tcMar')
                for margin_name in ['top', 'bottom', 'left', 'right']:
                    margin = OxmlElement(f'w:{margin_name}')
                    margin.set(qn('w:w'), '40')  
                    margin.set(qn('w:type'), 'dxa')
                    tcMar.append(margin)
                
                existing_mar = tcPr.find(qn('w:tcMar'))
                if existing_mar is not None:
                    tcPr.remove(existing_mar)
                tcPr.append(tcMar)

                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(1)
                    p.paragraph_format.space_after = Pt(1)
                    for run in p.runs:
                        run.font.name = 'Segoe UI'
                        if run.font.size is None:
                            run.font.size = Pt(10)


def _fix_typography_and_captions(doc):
    """Enforce global font, heading sizes, caption styling, page breaks, and text justification."""
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement

    style_rules = {
        'Heading 1': {'size': 15, 'bold': False},
        'Heading 2': {'size': 14, 'bold': False},
        'Heading 3': {'size': 13, 'bold': False},
        'Heading 4': {'size': 12, 'bold': False},
    }

    figure_counter = 1

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # --- A. FIX IMAGE CAPTIONS (Renumber & Center Safely) ---
        if text.lower().startswith("figure") and ":" in text:
            # Match either "Figure:" OR "Figure N:" — code controls the numbering
            new_text = re.sub(r'(?i)^figure\s*\d*\s*:', f'Figure {figure_counter}:', text)

            # SAFE TEXT REPLACEMENT: Protect inline images from being deleted!
            has_drawing = any(getattr(run._element, "drawing_lst", None) for run in para.runs)

            if has_drawing:
                for run in para.runs:
                    if run.text:
                        run.text = ""
                para.add_run("\n" + new_text)
            else:
                para.text = new_text

            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            for run in para.runs:
                if run.text.strip():
                    run.font.name = 'Segoe UI'
                    run.font.size = Pt(10)
                    run.font.italic = True
                    run.font.color.rgb = RGBColor(128, 128, 128)

            figure_counter += 1
            continue

        # --- B. STANDARD TEXT JUSTIFICATION ---
        if para.style.name == "Normal":
            # Skip justification if the paragraph contains an image to preserve image alignments
            has_drawing = any(getattr(run._element, "drawing_lst", None) for run in para.runs)
            if not has_drawing:
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # --- C. PAGE BREAKS FOR MAIN SECTIONS ---
        if para.style.name == "Heading 1" or any(text.startswith(f"{i}.0") for i in range(2, 21)):
            if not text.startswith("1.0") and not text.startswith("2.0"):
                para.paragraph_format.page_break_before = True

        # --- D. ENFORCE TYPOGRAPHY SCALING ---
        style_name = para.style.name
        rule = style_rules.get(style_name)

        for run in para.runs:
            run.font.name = 'Segoe UI'
            if rule:
                run.font.size = Pt(rule['size'])
                run.font.bold = rule['bold']
            else:
                if run.font.size is None:
                    run.font.size = Pt(11)

    # --- E. INSERT BLANK LINE BEFORE SUBHEADINGS ---
    # Rule: Insert a blank paragraph before H2/H3/H4 UNLESS the nearest
    # preceding heading is the direct parent (e.g., H1→H2, H2→H3).
    heading_styles = {'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4'}
    parent_of = {'Heading 2': 'Heading 1', 'Heading 3': 'Heading 2', 'Heading 4': 'Heading 3'}
    all_paragraphs = list(doc.paragraphs)

    for i, p in enumerate(all_paragraphs):
        if p.style.name not in parent_of:
            continue
        # Find the nearest preceding heading
        prev_heading_style = None
        for j in range(i - 1, -1, -1):
            if all_paragraphs[j].style.name in heading_styles:
                prev_heading_style = all_paragraphs[j].style.name
                break
        # Skip if the previous heading is the direct parent
        if prev_heading_style == parent_of[p.style.name]:
            continue
        p._element.addprevious(OxmlElement('w:p'))


def _fix_images(doc):
    """Center images and enforce max height with proportional scaling."""
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    # Center image paragraphs (unless explicitly left aligned)
    for p in doc.paragraphs:
        for run in p.runs:
            drawing = getattr(run._element, "drawing_lst", None)
            if drawing:
                if p.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Resize oversized images
    for shape in doc.inline_shapes:
        max_height = Inches(4.5)
        if shape.height > max_height:
            ratio = max_height / float(shape.height)
            shape.height = max_height
            shape.width = int(shape.width * ratio)


def _strip_heading_numbers(doc):
    """Remove hardcoded section numbers from headings so Word's multilevel list numbering takes over."""
    for para in doc.paragraphs:
        if para.style.name in ('Heading 1', 'Heading 2', 'Heading 3', 'Heading 4'):
            text = para.text
            match = re.match(r'^(\d+(\.\d+)+\s+)', text)
            if match:
                to_remove = len(match.group(1))
                for run in para.runs:
                    if to_remove <= 0:
                        break
                    run_len = len(run.text)
                    if run_len <= to_remove:
                        to_remove -= run_len
                        run.text = ""
                    else:
                        run.text = run.text[to_remove:]
                        to_remove = 0
