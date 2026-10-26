import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import fitz
import pandas as pd
import streamlit as st

from llm_helper import answer_with_gemini


# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

PROJECT_DIR = Path(__file__).resolve().parent
DEMO_DIR = PROJECT_DIR / "demo_dataset"
EXTRACTOR_PATH = PROJECT_DIR / "extract_iges.py"

FREECAD_CMD = Path(
    r"C:\Program Files\FreeCAD 1.1\bin\FreeCADCmd.exe"
)

st.set_page_config(
    page_title="CNC Part Intelligence",
    page_icon="⚙️",
    layout="wide",
)


# ------------------------------------------------------------
# General helper functions
# ------------------------------------------------------------

def clean_value(value):
    if pd.isna(value):
        return None

    text = str(value).strip()

    if text.lower() in {"", "nan", "none"}:
        return None

    if text.endswith(".0"):
        try:
            return str(int(float(text)))
        except ValueError:
            pass

    return text


def normalize_part_number(value):
    text = clean_value(value)

    if text is None:
        return ""

    return re.sub(r"\s+", "", text).lower()


def safe_float(value):
    if value is None:
        return None

    match = re.search(r"-?\d+(?:\.\d+)?", str(value))

    if not match:
        return None

    try:
        return float(match.group())
    except ValueError:
        return None


def save_uploaded_file(uploaded_file, destination):
    with open(destination, "wb") as output_file:
        output_file.write(uploaded_file.getbuffer())


# ------------------------------------------------------------
# Excel processing
# ------------------------------------------------------------

def find_part_column(columns):
    preferred_names = [
        "part number",
        "part no",
        "part no.",
        "part_number",
        "partnumber",
        "component number",
        "component id",
    ]

    normalized_columns = {
        str(column).strip().lower(): column
        for column in columns
    }

    for preferred_name in preferred_names:
        if preferred_name in normalized_columns:
            return normalized_columns[preferred_name]

    for column in columns:
        name = str(column).strip().lower()

        if "part" in name and (
            "number" in name
            or "no" in name
            or "id" in name
        ):
            return column

    return None


def read_excel_record(excel_path, requested_part_number):
    try:
        workbook = pd.ExcelFile(excel_path)
    except Exception as error:
        return {
            "status": "error",
            "source_file": Path(excel_path).name,
            "message": f"Could not open Excel file: {error}",
        }

    requested_normalized = normalize_part_number(
        requested_part_number
    )

    for sheet_name in workbook.sheet_names:
        try:
            dataframe = pd.read_excel(
                excel_path,
                sheet_name=sheet_name,
                dtype=str,
            )
        except Exception:
            continue

        if dataframe.empty:
            continue

        dataframe.columns = [
            str(column).strip()
            for column in dataframe.columns
        ]

        part_column = find_part_column(dataframe.columns)

        if part_column is not None:
            matches = dataframe[
                dataframe[part_column]
                .map(normalize_part_number)
                .eq(requested_normalized)
            ]
        else:
            matches = dataframe[
                dataframe.apply(
                    lambda row: any(
                        normalize_part_number(value)
                        == requested_normalized
                        for value in row
                    ),
                    axis=1,
                )
            ]

        if matches.empty:
            continue

        row_index = matches.index[0]
        matched_row = matches.iloc[0]

        record = {}

        for column, value in matched_row.items():
            cleaned = clean_value(value)

            if cleaned is not None:
                record[str(column)] = cleaned

        return {
            "status": "success",
            "source_file": Path(excel_path).name,
            "sheet": sheet_name,
            "excel_row": int(row_index) + 2,
            "part_column": (
                str(part_column)
                if part_column is not None
                else "Detected by complete row search"
            ),
            "record": record,
        }

    return {
        "status": "not_found",
        "source_file": Path(excel_path).name,
        "message": (
            f"Part number {requested_part_number} was not found "
            "in any Excel worksheet."
        ),
    }


# ------------------------------------------------------------
# PDF processing
# ------------------------------------------------------------

def extract_pdf_facts(text):
    facts = {}

    patterns = {
        "part_number": [
            r"PART\s*NUMBER\s*[:\-]?\s*([A-Za-z0-9_.\-]+)",
        ],
        "revision": [
            r"REVISION\s*[:\-]?\s*([A-Za-z0-9_.\-]+)",
        ],
        "material": [
            r"Material\s*:\s*([^\n.]+)",
        ],
        "surface_finish": [
            r"Surface\s*finish\s*:\s*([^\n]+)",
        ],
        "general_tolerance": [
            r"General\s+dimensional\s+tolerance\s*:\s*([^\n]+)",
            r"General\s+tolerance\s*:\s*([^\n]+)",
        ],
        "process": [
            r"Process\s*:\s*([^\n.]+)",
        ],
    }

    for field, field_patterns in patterns.items():
        for pattern in field_patterns:
            match = re.search(
                pattern,
                text,
                flags=re.IGNORECASE,
            )

            if match:
                facts[field] = match.group(1).strip()
                break

    dimension_match = re.search(
        r"Overall\s+dimensions?\s*:\s*"
        r"(\d+(?:\.\d+)?)\s*[xX×]\s*"
        r"(\d+(?:\.\d+)?)\s*[xX×]\s*"
        r"(\d+(?:\.\d+)?)\s*mm",
        text,
        flags=re.IGNORECASE,
    )

    if dimension_match:
        facts["overall_dimensions_mm"] = [
            float(dimension_match.group(1)),
            float(dimension_match.group(2)),
            float(dimension_match.group(3)),
        ]

    corner_hole_match = re.search(
        r"(\d+)\s*[xX]\s*DIA\s*"
        r"(\d+(?:\.\d+)?)\s*mm\s*THRU",
        text,
        flags=re.IGNORECASE,
    )

    if corner_hole_match:
        facts["corner_hole_count"] = int(
            corner_hole_match.group(1)
        )
        facts["corner_hole_diameter_mm"] = float(
            corner_hole_match.group(2)
        )

    center_hole_match = re.search(
        r"1\s*[xX]\s*DIA\s*"
        r"(\d+(?:\.\d+)?)\s*mm\s*THRU",
        text,
        flags=re.IGNORECASE,
    )

    if center_hole_match:
        facts["center_hole_diameter_mm"] = float(
            center_hole_match.group(1)
        )

    return facts


def read_pdf(pdf_path):
    pages = []

    try:
        document = fitz.open(pdf_path)

        for page_index, page in enumerate(document):
            text = page.get_text("text").strip()

            pages.append(
                {
                    "page": page_index + 1,
                    "text": text,
                }
            )

        document.close()

    except Exception as error:
        return {
            "status": "error",
            "source_file": Path(pdf_path).name,
            "message": f"Could not read PDF: {error}",
            "pages": [],
            "facts": {},
            "full_text": "",
        }

    combined_text = "\n".join(
        page["text"]
        for page in pages
        if page["text"]
    )

    if not combined_text.strip():
        return {
            "status": "warning",
            "source_file": Path(pdf_path).name,
            "page_count": len(pages),
            "message": (
                "No machine-readable text was found. "
                "The PDF may be scanned and require OCR."
            ),
            "pages": pages,
            "facts": {},
            "full_text": "",
        }

    return {
        "status": "success",
        "source_file": Path(pdf_path).name,
        "page_count": len(pages),
        "pages": pages,
        "full_text": combined_text,
        "facts": extract_pdf_facts(combined_text),
    }


# ------------------------------------------------------------
# IGES processing through FreeCAD
# ------------------------------------------------------------

def analyze_iges(iges_path, working_directory):
    if not FREECAD_CMD.exists():
        return {
            "status": "error",
            "source_file": Path(iges_path).name,
            "message": (
                "FreeCADCmd.exe was not found at "
                f"{FREECAD_CMD}"
            ),
        }

    if not EXTRACTOR_PATH.exists():
        return {
            "status": "error",
            "source_file": Path(iges_path).name,
            "message": "extract_iges.py was not found.",
        }

    output_path = (
        Path(working_directory) / "geometry_result.json"
    )

    environment = os.environ.copy()
    environment["IGES_INPUT"] = str(
        Path(iges_path).resolve()
    )
    environment["IGES_OUTPUT"] = str(
        output_path.resolve()
    )

    try:
        process = subprocess.run(
            [
                str(FREECAD_CMD),
                str(EXTRACTOR_PATH),
            ],
            cwd=str(PROJECT_DIR),
            env=environment,
            capture_output=True,
            text=True,
            timeout=120,
        )

    except subprocess.TimeoutExpired:
        return {
            "status": "error",
            "source_file": Path(iges_path).name,
            "message": (
                "IGES analysis exceeded the 120-second limit."
            ),
        }

    except Exception as error:
        return {
            "status": "error",
            "source_file": Path(iges_path).name,
            "message": f"Could not start FreeCAD: {error}",
        }

    if not output_path.exists():
        return {
            "status": "error",
            "source_file": Path(iges_path).name,
            "message": (
                "FreeCAD did not produce a geometry result."
            ),
            "freecad_output": (
                process.stdout + "\n" + process.stderr
            ).strip(),
        }

    try:
        with open(
            output_path,
            "r",
            encoding="utf-8",
        ) as result_file:
            result = json.load(result_file)

    except Exception as error:
        return {
            "status": "error",
            "source_file": Path(iges_path).name,
            "message": (
                f"Could not read the geometry result: {error}"
            ),
        }

    result["freecad_return_code"] = process.returncode
    return result


# ------------------------------------------------------------
# Cross-source validation
# ------------------------------------------------------------

def find_excel_value(record, possible_names):
    normalized_record = {
        re.sub(r"[^a-z0-9]", "", key.lower()): value
        for key, value in record.items()
    }

    for possible_name in possible_names:
        normalized_name = re.sub(
            r"[^a-z0-9]",
            "",
            possible_name.lower(),
        )

        if normalized_name in normalized_record:
            return normalized_record[normalized_name]

    return None


def obtain_excel_dimensions(excel_result):
    if excel_result.get("status") != "success":
        return None

    record = excel_result.get("record", {})

    length = safe_float(
        find_excel_value(
            record,
            [
                "Overall Length (mm)",
                "Length (mm)",
                "Length",
            ],
        )
    )

    width = safe_float(
        find_excel_value(
            record,
            [
                "Overall Width (mm)",
                "Width (mm)",
                "Width",
            ],
        )
    )

    thickness = safe_float(
        find_excel_value(
            record,
            [
                "Thickness (mm)",
                "Height (mm)",
                "Thickness",
                "Height",
            ],
        )
    )

    if None in {length, width, thickness}:
        return None

    return [length, width, thickness]


def compare_dimensions(
    excel_result,
    pdf_result,
    iges_result,
):
    sources = {}

    excel_dimensions = obtain_excel_dimensions(
        excel_result
    )

    if excel_dimensions:
        sources["Excel"] = excel_dimensions

    pdf_dimensions = (
        pdf_result.get("facts", {})
        .get("overall_dimensions_mm")
    )

    if pdf_dimensions:
        sources["PDF"] = pdf_dimensions

    if iges_result.get("status") == "success":
        box = iges_result.get("bounding_box_mm", {})

        model_dimensions = [
            box.get("length_x"),
            box.get("width_y"),
            box.get("height_z"),
        ]

        if all(
            value is not None
            for value in model_dimensions
        ):
            sources["IGES"] = model_dimensions

    rows = []
    source_names = list(sources.keys())

    for first_index in range(len(source_names)):
        for second_index in range(
            first_index + 1,
            len(source_names),
        ):
            first_name = source_names[first_index]
            second_name = source_names[second_index]

            first_dimensions = sorted(
                sources[first_name]
            )
            second_dimensions = sorted(
                sources[second_name]
            )

            differences = [
                abs(first - second)
                for first, second in zip(
                    first_dimensions,
                    second_dimensions,
                )
            ]

            maximum_difference = max(differences)

            if maximum_difference <= 0.1:
                status = "MATCH"
                explanation = (
                    "Dimensions agree within 0.10 mm."
                )
            else:
                status = "CONFLICT"
                explanation = (
                    "Maximum dimensional difference is "
                    f"{maximum_difference:.3f} mm."
                )

            rows.append(
                {
                    "Comparison": (
                        f"{first_name} vs {second_name}"
                    ),
                    "First source (mm)": " × ".join(
                        f"{value:g}"
                        for value in sources[first_name]
                    ),
                    "Second source (mm)": " × ".join(
                        f"{value:g}"
                        for value in sources[second_name]
                    ),
                    "Status": status,
                    "Explanation": explanation,
                }
            )

    if not rows:
        rows.append(
            {
                "Comparison": "Overall dimensions",
                "First source (mm)": "Not available",
                "Second source (mm)": "Not available",
                "Status": "INSUFFICIENT DATA",
                "Explanation": (
                    "At least two sources are required "
                    "for cross-checking."
                ),
            }
        )

    return rows


# ------------------------------------------------------------
# Deterministic offline fallback
# ------------------------------------------------------------

def answer_question(
    question,
    part_number,
    excel_result,
    pdf_result,
    iges_result,
    comparison_rows,
):
    question_lower = question.lower()
    evidence = []
    answer_parts = []

    excel_record = excel_result.get("record", {})

    excel_source = (
        f"{excel_result.get('source_file')}, "
        f"sheet {excel_result.get('sheet')}, "
        f"row {excel_result.get('excel_row')}"
    )

    pdf_source = (
        f"{pdf_result.get('source_file')}, page 1"
    )

    iges_source = iges_result.get(
        "source_file",
        "IGES model",
    )

    if "material" in question_lower:
        material = find_excel_value(
            excel_record,
            ["Material"],
        )

        if material:
            answer_parts.append(
                f"The material is {material}."
            )
            evidence.append(excel_source)

        pdf_material = (
            pdf_result.get("facts", {})
            .get("material")
        )

        if pdf_material:
            answer_parts.append(
                f"The drawing states: {pdf_material}."
            )
            evidence.append(pdf_source)

    elif (
        "dimension" in question_lower
        or "size" in question_lower
        or "bounding" in question_lower
    ):
        dimensions = obtain_excel_dimensions(
            excel_result
        )

        if dimensions:
            answer_parts.append(
                "The Excel record gives overall dimensions of "
                f"{dimensions[0]:g} × "
                f"{dimensions[1]:g} × "
                f"{dimensions[2]:g} mm."
            )
            evidence.append(excel_source)

        pdf_dimensions = (
            pdf_result.get("facts", {})
            .get("overall_dimensions_mm")
        )

        if pdf_dimensions:
            answer_parts.append(
                "The drawing gives overall dimensions of "
                f"{pdf_dimensions[0]:g} × "
                f"{pdf_dimensions[1]:g} × "
                f"{pdf_dimensions[2]:g} mm."
            )
            evidence.append(pdf_source)

        if iges_result.get("status") == "success":
            box = iges_result["bounding_box_mm"]

            answer_parts.append(
                "The IGES bounding box is "
                f"{box['length_x']:g} × "
                f"{box['width_y']:g} × "
                f"{box['height_z']:g} mm."
            )
            evidence.append(iges_source)

    elif "volume" in question_lower:
        if iges_result.get("status") == "success":
            volume = iges_result.get("volume_mm3")
            reliability = iges_result.get(
                "volume_reliability",
                "unknown",
            )

            answer_parts.append(
                f"The reported model volume is "
                f"{volume:,.3f} mm³. "
                f"Reliability is {reliability}."
            )
            evidence.append(iges_source)

            if iges_result.get("warning"):
                answer_parts.append(
                    f"Warning: {iges_result['warning']}"
                )

    elif (
        "surface finish" in question_lower
        or "roughness" in question_lower
    ):
        finish = find_excel_value(
            excel_record,
            ["Surface Finish"],
        )

        if finish:
            answer_parts.append(
                f"The specified surface finish is {finish}."
            )
            evidence.append(excel_source)

        pdf_finish = (
            pdf_result.get("facts", {})
            .get("surface_finish")
        )

        if pdf_finish:
            answer_parts.append(
                f"The drawing states: {pdf_finish}."
            )
            evidence.append(pdf_source)

    elif "quantity" in question_lower:
        quantity = find_excel_value(
            excel_record,
            ["Quantity", "Qty"],
        )

        if quantity:
            answer_parts.append(
                f"The manufacturing quantity is {quantity}."
            )
            evidence.append(excel_source)

    elif (
        "agree" in question_lower
        or "match" in question_lower
        or "conflict" in question_lower
        or "disagree" in question_lower
    ):
        for row in comparison_rows:
            answer_parts.append(
                f"{row['Comparison']}: "
                f"{row['Status']} — "
                f"{row['Explanation']}"
            )

        evidence.extend(
            [
                excel_result.get(
                    "source_file",
                    "Excel",
                ),
                pdf_result.get(
                    "source_file",
                    "PDF",
                ),
                iges_source,
            ]
        )

    elif "process" in question_lower:
        process = find_excel_value(
            excel_record,
            ["Manufacturing Process", "Process"],
        )

        if process:
            answer_parts.append(
                f"The manufacturing process is {process}."
            )
            evidence.append(excel_source)

    elif "revision" in question_lower:
        revision = find_excel_value(
            excel_record,
            ["Revision", "Rev"],
        )

        if revision:
            answer_parts.append(
                f"The Excel record revision is {revision}."
            )
            evidence.append(excel_source)

        pdf_revision = (
            pdf_result.get("facts", {})
            .get("revision")
        )

        if pdf_revision:
            answer_parts.append(
                f"The drawing revision is {pdf_revision}."
            )
            evidence.append(pdf_source)

    elif "part number" in question_lower:
        answer_parts.append(
            f"The selected part number is {part_number}."
        )

        if excel_result.get("status") == "success":
            evidence.append(excel_source)

    if not answer_parts:
        return {
            "answer": (
                "I cannot find that information in the "
                "supplied files. I will not guess."
            ),
            "evidence": [],
        }

    return {
        "answer": " ".join(answer_parts),
        "evidence": list(dict.fromkeys(evidence)),
    }


# ------------------------------------------------------------
# User interface
# ------------------------------------------------------------

st.title("⚙️ CNC Part Intelligence")

st.caption(
    "Gemini-powered, evidence-grounded analysis of Excel "
    "manufacturing records, PDF engineering drawings, "
    "and IGES 3D models."
)

with st.sidebar:
    st.header("Analysis settings")

    use_demo = st.checkbox(
        "Use team-created demonstration dataset",
        value=True,
    )

    part_number = st.text_input(
        "Part number",
        value="2280",
    )

    if not use_demo:
        uploaded_excel = st.file_uploader(
            "Upload Excel manufacturing record",
            type=["xlsx", "xls"],
        )

        uploaded_pdf = st.file_uploader(
            "Upload PDF engineering drawing",
            type=["pdf"],
        )

        uploaded_iges = st.file_uploader(
            "Upload IGES 3D model",
            type=["igs", "iges"],
        )
    else:
        uploaded_excel = None
        uploaded_pdf = None
        uploaded_iges = None

    analyze_button = st.button(
        "Analyze Part",
        type="primary",
        width="stretch",
    )

st.info(
    "The application reports missing information instead of "
    "inventing an answer. Responses identify their evidence."
)

if analyze_button:
    if not part_number.strip():
        st.error("Enter a part number.")
        st.stop()

    temporary_directory = tempfile.mkdtemp(
        prefix="cnc_analysis_"
    )

    try:
        temporary_path = Path(temporary_directory)

        if use_demo:
            excel_path = (
                DEMO_DIR / "manufacturing_records.xlsx"
            )
            pdf_path = (
                DEMO_DIR / "2280_engineering_drawing.pdf"
            )
            iges_path = (
                DEMO_DIR / "2280_model.iges"
            )

            missing_files = [
                str(path)
                for path in [
                    excel_path,
                    pdf_path,
                    iges_path,
                ]
                if not path.exists()
            ]

            if missing_files:
                st.error(
                    "Demonstration files are missing:\n"
                    + "\n".join(missing_files)
                )
                st.stop()

        else:
            if not all(
                [
                    uploaded_excel,
                    uploaded_pdf,
                    uploaded_iges,
                ]
            ):
                st.error(
                    "Upload the Excel, PDF, and IGES files."
                )
                st.stop()

            excel_path = (
                temporary_path / uploaded_excel.name
            )
            pdf_path = temporary_path / uploaded_pdf.name
            iges_path = temporary_path / uploaded_iges.name

            save_uploaded_file(
                uploaded_excel,
                excel_path,
            )
            save_uploaded_file(
                uploaded_pdf,
                pdf_path,
            )
            save_uploaded_file(
                uploaded_iges,
                iges_path,
            )

        with st.spinner(
            "Reading Excel manufacturing record..."
        ):
            excel_result = read_excel_record(
                excel_path,
                part_number,
            )

        with st.spinner(
            "Reading PDF engineering drawing..."
        ):
            pdf_result = read_pdf(pdf_path)

        with st.spinner(
            "Analyzing IGES geometry with FreeCAD..."
        ):
            iges_result = analyze_iges(
                iges_path,
                temporary_path,
            )

        comparison_rows = compare_dimensions(
            excel_result,
            pdf_result,
            iges_result,
        )

        st.session_state["analysis"] = {
            "part_number": part_number,
            "excel": excel_result,
            "pdf": pdf_result,
            "iges": iges_result,
            "comparisons": comparison_rows,
        }

    finally:
        shutil.rmtree(
            temporary_directory,
            ignore_errors=True,
        )


if "analysis" in st.session_state:
    analysis = st.session_state["analysis"]

    excel_result = analysis["excel"]
    pdf_result = analysis["pdf"]
    iges_result = analysis["iges"]
    comparison_rows = analysis["comparisons"]

    st.success(
        f"Analysis completed for Part "
        f"{analysis['part_number']}."
    )

    st.subheader("Source status")

    source_column_1, source_column_2, source_column_3 = (
        st.columns(3)
    )

    source_column_1.metric(
        "Excel",
        excel_result.get(
            "status",
            "unknown",
        ).upper(),
    )

    source_column_2.metric(
        "PDF",
        pdf_result.get(
            "status",
            "unknown",
        ).upper(),
    )

    source_column_3.metric(
        "IGES",
        iges_result.get(
            "status",
            "unknown",
        ).upper(),
    )

    tabs = st.tabs(
        [
            "Manufacturing Record",
            "Engineering Drawing",
            "3D Geometry",
            "Cross-Check",
            "Ask Questions",
        ]
    )

    with tabs[0]:
        st.subheader("Excel manufacturing record")

        if excel_result.get("status") == "success":
            st.dataframe(
                pd.DataFrame(
                    [excel_result["record"]]
                ),
                width="stretch",
            )

            st.caption(
                "Evidence: "
                f"{excel_result['source_file']}, "
                f"sheet {excel_result['sheet']}, "
                f"row {excel_result['excel_row']}"
            )
        else:
            st.warning(
                excel_result.get(
                    "message",
                    "Manufacturing record unavailable.",
                )
            )

    with tabs[1]:
        st.subheader("PDF engineering drawing")

        if pdf_result.get("facts"):
            st.json(pdf_result["facts"])
        else:
            st.warning(
                pdf_result.get(
                    "message",
                    "No structured drawing facts found.",
                )
            )

        with st.expander(
            "Show extracted drawing text"
        ):
            st.text(
                pdf_result.get(
                    "full_text",
                    "No text extracted.",
                )
            )

        st.caption(
            f"Evidence: "
            f"{pdf_result.get('source_file')}, "
            f"{pdf_result.get('page_count', 0)} page(s)"
        )

    with tabs[2]:
        st.subheader("IGES geometry")

        if iges_result.get("status") == "success":
            box = iges_result["bounding_box_mm"]

            geometry_columns = st.columns(3)

            geometry_columns[0].metric(
                "Bounding box",
                (
                    f"{box['length_x']:g} × "
                    f"{box['width_y']:g} × "
                    f"{box['height_z']:g} mm"
                ),
            )

            geometry_columns[1].metric(
                "Surface area",
                (
                    f"{iges_result['surface_area_mm2']:,.3f} "
                    "mm²"
                ),
            )

            geometry_columns[2].metric(
                "Reported volume",
                (
                    f"{iges_result['volume_mm3']:,.3f} "
                    "mm³"
                ),
            )

            st.json(iges_result)

            if iges_result.get("warning"):
                st.warning(
                    iges_result["warning"]
                )

            st.caption(
                "Evidence: "
                f"{iges_result.get('source_file')}"
            )
        else:
            st.error(
                iges_result.get(
                    "message",
                    "IGES analysis failed.",
                )
            )

    with tabs[3]:
        st.subheader("Cross-source validation")

        st.dataframe(
            pd.DataFrame(comparison_rows),
            width="stretch",
            hide_index=True,
        )

        conflicts = [
            row
            for row in comparison_rows
            if row["Status"] == "CONFLICT"
        ]

        if conflicts:
            st.error(
                f"{len(conflicts)} cross-source "
                "conflict(s) detected."
            )
        else:
            st.success(
                "No dimensional conflict was detected "
                "in the available sources."
            )

    with tabs[4]:
        st.subheader("Evidence-grounded questions")

        st.write("Try asking:")

        st.code(
            "What material is this part?\n"
            "What are the overall dimensions?\n"
            "Do the three sources agree?\n"
            "What is the surface finish?\n"
            "What is the model volume?\n"
            "What is the manufacturing process?"
        )

        question = st.text_input(
            "Ask a manufacturing or engineering question"
        )

        if st.button("Answer question"):
            if not question.strip():
                st.warning("Enter a question.")

            else:
                with st.spinner(
                    "Gemini is reviewing the extracted evidence..."
                ):
                    llm_response = answer_with_gemini(
                        question,
                        analysis["part_number"],
                        excel_result,
                        pdf_result,
                        iges_result,
                        comparison_rows,
                    )

                if (
                    llm_response.get("status")
                    == "success"
                ):
                    st.markdown("### LLM answer")
                    st.markdown(
                        llm_response["answer"]
                    )

                    model_used = llm_response.get(
                        "model",
                        "Gemini",
                    )

                    st.success(
                        f"Answered by {model_used} using "
                        "evidence extracted from the "
                        "supplied files."
                    )

                else:
                    fallback_response = answer_question(
                        question,
                        analysis["part_number"],
                        excel_result,
                        pdf_result,
                        iges_result,
                        comparison_rows,
                    )

                    st.markdown(
                        "### Offline fallback answer"
                    )
                    st.write(
                        fallback_response["answer"]
                    )

                    st.markdown("### Evidence")

                    if fallback_response["evidence"]:
                        for source in (
                            fallback_response["evidence"]
                        ):
                            st.write(f"- {source}")
                    else:
                        st.write(
                            "No supporting evidence was "
                            "found in the supplied files."
                        )

                    st.warning(
                        "Gemini was unavailable, so the "
                        "application used its deterministic "
                        "offline fallback. No answer was "
                        "invented."
                    )

                    with st.expander(
                        "Technical error details"
                    ):
                        st.code(
                            llm_response.get(
                                "message",
                                "Unknown Gemini error.",
                            )
                        )