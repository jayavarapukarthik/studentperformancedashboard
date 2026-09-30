import streamlit as st
import pandas as pd
from pathlib import Path
import html
import re
from io import BytesIO
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak
)
# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Student Academic Performance Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# FILE PATHS
# Keep all Excel files in the same folder as app.py
# ============================================================

BASE_DIR = Path(__file__).parent

INSEM_FILE = (
    BASE_DIR /
    "CSE-4-In-sem-1-Marks-all-regulations (2).xls"
)

MENTOR_FILE = (
    BASE_DIR /
    "CSE-4-Mentor-Mentee-data.xlsx"
)

Y23_FILE = (
    BASE_DIR /
    "Y23 Result Analysis.xlsx"
)

Y24_FILE = (
    BASE_DIR /
    "Y24 Result Analysis.xlsx"
)

Y25_FILE = (
    BASE_DIR /
    "Y25 Result Analysis.xlsx"
)

# NEW: CGPA FILE
CGPA_FILE = (
    BASE_DIR /
    "CGPA.xlsx"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.html("""
<style>

/* ============================================================
   MAIN APPLICATION
   ============================================================ */

.stApp {
    background:
        linear-gradient(
            135deg,
            #f8fafc 0%,
            #eef4f9 50%,
            #f8fafc 100%
        );
}

.block-container {
    max-width: 1250px;
    padding-top: 3.5rem;
    padding-bottom: 2rem;
}


/* ============================================================
   PREMIUM UNIVERSITY HEADER
   ============================================================ */

.university-header {
    position: relative;
    overflow: hidden;

    background:
        linear-gradient(
            135deg,
            #ffffff 0%,
            #f8fbff 45%,
            #eef5fb 100%
        );

    padding: 38px 35px 30px 35px;

    border-radius: 0 0 28px 28px;

    border: 1px solid #dbe5ef;

    box-shadow:
        0 12px 35px rgba(15, 52, 86, 0.10);

    margin-bottom: 22px;

    text-align: center;
}


/* TOP DECORATIVE LINE */

.university-header::before {
    content: "";

    position: absolute;

    top: 0;
    left: 0;

    width: 100%;
    height: 6px;

    background:
        linear-gradient(
            90deg,
            #0f3d68,
            #1769aa,
            #4f9bd1,
            #1769aa,
            #0f3d68
        );
}


/* DECORATIVE RIGHT GLOW */

.university-header::after {
    content: "";

    position: absolute;

    width: 320px;
    height: 320px;

    right: -140px;
    top: -180px;

    border-radius: 50%;

    background:
        radial-gradient(
            circle,
            rgba(23,105,170,0.13),
            rgba(23,105,170,0)
        );

    pointer-events: none;
}


/* UNIVERSITY NAME */

.university-name {
    position: relative;
    z-index: 2;

    font-size: 38px;

    font-weight: 900;

    color: #123f6b;

    letter-spacing: 1px;

    line-height: 1.15;

    text-transform: uppercase;

    margin-top: 5px;
}


/* HEADER DIVIDER */

.header-divider {
    width: 95px;

    height: 4px;

    margin: 15px auto 13px auto;

    border-radius: 10px;

    background:
        linear-gradient(
            90deg,
            #123f6b,
            #3c8ac4
        );
}


/* DEPARTMENT */

.department-name {
    position: relative;

    z-index: 2;

    font-size: 23px;

    font-weight: 800;

    color: #475569;

    letter-spacing: 1.8px;

    text-transform: uppercase;

    line-height: 1.3;
}


/* PORTAL NAME */

.portal-name {
    position: relative;

    z-index: 2;

    display: inline-block;

    margin-top: 18px;

    padding: 12px 30px;

    font-size: 29px;

    font-weight: 900;

    color: #ffffff;

    letter-spacing: 1px;

    line-height: 1.2;

    text-transform: uppercase;

    border-radius: 12px;

    background:
        linear-gradient(
            135deg,
            #0f3d68,
            #1d6096
        );

    box-shadow:
        0 7px 18px rgba(15,61,104,0.20);
}


/* HEADER SUBTITLE */

.header-subtitle {
    position: relative;

    z-index: 2;

    margin-top: 11px;

    font-size: 12px;

    font-weight: 650;

    color: #64748b;

    letter-spacing: 2px;

    text-transform: uppercase;
}


/* ============================================================
   SEARCH CARD
   ============================================================ */

.search-card {
    background: white;

    padding: 17px 25px 13px 25px;

    border-radius: 17px;

    border: 1px solid #e2e8f0;

    box-shadow:
        0 6px 22px rgba(15, 23, 42, 0.06);

    margin-bottom: 10px;
}

.search-title {
    font-size: 21px;

    font-weight: 800;

    color: #1e293b;

    margin-bottom: 3px;
}

.search-subtitle {
    font-size: 14px;

    color: #64748b;
}


/* ============================================================
   SEARCH INPUT
   ============================================================ */

div[data-baseweb="input"] {
    border-radius: 11px !important;

    border: 1px solid #cbd5e1 !important;

    background: white !important;
}

div[data-baseweb="input"]:focus-within {
    border-color: #123b68 !important;

    box-shadow:
        0 0 0 2px rgba(18, 59, 104, 0.12) !important;
}

div[data-baseweb="input"] input {
    font-size: 17px !important;

    padding: 12px !important;

    color: #0f172a !important;
}


/* ============================================================
   SEARCH BUTTON
   ============================================================ */

.stFormSubmitButton > button {
    width: 100%;

    height: 48px;

    border-radius: 11px;

    border: none;

    font-size: 16px;

    font-weight: 750;

    color: white;

    background:
        linear-gradient(
            135deg,
            #123b68,
            #1d5d91
        );

    box-shadow:
        0 5px 14px rgba(18, 59, 104, 0.20);

    transition: all 0.2s ease;
}

.stFormSubmitButton > button:hover {
    transform: translateY(-1px);

    box-shadow:
        0 8px 18px rgba(18, 59, 104, 0.26);
}


/* ============================================================
   STUDENT DETAILS
   ============================================================ */

.student-details-container {
    background: white;

    border-radius: 17px;

    border: 1px solid #e2e8f0;

    border-left: 6px solid #123b68;

    box-shadow:
        0 7px 24px rgba(15, 23, 42, 0.07);

    padding: 16px 22px;

    margin-top: 17px;

    margin-bottom: 17px;
}

.details-heading {
    font-size: 18px;

    font-weight: 800;

    color: #123b68;

    margin-bottom: 12px;
}

.detail-label {
    font-size: 11px;

    font-weight: 800;

    color: #64748b;

    text-transform: uppercase;

    letter-spacing: 1px;

    margin-bottom: 3px;
}

.detail-value {
    font-size: 18px;

    font-weight: 750;

    color: #0f172a;

    line-height: 1.3;
}

.detail-value-id {
    color: #123b68;
}

/* NEW: CGPA STYLE */

.cgpa-value {
    color: #047857;

    font-size: 19px;

    font-weight: 850;
}


/* ============================================================
   SECTION HEADINGS
   ============================================================ */

.section-heading {
    font-size: 23px;

    font-weight: 800;

    color: #0f172a;

    margin: 10px 0 13px 0;
}


/* ============================================================
   GENERAL TABLE
   ============================================================ */

.table-wrapper {
    width: 100%;

    overflow-x: auto;

    border-radius: 16px;

    box-shadow:
        0 7px 25px rgba(15, 23, 42, 0.08);

    margin-bottom: 18px;
}

.performance-table {
    width: 100%;

    border-collapse: separate;

    border-spacing: 0;

    background: white;

    border: 1px solid #e2e8f0;

    border-radius: 16px;

    overflow: hidden;
}

.performance-table th {
    background: #123b68;

    color: white;

    padding: 16px 15px;

    font-size: 15px;

    font-weight: 800;

    text-align: left;

    white-space: nowrap;
}

.performance-table th:first-child {
    text-align: center;
}

.performance-table td {
    padding: 15px;

    border-bottom: 1px solid #e8edf3;

    color: #334155;

    font-size: 16px;

    background: white;

    vertical-align: middle;
}

.performance-table tr:last-child td {
    border-bottom: none;
}

.performance-table tr:hover td {
    background: #f8fafc;
}


/* ============================================================
   TABLE VALUES
   ============================================================ */

.serial-number {
    text-align: center;

    font-weight: 750;

    color: #64748b;

    font-size: 16px;
}

.course-code {
    font-weight: 750;

    color: #123b68;

    white-space: nowrap;

    font-size: 16px;
}

.course-title {
    font-weight: 600;

    color: #334155;

    font-size: 16px;
}

.marks {
    font-weight: 850;

    color: #047857;

    white-space: nowrap;

    font-size: 20px;
}

.grade-value {
    font-weight: 800;

    color: #123b68;

    font-size: 17px;
}

.category-value {
    font-weight: 750;

    color: #334155;

    font-size: 16px;
}


/* ============================================================
   SEMESTER HEADER
   ============================================================ */

.semester-header {
    background: white;

    border-left: 6px solid #123b68;

    border-radius: 14px;

    padding: 13px 20px;

    margin-top: 24px;

    margin-bottom: 12px;

    border-top: 1px solid #e2e8f0;

    border-right: 1px solid #e2e8f0;

    border-bottom: 1px solid #e2e8f0;

    box-shadow:
        0 5px 18px rgba(15, 23, 42, 0.06);
}

.semester-title {
    font-size: 20px;

    font-weight: 800;

    color: #123b68;
}

.semester-subtitle {
    font-size: 13px;

    color: #64748b;

    margin-top: 2px;
}


/* ============================================================
   ACADEMIC PERFORMANCE HIGHLIGHTS
   ============================================================ */

.academic-highlights {
    margin-top: 12px;

    margin-bottom: 12px;
}

.highlights-title {
    font-size: 22px;

    font-weight: 850;

    color: #123b68;

    margin-bottom: 12px;
}

.highlight-table-wrapper {
    width: 100%;

    overflow-x: auto;

    border-radius: 15px;

    box-shadow:
        0 7px 22px rgba(15, 23, 42, 0.07);
}

.highlight-table {
    width: 100%;

    border-collapse: separate;

    border-spacing: 0;

    background: white;

    border: 1px solid #e2e8f0;

    border-radius: 15px;

    overflow: hidden;
}

.highlight-table th {
    background: #123b68;

    color: white;

    padding: 14px 11px;

    font-size: 12px;

    font-weight: 800;

    text-align: center;

    white-space: nowrap;
}

.highlight-table td {
    background: white;

    padding: 16px 11px;

    text-align: center;

    font-size: 24px;

    font-weight: 850;

    border-right: 1px solid #e2e8f0;
}

.highlight-table td:last-child {
    border-right: none;
}

.highlight-table .normal-value {
    color: #123b68;
}

.highlight-table .pass-value {
    color: #047857;
}

.highlight-table .backlog-value {
    color: #dc2626;
}


/* ============================================================
   SUMMARY CARDS
   ============================================================ */

.summary-card {
    background: white;

    padding: 16px 20px;

    border-radius: 15px;

    border: 1px solid #e2e8f0;

    box-shadow:
        0 6px 20px rgba(15, 23, 42, 0.06);

    margin-top: 18px;
}

.summary-label {
    font-size: 11px;

    color: #64748b;

    text-transform: uppercase;

    font-weight: 750;

    letter-spacing: 0.8px;
}

.summary-value {
    font-size: 22px;

    color: #123b68;

    font-weight: 800;

    margin-top: 4px;
}


/* ============================================================
   FOOTER
   ============================================================ */

.footer {
    text-align: center;

    color: #64748b;

    font-size: 12px;

    margin-top: 28px;

    padding-top: 14px;

    border-top: 1px solid #e2e8f0;
}


/* ============================================================
   RESPONSIVE
   ============================================================ */

@media (max-width: 900px) {

    .university-header {
        padding: 32px 20px 27px 20px;
    }

    .university-name {
        font-size: 29px;
    }

    .department-name {
        font-size: 21px;
    }

    .portal-name {
        font-size: 24px;

        padding: 10px 20px;
    }

    .highlight-table th {
        font-size: 11px;

        padding: 12px 9px;
    }

    .highlight-table td {
        font-size: 20px;

        padding: 14px 9px;
    }
}


@media (max-width: 600px) {

    .university-header {
        padding: 29px 15px 24px 15px;

        border-radius: 0 0 20px 20px;
    }

    .university-name {
        font-size: 22px;

        letter-spacing: 0.5px;
    }

    .department-name {
        font-size: 16px;

        letter-spacing: 1px;
    }

    .portal-name {
        font-size: 18px;

        padding: 9px 14px;
    }

    .header-subtitle {
        font-size: 10px;
    }

    .highlight-table th {
        font-size: 10px;
    }

    .highlight-table td {
        font-size: 18px;
    }
}

</style>
""")


# ============================================================
# GENERAL COLUMN FINDER
# ============================================================

def find_column(df, possible_names):

    normalized_columns = {}

    for column in df.columns:

        normalized = (
            str(column)
            .lower()
            .replace("\n", " ")
            .replace("\r", " ")
            .strip()
        )

        normalized_columns[normalized] = column

    # Exact match

    for name in possible_names:

        target = name.lower().strip()

        if target in normalized_columns:
            return normalized_columns[target]

    # Partial match

    for column in df.columns:

        normalized = (
            str(column)
            .lower()
            .replace("\n", " ")
            .replace("\r", " ")
            .strip()
        )

        for name in possible_names:

            target = name.lower().strip()

            if target in normalized:
                return column

    return None


# ============================================================
# CLEAN STUDENT ID
# ============================================================

def clean_student_id(series):

    return (
        series
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )


# ============================================================
# LOAD IN-SEMESTER DATA
# ============================================================

@st.cache_data
def load_insem_data():

    df = pd.read_excel(
        INSEM_FILE,
        sheet_name="Final-Insem-1-Marks",
        engine="xlrd"
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    id_column = find_column(
        df,
        [
            "University Id",
            "University ID",
            "ID Number",
            "Student ID"
        ]
    )

    if id_column is None:

        raise ValueError(
            "University ID column was not found "
            "in the In-Semester file."
        )

    df = df.rename(
        columns={
            id_column: "University Id"
        }
    )

    df["University Id"] = clean_student_id(
        df["University Id"]
    )

    total_column = find_column(
        df,
        [
            "Total",
            "Total Marks",
            "In Sem Total",
            "In-Sem Total"
        ]
    )

    if total_column is None:

        raise ValueError(
            "Total marks column was not found "
            "in the In-Semester file."
        )

    df = df.rename(
        columns={
            total_column: "Total"
        }
    )

    df["Total"] = pd.to_numeric(
        df["Total"],
        errors="coerce"
    )

    course_code_column = find_column(
        df,
        [
            "Course Code",
            "CourseCode"
        ]
    )

    if course_code_column is not None:

        df = df.rename(
            columns={
                course_code_column:
                "Course Code"
            }
        )

    course_name_column = find_column(
        df,
        [
            "Course Name",
            "Course Title",
            "CourseName"
        ]
    )

    if course_name_column is not None:

        df = df.rename(
            columns={
                course_name_column:
                "Course Name"
            }
        )

    return df


# ============================================================
# LOAD MENTOR-MENTEE DATA
# ============================================================

@st.cache_data
def load_mentor_data():

    df = pd.read_excel(
        MENTOR_FILE,
        sheet_name=0
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.replace(
            "\n",
            " ",
            regex=False
        )
        .str.replace(
            "\r",
            " ",
            regex=False
        )
        .str.strip()
    )

    student_id_column = find_column(
        df,
        [
            "Student Unique Enrollment ID",
            "Unique Enrollment ID",
            "Student ID"
        ]
    )

    student_name_column = find_column(
        df,
        [
            "Name of the student",
            "Student Name"
        ]
    )

    mentor_column = find_column(
        df,
        [
            "Name of the Mentor",
            "Mentor Name",
            "Mentor",
            "Counsellor"
        ]
    )

    emp_id_column = find_column(
        df,
        [
            "EMP ID",
            "Emp ID",
            "Employee ID",
            "Employee Id",
            "EMPID",
            "EmpID"
        ]
    )

    regulation_column = find_column(
        df,
        [
            "Reg",
            "Regulation"
        ]
    )

    rename_map = {}

    if student_id_column is not None:
        rename_map[student_id_column] = "Student ID"

    if student_name_column is not None:
        rename_map[student_name_column] = "Student Name"

    if mentor_column is not None:
        rename_map[mentor_column] = "Counsellor"

    if emp_id_column is not None:
        rename_map[emp_id_column] = "EMP ID"

    if regulation_column is not None:
        rename_map[regulation_column] = "Regulation"

    df = df.rename(
        columns=rename_map
    )

    # Current Excel structure fallback:
    # A = Reg
    # B = Student Unique Enrollment ID
    # C = Name of the student
    # D = EMP ID
    # E = Name of the Mentor

    if "Student ID" not in df.columns:

        if len(df.columns) >= 5:

            columns = list(df.columns)

            df = df.rename(
                columns={
                    columns[0]:
                    "Regulation",

                    columns[1]:
                    "Student ID",

                    columns[2]:
                    "Student Name",

                    columns[3]:
                    "EMP ID",

                    columns[4]:
                    "Counsellor"
                }
            )

    if "Student ID" in df.columns:

        df["Student ID"] = clean_student_id(
            df["Student ID"]
        )

    return df


# ============================================================
# LOAD RESULT FILE
# ============================================================

@st.cache_data
def load_result_file(
    file_path,
    regulation
):

    df = pd.read_excel(
        file_path,
        sheet_name=0
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.replace(
            "\n",
            " ",
            regex=False
        )
        .str.replace(
            "\r",
            " ",
            regex=False
        )
        .str.strip()
    )

    id_column = find_column(
        df,
        [
            "ID Number",
            "ID",
            "University ID",
            "University Id",
            "Student ID"
        ]
    )

    name_column = find_column(
        df,
        [
            "Name",
            "Student Name",
            "Name of the student"
        ]
    )

    course_code_column = find_column(
        df,
        [
            "Course Code"
        ]
    )

    course_name_column = find_column(
        df,
        [
            "Course Name",
            "Course Title"
        ]
    )

    grade_column = find_column(
        df,
        [
            "Grade"
        ]
    )

    category_column = find_column(
        df,
        [
            "Category"
        ]
    )

    ay_column = find_column(
        df,
        [
            "AY",
            "Academic Year",
            "AcademicYear"
        ]
    )

    semester_column = find_column(
        df,
        [
            "Semester",
            "Sem"
        ]
    )

    missing = []

    if id_column is None:
        missing.append("ID Number")

    if course_code_column is None:
        missing.append("Course Code")

    if course_name_column is None:
        missing.append("Course Name")

    if grade_column is None:
        missing.append("Grade")

    if category_column is None:
        missing.append("Category")

    if ay_column is None:
        missing.append("AY")

    if semester_column is None:
        missing.append("Semester")

    if missing:

        raise ValueError(
            f"{regulation} result file is missing columns: "
            + ", ".join(missing)
        )

    rename_map = {

        id_column:
        "Student ID",

        course_code_column:
        "Course Code",

        course_name_column:
        "Course Name",

        grade_column:
        "Grade",

        category_column:
        "Category",

        ay_column:
        "AY",

        semester_column:
        "Semester"
    }

    if name_column is not None:

        rename_map[
            name_column
        ] = "Name"

    df = df.rename(
        columns=rename_map
    )

    df["Student ID"] = clean_student_id(
        df["Student ID"]
    )

    df["AY"] = (
        df["AY"]
        .astype(str)
        .str.strip()
    )

    df["Semester"] = (
        df["Semester"]
        .astype(str)
        .str.strip()
    )

    df["Grade"] = (
        df["Grade"]
        .astype(str)
        .str.strip()
    )

    df["Category"] = (
        df["Category"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["Course Code"] = (
        df["Course Code"]
        .astype(str)
        .str.strip()
    )

    df["Course Name"] = (
        df["Course Name"]
        .astype(str)
        .str.strip()
    )

    df["Regulation"] = regulation

    return df


# ============================================================
# LOAD ALL RESULT FILES
# ============================================================

def load_all_results():

    frames = []

    if Y23_FILE.exists():

        frames.append(
            load_result_file(
                Y23_FILE,
                "Y-23"
            )
        )

    if Y24_FILE.exists():

        frames.append(
            load_result_file(
                Y24_FILE,
                "Y-24"
            )
        )

    if Y25_FILE.exists():

        frames.append(
            load_result_file(
                Y25_FILE,
                "Y-25"
            )
        )

    if not frames:

        raise FileNotFoundError(
            "No Result Analysis Excel files were found."
        )

    return pd.concat(
        frames,
        ignore_index=True
    )


# ============================================================
# LOAD CGPA DATA
# ============================================================

@st.cache_data
def load_cgpa_data():

    if not CGPA_FILE.exists():

        return pd.DataFrame(
            columns=[
                "Student ID",
                "CGPA"
            ]
        )

    df = pd.read_excel(
        CGPA_FILE,
        sheet_name=0
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.replace(
            "\n",
            " ",
            regex=False
        )
        .str.replace(
            "\r",
            " ",
            regex=False
        )
        .str.strip()
    )

    student_id_column = find_column(
        df,
        [
            "Student ID",
            "StudentID",
            "ID Number",
            "University ID",
            "University Id"
        ]
    )

    cgpa_column = find_column(
        df,
        [
            "CGPA",
            "C.G.P.A"
        ]
    )

    if student_id_column is None:

        raise ValueError(
            "Student ID column was not found "
            "in CGPA.xlsx."
        )

    if cgpa_column is None:

        raise ValueError(
            "CGPA column was not found "
            "in CGPA.xlsx."
        )

    df = df.rename(
        columns={
            student_id_column:
            "Student ID",

            cgpa_column:
            "CGPA"
        }
    )

    df["Student ID"] = clean_student_id(
        df["Student ID"]
    )

    df["CGPA"] = pd.to_numeric(
        df["CGPA"],
        errors="coerce"
    )

    return (
        df[
            [
                "Student ID",
                "CGPA"
            ]
        ]
        .drop_duplicates(
            subset=["Student ID"],
            keep="first"
        )
    )

# ============================================================
# PARENT PDF REPORT GENERATOR
# ============================================================

def _pdf_text(value):

    if value is None:
        return "-"

    try:
        if pd.isna(value):
            return "-"
    except Exception:
        pass

    return str(value).strip()


def _pdf_paragraph(value, style):

    safe_text = html.escape(
        _pdf_text(value)
    )

    return Paragraph(
        safe_text,
        style
    )


def build_parent_report_pdf(
    student_id,
    student_name,
    counsellor_name,
    emp_id,
    regulation,
    cgpa_value,
    student_data,
    result_data,
    total_courses,
    passed_courses,
    backlog_courses,
    category_counts,
    ordered_categories
):

    """
    Generate complete Parent Academic Report PDF.
    """

    buffer = BytesIO()

    # ========================================================
    # PDF DOCUMENT
    # ========================================================

    doc = SimpleDocTemplate(

        buffer,

        pagesize=A4,

        rightMargin=12 * mm,

        leftMargin=12 * mm,

        topMargin=12 * mm,

        bottomMargin=14 * mm,

        title=(
            f"Parent Academic Report - "
            f"{student_id}"
        ),

        author=(
            "Department of CSE-4, KLEF"
        )
    )

    styles = getSampleStyleSheet()

    # ========================================================
    # PDF STYLES
    # ========================================================

    title_style = ParagraphStyle(

        "KlefTitle",

        parent=styles["Title"],

        fontName="Helvetica-Bold",

        fontSize=16,

        leading=20,

        textColor=colors.HexColor(
            "#123b68"
        ),

        alignment=TA_CENTER,

        spaceAfter=3
    )

    dept_style = ParagraphStyle(

        "Dept",

        parent=styles["Normal"],

        fontName="Helvetica-Bold",

        fontSize=10,

        leading=13,

        textColor=colors.HexColor(
            "#475569"
        ),

        alignment=TA_CENTER,

        spaceAfter=4
    )

    portal_style = ParagraphStyle(

        "Portal",

        parent=styles["Normal"],

        fontName="Helvetica-Bold",

        fontSize=11,

        leading=14,

        textColor=colors.white,

        alignment=TA_CENTER
    )

    section_style = ParagraphStyle(

        "Section",

        parent=styles["Heading2"],

        fontName="Helvetica-Bold",

        fontSize=11,

        leading=14,

        textColor=colors.HexColor(
            "#123b68"
        ),

        spaceBefore=8,

        spaceAfter=6
    )

    body_style = ParagraphStyle(

        "Body",

        parent=styles["Normal"],

        fontName="Helvetica",

        fontSize=8.5,

        leading=11,

        textColor=colors.HexColor(
            "#334155"
        )
    )

    small_style = ParagraphStyle(

        "Small",

        parent=body_style,

        fontSize=7.5,

        leading=9.5
    )

    small_bold_style = ParagraphStyle(

        "SmallBold",

        parent=small_style,

        fontName="Helvetica-Bold"
    )

    story = []

    # ========================================================
    # KLEF HEADER
    # ========================================================

    story.append(

        Paragraph(

            "KONERU LAKSHMAIAH EDUCATION FOUNDATION",

            title_style

        )

    )

    story.append(

        Paragraph(

            "DEPARTMENT OF CSE-4",

            dept_style

        )

    )

    portal_table = Table(

        [[

            Paragraph(

                "STUDENT ACADEMIC PERFORMANCE PORTAL",

                portal_style

            )

        ]],

        colWidths=[
            186 * mm
        ]

    )

    portal_table.setStyle(

        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (-1, -1),
                colors.HexColor(
                    "#123b68"
                )
            ),

            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor(
                    "#123b68"
                )
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                7
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                7
            )

        ])

    )

    story.append(
        portal_table
    )

    story.append(
        Spacer(
            1,
            2 * mm
        )
    )

    story.append(

        Paragraph(

            "Academic Excellence • "
            "Performance Monitoring • "
            "Student Success",

            ParagraphStyle(

                "Tagline",

                parent=body_style,

                alignment=TA_CENTER,

                fontSize=7.5,

                textColor=colors.HexColor(
                    "#64748b"
                )

            )

        )

    )

    story.append(
        Spacer(
            1,
            6 * mm
        )
    )

    # ========================================================
    # STUDENT DETAILS
    # ========================================================

    story.append(

        Paragraph(

            "STUDENT DETAILS",

            section_style

        )

    )

    detail_data = [

        [

            _pdf_paragraph(
                "Student ID",
                small_bold_style
            ),

            _pdf_paragraph(
                student_id,
                body_style
            ),

            _pdf_paragraph(
                "Student Name",
                small_bold_style
            ),

            _pdf_paragraph(
                student_name,
                body_style
            )

        ],

        [

            _pdf_paragraph(
                "Counsellor",
                small_bold_style
            ),

            _pdf_paragraph(
                counsellor_name,
                body_style
            ),

            _pdf_paragraph(
                "Emp ID",
                small_bold_style
            ),

            _pdf_paragraph(
                emp_id,
                body_style
            )

        ],

        [

            _pdf_paragraph(
                "Regulation",
                small_bold_style
            ),

            _pdf_paragraph(
                regulation,
                body_style
            ),

            _pdf_paragraph(
                "CGPA",
                small_bold_style
            ),

            _pdf_paragraph(
                cgpa_value,
                body_style
            )

        ]

    ]

    detail_table = Table(

        detail_data,

        colWidths=[

            28 * mm,
            63 * mm,
            28 * mm,
            67 * mm

        ]

    )

    detail_table.setStyle(

        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (0, -1),
                colors.HexColor(
                    "#eef4f9"
                )
            ),

            (
                "BACKGROUND",
                (2, 0),
                (2, -1),
                colors.HexColor(
                    "#eef4f9"
                )
            ),

            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.5,
                colors.HexColor(
                    "#cbd5e1"
                )
            ),

            (
                "INNERGRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.HexColor(
                    "#dbe5ef"
                )
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                6
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                6
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                5
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                5
            )

        ])

    )

    story.append(
        detail_table
    )

    # ========================================================
    # ACADEMIC PERFORMANCE HIGHLIGHTS
    # ========================================================

    story.append(

        Paragraph(

            "ACADEMIC PERFORMANCE HIGHLIGHTS",

            section_style

        )

    )

    highlight_headers = [

        "Total Courses",
        "Passed",
        "Backlogs"

    ] + [

        _pdf_text(category)

        for category
        in ordered_categories

    ]

    highlight_values = [

        str(total_courses),
        str(passed_courses),
        str(backlog_courses)

    ] + [

        str(
            category_counts.get(
                category,
                0
            )
        )

        for category
        in ordered_categories

    ]

    highlight_count = len(
        highlight_headers
    )

    highlight_width = (
        186 * mm
        / highlight_count
    )

    highlight_table = Table(

        [

            [

                _pdf_paragraph(
                    value,
                    small_bold_style
                )

                for value
                in highlight_headers

            ],

            [

                _pdf_paragraph(
                    value,
                    body_style
                )

                for value
                in highlight_values

            ]

        ],

        colWidths=[
            highlight_width
        ] * highlight_count

    )

    highlight_table.setStyle(

        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor(
                    "#123b68"
                )
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.HexColor(
                    "#dbe5ef"
                )
            ),

            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER"
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                5
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                5
            )

        ])

    )

    story.append(
        highlight_table
    )

    # ========================================================
    # IN-SEMESTER PERFORMANCE
    # ========================================================

    if (

        isinstance(
            student_data,
            pd.DataFrame
        )

        and

        not student_data.empty

        and

        all(

            column in student_data.columns

            for column in [

                "Course Code",
                "Course Name",
                "Total"

            ]

        )

    ):

        story.append(

            Paragraph(

                "IN-SEMESTER PERFORMANCE",

                section_style

            )

        )

        insem_rows = [[

            _pdf_paragraph(
                "S.No",
                small_bold_style
            ),

            _pdf_paragraph(
                "Course Code",
                small_bold_style
            ),

            _pdf_paragraph(
                "Course Title",
                small_bold_style
            ),

            _pdf_paragraph(
                "In-Sem Marks / 50",
                small_bold_style
            )

        ]]

        for index, (_, row) in enumerate(

            student_data.iterrows(),

            start=1

        ):

            total = row["Total"]

            if pd.isna(total):

                marks_text = "-"

            else:

                try:

                    marks = float(total)

                    if marks.is_integer():

                        marks_text = (
                            f"{int(marks)} / 50"
                        )

                    else:

                        marks_text = (
                            f"{marks:g} / 50"
                        )

                except Exception:

                    marks_text = (
                        f"{_pdf_text(total)} / 50"
                    )

            insem_rows.append([

                _pdf_paragraph(
                    index,
                    small_style
                ),

                _pdf_paragraph(
                    row["Course Code"],
                    small_style
                ),

                _pdf_paragraph(
                    row["Course Name"],
                    small_style
                ),

                _pdf_paragraph(
                    marks_text,
                    small_style
                )

            ])

        insem_table = Table(

            insem_rows,

            colWidths=[

                14 * mm,
                34 * mm,
                98 * mm,
                40 * mm

            ],

            repeatRows=1

        )

        insem_table.setStyle(

            TableStyle([

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        "#123b68"
                    )
                ),

                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    colors.HexColor(
                        "#dbe5ef"
                    )
                ),

                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [

                        colors.white,

                        colors.HexColor(
                            "#f8fafc"
                        )

                    ]
                ),

                (
                    "ALIGN",
                    (0, 0),
                    (0, -1),
                    "CENTER"
                ),

                (
                    "ALIGN",
                    (3, 1),
                    (3, -1),
                    "CENTER"
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                )

            ])

        )

        story.append(
            insem_table
        )

    # ========================================================
    # SEMESTER-WISE PERFORMANCE
    # ========================================================

    story.append(
        PageBreak()
    )

    story.append(

        Paragraph(

            "SEMESTER-WISE ACADEMIC PERFORMANCE",

            section_style

        )

    )

    if (

        isinstance(
            result_data,
            pd.DataFrame
        )

        and

        not result_data.empty

    ):

        pdf_result = (
            result_data.copy()
        )

        if (
            "Semester Number"
            not in pdf_result.columns
        ):

            pdf_result[
                "Semester Number"
            ] = pdf_result.apply(

                lambda row:

                get_semester_number(

                    regulation,

                    row.get(
                        "AY",
                        ""
                    ),

                    row.get(
                        "Semester",
                        ""
                    )

                ),

                axis=1

            )

        semester_order = [

            "1-1",
            "1-2",
            "2-1",
            "2-2",
            "3-1",
            "3-2",
            "4-1",
            "4-2"

        ]

        for semester_number in semester_order:

            semester_data = (

                pdf_result[

                    pdf_result[
                        "Semester Number"
                    ]

                    == semester_number

                ]

                .copy()

            )

            if semester_data.empty:

                continue

            ay_value = (
                semester_data.iloc[0].get(
                    "AY",
                    "-"
                )
            )

            term_value = (
                semester_data.iloc[0].get(
                    "Semester",
                    "-"
                )
            )

            story.append(

                Paragraph(

                    f"SEMESTER "
                    f"{html.escape(_pdf_text(semester_number))}",

                    section_style

                )

            )

            story.append(

                Paragraph(

                    "Academic Year: "
                    f"{html.escape(_pdf_text(ay_value))}"
                    "  |  "
                    f"{html.escape(_pdf_text(term_value))}",

                    small_style

                )

            )

            semester_rows = [[

                _pdf_paragraph(
                    "S.No",
                    small_bold_style
                ),

                _pdf_paragraph(
                    "Course Code",
                    small_bold_style
                ),

                _pdf_paragraph(
                    "Course Name",
                    small_bold_style
                ),

                _pdf_paragraph(
                    "Grade",
                    small_bold_style
                ),

                _pdf_paragraph(
                    "Category",
                    small_bold_style
                )

            ]]

            for index, (_, row) in enumerate(

                semester_data.iterrows(),

                start=1

            ):

                semester_rows.append([

                    _pdf_paragraph(
                        index,
                        small_style
                    ),

                    _pdf_paragraph(
                        row.get(
                            "Course Code",
                            "-"
                        ),
                        small_style
                    ),

                    _pdf_paragraph(
                        row.get(
                            "Course Name",
                            "-"
                        ),
                        small_style
                    ),

                    _pdf_paragraph(
                        row.get(
                            "Grade",
                            "-"
                        ),
                        small_style
                    ),

                    _pdf_paragraph(
                        row.get(
                            "Category",
                            "-"
                        ),
                        small_style
                    )

                ])

            semester_table = Table(

                semester_rows,

                colWidths=[

                    12 * mm,
                    34 * mm,
                    93 * mm,
                    22 * mm,
                    25 * mm

                ],

                repeatRows=1

            )

            semester_table.setStyle(

                TableStyle([

                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.HexColor(
                            "#123b68"
                        )
                    ),

                    (
                        "TEXTCOLOR",
                        (0, 0),
                        (-1, 0),
                        colors.white
                    ),

                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.35,
                        colors.HexColor(
                            "#dbe5ef"
                        )
                    ),

                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [

                            colors.white,

                            colors.HexColor(
                                "#f8fafc"
                            )

                        ]
                    ),

                    (
                        "ALIGN",
                        (0, 0),
                        (0, -1),
                        "CENTER"
                    ),

                    (
                        "ALIGN",
                        (3, 1),
                        (-1, -1),
                        "CENTER"
                    ),

                    (
                        "VALIGN",
                        (0, 0),
                        (-1, -1),
                        "MIDDLE"
                    ),

                    (
                        "TOPPADDING",
                        (0, 0),
                        (-1, -1),
                        3
                    ),

                    (
                        "BOTTOMPADDING",
                        (0, 0),
                        (-1, -1),
                        3
                    )

                ])

            )

            story.append(
                semester_table
            )

            story.append(
                Spacer(
                    1,
                    3 * mm
                )
            )

    # ========================================================
    # COUNSELLING / REPORT INFORMATION
    # ========================================================

    story.append(

        Paragraph(

            "COUNSELLING / REPORT INFORMATION",

            section_style

        )

    )

    report_info_table = Table(

        [

            [

                _pdf_paragraph(
                    "Counsellor",
                    small_bold_style
                ),

                _pdf_paragraph(
                    counsellor_name,
                    body_style
                )

            ],

            [

                _pdf_paragraph(
                    "Emp ID",
                    small_bold_style
                ),

                _pdf_paragraph(
                    emp_id,
                    body_style
                )

            ],

            [

                _pdf_paragraph(
                    "Report Generated",
                    small_bold_style
                ),

                _pdf_paragraph(

                    datetime.now().strftime(
                        "%d-%m-%Y %I:%M %p"
                    ),

                    body_style

                )

            ]

        ],

        colWidths=[

            40 * mm,
            146 * mm

        ]

    )

    report_info_table.setStyle(

        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (0, -1),
                colors.HexColor(
                    "#eef4f9"
                )
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.HexColor(
                    "#dbe5ef"
                )
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                5
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                5
            )

        ])

    )

    story.append(
        report_info_table
    )

    story.append(
        Spacer(
            1,
            6 * mm
        )
    )

    story.append(

        Paragraph(

            "This report is generated from the "
            "academic records available in the "
            "Student Academic Performance Portal.",

            ParagraphStyle(

                "Disclaimer",

                parent=small_style,

                alignment=TA_CENTER,

                textColor=colors.HexColor(
                    "#64748b"
                )

            )

        )

    )

    # ========================================================
    # PAGE NUMBER
    # ========================================================

    def add_page_number(
        canvas,
        doc
    ):

        canvas.saveState()

        canvas.setFont(
            "Helvetica",
            7
        )

        canvas.setFillColor(
            colors.HexColor(
                "#64748b"
            )
        )

        canvas.drawCentredString(

            A4[0] / 2,

            7 * mm,

            "KLEF • Department of CSE-4 • "
            "Student Academic Performance Portal • "
            f"Page {doc.page}"

        )

        canvas.restoreState()

    # ========================================================
    # BUILD PDF
    # ========================================================

    doc.build(

        story,

        onFirstPage=add_page_number,

        onLaterPages=add_page_number

    )

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# SEMESTER MAPPING
# ============================================================

SEMESTER_MAPPING = {

    "Y-23": {

        "2023-2024 Odd Sem":
            "1-1",

        "2023-2024 Even Sem":
            "1-2",

        "2024-2025 Odd Sem":
            "2-1",

        "2024-2025 Even Sem":
            "2-2",

        "2025-2026 Odd Sem":
            "3-1",

        "2025-2026 Even Sem":
            "3-2",

        "2026-2027 Odd Sem":
            "4-1",

        "2026-2027 Even Sem":
            "4-2"
    },

    "Y-24": {

        "2024-2025 Odd Sem":
            "1-1",

        "2024-2025 Even Sem":
            "1-2",

        "2025-2026 Odd Sem":
            "2-1",

        "2025-2026 Even Sem":
            "2-2",

        "2026-2027 Odd Sem":
            "3-1",

        "2026-2027 Even Sem":
            "3-2",

        "2027-2028 Odd Sem":
            "4-1",

        "2027-2028 Even Sem":
            "4-2"
    },

    "Y-25": {

        "2025-2026 Odd Sem":
            "1-1",

        "2025-2026 Even Sem":
            "1-2",

        "2026-2027 Odd Sem":
            "2-1",

        "2026-2027 Even Sem":
            "2-2",

        "2027-2028 Odd Sem":
            "3-1",

        "2027-2028 Even Sem":
            "3-2",

        "2028-2029 Odd Sem":
            "4-1",

        "2028-2029 Even Sem":
            "4-2"
    }
}


# ============================================================
# GET SEMESTER NUMBER
# ============================================================

def get_semester_number(
    regulation,
    ay,
    semester
):

    ay_text = (
        str(ay)
        .strip()
        .lower()
        .replace("_", "-")
    )

    semester_text = (
        str(semester)
        .strip()
        .lower()
    )

    if "odd" in semester_text:

        term = "Odd Sem"

    elif "even" in semester_text:

        term = "Even Sem"

    elif "odd" in ay_text:

        term = "Odd Sem"

    elif "even" in ay_text:

        term = "Even Sem"

    else:

        term = ""

    match = re.search(
        r"(20\d{2})\D+(20\d{2})",
        ay_text
    )

    if match:

        y1 = match.group(1)
        y2 = match.group(2)

        normalized_ay = (
            f"{y1}-{y2}"
        )

    else:

        normalized_ay = (
            str(ay).strip()
        )

    lookup_key = (
        f"{normalized_ay} {term}"
    )

    regulation_map = (
        SEMESTER_MAPPING
        .get(
            regulation,
            {}
        )
    )

    return regulation_map.get(
        lookup_key,
        None
    )


# ============================================================
# LOAD DATA
# ============================================================

try:

    insem_df = load_insem_data()

except Exception as e:

    st.error(
        f"❌ Unable to load In-Semester data.\n\n{e}"
    )

    st.stop()


try:

    mentor_df = load_mentor_data()

except Exception as e:

    st.error(
        f"❌ Unable to load Mentor-Mentee data.\n\n{e}"
    )

    st.stop()


try:

    results_df = load_all_results()

except Exception as e:

    st.error(
        f"❌ Unable to load Result Analysis files.\n\n{e}"
    )

    st.stop()


try:

    cgpa_df = load_cgpa_data()

except Exception as e:

    st.error(
        f"❌ Unable to load CGPA.xlsx.\n\n{e}"
    )

    st.stop()


# ============================================================
# PREMIUM HEADER
# ============================================================

st.html("""
<div class="university-header">

    <div class="university-name">
        KONERU LAKSHMAIAH EDUCATION FOUNDATION
    </div>

    <div class="header-divider"></div>

    <div class="department-name">
        DEPARTMENT OF CSE-4
    </div>

    <div class="portal-name">
        STUDENT ACADEMIC PERFORMANCE PORTAL
    </div>

    <div class="header-subtitle">
        Academic Excellence • Performance Monitoring • Student Success
    </div>

</div>
""")


# ============================================================
# SEARCH CARD
# ============================================================

st.html("""
<div class="search-card">

    <div class="search-title">
        🔎 SEARCH STUDENT ACADEMIC DETAILS
    </div>

    <div class="search-subtitle">
        Enter the student's University ID to view
        academic performance and counselling details.
    </div>

</div>
""")


# ============================================================
# SEARCH FORM
# ============================================================

with st.form(
    "student_search_form"
):

    col1, col2 = st.columns(
        [5, 1],
        vertical_alignment="bottom"
    )

    with col1:

        student_id = st.text_input(
            "Student ID",
            placeholder="Enter University ID",
            label_visibility="collapsed"
        )

    with col2:

        search_clicked = st.form_submit_button(
            "🔍 Search",
            use_container_width=True
        )


# ============================================================
# SEARCH
# ============================================================

if search_clicked:

    student_id = (
        student_id
        .strip()
    )

    if not student_id:

        st.warning(
            "Please enter a Student ID before searching."
        )

    else:

        # ====================================================
        # FIND IN-SEM DATA
        # ====================================================

        student_data = insem_df[
            insem_df[
                "University Id"
            ] == student_id
        ].copy()


        # ====================================================
        # FIND MENTOR DATA
        # ====================================================

        if "Student ID" in mentor_df.columns:

            mentor_data = mentor_df[
                mentor_df[
                    "Student ID"
                ] == student_id
            ].copy()

        else:

            mentor_data = pd.DataFrame()


        # ====================================================
        # FIND RESULT DATA
        # ====================================================

        result_data = results_df[
            results_df[
                "Student ID"
            ] == student_id
        ].copy()


        # ====================================================
        # FIND CGPA DATA
        # ====================================================

        cgpa_data = cgpa_df[
            cgpa_df[
                "Student ID"
            ] == student_id
        ].copy()


        # ====================================================
        # STUDENT FOUND
        # ====================================================

        if (
            not student_data.empty
            or not result_data.empty
        ):

            student_name = (
                "NOT AVAILABLE"
            )

            counsellor_name = (
                "NOT AVAILABLE"
            )

            regulation = (
                "NOT AVAILABLE"
            )

            emp_id = (
                "NOT AVAILABLE"
            )

            cgpa_value = (
                "NOT AVAILABLE"
            )


            # =================================================
            # MENTOR DATA
            # =================================================

            if not mentor_data.empty:

                if (
                    "Student Name"
                    in mentor_data.columns
                ):

                    value = (
                        mentor_data.iloc[0][
                            "Student Name"
                        ]
                    )

                    if pd.notna(value):

                        student_name = (
                            str(value)
                            .strip()
                            .upper()
                        )


                if (
                    "Counsellor"
                    in mentor_data.columns
                ):

                    value = (
                        mentor_data.iloc[0][
                            "Counsellor"
                        ]
                    )

                    if pd.notna(value):

                        counsellor_name = (
                            str(value)
                            .strip()
                            .upper()
                        )


                if (
                    "Regulation"
                    in mentor_data.columns
                ):

                    value = (
                        mentor_data.iloc[0][
                            "Regulation"
                        ]
                    )

                    if pd.notna(value):

                        regulation = (
                            str(value)
                            .strip()
                            .upper()
                        )


                if (
                    "EMP ID"
                    in mentor_data.columns
                ):

                    value = (
                        mentor_data.iloc[0][
                            "EMP ID"
                        ]
                    )

                    if pd.notna(value):

                        emp_id = str(value).strip()


            # =================================================
            # GET REGULATION FROM RESULT FILE
            # =================================================

            if (
                regulation == "NOT AVAILABLE"
                and not result_data.empty
            ):

                regulation = (
                    str(
                        result_data.iloc[0][
                            "Regulation"
                        ]
                    )
                    .strip()
                    .upper()
                )


            # =================================================
            # GET CGPA
            # =================================================

            if not cgpa_data.empty:

                raw_cgpa = (
                    cgpa_data.iloc[0]["CGPA"]
                )

                if pd.notna(raw_cgpa):

                    try:

                        cgpa_value = (
                            f"{float(raw_cgpa):.2f}"
                        )

                    except Exception:

                        cgpa_value = (
                            str(raw_cgpa).strip()
                        )


            # =================================================
            # STUDENT DETAILS
            # =================================================

            st.html(f"""

            <div class="student-details-container">

                <div class="details-heading">
                    👤 STUDENT DETAILS
                </div>

                <div style="
                    display:grid;
                    grid-template-columns:
                        1.05fr 1.8fr 2fr 1fr 0.9fr 0.8fr;
                    gap:22px;
                    align-items:start;
                ">


                    <div>

                        <div class="detail-label">
                            STUDENT ID
                        </div>

                        <div class="detail-value
                                    detail-value-id">

                            {html.escape(student_id)}

                        </div>

                    </div>


                    <div>

                        <div class="detail-label">
                            STUDENT NAME
                        </div>

                        <div class="detail-value">

                            {html.escape(student_name)}

                        </div>

                    </div>


                    <div>

                        <div class="detail-label">
                            COUNSELLOR
                        </div>

                        <div class="detail-value">

                            {html.escape(counsellor_name)}

                        </div>

                    </div>


                    <div>

                        <div class="detail-label">
                            EMP ID
                        </div>

                        <div class="detail-value">

                            {html.escape(emp_id)}

                        </div>

                    </div>


                    <div>

                        <div class="detail-label">
                            REGULATION
                        </div>

                        <div class="detail-value">

                            {html.escape(regulation)}

                        </div>

                    </div>


                    <div>

                        <div class="detail-label">
                            CGPA
                        </div>

                        <div class="detail-value cgpa-value">

                            {html.escape(cgpa_value)}

                        </div>

                    </div>

                </div>

            </div>

            """)


            # =================================================
            # IN-SEMESTER PERFORMANCE
            # =================================================

            if not student_data.empty:

                st.html("""
                <div class="section-heading">
                    📊 IN-SEMESTER PERFORMANCE
                </div>
                """)

                required_columns = [
                    "Course Code",
                    "Course Name",
                    "Total"
                ]

                missing_columns = [
                    column
                    for column in required_columns
                    if column not in student_data.columns
                ]

                if missing_columns:

                    st.error(
                        "Missing In-Semester columns: "
                        + ", ".join(
                            missing_columns
                        )
                    )

                else:

                    display_data = student_data[
                        required_columns
                    ].copy()

                    display_data.insert(
                        0,
                        "S.No",
                        range(
                            1,
                            len(display_data) + 1
                        )
                    )

                    table_rows = ""

                    for _, row in display_data.iterrows():

                        course_code = html.escape(
                            str(
                                row[
                                    "Course Code"
                                ]
                            )
                        )

                        course_name = html.escape(
                            str(
                                row[
                                    "Course Name"
                                ]
                            )
                        )

                        if pd.isna(
                            row["Total"]
                        ):

                            marks_text = "—"

                        else:

                            marks = float(
                                row["Total"]
                            )

                            if marks.is_integer():

                                marks_text = (
                                    f"{int(marks)} / 50"
                                )

                            else:

                                marks_text = (
                                    f"{marks:g} / 50"
                                )

                        table_rows += f"""

                        <tr>

                            <td class="serial-number">
                                {int(row["S.No"])}
                            </td>

                            <td class="course-code">
                                {course_code}
                            </td>

                            <td class="course-title">
                                {course_name}
                            </td>

                            <td class="marks">
                                {marks_text}
                            </td>

                        </tr>

                        """


                    # =================================================
                    # IN-SEM TABLE
                    # =================================================

                    st.html(f"""

                    <div class="table-wrapper">

                        <table class="performance-table">

                            <thead>

                                <tr>

                                    <th style="width:8%;">
                                        S.No
                                    </th>

                                    <th style="width:18%;">
                                        Course Code
                                    </th>

                                    <th>
                                        Course Title
                                    </th>

                                    <th style="width:18%;">
                                        In-Sem Marks / 50
                                    </th>

                                </tr>

                            </thead>

                            <tbody>

                                {table_rows}

                            </tbody>

                        </table>

                    </div>

                    """)


                    # =================================================
                    # IN-SEM SUMMARY
                    # =================================================

                    valid_marks = (
                        display_data[
                            "Total"
                        ]
                        .dropna()
                    )

                    if not valid_marks.empty:

                        total_marks = (
                            valid_marks.sum()
                        )

                        maximum_marks = (
                            len(valid_marks)
                            * 50
                        )

                        average_marks = (
                            valid_marks.mean()
                        )

                        summary_col1, summary_col2, summary_col3 = (
                            st.columns(3)
                        )

                        with summary_col1:

                            st.html(f"""

                            <div class="summary-card">

                                <div class="summary-label">
                                    TOTAL COURSES
                                </div>

                                <div class="summary-value">
                                    {len(valid_marks)}
                                </div>

                            </div>

                            """)

                        with summary_col2:

                            st.html(f"""

                            <div class="summary-card">

                                <div class="summary-label">
                                    TOTAL MARKS
                                </div>

                                <div class="summary-value">
                                    {total_marks:g}
                                    /
                                    {maximum_marks}
                                </div>

                            </div>

                            """)

                        with summary_col3:

                            st.html(f"""

                            <div class="summary-card">

                                <div class="summary-label">
                                    AVERAGE MARKS
                                </div>

                                <div class="summary-value">
                                    {average_marks:.2f}
                                    / 50
                                </div>

                            </div>

                            """)


            # =================================================
            # COMPLETE ACADEMIC PERFORMANCE
            # =================================================

            st.html("""
            <div class="section-heading"
                 style="margin-top:32px;">

                🎓 COMPLETE ACADEMIC PERFORMANCE

            </div>
            """)


            # =================================================
            # ACADEMIC PERFORMANCE HIGHLIGHTS
            # =================================================

            if not result_data.empty:

                result_data[
                    "Category Clean"
                ] = (
                    result_data[
                        "Category"
                    ]
                    .fillna("-")
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                result_data.loc[
                    result_data[
                        "Category Clean"
                    ].isin(
                        [
                            "",
                            "NAN",
                            "NONE"
                        ]
                    ),
                    "Category Clean"
                ] = "-"


                total_courses = len(
                    result_data
                )


                category_counts = (
                    result_data[
                        "Category Clean"
                    ]
                    .value_counts()
                    .to_dict()
                )


                # ONLY P IS PASSED

                passed_courses = (
                    category_counts.get(
                        "P",
                        0
                    )
                )


                # EVERYTHING EXCEPT P

                backlog_courses = (
                    total_courses
                    - passed_courses
                )


                all_categories = list(
                    category_counts.keys()
                )


                ordered_categories = []


                if "P" in all_categories:

                    ordered_categories.append(
                        "P"
                    )


                remaining_categories = sorted(
                    [
                        category
                        for category
                        in all_categories
                        if category != "P"
                    ]
                )


                ordered_categories.extend(
                    remaining_categories
                )


                # =================================================
                # HIGHLIGHTS TITLE
                # =================================================

                st.html("""
                <div class="academic-highlights">

                    <div class="highlights-title">
                        📌 ACADEMIC PERFORMANCE HIGHLIGHTS
                    </div>

                </div>
                """)


                # =================================================
                # HORIZONTAL TABLE
                # =================================================

                header_html = """

                <th>
                    TOTAL COURSES
                </th>

                <th>
                    PASSED
                </th>

                <th>
                    BACKLOGS
                </th>

                """


                value_html = f"""

                <td class="normal-value">
                    {total_courses}
                </td>

                <td class="pass-value">
                    {passed_courses}
                </td>

                <td class="backlog-value">
                    {backlog_courses}
                </td>

                """


                # ALL CATEGORIES FROM EXCEL

                for category in ordered_categories:

                    count = (
                        category_counts
                        .get(
                            category,
                            0
                        )
                    )

                    safe_category = html.escape(
                        str(category)
                    )

                    header_html += f"""

                    <th>
                        {safe_category}
                    </th>

                    """

                    if category == "P":

                        value_class = (
                            "pass-value"
                        )

                    else:

                        value_class = (
                            "backlog-value"
                        )

                    value_html += f"""

                    <td class="{value_class}">
                        {count}
                    </td>

                    """


                st.html(f"""

                <div class="highlight-table-wrapper">

                    <table class="highlight-table">

                        <thead>

                            <tr>

                                {header_html}

                            </tr>

                        </thead>

                        <tbody>

                            <tr>

                                {value_html}

                            </tr>

                        </tbody>

                    </table>

                </div>

                """)


                # =================================================
                # ADD SEMESTER NUMBER
                # =================================================

                result_data[
                    "Semester Number"
                ] = result_data.apply(
                    lambda row:
                    get_semester_number(
                        regulation,
                        row["AY"],
                        row["Semester"]
                    ),
                    axis=1
                )


                # =================================================
                # KEEP ONLY MAPPED SEMESTERS
                # =================================================

                result_data = result_data[
                    result_data[
                        "Semester Number"
                    ].notna()
                ].copy()


                # =================================================
                # SEMESTER ORDER
                # =================================================

                semester_order = [

                    "1-1",
                    "1-2",

                    "2-1",
                    "2-2",

                    "3-1",
                    "3-2",

                    "4-1",
                    "4-2"
                ]


                result_data[
                    "Semester Sort"
                ] = (
                    result_data[
                        "Semester Number"
                    ]
                    .apply(
                        lambda x:
                        semester_order.index(x)
                        if x in semester_order
                        else 999
                    )
                )


                result_data = (
                    result_data
                    .sort_values(
                        by=[
                            "Semester Sort",
                            "AY",
                            "Semester"
                        ]
                    )
                )


                # =================================================
                # DISPLAY EACH SEMESTER
                # =================================================

                for semester_number in semester_order:

                    semester_data = result_data[
                        result_data[
                            "Semester Number"
                        ]
                        == semester_number
                    ].copy()


                    if semester_data.empty:
                        continue


                    ay_value = (
                        semester_data.iloc[0][
                            "AY"
                        ]
                    )


                    term_value = (
                        semester_data.iloc[0][
                            "Semester"
                        ]
                    )


                    # =================================================
                    # SEMESTER HEADER
                    # =================================================

                    st.html(f"""

                    <div class="semester-header">

                        <div class="semester-title">

                            📘 SEMESTER
                            {semester_number}

                        </div>

                        <div class="semester-subtitle">

                            Academic Year:
                            {html.escape(
                                str(ay_value)
                            )}

                            &nbsp;&nbsp;|&nbsp;&nbsp;

                            {html.escape(
                                str(term_value)
                            )}

                        </div>

                    </div>

                    """)


                    # =================================================
                    # SEMESTER TABLE
                    # =================================================

                    semester_display = semester_data[
                        [
                            "Course Code",
                            "Course Name",
                            "Grade",
                            "Category"
                        ]
                    ].copy()


                    semester_display.insert(
                        0,
                        "S.No",
                        range(
                            1,
                            len(
                                semester_display
                            ) + 1
                        )
                    )


                    semester_rows = ""


                    for _, row in semester_display.iterrows():

                        course_code = html.escape(
                            str(
                                row[
                                    "Course Code"
                                ]
                            )
                        )


                        course_name = html.escape(
                            str(
                                row[
                                    "Course Name"
                                ]
                            )
                        )


                        grade = html.escape(
                            str(
                                row[
                                    "Grade"
                                ]
                            )
                        )


                        category = html.escape(
                            str(
                                row[
                                    "Category"
                                ]
                            )
                        )


                        semester_rows += f"""

                        <tr>

                            <td class="serial-number">
                                {int(row["S.No"])}
                            </td>

                            <td class="course-code">
                                {course_code}
                            </td>

                            <td class="course-title">
                                {course_name}
                            </td>

                            <td class="grade-value">
                                {grade}
                            </td>

                            <td class="category-value">
                                {category}
                            </td>

                        </tr>

                        """


                    # =================================================
                    # DISPLAY SEMESTER TABLE
                    # =================================================

                    st.html(f"""

                    <div class="table-wrapper">

                        <table class="performance-table">

                            <thead>

                                <tr>

                                    <th style="width:8%;">
                                        S.No
                                    </th>

                                    <th style="width:17%;">
                                        Course Code
                                    </th>

                                    <th>
                                        Course Name
                                    </th>

                                    <th style="width:12%;">
                                        Grade
                                    </th>

                                    <th style="width:18%;">
                                        Category
                                    </th>

                                </tr>

                            </thead>

                            <tbody>

                                {semester_rows}

                            </tbody>

                        </table>

                    </div>

                    """)


            else:

                st.info(
                    "No academic result records were found "
                    "for this student."
                )


        # ====================================================
        # STUDENT NOT FOUND
        # ====================================================
        # ============================================================
        # PARENT ACADEMIC REPORT
        # ============================================================

        st.html("""
        <div class="parent-report-card">

            <div class="parent-report-title">
                📄 PARENT ACADEMIC REPORT
            </div>

            <div class="parent-report-subtitle">
                Download a printable PDF containing the student's
                complete academic details, in-semester performance,
                academic highlights, semester-wise results,
                CGPA and counselling information.
            </div>

            <div class="parent-report-note">
                The report is generated for the student currently
                displayed above.
            </div>

        </div>
        """)


        # ============================================================
        # GENERATE PDF
        # ============================================================

        try:

            parent_pdf = build_parent_report_pdf(

                student_id=student_id,

                student_name=student_name,

                counsellor_name=counsellor_name,

                emp_id=emp_id,

                regulation=regulation,

                cgpa_value=cgpa_value,

                student_data=student_data,

                result_data=result_data,

                total_courses=total_courses,

                passed_courses=passed_courses,

                backlog_courses=backlog_courses,

                category_counts=category_counts,

                ordered_categories=ordered_categories

            )


            # ========================================================
            # DOWNLOAD BUTTON
            # ========================================================

            st.download_button(

                label=(
                    "🖨️ DOWNLOAD PARENT REPORT (PDF)"
                ),

                data=parent_pdf,

                file_name=(
                    f"{student_id}_Parent_Academic_Report.pdf"
                ),

                mime="application/pdf",

                use_container_width=True,

                key=(
                    f"parent_pdf_{student_id}"
                )

            )


        except Exception as pdf_error:

            st.error(

                "Unable to generate the Parent Academic Report. "
                f"Details: {pdf_error}"

            )


        else:

            st.error(
                f"Student ID {student_id} was not found "
                "in the available academic records."
            )


# ============================================================
# FOOTER
# ============================================================

st.html("""
<div class="footer">

    STUDENT ACADEMIC PERFORMANCE PORTAL
    •
    DEPARTMENT OF CSE-4

</div>
""")
