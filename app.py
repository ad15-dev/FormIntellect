
"""
FormIntellect — Read-Only Response Analyzer

Features:
- Responsive Streamlit dashboard
- XLSX, XLS, CSV, TSV, ODS and JSON support
- Per-field response tables and pie charts
- Data-quality summary
- Downloadable Excel report with embedded pie charts
- Original source file is never modified

Run:
    python -m streamlit run app.py
"""

from __future__ import annotations

import io
import json
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# --------------------------------------------------
# Configuration
# --------------------------------------------------

APP_NAME = "FormIntellect"
IST = ZoneInfo("Asia/Kolkata")

SUPPORTED_FORMATS = [
    "xlsx", "xls", "csv", "tsv", "ods", "json"
]

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --------------------------------------------------
# UI styling
# --------------------------------------------------

st.markdown(
    """
    <style>
    .block-container {
        max-width: 1500px;
        padding-top: 1.5rem;
        padding-bottom: 2.5rem;
    }

    .hero {
        padding: 1.5rem;
        border-radius: 18px;
        background: linear-gradient(
            120deg, #edf4ff 0%, #f3efff 100%
        );
        border: 1px solid #dce5fb;
        margin-bottom: 1.2rem;
    }

    .hero h1 {
        margin: 0;
        color: #17233d;
        font-size: 2rem;
    }

    .hero p {
        color: #52617b;
        margin: 0.45rem 0 0;
    }

    div[data-testid="stMetric"] {
        padding: 14px 16px;
        border-radius: 14px;
        border: 1px solid rgba(128,128,128,0.2);
    }

    @media (max-width: 700px) {
        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }

        .hero h1 {
            font-size: 1.6rem;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>📊 FormIntellect</h1>
        <p>
            Response analytics · Section-wise charts ·
            Data quality · Excel reporting
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------
# Read input files without modifying them
# --------------------------------------------------

def load_dataframe(uploaded_file) -> pd.DataFrame:
    """Load the uploaded file into memory without writing to it."""

    extension = uploaded_file.name.rsplit(".", 1)[-1].lower()
    file_bytes = uploaded_file.getvalue()

    if extension in ("xlsx", "xls"):
        df = pd.read_excel(io.BytesIO(file_bytes))

    elif extension == "csv":
        try:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding="utf-8-sig"
            )
        except UnicodeDecodeError:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding="latin-1"
            )

    elif extension == "tsv":
        df = pd.read_csv(
            io.BytesIO(file_bytes),
            sep="\t",
            encoding="utf-8-sig"
        )

    elif extension == "ods":
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            engine="odf"
        )

    elif extension == "json":
        obj = json.loads(file_bytes.decode("utf-8-sig"))

        if isinstance(obj, list):
            df = pd.json_normalize(obj)

        elif isinstance(obj, dict):
            candidate = obj.get(
                "data",
                obj.get("responses", obj)
            )

            if isinstance(candidate, (list, dict)):
                df = pd.json_normalize(candidate)
            else:
                df = pd.json_normalize(obj)

        else:
            raise ValueError(
                "JSON must contain a list or object."
            )

    else:
        raise ValueError(
            f"Unsupported file extension: {extension}"
        )

    if df.empty and len(df.columns) == 0:
        raise ValueError("The file contains no tabular columns.")

    # Normalize column labels in memory only.
    df.columns = [
        str(column).strip() or f"Column_{i + 1}"
        for i, column in enumerate(df.columns)
    ]

    return df


# --------------------------------------------------
# Response analysis
# --------------------------------------------------

def clean_series(series: pd.Series) -> pd.Series:
    """Return non-empty values as strings for category analysis."""

    values = series.dropna().astype(str).str.strip()
    return values[values != ""]


def response_frequency(
    series: pd.Series,
    category_limit: int = 10
) -> pd.DataFrame:
    """Create response counts and percentages."""

    values = clean_series(series)
    counts = values.value_counts()

    if len(counts) > category_limit:
        top_counts = counts.iloc[:category_limit - 1]
        other_count = counts.iloc[category_limit - 1:].sum()

        counts = pd.concat([
            top_counts,
            pd.Series({"Other": other_count})
        ])

    total = int(counts.sum())

    result = pd.DataFrame({
        "Response": counts.index.astype(str),
        "Count": counts.values.astype(int),
    })

    if total:
        result["Percentage (%)"] = (
            result["Count"] / total * 100
        ).round(2)
    else:
        result["Percentage (%)"] = 0.0

    return result


def show_pie_chart(
    frequency_df: pd.DataFrame,
    title: str
) -> None:
    """Display one pie chart for a field."""

    if (
        frequency_df.empty
        or frequency_df["Count"].sum() == 0
    ):
        st.info("No non-empty responses available for this chart.")
        return

    fig, ax = plt.subplots(figsize=(6, 4.3))

    ax.pie(
        frequency_df["Count"],
        labels=frequency_df["Response"],
        autopct=lambda pct: (
            f"{pct:.1f}%" if pct >= 4 else ""
        ),
        startangle=90,
        wedgeprops={
            "linewidth": 1,
            "edgecolor": "white"
        },
        textprops={"fontsize": 8},
    )

    ax.set_title(title, fontsize=12, pad=12)
    ax.axis("equal")
    fig.tight_layout()

    st.pyplot(fig, use_container_width=True)
    plt.close(fig)


# --------------------------------------------------
# Excel report generation
# --------------------------------------------------

def unique_sheet_name(
    original_name: str,
    used_names: set[str]
) -> str:
    """Create a valid, unique Excel worksheet name."""

    name = re.sub(
        r"[\[\]:*?/\\]",
        "_",
        str(original_name)
    ).strip("' ")

    if not name:
        name = "Field"

    name = name[:31]
    base_name = name
    counter = 1

    while name.lower() in used_names:
        suffix = f"_{counter}"
        name = base_name[:31 - len(suffix)] + suffix
        counter += 1

    used_names.add(name.lower())
    return name


def style_worksheet(ws) -> None:
    """Apply consistent formatting to an Excel worksheet."""

    ws.sheet_view.showGridLines = False

    for cell in ws[1]:
        cell.font = Font(
            bold=True,
            size=14,
            color="FFFFFF"
        )
        cell.fill = PatternFill(
            "solid",
            fgColor="253858"
        )

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )


def build_excel_report(
    df: pd.DataFrame,
    source_filename: str
) -> bytes:
    """
    Generate a new Excel report.
    The source file is never modified.
    """

    workbook = Workbook()
    overview = workbook.active
    overview.title = "Overview"

    overview.append(["FormIntellect Analysis Report"])
    overview.append(["Source file", source_filename])
    overview.append([
        "Generated at (IST)",
        datetime.now(IST).strftime(
            "%d-%m-%Y %I:%M:%S %p IST"
        )
    ])
    overview.append(["Total records", int(len(df))])
    overview.append(["Total fields", int(len(df.columns))])
    overview.append([])

    overview.append([
        "Field",
        "Non-empty responses",
        "Unique responses",
        "Missing / blank"
    ])

    for column in df.columns:
        cleaned = clean_series(df[column])

        overview.append([
            str(column),
            int(len(cleaned)),
            int(cleaned.nunique()),
            int(len(df) - len(cleaned))
        ])

    used_names = {"overview"}

    # Each source field gets its own worksheet.
    for column in df.columns:
        field_name = str(column)
        counts_df = response_frequency(
            df[column],
            category_limit=100000
        )

        sheet_name = unique_sheet_name(
            field_name,
            used_names
        )

        ws = workbook.create_sheet(sheet_name)

        ws.append([f"Field analysis: {field_name}"])
        ws.append(["Source file", source_filename])
        ws.append(["Total records", int(len(df))])
        ws.append([
            "Non-empty responses",
            int(len(clean_series(df[column])))
        ])
        ws.append([])
        ws.append(["Response", "Count", "Percentage (%)"])

        for response, count, percentage in (
            counts_df.itertuples(index=False, name=None)
        ):
            ws.append([
                str(response),
                int(count),
                float(percentage)
            ])

        # Add a native Excel pie chart.
        if not counts_df.empty:
            chart = PieChart()
            chart.title = (
                f"Response distribution: {field_name[:40]}"
            )
            chart.height = 8
            chart.width = 13
            chart.style = 10

            # Row 6 contains the table headers.
            data = Reference(
                ws,
                min_col=2,
                min_row=6,
                max_row=6 + len(counts_df)
            )

            labels = Reference(
                ws,
                min_col=1,
                min_row=7,
                max_row=6 + len(counts_df)
            )

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
            ws.add_chart(chart, "E2")

        ws.freeze_panes = "A7"
        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 20

        style_worksheet(ws)

        for cell in ws[6]:
            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cell.fill = PatternFill(
                "solid",
                fgColor="4255A4"
            )

    # Format overview sheet.
    overview.column_dimensions["A"].width = 34
    overview.column_dimensions["B"].width = 24
    overview.column_dimensions["C"].width = 22
    overview.column_dimensions["D"].width = 20
    overview.freeze_panes = "A7"

    style_worksheet(overview)

    for cell in overview[7]:
        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )
        cell.fill = PatternFill(
            "solid",
            fgColor="4255A4"
        )

    output = io.BytesIO()
    workbook.save(output)

    return output.getvalue()


# --------------------------------------------------
# Sidebar settings
# --------------------------------------------------

with st.sidebar:
    st.header("Analysis settings")

    category_limit = st.slider(
        "Maximum categories per pie chart",
        min_value=3,
        max_value=20,
        value=10
    )

    show_preview = st.checkbox(
        "Show source data preview",
        value=True
    )

    st.divider()

    st.markdown("**Supported formats**")
    st.caption("XLSX · XLS · CSV · TSV · ODS · JSON")

    st.info(
        "Read-only mode: the uploaded source file "
        "is never overwritten."
    )


# --------------------------------------------------
# File upload
# --------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload your response file",
    type=SUPPORTED_FORMATS,
    help=(
        "Upload an Excel, CSV, TSV, ODS or JSON file. "
        "The original file is not modified."
    )
)

if uploaded_file is None:
    st.info(
        "Upload a response file to generate "
        "field-wise tables and pie charts."
    )

    col1, col2, col3 = st.columns(3)

    col1.metric("Source protection", "Read-only")
    col2.metric("Analysis", "Field-wise")
    col3.metric("Export", "Excel + charts")

    st.stop()


# --------------------------------------------------
# Load and validate
# --------------------------------------------------

try:
    source_df = load_dataframe(uploaded_file)

except Exception as error:
    st.error(f"Unable to read the file: {error}")
    st.stop()


# Work with a separate in-memory analysis copy.
df = source_df.copy(deep=True)

# Exclude completely blank rows from analysis only.
# No changes are written to the uploaded file.
blank_rows = df.isna().all(axis=1)
df = df.loc[~blank_rows].copy()

if len(df.columns) == 0:
    st.warning("No usable columns were found.")
    st.stop()


# --------------------------------------------------
# Overview dashboard
# --------------------------------------------------

st.subheader("Overview")

total_records = len(df)
total_fields = len(df.columns)

non_empty_cells = sum(
    len(clean_series(df[column]))
    for column in df.columns
)

missing_cells = (
    total_records * total_fields - non_empty_cells
)

metric1, metric2, metric3, metric4 = st.columns(4)

metric1.metric("Records", f"{total_records:,}")
metric2.metric("Fields", f"{total_fields:,}")
metric3.metric("Non-empty cells", f"{non_empty_cells:,}")
metric4.metric("Missing / blank", f"{missing_cells:,}")


# --------------------------------------------------
# Data quality
# --------------------------------------------------

with st.expander("Data quality summary"):
    quality_rows = []

    for column in df.columns:
        values = clean_series(df[column])

        quality_rows.append({
            "Field": str(column),
            "Non-empty": int(len(values)),
            "Unique values": int(values.nunique()),
            "Missing / blank": int(
                total_records - len(values)
            )
        })

    quality_df = pd.DataFrame(quality_rows)

    st.dataframe(
        quality_df,
        use_container_width=True,
        hide_index=True
    )


# --------------------------------------------------
# Optional raw data preview
# --------------------------------------------------

if show_preview:
    with st.expander("Source data preview — read-only"):
        st.dataframe(
            df.head(100),
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            "This is a preview only. "
            "No edits are written to the source file."
        )


# --------------------------------------------------
# Field-wise charts and tables
# --------------------------------------------------

st.divider()
st.subheader("Section-wise tables and pie charts")

st.caption(
    "By default, each input column is treated as one "
    "analysis field. If your data contains a dedicated "
    "section column, use the section grouping option below."
)

# Optional actual-section grouping.
section_column = st.selectbox(
    "Optional: choose a column that identifies form sections",
    options=["No section column"] + list(df.columns),
    index=0
)

if section_column != "No section column":
    section_values = (
        df[section_column]
        .fillna("Unspecified section")
        .astype(str)
        .str.strip()
        .replace("", "Unspecified section")
    )

    section_names = sorted(section_values.unique())

    selected_sections = st.multiselect(
        "Choose sections to display",
        options=section_names,
        default=section_names
    )

    section_fields = st.multiselect(
        "Choose response fields to analyze",
        options=[
            column for column in df.columns
            if column != section_column
        ],
        default=[
            column for column in df.columns
            if column != section_column
        ]
    )

    for section in selected_sections:
        section_mask = section_values == section
        section_df = df.loc[section_mask]

        st.markdown(f"### {section}")
        st.caption(f"{len(section_df):,} records in this section")

        for column in section_fields:
            with st.container(border=True):
                st.markdown(f"#### {column}")

                counts_df = response_frequency(
                    section_df[column],
                    category_limit=category_limit
                )

                chart_col, table_col = st.columns([1, 1])

                with chart_col:
                    show_pie_chart(
                        counts_df,
                        f"{section}: {column}"
                    )

                with table_col:
                    st.markdown("**Response frequency table**")

                    st.dataframe(
                        counts_df,
                        use_container_width=True,
                        hide_index=True
                    )

else:
    selected_fields = st.multiselect(
        "Choose fields to analyze",
        options=list(df.columns),
        default=list(df.columns)
    )

    if not selected_fields:
        st.warning("Select at least one field.")
    else:
        for index, column in enumerate(
            selected_fields,
            start=1
        ):
            with st.container(border=True):
                st.markdown(f"#### {index}. {column}")

                counts_df = response_frequency(
                    df[column],
                    category_limit=category_limit
                )

                chart_col, table_col = st.columns([1, 1])

                with chart_col:
                    show_pie_chart(
                        counts_df,
                        f"Response distribution: {column}"
                    )

                with table_col:
                    st.markdown("**Response frequency table**")

                    st.dataframe(
                        counts_df,
                        use_container_width=True,
                        hide_index=True
                    )

                    response_total = (
                        int(counts_df["Count"].sum())
                        if not counts_df.empty
                        else 0
                    )

                    st.caption(
                        f"{response_total:,} non-empty responses"
                    )


# --------------------------------------------------
# Excel export
# --------------------------------------------------

st.divider()
st.subheader("Download Excel report")

st.write(
    "The generated workbook contains an overview sheet "
    "and a separate worksheet for every source field. "
    "Each field worksheet includes response counts, "
    "percentages, and an embedded Excel pie chart."
)

try:
    report_data = build_excel_report(
        df,
        uploaded_file.name
    )

    source_stem = uploaded_file.name.rsplit(".", 1)[0]

    st.download_button(
        label="⬇️ Download section-wise Excel report",
        data=report_data,
        file_name=f"{source_stem}_FormIntellect_Analysis.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True
    )

except Exception as error:
    st.error(f"Unable to generate the Excel report: {error}")


# --------------------------------------------------
# Footer
# --------------------------------------------------

st.caption(
    "FormIntellect · Read-only analytics · "
    "No form submissions are performed by this version."
)

st.caption(
    "Generated at: "
    + datetime.now(IST).strftime(
        "%d-%m-%Y %I:%M:%S %p IST"
    )
)
