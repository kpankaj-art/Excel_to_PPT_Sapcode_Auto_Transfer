import io
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Pt


# =========================================================
# APP
# =========================================================

st.set_page_config(
    page_title="Excel to PowerPoint Dealer Data Automation",
    layout="wide"
)

st.title("Excel → PowerPoint Dealer Data Automation")

st.caption(
    "Safe matching: Name + Contact → duplicate होने पर Size → SAP Code transfer"
)


# =========================================================
# EXCEL COLUMN ALIASES
# =========================================================

ALIASES = {

    "outlet_name": [
        "outlet name",
        "outlet",
        "dealer name",
        "dealer / name",
        "dealer/name",
        "dealer",
        "customer name",
        "customer",
        "retailer name",
        "retailer",
        "shop name",
        "party name"
    ],

    "address": [
        "address",
        "dealer address",
        "dealer / address",
        "dealer / adderess",
        "dealer/address",
        "outlet address",
        "customer address",
        "retailer address",
        "shop address",
        "location"
    ],

    "contact_no": [
        "contact",
        "contact no",
        "contact no.",
        "contact number",
        "dealer contact",
        "dealer / contact",
        "dealer/contact",
        "mobile",
        "mobile no",
        "mobile number",
        "phone",
        "phone no",
        "phone number",
        "telephone"
    ],

    "district": [
        "district",
        "district name",
        "dist",
        "dist name"
    ],

    "sapcode": [
        "sapcode",
        "sap code",
        "sap-code",
        "dealer code",
        "dealercode",
        "dealer_code",
        "customer code",
        "customercode",
        "customer_code",
        "customer id",
        "dealer id",
        "sap id",
        "sap number"
    ],

    "size": [
        "size",
        "size inches",
        "dimension",
        "dimensions"
    ],

    "media_type": [
        "media type",
        "media",
        "board type",
        "material type",
        "type"
    ],

    "remarks": [
        "remarks",
        "remark",
        "comments",
        "comment",
        "note",
        "notes"
    ],

    "qty": [
        "qty",
        "quantity",
        "qnty"
    ]
}


# =========================================================
# PPT LABELS
# =========================================================

PPT_LABELS = {

    "outlet_name": [
        "outlet name",
        "outlet",
        "dealer name",
        "dealer",
        "customer name",
        "customer",
        "retailer name",
        "retailer"
    ],

    "address": [
        "address",
        "dealer address",
        "outlet address",
        "customer address"
    ],

    "contact_no": [
        "contact no",
        "contact",
        "mobile",
        "phone",
        "contact number",
        "mobile number"
    ],

    "district": [
        "district",
        "district name",
        "dist"
    ],

    "sapcode": [
        "sapcode",
        "sap code",
        "dealer code",
        "dealercode",
        "customer code",
        "customercode"
    ],

    "size": [
        "size",
        "dimension",
        "dimensions",
        "type and size",
        "type & size"
    ],

    "media_type": [
        "media type",
        "media"
    ],

    "remarks": [
        "remarks",
        "remark",
        "comments",
        "comment"
    ],

    "qty": [
        "qty",
        "quantity"
    ]
}


# =========================================================
# NORMALIZATION
# =========================================================

def norm(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value).lower().strip()

    text = text.replace("×", "x")
    text = text.replace("*", "x")

    text = re.sub(
        r"[\u2013\u2014_/().,:-]+",
        " ",
        text
    )

    text = re.sub(
        r"[^a-z0-9x ]+",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def compact(value):

    return re.sub(
        r"[^a-z0-9]",
        "",
        norm(value)
    )


def similarity(a, b):

    a = compact(a)
    b = compact(b)

    if not a or not b:
        return 0

    return SequenceMatcher(
        None,
        a,
        b
    ).ratio()


# =========================================================
# PHONE
# =========================================================

def extract_phones(value):

    if value is None:
        return set()

    try:
        if pd.isna(value):
            return set()
    except Exception:
        pass

    text = str(value)

    numbers = re.findall(
        r"\d{7,15}",
        text
    )

    result = set()

    for number in numbers:

        number = re.sub(
            r"\D",
            "",
            number
        )

        if len(number) >= 10:
            result.add(
                number[-10:]
            )
        elif number:
            result.add(number)

    return result


def phone_match(a, b):

    a_numbers = extract_phones(a)
    b_numbers = extract_phones(b)

    if not a_numbers or not b_numbers:
        return False

    return bool(
        a_numbers.intersection(
            b_numbers
        )
    )


# =========================================================
# SIZE
# =========================================================

def normalize_size(value):

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value)

    text = (
        text
        .replace("×", "x")
        .replace("*", "x")
        .replace(" ", "")
    )

    numbers = re.findall(
        r"\d+(?:\.\d+)?",
        text
    )

    if len(numbers) >= 2:

        return (
            f"{numbers[0]}x"
            f"{numbers[1]}"
        )

    return ""


# =========================================================
# EXCEL COLUMN DETECTION
# =========================================================

def detect_columns(df):

    columns = list(
        df.columns
    )

    mapping = {}

    # ---------------------------------------------
    # Explicit W / H detection FIRST
    # ---------------------------------------------

    for column in columns:

        n = norm(column)

        if n in {
            "w",
            "width",
            "width inches",
            "width inch"
        }:

            mapping["width"] = column

        if n in {
            "h",
            "height",
            "height inches",
            "height inch"
        }:

            mapping["height"] = column

    # ---------------------------------------------
    # Other fields
    # ---------------------------------------------

    for field, aliases in ALIASES.items():

        # Size must NOT select TYPE
        if field == "size":

            valid_columns = []

            for column in columns:

                n = norm(column)

                if n not in {
                    "type",
                    "media type",
                    "media"
                }:

                    valid_columns.append(
                        column
                    )

        else:

            valid_columns = columns

        best_column = None
        best_score = 0

        for column in valid_columns:

            header = norm(column)

            for alias in aliases:

                alias_norm = norm(alias)

                if header == alias_norm:

                    score = 1.0

                elif (
                    alias_norm in header
                    or header in alias_norm
                ):

                    score = 0.94

                else:

                    score = similarity(
                        header,
                        alias_norm
                    )

                if score > best_score:

                    best_score = score
                    best_column = column

        if (
            best_column is not None
            and best_score >= 0.70
        ):

            mapping[field] = best_column

    return mapping


# =========================================================
# GET EXCEL VALUE
# =========================================================

def get_value(
    row,
    mapping,
    field
):

    column = mapping.get(
        field
    )

    if not column:
        return ""

    value = row.get(
        column,
        ""
    )

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    if (
        isinstance(value, float)
        and value.is_integer()
    ):

        return str(
            int(value)
        )

    return str(value).strip()


# =========================================================
# EXCEL SIZE
# =========================================================

def get_excel_size(
    row,
    mapping
):

    # ---------------------------------------------
    # If real Size column exists
    # ---------------------------------------------

    size_column = mapping.get(
        "size"
    )

    if size_column:

        value = row.get(
            size_column,
            ""
        )

        size = normalize_size(
            value
        )

        if size:

            return size

    # ---------------------------------------------
    # Otherwise W + H
    # ---------------------------------------------

    width_column = mapping.get(
        "width"
    )

    height_column = mapping.get(
        "height"
    )

    if (
        width_column
        and height_column
    ):

        width = row.get(
            width_column,
            ""
        )

        height = row.get(
            height_column,
            ""
        )

        return normalize_size(
            f"{width}x{height}"
        )

    return ""


# =========================================================
# PPT LABEL PARSER
# =========================================================

def parse_ppt_line(line):

    match = re.match(
        r"^\s*(.+?)\s*:\s*(.*?)\s*$",
        line
    )

    if not match:

        return None, None

    label = norm(
        match.group(1)
    )

    value = match.group(2).strip()

    best_field = None
    best_score = 0

    for field, aliases in PPT_LABELS.items():

        for alias in aliases:

            alias_norm = norm(
                alias
            )

            if label == alias_norm:

                score = 1.0

            elif (
                label.startswith(
                    alias_norm
                )
                or alias_norm.startswith(
                    label
                )
            ):

                score = 0.95

            else:

                score = similarity(
                    label,
                    alias_norm
                )

            if score > best_score:

                best_score = score
                best_field = field

    if best_score >= 0.70:

        return (
            best_field,
            value
        )

    return None, None


# =========================================================
# READ PPT
# =========================================================

def extract_ppt_fields(slide):

    data = {}

    for shape in slide.shapes:

        if not hasattr(
            shape,
            "text"
        ):

            continue

        if not shape.text.strip():

            continue

        for line in shape.text.splitlines():

            field, value = (
                parse_ppt_line(
                    line
                )
            )

            if field:

                data[field] = value

    return data


# =========================================================
# PPT INDEX
# =========================================================

def build_ppt_index(prs):

    result = []

    for slide_no, slide in enumerate(
        prs.slides,
        start=1
    ):

        result.append(
            {
                "slide_no": slide_no,
                "fields":
                    extract_ppt_fields(
                        slide
                    )
            }
        )

    return result


# =========================================================
# NAME MATCH
# =========================================================

def name_match(
    excel_name,
    ppt_name
):

    a = compact(
        excel_name
    )

    b = compact(
        ppt_name
    )

    if not a or not b:

        return False

    # Exact
    if a == b:

        return True

    # Small spelling difference
    return (
        SequenceMatcher(
            None,
            a,
            b
        ).ratio()
        >= 0.88
    )


# =========================================================
# MATCH LOGIC
#
# 1. Name
# 2. Contact
# 3. If multiple -> Size
# =========================================================

def match_row(
    row,
    mapping,
    ppt_index
):

    excel_name = get_value(
        row,
        mapping,
        "outlet_name"
    )

    excel_contact = get_value(
        row,
        mapping,
        "contact_no"
    )

    excel_size = get_excel_size(
        row,
        mapping
    )

    if not excel_name:

        return (
            "NO_MATCH",
            "Excel outlet/dealer name missing",
            ""
        )

    if not excel_contact:

        return (
            "NO_MATCH",
            "Excel contact number missing",
            ""
        )

    # ---------------------------------------------
    # NAME + CONTACT
    # ---------------------------------------------

    candidates = []

    for item in ppt_index:

        fields = item[
            "fields"
        ]

        ppt_name = fields.get(
            "outlet_name",
            ""
        )

        ppt_contact = fields.get(
            "contact_no",
            ""
        )

        if (
            name_match(
                excel_name,
                ppt_name
            )
            and
            phone_match(
                excel_contact,
                ppt_contact
            )
        ):

            candidates.append(
                item
            )

    # ---------------------------------------------
    # Nothing found
    # ---------------------------------------------

    if not candidates:

        return (
            "NO_MATCH",
            "Name + contact not found in PPT",
            ""
        )

    # ---------------------------------------------
    # Only one PPT record
    # ---------------------------------------------

    if len(candidates) == 1:

        return (
            "MATCHED",
            "Unique Name + Contact match",
            candidates[0][
                "slide_no"
            ]
        )

    # ---------------------------------------------
    # Multiple records
    # Size becomes compulsory
    # ---------------------------------------------

    if not excel_size:

        return (
            "AMBIGUOUS",
            "Multiple Name + Contact matches but Excel size missing",
            ""
        )

    size_candidates = []

    for candidate in candidates:

        ppt_size = normalize_size(
            candidate[
                "fields"
            ].get(
                "size",
                ""
            )
        )

        if ppt_size == excel_size:

            size_candidates.append(
                candidate
            )

    if len(size_candidates) == 1:

        return (
            "MATCHED",
            "Name + Contact + Size match",
            size_candidates[0][
                "slide_no"
            ]
        )

    if len(size_candidates) > 1:

        return (
            "AMBIGUOUS",
            "Name + Contact + Size matches multiple PPT slides",
            ""
        )

    return (
        "NO_MATCH",
        f"Name + Contact matched but Size {excel_size} did not match",
        ""
    )


# =========================================================
# FIND MAIN INFORMATION BOX
# =========================================================

def find_info_shape(slide):

    best_shape = None
    best_score = -1

    required = {
        "outlet_name",
        "address",
        "contact_no",
        "district"
    }

    for shape in slide.shapes:

        if not hasattr(
            shape,
            "text"
        ):

            continue

        if not shape.text.strip():

            continue

        found = set()

        for line in shape.text.splitlines():

            field, _ = (
                parse_ppt_line(
                    line
                )
            )

            if field:

                found.add(
                    field
                )

        score = len(
            found.intersection(
                required
            )
        )

        if score > best_score:

            best_score = score
            best_shape = shape

    return best_shape


# =========================================================
# ADD SAPCODE WITH AUTO SPACE ADJUSTMENT
# =========================================================

def write_text_with_auto_fit(
    shape,
    lines
):

    text_frame = shape.text_frame

    text_frame.clear()

    for index, line in enumerate(
        lines
    ):

        if index == 0:

            paragraph = (
                text_frame.paragraphs[0]
            )

        else:

            paragraph = (
                text_frame.add_paragraph()
            )

        paragraph.text = line

    text_frame.word_wrap = True

    # ---------------------------------------------
    # Existing template has 7 lines.
    # Adding SAP makes 8.
    # Reduce font automatically.
    # ---------------------------------------------

    if len(lines) >= 8:

        font_size = 12

    else:

        font_size = 15

    for paragraph in (
        text_frame.paragraphs
    ):

        for run in paragraph.runs:

            run.font.size = Pt(
                font_size
            )


def add_sapcode(
    slide,
    sapcode
):

    if not sapcode:

        return False

    sapcode = str(
        sapcode
    ).strip()

    # ---------------------------------------------
    # If Sapcode already exists
    # ---------------------------------------------

    for shape in slide.shapes:

        if not hasattr(
            shape,
            "text"
        ):

            continue

        if not shape.text.strip():

            continue

        lines = shape.text.splitlines()

        for index, line in enumerate(
            lines
        ):

            field, _ = (
                parse_ppt_line(
                    line
                )
            )

            if field == "sapcode":

                label = line.split(
                    ":",
                    1
                )[0].strip()

                lines[index] = (
                    f"{label} : {sapcode}"
                )

                write_text_with_auto_fit(
                    shape,
                    lines
                )

                return True

    # ---------------------------------------------
    # Find main info box
    # ---------------------------------------------

    shape = find_info_shape(
        slide
    )

    if shape is not None:

        lines = shape.text.splitlines()

        insert_at = len(
            lines
        )

        for index, line in enumerate(
            lines
        ):

            field, _ = (
                parse_ppt_line(
                    line
                )
            )

            if field == "district":

                insert_at = (
                    index + 1
                )

        lines.insert(
            insert_at,
            f"Sapcode     : {sapcode}"
        )

        write_text_with_auto_fit(
            shape,
            lines
        )

        return True

    return False


# =========================================================
# PROCESS
# =========================================================

def process_transfer(
    ppt_bytes,
    df,
    mapping
):

    prs = Presentation(
        io.BytesIO(
            ppt_bytes
        )
    )

    ppt_index = build_ppt_index(
        prs
    )

    used_slides = {}

    results = []

    for excel_index, row in (
        df.iterrows()
    ):

        status, reason, slide_no = (
            match_row(
                row,
                mapping,
                ppt_index
            )
        )

        sapcode = get_value(
            row,
            mapping,
            "sapcode"
        )

        result = {

            "Excel Row":
                excel_index + 2,

            "Outlet Name":
                get_value(
                    row,
                    mapping,
                    "outlet_name"
                ),

            "Contact No":
                get_value(
                    row,
                    mapping,
                    "contact_no"
                ),

            "Size":
                get_excel_size(
                    row,
                    mapping
                ),

            "Sapcode":
                sapcode,

            "Status":
                status,

            "Reason":
                reason,

            "PPT Slide":
                slide_no
        }

        # -----------------------------------------
        # MATCHED
        # -----------------------------------------

        if status == "MATCHED":

            # Same PPT slide cannot be reused
            if slide_no in used_slides:

                result["Status"] = (
                    "AMBIGUOUS"
                )

                result["Reason"] = (
                    "This PPT slide is already "
                    f"matched with Excel row "
                    f"{used_slides[slide_no]}"
                )

            elif not sapcode:

                result["Status"] = (
                    "NO_MATCH"
                )

                result["Reason"] = (
                    "SAP code is blank"
                )

            else:

                add_sapcode(
                    prs.slides[
                        slide_no - 1
                    ],
                    sapcode
                )

                used_slides[
                    slide_no
                ] = excel_index + 2

        results.append(
            result
        )

    # ---------------------------------------------
    # SAVE PPT
    # ---------------------------------------------

    output = io.BytesIO()

    prs.save(
        output
    )

    output.seek(0)

    return (
        output.getvalue(),
        pd.DataFrame(
            results
        )
    )


# =========================================================
# REPORT
# =========================================================

def create_report_excel(
    report
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        report.to_excel(
            writer,
            index=False,
            sheet_name="Matching Report"
        )

    output.seek(0)

    return output.getvalue()


# =========================================================
# UI
# =========================================================

excel_file = st.file_uploader(
    "1. Upload Excel",
    type=["xlsx", "xls"]
)

ppt_file = st.file_uploader(
    "2. Upload PowerPoint",
    type=["pptx"]
)


if excel_file and ppt_file:

    try:

        # ---------------------------------------------
        # READ EXCEL
        # ---------------------------------------------

        df = pd.read_excel(
            excel_file
        )

        mapping = detect_columns(
            df
        )

        st.subheader(
            "Detected Excel Columns"
        )

        detected = {

            "Outlet Name":
                mapping.get(
                    "outlet_name",
                    "❌ Not Found"
                ),

            "Address":
                mapping.get(
                    "address",
                    "❌ Not Found"
                ),

            "Contact No":
                mapping.get(
                    "contact_no",
                    "❌ Not Found"
                ),

            "District":
                mapping.get(
                    "district",
                    "❌ Not Found"
                ),

            "Sapcode":
                mapping.get(
                    "sapcode",
                    "❌ Not Found"
                ),

            "Size":
                mapping.get(
                    "size",
                    "Using W + H"
                ),

            "Width":
                mapping.get(
                    "width",
                    "-"
                ),

            "Height":
                mapping.get(
                    "height",
                    "-"
                ),

            "Media Type":
                mapping.get(
                    "media_type",
                    "-"
                )
        }

        st.dataframe(
            pd.DataFrame(
                detected.items(),
                columns=[
                    "Required Field",
                    "Detected Excel Column"
                ]
            ),
            use_container_width=True,
            hide_index=True
        )

        # ---------------------------------------------
        # PPT INFO
        # ---------------------------------------------

        preview_prs = Presentation(
            io.BytesIO(
                ppt_file.getvalue()
            )
        )

        st.info(
            f"Excel data rows: {len(df)} | "
            f"PPT slides: {len(preview_prs.slides)}"
        )

        # ---------------------------------------------
        # BUTTON
        # ---------------------------------------------

        if st.button(
            "Analyse & Transfer SAP Codes",
            type="primary"
        ):

            with st.spinner(
                "Matching Name → Contact → Size..."
            ):

                final_ppt, report = (
                    process_transfer(
                        ppt_file.getvalue(),
                        df,
                        mapping
                    )
                )

            matched = int(
                (
                    report["Status"]
                    == "MATCHED"
                ).sum()
            )

            ambiguous = int(
                (
                    report["Status"]
                    == "AMBIGUOUS"
                ).sum()
            )

            not_matched = int(
                (
                    report["Status"]
                    == "NO_MATCH"
                ).sum()
            )

            st.success(
                f"Completed: "
                f"{matched} matched | "
                f"{ambiguous} ambiguous | "
                f"{not_matched} not matched"
            )

            st.subheader(
                "Matching Report"
            )

            st.dataframe(
                report,
                use_container_width=True,
                hide_index=True
            )

            # -----------------------------------------
            # FINAL PPT
            # -----------------------------------------

            st.download_button(
                "Download Final PPT",
                data=final_ppt,
                file_name="Final_SAP_Transfer.pptx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                )
            )

            # -----------------------------------------
            # REPORT
            # -----------------------------------------

            st.download_button(
                "Download Matching Report",
                data=create_report_excel(
                    report
                ),
                file_name="PPT_Matching_Report.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )

    except Exception as error:

        st.error(
            f"Error: {error}"
        )
