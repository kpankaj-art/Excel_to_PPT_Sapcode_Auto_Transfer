import io
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Pt


# =========================================================
# APP SETTINGS
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
        "material type"
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
# PPT LABEL ALIASES
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
# NORMALIZE TEXT
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
# PHONE NUMBER
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

            result.add(
                number
            )

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
# DETECT EXCEL COLUMNS
# =========================================================

def detect_columns(df):

    columns = list(
        df.columns
    )

    mapping = {}

    # -----------------------------------------------------
    # W / WIDTH
    # -----------------------------------------------------

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

    # -----------------------------------------------------
    # OTHER COLUMNS
    # -----------------------------------------------------

    for field, aliases in ALIASES.items():

        valid_columns = columns

        # TYPE को Size नहीं मानना
        if field == "size":

            valid_columns = [
                column
                for column in columns
                if norm(column) not in {
                    "type",
                    "media",
                    "media type"
                }
            ]

        best_column = None
        best_score = 0

        for column in valid_columns:

            header = norm(column)

            for alias in aliases:

                alias_norm = norm(
                    alias
                )

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

            mapping[field] = (
                best_column
            )

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
# GET EXCEL SIZE
# =========================================================

def get_excel_size(
    row,
    mapping
):

    # ---------------------------------------------
    # First: Size column
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
    # Second: W + H
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
# PPT LINE PARSER
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
# EXTRACT PPT FIELDS
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
# BUILD PPT INDEX
# =========================================================

def build_ppt_index(prs):

    index = []

    for slide_no, slide in enumerate(
        prs.slides,
        start=1
    ):

        index.append(
            {
                "slide_no": slide_no,
                "fields":
                    extract_ppt_fields(
                        slide
                    )
            }
        )

    return index


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

    # Exact match
    if a == b:
        return True

    # Minor spelling mistake
    return (
        SequenceMatcher(
            None,
            a,
            b
        ).ratio()
        >= 0.88
    )


# =========================================================
# MATCH EXCEL ROW WITH PPT
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

    # -----------------------------------------------------
    # NAME REQUIRED
    # -----------------------------------------------------

    if not excel_name:

        return (
            "NO_MATCH",
            "Excel outlet/dealer name missing",
            ""
        )

    # -----------------------------------------------------
    # CONTACT REQUIRED
    # -----------------------------------------------------

    if not excel_contact:

        return (
            "NO_MATCH",
            "Excel contact number missing",
            ""
        )

    # -----------------------------------------------------
    # NAME + CONTACT
    # -----------------------------------------------------

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

    # -----------------------------------------------------
    # NO MATCH
    # -----------------------------------------------------

    if not candidates:

        return (
            "NO_MATCH",
            "Name + contact not found in PPT",
            ""
        )

    # -----------------------------------------------------
    # UNIQUE NAME + CONTACT
    # -----------------------------------------------------

    if len(candidates) == 1:

        return (
            "MATCHED",
            "Unique Name + Contact match",
            candidates[0][
                "slide_no"
            ]
        )

    # -----------------------------------------------------
    # DUPLICATE NAME + CONTACT
    # SIZE REQUIRED
    # -----------------------------------------------------

    if not excel_size:

        return (
            "AMBIGUOUS",
            "Multiple Name + Contact matches but Excel size is missing",
            ""
        )

    # -----------------------------------------------------
    # SIZE MATCH
    # -----------------------------------------------------

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

        if (
            ppt_size
            and
            ppt_size == excel_size
        ):

            size_candidates.append(
                candidate
            )

    # Exactly one
    if len(size_candidates) == 1:

        return (
            "MATCHED",
            "Name + Contact + Size match",
            size_candidates[0][
                "slide_no"
            ]
        )

    # More than one
    if len(size_candidates) > 1:

        return (
            "AMBIGUOUS",
            "Name + Contact + Size matches multiple PPT slides",
            ""
        )

    # Size mismatch
    return (
        "NO_MATCH",
        (
            "Name + Contact matched but "
            f"Size {excel_size} did not match"
        ),
        ""
    )


# =========================================================
# FIND MAIN INFORMATION TEXT BOX
# =========================================================

def find_info_shape(slide):

    best_shape = None
    best_score = -1

    required_fields = {
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
                required_fields
            )
        )

        if score > best_score:

            best_score = score
            best_shape = shape

    return best_shape


# =========================================================
# GET ORIGINAL FONT SIZE
# =========================================================

def get_existing_font_size(shape):

    sizes = []

    try:

        for paragraph in (
            shape.text_frame.paragraphs
        ):

            for run in paragraph.runs:

                if run.font.size:

                    sizes.append(
                        run.font.size.pt
                    )

    except Exception:

        pass

    if sizes:

        return min(
            sizes
        )

    return 15


# =========================================================
# WRITE ALL TEXT AS BOLD
# =========================================================

def write_text_with_auto_fit(
    shape,
    lines
):

    text_frame = shape.text_frame

    original_size = (
        get_existing_font_size(
            shape
        )
    )

    # ---------------------------------------------
    # Clear old text
    # ---------------------------------------------

    text_frame.clear()

    text_frame.word_wrap = True

    # ---------------------------------------------
    # SAP code added = one extra line
    # Reduce font if necessary
    # ---------------------------------------------

    if len(lines) >= 8:

        font_size = max(
            9,
            min(
                original_size,
                12
            )
        )

    else:

        font_size = original_size

    # ---------------------------------------------
    # Write every line
    # ---------------------------------------------

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

        run = paragraph.add_run()

        run.text = str(
            line
        )

        # =========================================
        # IMPORTANT
        # ALL TEXT BOLD
        # =========================================

        run.font.bold = True

        run.font.size = Pt(
            font_size
        )


# =========================================================
# ADD SAP CODE
# =========================================================

def add_sapcode(
    slide,
    sapcode
):

    if not sapcode:
        return False

    sapcode = str(
        sapcode
    ).strip()

    # -----------------------------------------------------
    # EXISTING SAPCODE FIELD
    # -----------------------------------------------------

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

    # -----------------------------------------------------
    # FIND MAIN INFORMATION BOX
    # -----------------------------------------------------

    shape = find_info_shape(
        slide
    )

    if shape is not None:

        lines = shape.text.splitlines()

        # Put SAP code after District
        insert_position = len(
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

                insert_position = (
                    index + 1
                )

        lines.insert(
            insert_position,
            f"Sapcode     : {sapcode}"
        )

        # ---------------------------------------------
        # ALL TEXT INCLUDING SAPCODE WILL BE BOLD
        # ---------------------------------------------

        write_text_with_auto_fit(
            shape,
            lines
        )

        return True

    # -----------------------------------------------------
    # Don't put SAP code randomly
    # -----------------------------------------------------

    return False


# =========================================================
# PROCESS TRANSFER
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

    # -----------------------------------------------------
    # Process every Excel row
    # -----------------------------------------------------

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

        # -------------------------------------------------
        # MATCHED
        # -------------------------------------------------

        if status == "MATCHED":

            # Same slide cannot be used twice
            if slide_no in used_slides:

                result["Status"] = (
                    "AMBIGUOUS"
                )

                result["Reason"] = (
                    "PPT slide already used by "
                    f"Excel row "
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

                success = add_sapcode(
                    prs.slides[
                        slide_no - 1
                    ],
                    sapcode
                )

                if success:

                    used_slides[
                        slide_no
                    ] = excel_index + 2

                else:

                    result["Status"] = (
                        "NO_MATCH"
                    )

                    result["Reason"] = (
                        "Match found but PPT "
                        "information box could "
                        "not be located"
                    )

        results.append(
            result
        )

    # -----------------------------------------------------
    # SAVE FINAL PPT
    # -----------------------------------------------------

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
# CREATE REPORT EXCEL
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
# FILE UPLOAD
# =========================================================

st.divider()

excel_file = st.file_uploader(
    "1. Upload Excel",
    type=[
        "xlsx",
        "xls"
    ]
)

ppt_file = st.file_uploader(
    "2. Upload PowerPoint",
    type=[
        "pptx"
    ]
)


# =========================================================
# MAIN PROCESS
# =========================================================

if excel_file and ppt_file:

    try:

        # -------------------------------------------------
        # READ EXCEL
        # -------------------------------------------------

        df = pd.read_excel(
            excel_file
        )

        # -------------------------------------------------
        # AUTO DETECT COLUMNS
        # -------------------------------------------------

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

        # -------------------------------------------------
        # PPT INFORMATION
        # -------------------------------------------------

        preview_prs = Presentation(
            io.BytesIO(
                ppt_file.getvalue()
            )
        )

        st.info(
            f"Excel data rows: {len(df)} | "
            f"PPT slides: {len(preview_prs.slides)}"
        )

        # -------------------------------------------------
        # BUTTON
        # -------------------------------------------------

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

            # -------------------------------------------------
            # SUMMARY
            # -------------------------------------------------

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

            # -------------------------------------------------
            # REPORT
            # -------------------------------------------------

            st.subheader(
                "Matching Report"
            )

            st.dataframe(
                report,
                use_container_width=True,
                hide_index=True
            )

            # -------------------------------------------------
            # FINAL PPT
            # -------------------------------------------------

            st.download_button(
                label="Download Final PPT",
                data=final_ppt,
                original_name = ppt_file.name

if original_name.lower().endswith(".pptx"):
    original_name = original_name[:-5]

updated_name = original_name + "_Update.pptx"
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                )
            )

            # -------------------------------------------------
            # REPORT EXCEL
            # -------------------------------------------------

           st.download_button(
    label="Download Updated PPT",
    data=final_ppt,
    file_name=updated_name,
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "presentationml.presentation"
    )
)
    except Exception as error:

        st.error(
            f"Error: {error}"
        )
