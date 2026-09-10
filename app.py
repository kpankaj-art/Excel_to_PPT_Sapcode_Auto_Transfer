import io
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches


# ============================================================
# APP SETTINGS
# ============================================================

st.set_page_config(
    page_title="Excel → PowerPoint Dealer Data Automation",
    layout="wide"
)

st.title("Excel → PowerPoint Dealer Data Automation")
st.caption(
    "Safe matching: Name + Contact → duplicate होने पर Size → तभी SAP Code transfer"
)


# ============================================================
# FIELD ALIASES
# ============================================================

FIELD_ALIASES = {

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
        "type and size",
        "type & size",
        "dimension",
        "dimensions",
        "w x h",
        "width x height"
    ],

    "media_type": [
        "media type",
        "media",
        "type",
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


# ============================================================
# PPT LABELS
# ============================================================

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
        "type and size",
        "type & size",
        "dimension"
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


# ============================================================
# NORMALIZATION
# ============================================================

def norm(value):
    if value is None:
        return ""

    if pd.isna(value):
        return ""

    text = str(value).lower().strip()

    text = text.replace("&", " and ")
    text = text.replace("×", "x")
    text = text.replace("*", "x")

    text = re.sub(r"[\u2013\u2014_/().:,-]+", " ", text)
    text = re.sub(r"[^a-z0-9x ]+", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def compact(value):
    return re.sub(r"[^a-z0-9]", "", norm(value))


def similarity(a, b):
    a = compact(a)
    b = compact(b)

    if not a or not b:
        return 0

    return SequenceMatcher(None, a, b).ratio()


# ============================================================
# PHONE NORMALIZATION
# ============================================================

def phone_numbers(value):

    if value is None or pd.isna(value):
        return set()

    text = str(value)

    # Handles:
    # 9569757263
    # 9451736008/8004119380
    # 94517 36008
    numbers = re.findall(r"\d{7,15}", text)

    if not numbers:
        digits = re.sub(r"\D", "", text)

        if len(digits) >= 7:
            numbers = [digits]

    result = set()

    for number in numbers:

        number = re.sub(r"\D", "", number)

        if len(number) >= 10:
            result.add(number[-10:])
        else:
            result.add(number)

    return result


def phone_matches(excel_phone, ppt_phone):

    excel_numbers = phone_numbers(excel_phone)
    ppt_numbers = phone_numbers(ppt_phone)

    if not excel_numbers or not ppt_numbers:
        return False

    return bool(excel_numbers.intersection(ppt_numbers))


# ============================================================
# SIZE
# ============================================================

def normal_size(value):

    if value is None or pd.isna(value):
        return ""

    text = str(value)

    text = (
        text.lower()
        .replace("×", "x")
        .replace("*", "x")
        .replace(" ", "")
    )

    numbers = re.findall(r"\d+(?:\.\d+)?", text)

    if len(numbers) >= 2:
        return f"{numbers[0]}x{numbers[1]}"

    return ""


def size_matches(a, b):

    a = normal_size(a)
    b = normal_size(b)

    if not a or not b:
        return False

    return a == b


# ============================================================
# EXCEL SIZE
# ============================================================

def make_excel_size(row, mapping):

    # First preference: Size column
    size_column = mapping.get("size")

    if size_column:

        value = row.get(size_column, "")

        if value is not None and not pd.isna(value):

            text = str(value).strip()

            if text:
                return normal_size(text)

    # Second preference: W + H
    width_column = mapping.get("width")
    height_column = mapping.get("height")

    if width_column and height_column:

        width = row.get(width_column, "")
        height = row.get(height_column, "")

        if (
            width is not None
            and height is not None
            and not pd.isna(width)
            and not pd.isna(height)
        ):

            return normal_size(f"{width}x{height}")

    return ""


# ============================================================
# VALUE FROM EXCEL
# ============================================================

def get_value(row, mapping, field):

    column = mapping.get(field)

    if not column:
        return ""

    value = row.get(column, "")

    if value is None or pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


# ============================================================
# EXCEL COLUMN DETECTION
# ============================================================

def detect_columns(df):

    mapping = {}

    headers = list(df.columns)

    normalized_headers = {
        column: norm(column)
        for column in headers
    }

    for field, aliases in FIELD_ALIASES.items():

        best_column = None
        best_score = 0

        for column, normalized_header in normalized_headers.items():

            for alias in aliases:

                alias_norm = norm(alias)

                score = similarity(
                    normalized_header,
                    alias_norm
                )

                # Exact match
                if normalized_header == alias_norm:
                    score = 1.0

                # Partial match
                elif (
                    alias_norm in normalized_header
                    or normalized_header in alias_norm
                ):
                    score = max(score, 0.92)

                if score > best_score:
                    best_score = score
                    best_column = column

        if best_column and best_score >= 0.58:
            mapping[field] = best_column

    # IMPORTANT:
    # Your Excel can have W and H instead of Size.

    width_aliases = [
        "w",
        "width",
        "width inches",
        "width inch"
    ]

    height_aliases = [
        "h",
        "height",
        "height inches",
        "height inch"
    ]

    for column in headers:

        n = norm(column)

        if n in width_aliases:
            mapping["width"] = column

        if n in height_aliases:
            mapping["height"] = column

    return mapping


# ============================================================
# PPT LABEL DETECTION
# ============================================================

def detect_ppt_field(text):

    normalized_text = norm(text)

    for field, labels in PPT_LABELS.items():

        for label in labels:

            label_norm = norm(label)

            if normalized_text.startswith(label_norm):

                return field

    return None


# ============================================================
# PPT TEXT EXTRACTION
# ============================================================

def extract_ppt_fields(slide):

    data = {}

    for shape in slide.shapes:

        if not hasattr(shape, "text"):
            continue

        text = shape.text

        if not text or not text.strip():
            continue

        lines = text.splitlines()

        for line in lines:

            line = line.strip()

            if not line:
                continue

            # ------------------------------------------
            # LABEL : VALUE
            # ------------------------------------------

            for field, labels in PPT_LABELS.items():

                for label in labels:

                    pattern = (
                        r"^\s*"
                        + re.escape(label)
                        + r"\s*:\s*(.*?)\s*$"
                    )

                    match = re.match(
                        pattern,
                        line,
                        re.IGNORECASE
                    )

                    if match:

                        data[field] = match.group(1).strip()

                        break

                if field in data:
                    break

    return data


# ============================================================
# BUILD PPT INDEX
# ============================================================

def build_ppt_index(prs):

    index = []

    for slide_number, slide in enumerate(
        prs.slides,
        start=1
    ):

        fields = extract_ppt_fields(slide)

        index.append(
            {
                "slide_no": slide_number,
                "fields": fields
            }
        )

    return index


# ============================================================
# NAME MATCH
# ============================================================

def name_matches(excel_name, ppt_name):

    excel_name = compact(excel_name)
    ppt_name = compact(ppt_name)

    if not excel_name or not ppt_name:
        return False

    # Exact
    if excel_name == ppt_name:
        return True

    # Conservative fuzzy match
    score = similarity(
        excel_name,
        ppt_name
    )

    return score >= 0.90


# ============================================================
# MATCH ONE EXCEL ROW
# ============================================================

def match_excel_row(
    row,
    mapping,
    ppt_index
):

    excel_name = get_value(
        row,
        mapping,
        "outlet_name"
    )

    excel_phone = get_value(
        row,
        mapping,
        "contact_no"
    )

    excel_size = make_excel_size(
        row,
        mapping
    )

    # --------------------------------------------------------
    # Basic validation
    # --------------------------------------------------------

    if not excel_name:

        return {
            "status": "NO_MATCH",
            "reason": "Excel outlet/dealer name missing",
            "slide_no": ""
        }

    if not excel_phone:

        return {
            "status": "NO_MATCH",
            "reason": "Excel contact number missing",
            "slide_no": ""
        }

    # --------------------------------------------------------
    # STEP 1
    # NAME + CONTACT
    # --------------------------------------------------------

    candidates = []

    for item in ppt_index:

        ppt_fields = item["fields"]

        ppt_name = ppt_fields.get(
            "outlet_name",
            ""
        )

        ppt_phone = ppt_fields.get(
            "contact_no",
            ""
        )

        if (
            name_matches(
                excel_name,
                ppt_name
            )
            and
            phone_matches(
                excel_phone,
                ppt_phone
            )
        ):

            candidates.append(item)

    # --------------------------------------------------------
    # No name + phone match
    # --------------------------------------------------------

    if not candidates:

        return {
            "status": "NO_MATCH",
            "reason": "Name + contact not found in PPT",
            "slide_no": ""
        }

    # --------------------------------------------------------
    # UNIQUE MATCH
    # --------------------------------------------------------

    if len(candidates) == 1:

        return {
            "status": "MATCHED",
            "reason": "Unique Name + Contact match",
            "slide_no": candidates[0]["slide_no"]
        }

    # --------------------------------------------------------
    # DUPLICATE NAME + CONTACT
    #
    # Now SIZE becomes mandatory
    # --------------------------------------------------------

    if not excel_size:

        return {
            "status": "AMBIGUOUS",
            "reason": (
                "Multiple Name + Contact matches "
                "but Excel size is missing"
            ),
            "slide_no": ""
        }

    # --------------------------------------------------------
    # SIZE MATCH
    # --------------------------------------------------------

    size_candidates = []

    for candidate in candidates:

        ppt_size = candidate["fields"].get(
            "size",
            ""
        )

        if size_matches(
            excel_size,
            ppt_size
        ):

            size_candidates.append(candidate)

    # Exact one size match
    if len(size_candidates) == 1:

        return {
            "status": "MATCHED",
            "reason": "Name + Contact + Size match",
            "slide_no": size_candidates[0]["slide_no"]
        }

    # Multiple same size
    if len(size_candidates) > 1:

        return {
            "status": "AMBIGUOUS",
            "reason": (
                "Name + Contact + Size still matches "
                "multiple PPT slides"
            ),
            "slide_no": ""
        }

    # Name + phone matched but size didn't
    return {
        "status": "NO_MATCH",
        "reason": (
            f"Name + Contact matched, "
            f"but Size {excel_size} did not match"
        ),
        "slide_no": ""
    }


# ============================================================
# FIND MAIN INFORMATION TEXT BOX
# ============================================================

def find_main_info_shape(slide):

    best_shape = None
    best_score = -1

    required_fields = {
        "outlet_name",
        "address",
        "contact_no",
        "district"
    }

    for shape in slide.shapes:

        if not hasattr(shape, "text"):
            continue

        text = shape.text

        if not text.strip():
            continue

        found = set()

        for line in text.splitlines():

            line = line.strip()

            for field, labels in PPT_LABELS.items():

                for label in labels:

                    if re.match(
                        r"^\s*"
                        + re.escape(label)
                        + r"\s*:",
                        line,
                        re.IGNORECASE
                    ):

                        found.add(field)

        score = len(
            found.intersection(
                required_fields
            )
        )

        if score > best_score:

            best_score = score
            best_shape = shape

    return best_shape


# ============================================================
# UPDATE SAPCODE
# ============================================================

def update_slide_with_sapcode(
    slide,
    sapcode
):

    if not sapcode:
        return False

    # --------------------------------------------------------
    # 1. Existing Sapcode field
    # --------------------------------------------------------

    for shape in slide.shapes:

        if not hasattr(shape, "text"):
            continue

        text = shape.text

        if not text.strip():
            continue

        lines = text.splitlines()

        for i, line in enumerate(lines):

            for label in PPT_LABELS["sapcode"]:

                pattern = (
                    r"^\s*"
                    + re.escape(label)
                    + r"\s*:?"
                )

                if re.match(
                    pattern,
                    line,
                    re.IGNORECASE
                ):

                    colon_position = line.find(":")

                    if colon_position >= 0:

                        prefix = line[
                            :colon_position + 1
                        ]

                        lines[i] = (
                            prefix
                            + " "
                            + str(sapcode)
                        )

                    else:

                        lines[i] = (
                            line.strip()
                            + " : "
                            + str(sapcode)
                        )

                    shape.text = "\n".join(lines)

                    return True

    # --------------------------------------------------------
    # 2. Find main information box
    # --------------------------------------------------------

    main_shape = find_main_info_shape(
        slide
    )

    if main_shape is not None:

        lines = main_shape.text.splitlines()

        # Find District
        insert_position = len(lines)

        for i, line in enumerate(lines):

            for label in PPT_LABELS["district"]:

                if re.match(
                    r"^\s*"
                    + re.escape(label)
                    + r"\s*:",
                    line,
                    re.IGNORECASE
                ):

                    insert_position = i + 1

                    break

        lines.insert(
            insert_position,
            f"Sapcode     : {sapcode}"
        )

        main_shape.text = "\n".join(lines)

        return True

    # --------------------------------------------------------
    # 3. Last resort
    # --------------------------------------------------------

    textbox = slide.shapes.add_textbox(
        Inches(0.5),
        Inches(1.5),
        Inches(5),
        Inches(0.4)
    )

    textbox.text = (
        f"Sapcode     : {sapcode}"
    )

    for paragraph in textbox.text_frame.paragraphs:

        for run in paragraph.runs:

            run.font.size = Pt(12)

    return True


# ============================================================
# PROCESS
# ============================================================

def process_transfer(
    ppt_bytes,
    df,
    mapping
):

    prs = Presentation(
        io.BytesIO(ppt_bytes)
    )

    if len(prs.slides) == 0:

        raise ValueError(
            "PPT has no slides."
        )

    ppt_index = build_ppt_index(
        prs
    )

    results = []

    used_slides = {}

    for excel_index, row in df.iterrows():

        match = match_excel_row(
            row,
            mapping,
            ppt_index
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
                make_excel_size(
                    row,
                    mapping
                ),

            "Sapcode":
                sapcode,

            "Status":
                match["status"],

            "Reason":
                match["reason"],

            "PPT Slide":
                match["slide_no"]
        }

        # ----------------------------------------------------
        # MATCHED
        # ----------------------------------------------------

        if match["status"] == "MATCHED":

            slide_number = match["slide_no"]

            # Same slide already used
            if slide_number in used_slides:

                result["Status"] = "AMBIGUOUS"

                result["Reason"] = (
                    "PPT slide already matched "
                    f"to Excel row "
                    f"{used_slides[slide_number]}"
                )

            # SAP missing
            elif not sapcode:

                result["Status"] = "NO_MATCH"

                result["Reason"] = (
                    "Matched PPT slide but "
                    "SAP code is blank"
                )

            else:

                update_slide_with_sapcode(
                    prs.slides[
                        slide_number - 1
                    ],
                    sapcode
                )

                used_slides[
                    slide_number
                ] = excel_index + 2

        results.append(result)

    # --------------------------------------------------------
    # SAVE PPT
    # --------------------------------------------------------

    output = io.BytesIO()

    prs.save(output)

    output.seek(0)

    return (
        output.getvalue(),
        pd.DataFrame(results)
    )


# ============================================================
# MATCHING REPORT EXCEL
# ============================================================

def report_to_excel(report):

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


# ============================================================
# USER INTERFACE
# ============================================================

st.divider()

col1, col2 = st.columns(2)

with col1:

    excel_file = st.file_uploader(
        "1. Upload Excel",
        type=["xlsx", "xls"]
    )

with col2:

    ppt_file = st.file_uploader(
        "2. Upload PowerPoint Template",
        type=["pptx"]
    )


# ============================================================
# AFTER UPLOAD
# ============================================================

if excel_file and ppt_file:

    try:

        # Read Excel
        df = pd.read_excel(
            excel_file
        )

        # Detect columns
        mapping = detect_columns(
            df
        )

        st.subheader(
            "Detected Excel Columns"
        )

        display_mapping = {

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

        mapping_df = pd.DataFrame(
            list(
                display_mapping.items()
            ),
            columns=[
                "Required Field",
                "Detected Excel Column"
            ]
        )

        st.dataframe(
            mapping_df,
            use_container_width=True,
            hide_index=True
        )

        ppt_preview = Presentation(
            io.BytesIO(
                ppt_file.getvalue()
            )
        )

        st.info(
            f"Excel rows: {len(df)}  |  "
            f"PPT slides: {len(ppt_preview.slides)}"
        )

        # ----------------------------------------------------
        # BUTTON
        # ----------------------------------------------------

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

            no_match = int(
                (
                    report["Status"]
                    == "NO_MATCH"
                ).sum()
            )

            # ------------------------------------------------
            # SUMMARY
            # ------------------------------------------------

            st.success(
                f"Completed: "
                f"{matched} matched | "
                f"{ambiguous} ambiguous | "
                f"{no_match} not matched"
            )

            # ------------------------------------------------
            # REPORT
            # ------------------------------------------------

            st.subheader(
                "Matching Report"
            )

            st.dataframe(
                report,
                use_container_width=True,
                hide_index=True
            )

            # ------------------------------------------------
            # DOWNLOAD FINAL PPT
            # ------------------------------------------------

            st.download_button(
                label="Download Final PPT",
                data=final_ppt,
                file_name="Final_SAP_Transfer.pptx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                )
            )

            # ------------------------------------------------
            # DOWNLOAD REPORT
            # ------------------------------------------------

            st.download_button(
                label="Download Matching Report",
                data=report_to_excel(
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
