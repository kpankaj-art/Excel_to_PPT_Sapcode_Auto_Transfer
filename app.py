import streamlit as st
import pandas as pd
import re
import io
from pptx import Presentation
from pptx.util import Pt
from difflib import SequenceMatcher


# =========================================================
# PAGE SETTINGS
# =========================================================

st.set_page_config(
    page_title="SAP Code Transfer Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 SAP Code Transfer Tool")
st.write(
    "Excel se SAP Code ko existing PowerPoint template me "
    "safe matching ke saath transfer karein."
)


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def safe_filename(name):
    """
    Invalid Windows filename characters ko replace karta hai.
    """
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip()


def normalize_text(value):
    """
    Text ko matching ke liye normalize karta hai.
    """
    if value is None:
        return ""

    if pd.isna(value):
        return ""

    value = str(value).strip().upper()

    value = value.replace("\n", " ")
    value = value.replace("\r", " ")

    # Multiple spaces remove
    value = re.sub(r"\s+", " ", value)

    # Special characters ko space
    value = re.sub(r"[^A-Z0-9 ]", " ", value)

    value = re.sub(r"\s+", " ", value).strip()

    return value


def normalize_phone(value):
    """
    Contact number ko normalize karta hai.

    Example:
    9451736008/8004119380
    9451736008 / 8004119380
    """
    if value is None:
        return []

    if pd.isna(value):
        return []

    value = str(value).strip()

    # Excel kabhi-kabhi number ko 9.561E+09 format me de sakta hai
    if "E+" in value.upper():
        try:
            value = str(int(float(value)))
        except:
            pass

    # Decimal .0 remove
    if value.endswith(".0"):
        value = value[:-2]

    numbers = re.findall(r"\d{10,}", value)

    # Agar direct 10 digit number hai
    if not numbers:
        digits = re.sub(r"\D", "", value)

        if len(digits) >= 10:
            numbers = [digits]

    return list(dict.fromkeys(numbers))


def normalize_size(value):
    """
    Size ko standard format me convert karta hai.

    600 X 48
    600x48
    600*48

    sab:
    600X48
    """

    if value is None:
        return ""

    if pd.isna(value):
        return ""

    value = str(value).upper().strip()

    value = value.replace(" ", "")
    value = value.replace("*", "X")
    value = value.replace("×", "X")

    # Sirf dimensions extract karo
    match = re.search(r"(\d+(?:\.\d+)?)X(\d+(?:\.\d+)?)", value)

    if match:
        a = match.group(1)
        b = match.group(2)

        return f"{a}X{b}"

    return value


def sizes_equal(size1, size2):
    """
    Size compare karta hai.
    """
    s1 = normalize_size(size1)
    s2 = normalize_size(size2)

    return s1 != "" and s2 != "" and s1 == s2


def names_similar(name1, name2, threshold=0.88):
    """
    Exact ya minor spelling difference ke liye fuzzy matching.
    """

    n1 = normalize_text(name1)
    n2 = normalize_text(name2)

    if not n1 or not n2:
        return False

    if n1 == n2:
        return True

    score = SequenceMatcher(None, n1, n2).ratio()

    return score >= threshold


# =========================================================
# EXCEL COLUMN DETECTION
# =========================================================

def clean_column_name(col):
    return normalize_text(col)


def detect_columns(df):

    columns = list(df.columns)

    normalized_columns = {
        col: clean_column_name(col)
        for col in columns
    }

    mapping = {}

    # -----------------------------------------------------
    # NAME
    # -----------------------------------------------------

    name_aliases = [
        "DEALER / NAME",
        "DEALER NAME",
        "OUTLET NAME",
        "OUTLET",
        "DEALER",
        "CUSTOMER NAME",
        "NAME"
    ]

    # -----------------------------------------------------
    # CONTACT
    # -----------------------------------------------------

    contact_aliases = [
        "DEALER / CONTACT",
        "DEALER CONTACT",
        "CONTACT NO",
        "CONTACT NUMBER",
        "CONTACT",
        "PHONE",
        "MOBILE",
        "MOBILE NO"
    ]

    # -----------------------------------------------------
    # SAP
    # -----------------------------------------------------

    sap_aliases = [
        "SAPCODE",
        "SAP CODE",
        "DEALER CODE",
        "CUSTOMER CODE",
        "CUSTOMERCODE",
        "CUSTOMER ID",
        "SAP"
    ]

    # -----------------------------------------------------
    # ADDRESS
    # -----------------------------------------------------

    address_aliases = [
        "DEALER / ADDERSS",
        "DEALER / ADDRESS",
        "DEALER ADDRESS",
        "ADDRESS"
    ]

    # -----------------------------------------------------
    # DISTRICT
    # -----------------------------------------------------

    district_aliases = [
        "DISTRICT NAME",
        "DISTRICT"
    ]

    # -----------------------------------------------------
    # TYPE
    # -----------------------------------------------------

    type_aliases = [
        "TYPE",
        "MEDIA TYPE",
        "MEDIA"
    ]

    # -----------------------------------------------------
    # WIDTH
    # -----------------------------------------------------

    width_aliases = [
        "W",
        "WIDTH"
    ]

    # -----------------------------------------------------
    # HEIGHT
    # -----------------------------------------------------

    height_aliases = [
        "H",
        "HEIGHT"
    ]

    # -----------------------------------------------------
    # SIZE
    # -----------------------------------------------------

    size_aliases = [
        "SIZE",
        "DIMENSION",
        "DIMENSIONS"
    ]


    def find_column(aliases, exclude=None):

        exclude = exclude or []

        # First exact match
        for col, norm in normalized_columns.items():

            if col in exclude:
                continue

            if norm in [clean_column_name(x) for x in aliases]:
                return col

        # Then partial match
        for col, norm in normalized_columns.items():

            if col in exclude:
                continue

            for alias in aliases:

                alias_norm = clean_column_name(alias)

                if alias_norm and alias_norm in norm:
                    return col

        return None


    # Name
    mapping["name"] = find_column(name_aliases)

    # Contact
    mapping["contact"] = find_column(
        contact_aliases,
        exclude=[mapping["name"]] if mapping["name"] else []
    )

    # SAP
    mapping["sap"] = find_column(sap_aliases)

    # Address
    mapping["address"] = find_column(address_aliases)

    # District
    mapping["district"] = find_column(district_aliases)

    # Type
    mapping["type"] = find_column(type_aliases)

    # Width
    mapping["width"] = find_column(width_aliases)

    # Height
    mapping["height"] = find_column(height_aliases)


    # -----------------------------------------------------
    # IMPORTANT:
    # Agar W + H available hai to wahi Size banega.
    # TYPE ko kabhi Size nahi banayenge.
    # -----------------------------------------------------

    if mapping["width"] and mapping["height"]:

        mapping["size"] = None

    else:

        mapping["size"] = find_column(
            size_aliases,
            exclude=[
                mapping["type"]
            ] if mapping["type"] else []
        )


    return mapping


# =========================================================
# PPT FIELD PARSING
# =========================================================

def parse_label_line(line):

    """
    Example:
    Outlet Name : AZMAT TRADERS
    Contact No : 9569757263
    """

    if not line:
        return None, None

    line = str(line).strip()

    patterns = [

        (r"^\s*Outlet\s*Name\s*:\s*(.*)$", "name"),

        (r"^\s*Address\s*:\s*(.*)$", "address"),

        (r"^\s*Contact\s*No\s*:\s*(.*)$", "contact"),

        (r"^\s*Contact\s*Number\s*:\s*(.*)$", "contact"),

        (r"^\s*District\s*:\s*(.*)$", "district"),

        (r"^\s*Sap\s*code\s*:\s*(.*)$", "sap"),

        (r"^\s*Sapcode\s*:\s*(.*)$", "sap"),

        (r"^\s*SAP\s*Code\s*:\s*(.*)$", "sap"),

        (r"^\s*Size\s*:\s*(.*)$", "size"),

        (r"^\s*Media\s*Type\s*:\s*(.*)$", "type"),

        (r"^\s*Type\s*:\s*(.*)$", "type"),

        (r"^\s*Remarks\s*:\s*(.*)$", "remarks"),

        (r"^\s*Qty\s*:\s*(.*)$", "qty")
    ]

    for pattern, field in patterns:

        match = re.match(
            pattern,
            line,
            flags=re.IGNORECASE
        )

        if match:

            return field, match.group(1).strip()

    return None, None


def extract_shape_text(shape):

    try:

        if hasattr(shape, "text"):

            return shape.text or ""

    except:

        pass

    return ""


def extract_ppt_fields(prs):

    slides_data = []


    for slide_number, slide in enumerate(prs.slides, start=1):

        data = {
            "slide": slide_number,
            "name": "",
            "address": "",
            "contact": "",
            "district": "",
            "sap": "",
            "size": "",
            "type": "",
            "remarks": "",
            "qty": "",
            "info_shape_index": None
        }


        # -------------------------------------------------
        # First pass: all shape text
        # -------------------------------------------------

        for shape_index, shape in enumerate(slide.shapes):

            text = extract_shape_text(shape)

            if not text:
                continue

            lines = text.splitlines()


            # ---------------------------------------------
            # Parse EVERY line independently
            # ---------------------------------------------

            for line in lines:

                field, value = parse_label_line(line)

                if field:

                    data[field] = value


            # ---------------------------------------------
            # Info shape identify karo
            # ---------------------------------------------

            lower_text = text.lower()

            if (
                "outlet name" in lower_text
                and "contact" in lower_text
                and "district" in lower_text
            ):

                data["info_shape_index"] = shape_index


        slides_data.append(data)


    return slides_data


# =========================================================
# EXCEL SIZE CREATION
# =========================================================

def get_excel_size(row, mapping):

    width_col = mapping.get("width")
    height_col = mapping.get("height")
    size_col = mapping.get("size")

    # W + H preferred
    if width_col and height_col:

        width = row.get(width_col, "")
        height = row.get(height_col, "")

        if (
            str(width).strip() != ""
            and str(height).strip() != ""
            and str(width).lower() != "nan"
            and str(height).lower() != "nan"
        ):

            return normalize_size(
                f"{width}X{height}"
            )

    # Fallback: Size column
    if size_col:

        return normalize_size(
            row.get(size_col, "")
        )

    return ""


# =========================================================
# SAFE MATCHING
# =========================================================

def find_best_slide(excel_row, excel_index, ppt_slides, mapping, used_slides):

    excel_name = excel_row.get(
        mapping.get("name"),
        ""
    )

    excel_contact = excel_row.get(
        mapping.get("contact"),
        ""
    )

    excel_size = get_excel_size(
        excel_row,
        mapping
    )


    excel_phones = normalize_phone(
        excel_contact
    )


    # -----------------------------------------------------
    # REQUIRED
    # -----------------------------------------------------

    if not excel_name:
        return None, "Excel name missing"

    if not excel_phones:
        return None, "Excel contact missing/invalid"


    # -----------------------------------------------------
    # Step 1:
    # Name + Contact match
    # -----------------------------------------------------

    candidates = []


    for ppt in ppt_slides:

        slide_no = ppt["slide"]

        # Same slide ko dobara use nahi karna
        if slide_no in used_slides:
            continue


        ppt_name = ppt.get("name", "")

        ppt_contact = ppt.get("contact", "")


        # Name check
        if not names_similar(
            excel_name,
            ppt_name
        ):
            continue


        # Contact check
        ppt_phones = normalize_phone(
            ppt_contact
        )


        if not ppt_phones:
            continue


        # At least one contact number same hona chahiye
        contact_match = any(
            phone in ppt_phones
            for phone in excel_phones
        )


        if not contact_match:
            continue


        candidates.append(ppt)


    # -----------------------------------------------------
    # No Name + Contact match
    # -----------------------------------------------------

    if not candidates:

        return (
            None,
            "Name + Contact match nahi mila"
        )


    # -----------------------------------------------------
    # Only ONE candidate
    #
    # User rule:
    # Name + Contact unique hai to match allowed.
    # -----------------------------------------------------

    if len(candidates) == 1:

        return (
            candidates[0],
            "Matched by Name + Contact"
        )


    # -----------------------------------------------------
    # Multiple candidates
    #
    # Size tie-breaker mandatory
    # -----------------------------------------------------

    if not excel_size:

        return (
            None,
            "Name + Contact multiple matches; Size available nahi hai"
        )


    size_candidates = []

    for ppt in candidates:

        ppt_size = ppt.get(
            "size",
            ""
        )

        if sizes_equal(
            excel_size,
            ppt_size
        ):

            size_candidates.append(ppt)


    # -----------------------------------------------------
    # Exactly ONE size match
    # -----------------------------------------------------

    if len(size_candidates) == 1:

        return (
            size_candidates[0],
            "Matched by Name + Contact + Size"
        )


    # -----------------------------------------------------
    # Multiple same-size candidates
    # -----------------------------------------------------

    if len(size_candidates) > 1:

        return (
            None,
            "Multiple Name + Contact + Size matches; SAP transfer nahi kiya"
        )


    # -----------------------------------------------------
    # Size not matched
    # -----------------------------------------------------

    return (
        None,
        "Name + Contact multiple matches hain, lekin Size match nahi hua"
    )


# =========================================================
# ADD SAP CODE TO PPT
# =========================================================

def add_sap_to_info_shape(shape, sap_code):

    """
    Existing information block ko rewrite karta hai.

    Output:
    Outlet Name :
    Address     :
    Contact No  :
    District    :
    Sapcode     :

    Sabhi text BOLD.
    """

    old_text = shape.text or ""

    lines = old_text.splitlines()


    # -----------------------------------------------------
    # Existing fields extract karo
    # -----------------------------------------------------

    fields = []

    found_sap = False

    for line in lines:

        field, value = parse_label_line(line)

        if field == "sap":

            found_sap = True

            # Existing SAP ko replace
            fields.append(
                ("sap", value)
            )

        elif field in [
            "name",
            "address",
            "contact",
            "district"
        ]:

            fields.append(
                (field, value)
            )


    # -----------------------------------------------------
    # Agar SAP already present hai to replace
    # -----------------------------------------------------

    if found_sap:

        final_lines = []

        for field, value in fields:

            if field == "sap":

                final_lines.append(
                    f"Sapcode     : {sap_code}"
                )

            else:

                label_map = {
                    "name": "Outlet Name",
                    "address": "Address",
                    "contact": "Contact No",
                    "district": "District"
                }

                final_lines.append(
                    f"{label_map[field]} : {value}"
                )

    else:

        final_lines = []

        inserted = False

        for field, value in fields:

            label_map = {
                "name": "Outlet Name",
                "address": "Address",
                "contact": "Contact No",
                "district": "District"
            }

            final_lines.append(
                f"{label_map[field]} : {value}"
            )

            # District ke turant baad SAP
            if field == "district":

                final_lines.append(
                    f"Sapcode     : {sap_code}"
                )

                inserted = True


        # Agar District nahi mila
        if not inserted:

            final_lines.append(
                f"Sapcode     : {sap_code}"
            )


    # -----------------------------------------------------
    # Shape text rewrite
    # -----------------------------------------------------

    shape.text_frame.clear()


    # Font size:
    # 8+ lines ho to 12 pt
    # otherwise 15 pt
    # -----------------------------------------------------

    if len(final_lines) >= 8:

        font_size = 12

    elif len(final_lines) >= 7:

        font_size = 13

    else:

        font_size = 15


    for i, line in enumerate(final_lines):

        if i == 0:

            paragraph = shape.text_frame.paragraphs[0]

        else:

            paragraph = shape.text_frame.add_paragraph()


        paragraph.text = ""


        run = paragraph.add_run()

        run.text = line

        run.font.size = Pt(font_size)

        run.font.bold = True


        # Paragraph spacing
        paragraph.space_before = Pt(0)

        paragraph.space_after = Pt(0)


    # -----------------------------------------------------
    # Vertical alignment
    # -----------------------------------------------------

    try:

        shape.text_frame.vertical_anchor = 1

    except:

        pass


# =========================================================
# UPDATE PPT
# =========================================================

def update_ppt(prs, matching_results):

    updated_count = 0

    failed_count = 0


    for result in matching_results:

        if not result["matched"]:
            continue


        slide_number = result["slide"]

        sap_code = result["sap"]


        slide = prs.slides[
            slide_number - 1
        ]


        # Find info shape
        target_shape = None


        # First use saved shape index
        info_index = result.get(
            "info_shape_index"
        )


        if info_index is not None:

            if info_index < len(slide.shapes):

                target_shape = slide.shapes[
                    info_index
                ]


        # Fallback: search shape
        if target_shape is None:

            for shape in slide.shapes:

                text = extract_shape_text(
                    shape
                )

                lower_text = text.lower()

                if (
                    "outlet name" in lower_text
                    and "contact" in lower_text
                    and "district" in lower_text
                ):

                    target_shape = shape
                    break


        # If shape not found
        if target_shape is None:

            failed_count += 1

            continue


        # Add SAP
        add_sap_to_info_shape(
            target_shape,
            sap_code
        )

        updated_count += 1


    return updated_count, failed_count


# =========================================================
# MATCHING REPORT
# =========================================================

def create_report_excel(results):

    report_rows = []

    for result in results:

        report_rows.append({

            "Excel Row":
                result.get("excel_row", ""),

            "Excel Name":
                result.get("excel_name", ""),

            "Excel Contact":
                result.get("excel_contact", ""),

            "Excel Size":
                result.get("excel_size", ""),

            "SAP Code":
                result.get("sap", ""),

            "PPT Slide":
                result.get("slide", ""),

            "PPT Name":
                result.get("ppt_name", ""),

            "PPT Contact":
                result.get("ppt_contact", ""),

            "PPT Size":
                result.get("ppt_size", ""),

            "Status":
                "MATCHED"
                if result.get("matched")
                else "NOT MATCHED",

            "Reason":
                result.get("reason", "")
        })


    df_report = pd.DataFrame(
        report_rows
    )


    output = io.BytesIO()


    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df_report.to_excel(
            writer,
            index=False,
            sheet_name="Matching Report"
        )


    output.seek(0)

    return output.getvalue()


# =========================================================
# FILE UPLOAD
# =========================================================

st.subheader("1️⃣ Files Upload Karein")

excel_file = st.file_uploader(
    "Excel File Upload Karein",
    type=["xlsx", "xls"]
)

ppt_file = st.file_uploader(
    "PowerPoint Template Upload Karein",
    type=["pptx"]
)


# =========================================================
# PROCESS
# =========================================================

if excel_file and ppt_file:

    st.divider()

    try:

        # -------------------------------------------------
        # Read Excel
        # -------------------------------------------------

        df = pd.read_excel(
            excel_file
        )


        # -------------------------------------------------
        # Detect columns
        # -------------------------------------------------

        mapping = detect_columns(
            df
        )


        st.subheader(
            "2️⃣ Detected Excel Columns"
        )


        display_mapping = {

            "Outlet / Dealer Name":
                mapping.get("name"),

            "Contact":
                mapping.get("contact"),

            "SAP Code":
                mapping.get("sap"),

            "Address":
                mapping.get("address"),

            "District":
                mapping.get("district"),

            "Media Type":
                mapping.get("type"),

            "Width":
                mapping.get("width"),

            "Height":
                mapping.get("height"),

            "Size":
                "W + H"
                if mapping.get("width")
                and mapping.get("height")
                else mapping.get("size")
        }


        mapping_df = pd.DataFrame(
            list(display_mapping.items()),
            columns=[
                "Field",
                "Excel Column"
            ]
        )


        st.dataframe(
            mapping_df,
            use_container_width=True,
            hide_index=True
        )


        # -------------------------------------------------
        # Required columns validation
        # -------------------------------------------------

        missing = []


        if not mapping.get("name"):

            missing.append(
                "Dealer / Outlet Name"
            )


        if not mapping.get("contact"):

            missing.append(
                "Contact"
            )


        if not mapping.get("sap"):

            missing.append(
                "SAP Code / Customer Code"
            )


        if missing:

            st.error(
                "Required Excel columns detect nahi ho paaye: "
                + ", ".join(missing)
            )

            st.info(
                "Excel ke columns check karein."
            )

            st.stop()


        # -------------------------------------------------
        # Load PPT
        # -------------------------------------------------

        ppt_bytes = ppt_file.getvalue()

        prs = Presentation(
            io.BytesIO(ppt_bytes)
        )


        # -------------------------------------------------
        # Extract PPT
        # -------------------------------------------------

        ppt_slides = extract_ppt_fields(
            prs
        )


        st.subheader(
            "3️⃣ PPT Data Detected"
        )


        ppt_preview = []


        for item in ppt_slides:

            ppt_preview.append({

                "Slide":
                    item["slide"],

                "Outlet Name":
                    item["name"],

                "Contact":
                    item["contact"],

                "Size":
                    item["size"],

                "Media Type":
                    item["type"],

                "District":
                    item["district"]
            })


        st.dataframe(
            pd.DataFrame(ppt_preview),
            use_container_width=True,
            hide_index=True
        )


        # -------------------------------------------------
        # Start Matching
        # -------------------------------------------------

        st.subheader(
            "4️⃣ Safe Matching"
        )


        used_slides = set()

        matching_results = []


        for excel_index, row in df.iterrows():

            excel_name = row.get(
                mapping.get("name"),
                ""
            )

            excel_contact = row.get(
                mapping.get("contact"),
                ""
            )

            sap_code = row.get(
                mapping.get("sap"),
                ""
            )


            excel_size = get_excel_size(
                row,
                mapping
            )


            matched_slide, reason = find_best_slide(

                row,

                excel_index,

                ppt_slides,

                mapping,

                used_slides
            )


            # -------------------------------------------------
            # Matched
            # -------------------------------------------------

            if matched_slide is not None:

                slide_no = matched_slide[
                    "slide"
                ]

                used_slides.add(
                    slide_no
                )


                matching_results.append({

                    "excel_row":
                        excel_index + 2,

                    "excel_name":
                        str(excel_name),

                    "excel_contact":
                        str(excel_contact),

                    "excel_size":
                        excel_size,

                    "sap":
                        str(sap_code),

                    "slide":
                        slide_no,

                    "ppt_name":
                        matched_slide["name"],

                    "ppt_contact":
                        matched_slide["contact"],

                    "ppt_size":
                        matched_slide["size"],

                    "info_shape_index":
                        matched_slide[
                            "info_shape_index"
                        ],

                    "matched":
                        True,

                    "reason":
                        reason
                })


            # -------------------------------------------------
            # Not matched
            # -------------------------------------------------

            else:

                matching_results.append({

                    "excel_row":
                        excel_index + 2,

                    "excel_name":
                        str(excel_name),

                    "excel_contact":
                        str(excel_contact),

                    "excel_size":
                        excel_size,

                    "sap":
                        str(sap_code),

                    "slide":
                        "",

                    "ppt_name":
                        "",

                    "ppt_contact":
                        "",

                    "ppt_size":
                        "",

                    "info_shape_index":
                        None,

                    "matched":
                        False,

                    "reason":
                        reason
                })


        # -------------------------------------------------
        # Summary
        # -------------------------------------------------

        matched_count = sum(
            1
            for x in matching_results
            if x["matched"]
        )


        unmatched_count = (
            len(matching_results)
            - matched_count
        )


        col1, col2, col3 = st.columns(3)


        col1.metric(
            "Excel Rows",
            len(df)
        )


        col2.metric(
            "Matched",
            matched_count
        )


        col3.metric(
            "Not Matched",
            unmatched_count
        )


        # -------------------------------------------------
        # Matching result table
        # -------------------------------------------------

        result_preview = []


        for result in matching_results:

            result_preview.append({

                "Excel Row":
                    result["excel_row"],

                "Excel Name":
                    result["excel_name"],

                "Contact":
                    result["excel_contact"],

                "Size":
                    result["excel_size"],

                "SAP Code":
                    result["sap"],

                "PPT Slide":
                    result["slide"],

                "Status":
                    "✅ MATCHED"
                    if result["matched"]
                    else "❌ NOT MATCHED",

                "Reason":
                    result["reason"]
            })


        st.dataframe(
            pd.DataFrame(result_preview),
            use_container_width=True,
            hide_index=True
        )


        # -------------------------------------------------
        # Process Button
        # -------------------------------------------------

        st.divider()

        if st.button(
            "🚀 Transfer SAP Codes",
            type="primary",
            use_container_width=True
        ):

            with st.spinner(
                "SAP Codes PowerPoint me transfer ho rahe hain..."
            ):

                updated_count, failed_count = update_ppt(
                    prs,
                    matching_results
                )


                # -------------------------------------------------
                # Save PPT
                # -------------------------------------------------

                ppt_output = io.BytesIO()

                prs.save(
                    ppt_output
                )

                ppt_output.seek(0)

                final_ppt_bytes = (
                    ppt_output.getvalue()
                )


                # -------------------------------------------------
                # Filename:
                #
                # Dealer Data September.pptx
                # ↓
                # Dealer Data September_Update.pptx
                # -------------------------------------------------

                original_ppt_name = (
                    ppt_file.name
                )


                base_ppt_name = re.sub(
                    r"\.pptx$",
                    "",
                    original_ppt_name,
                    flags=re.IGNORECASE
                )


                updated_name = (
                    base_ppt_name
                    + "_Update.pptx"
                )


                updated_name = safe_filename(
                    updated_name
                )


                # -------------------------------------------------
                # Excel report filename:
                #
                # Ayodhya Dealer List.xlsx
                # ↓
                # Ayodhya Dealer List_Matching_Report.xlsx
                # -------------------------------------------------

                original_excel_name = (
                    excel_file.name
                )


                excel_base_name = re.sub(
                    r"\.(xlsx|xls)$",
                    "",
                    original_excel_name,
                    flags=re.IGNORECASE
                )


                report_name = (
                    excel_base_name
                    + "_Matching_Report.xlsx"
                )


                report_name = safe_filename(
                    report_name
                )


                # -------------------------------------------------
                # Success
                # -------------------------------------------------

                st.success(
                    f"Done! {updated_count} SAP Code successfully transfer hue."
                )


                if failed_count > 0:

                    st.warning(
                        f"{failed_count} slide update nahi ho paayi."
                    )


                # -------------------------------------------------
                # Download PPT
                # -------------------------------------------------

                st.download_button(

                    label="📥 Download Updated PPT",

                    data=final_ppt_bytes,

                    file_name=updated_name,

                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.presentation"
                    ),

                    use_container_width=True
                )


                # -------------------------------------------------
                # Download Report
                # -------------------------------------------------

                report_bytes = (
                    create_report_excel(
                        matching_results
                    )
                )


                st.download_button(

                    label="📊 Download Matching Report",

                    data=report_bytes,

                    file_name=report_name,

                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),

                    use_container_width=True
                )


                # -------------------------------------------------
                # Important information
                # -------------------------------------------------

                st.info(
                    "Safety rule: jis Excel row ka unique Name + Contact "
                    "match nahi mila, uska SAP Code PPT me transfer nahi kiya gaya."
                )


except Exception as e:

    st.error(
        "Error aa gaya:"
    )

    st.exception(e)
