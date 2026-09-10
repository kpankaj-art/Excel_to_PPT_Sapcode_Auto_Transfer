import io
import re
import copy
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Pt

st.set_page_config(page_title="Excel → PPT Safe SAP Transfer", layout="wide")

# ============================================================
# 1) Excel column aliases
# ============================================================
FIELD_ALIASES = {
    "outlet_name": [
        "outlet name", "outlet", "dealer name", "dealer / name", "dealer/name",
        "dealer", "customer name", "retailer name", "retailer", "shop name",
        "party name", "customer"
    ],
    "address": [
        "address", "dealer address", "dealer / address", "dealer / adderess",
        "outlet address", "customer address", "retailer address", "location", "shop address"
    ],
    "contact_no": [
        "contact", "contact no", "contact no.", "dealer contact", "dealer / contact",
        "mobile", "mobile no", "phone", "phone no", "contact number", "mobile number", "telephone"
    ],
    "district": ["district", "district name", "dist", "dist name"],
    "sapcode": [
        "sapcode", "sap code", "dealer code", "dealercode", "customer code",
        "customercode", "customer id", "dealer id", "sap id", "sap number"
    ],
    "size": ["size", "size inches", "type and size", "type & size", "dimension", "dimensions", "w x h", "width x height"],
    "media_type": ["media type", "media", "type", "board type", "material type"],
    "remarks": ["remarks", "remark", "comments", "comment", "note", "notes"],
    "qty": ["qty", "quantity", "qnty"],
}

PPT_LABELS = {
    "outlet_name": ["outlet name", "outlet", "dealer name", "dealer", "customer name", "retailer name"],
    "address": ["address", "dealer address", "outlet address", "customer address"],
    "contact_no": ["contact no", "contact", "mobile", "phone", "contact number"],
    "district": ["district", "district name", "dist"],
    "sapcode": ["sapcode", "sap code", "dealer code", "customer code", "customercode"],
    "size": ["size", "type and size", "type & size", "dimension"],
    "media_type": ["media type", "media"],
    "remarks": ["remarks", "remark", "comments", "comment"],
    "qty": ["qty", "quantity"],
}

CANONICAL_LABELS = {
    "outlet_name": "Outlet Name",
    "address": "Address",
    "contact_no": "Contact No",
    "district": "District",
    "sapcode": "Sapcode",
    "size": "Size",
    "media_type": "Media Type",
    "remarks": "Remarks",
    "qty": "Qty",
}

# ============================================================
# 2) Normalisation helpers
# ============================================================
def norm(s):
    s = "" if s is None else str(s)
    s = s.lower().strip()
    s = s.replace("&", " and ")
    s = re.sub(r"[\u2013\u2014_/().:-]+", " ", s)
    s = re.sub(r"[^a-z0-9x ]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def compact(s):
    return re.sub(r"[^a-z0-9]", "", norm(s))


def similarity(a, b):
    return SequenceMatcher(None, compact(a), compact(b)).ratio()


def phone_numbers(s):
    """Return usable phone-number digit strings from values such as 945.../800...."""
    s = "" if s is None else str(s)
    nums = re.findall(r"\d{7,15}", s)
    # Also handle Excel numeric values and numbers embedded with spaces.
    if not nums:
        digits = re.sub(r"\D", "", s)
        if len(digits) >= 7:
            nums = [digits]
    return {n[-10:] if len(n) >= 10 else n for n in nums}


def normal_size(s):
    """Convert 600X48, 600 x 48, 600*48 into 600x48."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    text = str(s).lower().replace("×", "x").replace("*", "x")
    nums = re.findall(r"\d+(?:\.\d+)?", text)
    if len(nums) >= 2:
        return f"{nums[0]}x{nums[1]}"
    return compact(text)


def make_excel_size(row, mapping):
    size_col = mapping.get("size")
    if size_col:
        s = row.get(size_col, "")
        if not pd.isna(s) and str(s).strip():
            return normal_size(s)

    w_col = mapping.get("width")
    h_col = mapping.get("height")
    if w_col and h_col:
        w, h = row.get(w_col, ""), row.get(h_col, "")
        if not pd.isna(w) and not pd.isna(h):
            return normal_size(f"{w}x{h}")
    return ""


def value(row, mapping, field, default=""):
    col = mapping.get(field)
    if not col:
        return default
    v = row.get(col, default)
    if pd.isna(v):
        return default
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

# ============================================================
# 3) Excel column detection
# ============================================================
def detect_columns(df):
    result = {}
    normalized_headers = {col: norm(col) for col in df.columns}

    for field, aliases in FIELD_ALIASES.items():
        best_col, best_score = None, 0
        for col, ncol in normalized_headers.items():
            scores = [similarity(ncol, alias) for alias in aliases]
            score = max(scores) if scores else 0
            for alias in aliases:
                na = norm(alias)
                if na == ncol:
                    score = 1.0
                elif na in ncol or ncol in na:
                    score = max(score, 0.92)
            if score > best_score:
                best_col, best_score = col, score
        if best_col is not None and best_score >= 0.58:
            result[field] = best_col

    # W/H are intentionally separate because the sample Excel has W and H.
    for key, aliases in {
        "width": ["w", "width", "width inches"],
        "height": ["h", "height", "height inches"],
    }.items():
        for col in df.columns:
            n = norm(col)
            if n in aliases:
                result[key] = col
                break
    return result

# ============================================================
# 4) PPT text extraction
# ============================================================
def detect_ppt_field(text):
    ntext = norm(text)
    for field, labels in PPT_LABELS.items():
        for label in labels:
            nl = norm(label)
            if re.search(rf"(^|\n)\s*{re.escape(nl)}\s*:?", ntext):
                return field
    return None


def extract_labeled_values(slide):
    """Extract label:value pairs from all text boxes on a PPT slide."""
    data = {}
    for shape in slide.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
        for line in shape.text.splitlines():
            raw = line.strip()
            if not raw:
                continue
            for field, labels in PPT_LABELS.items():
                for label in labels:
                    m = re.match(rf"^\s*{re.escape(label)}\s*:\s*(.*?)\s*$", raw, re.I)
                    if m:
                        data[field] = m.group(1).strip()
                        break
                if field in data:
                    break
    return data

# ============================================================
# 5) Safe matching: NAME + PHONE, then SIZE only for duplicates
# ============================================================
def name_matches(excel_name, ppt_name):
    a, b = compact(excel_name), compact(ppt_name)
    if not a or not b:
        return False
    if a == b:
        return True
    # Conservative fuzzy match for small OCR/typing differences such as
    # TRADRES vs TRADERS. Never use fuzzy name without phone verification.
    return similarity(excel_name, ppt_name) >= 0.92


def phone_matches(excel_phone, ppt_phone):
    a, b = phone_numbers(excel_phone), phone_numbers(ppt_phone)
    return bool(a and b and a.intersection(b))


def size_matches(a, b):
    na, nb = normal_size(a), normal_size(b)
    return bool(na and nb and na == nb)


def build_ppt_index(prs):
    index = []
    for idx, slide in enumerate(prs.slides, start=1):
        fields = extract_labeled_values(slide)
        index.append({"slide_no": idx, "fields": fields})
    return index


def match_excel_row(row, mapping, ppt_index):
    excel_name = value(row, mapping, "outlet_name")
    excel_phone = value(row, mapping, "contact_no")
    excel_size = make_excel_size(row, mapping)

    if not excel_name or not excel_phone:
        return {"status": "NO_MATCH", "reason": "Name or contact missing", "slide_no": ""}

    # Primary match: NAME + CONTACT.
    candidates = []
    for item in ppt_index:
        f = item["fields"]
        if name_matches(excel_name, f.get("outlet_name", "")) and phone_matches(excel_phone, f.get("contact_no", "")):
            candidates.append(item)

    if not candidates:
        return {"status": "NO_MATCH", "reason": "Name + contact not found", "slide_no": ""}

    if len(candidates) == 1:
        # Name + phone uniquely identifies the slide; size is not needed.
        return {"status": "MATCHED", "reason": "Unique name + contact", "slide_no": candidates[0]["slide_no"]}

    # Duplicate NAME + CONTACT: size becomes mandatory.
    if not excel_size:
        return {"status": "AMBIGUOUS", "reason": "Multiple name + contact matches; size missing", "slide_no": ""}

    size_candidates = [c for c in candidates if size_matches(excel_size, c["fields"].get("size", ""))]
    if len(size_candidates) == 1:
        return {"status": "MATCHED", "reason": "Name + contact + size", "slide_no": size_candidates[0]["slide_no"]}
    if len(size_candidates) > 1:
        return {"status": "AMBIGUOUS", "reason": "Name + contact + size still has multiple matches", "slide_no": ""}
    return {"status": "NO_MATCH", "reason": "Name + contact matched, but size did not", "slide_no": ""}

# ============================================================
# 6) PPT update
# ============================================================
def line_field(line):
    nline = norm(line)
    for field, labels in PPT_LABELS.items():
        for label in labels:
            nl = norm(label)
            if nline.startswith(nl + " ") or nline.startswith(nl + ":") or nline == nl:
                return field
    return None


def replace_field_in_text(text, field, new_value):
    lines = text.splitlines()
    changed = False
    for i, line in enumerate(lines):
        if line_field(line) != field:
            continue
        colon = line.find(":")
        if colon >= 0:
            prefix = line[:colon + 1]
            spaces = re.match(r"\s*", line[colon + 1:]).group(0)
            lines[i] = prefix + spaces + str(new_value)
        else:
            # Preserve the existing label if possible.
            label = line.strip()
            lines[i] = f"{label} : {new_value}"
        changed = True
    return "\n".join(lines), changed


def update_slide_with_sap(slide, sapcode):
    """Update existing Sapcode field, or add it after District in the main info block."""
    if not sapcode:
        return False

    # First: if Sapcode already exists, replace it.
    for shape in slide.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
        if detect_ppt_field(shape.text) == "sapcode":
            new_text, changed = replace_field_in_text(shape.text, "sapcode", sapcode)
            if changed:
                shape.text = new_text
                return True

    # Find the text box containing the main dealer fields.
    best_shape = None
    best_count = -1
    for shape in slide.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
        fields = {line_field(line) for line in shape.text.splitlines()}
        fields.discard(None)
        score = len(fields.intersection({"outlet_name", "address", "contact_no", "district"}))
        if score > best_count:
            best_count = score
            best_shape = shape

    if best_shape is not None and best_count > 0:
        lines = best_shape.text.splitlines()
        # Avoid duplicate insertion.
        if any(line_field(x) == "sapcode" for x in lines):
            return False
        insert_at = len(lines)
        for i, line in enumerate(lines):
            if line_field(line) == "district":
                insert_at = i + 1
                break
        lines.insert(insert_at, f"Sapcode : {sapcode}")
        best_shape.text = "\n".join(lines)
        try:
            for paragraph in best_shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.size is None:
                        run.font.size = Pt(12)
        except Exception:
            pass
        return True

    # Last resort: add a small text box.
    from pptx.util import Inches
    box = slide.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(4), Inches(0.35))
    box.text = f"Sapcode : {sapcode}"
    for p in box.text_frame.paragraphs:
        for run in p.runs:
            run.font.size = Pt(12)
    return True

# ============================================================
# 7) Main transfer
# ============================================================
def process_transfer(template_bytes, df, mapping):
    prs = Presentation(io.BytesIO(template_bytes))
    if not prs.slides:
        raise ValueError("PPT template has no slides.")

    ppt_index = build_ppt_index(prs)
    results = []
    used_slides = {}

    for excel_idx, row in df.iterrows():
        match = match_excel_row(row, mapping, ppt_index)
        sapcode = value(row, mapping, "sapcode")
        result = {
            "Excel Row": int(excel_idx) + 2,
            "Outlet Name": value(row, mapping, "outlet_name"),
            "Contact No": value(row, mapping, "contact_no"),
            "Size": make_excel_size(row, mapping),
            "Sapcode": sapcode,
            "Status": match["status"],
            "Reason": match["reason"],
            "PPT Slide": match["slide_no"],
        }

        if match["status"] == "MATCHED":
            slide_no = match["slide_no"]
            # Prevent two different Excel rows from overwriting the same PPT slide.
            if slide_no in used_slides:
                result["Status"] = "AMBIGUOUS"
                result["Reason"] = f"PPT slide already matched to Excel row {used_slides[slide_no]}"
            elif not sapcode:
                result["Status"] = "NO_MATCH"
                result["Reason"] = "Matched slide, but SAP code is blank"
            else:
                update_slide_with_sap(prs.slides[slide_no - 1], sapcode)
                used_slides[slide_no] = result["Excel Row"]

        results.append(result)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.getvalue(), pd.DataFrame(results)

def to_excel_bytes(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Matching Report")
    out.seek(0)
    return out.getvalue()

# ============================================================
# 8) Streamlit UI
# ============================================================
st.title("Excel → PPT Safe SAP Transfer")
st.caption("Safe matching: Name + Contact first. If duplicate, Size is mandatory. Uncertain matches are never transferred.")

c1, c2 = st.columns(2)
with c1:
    excel_file = st.file_uploader("1. Upload Excel", type=["xlsx", "xls"])
with c2:
    ppt_file = st.file_uploader("2. Upload PPT", type=["pptx"])

if excel_file and ppt_file:
    try:
        df = pd.read_excel(excel_file)
        mapping = detect_columns(df)

        st.subheader("Detected Excel Mapping")
        pretty = {
            "Outlet Name": mapping.get("outlet_name", "Not found"),
            "Address": mapping.get("address", "Not found"),
            "Contact No": mapping.get("contact_no", "Not found"),
            "District": mapping.get("district", "Not found"),
            "Sapcode": mapping.get("sapcode", "Not found"),
            "Size / W+H": mapping.get("size", mapping.get("width", "Not found")),
            "Media Type": mapping.get("media_type", "Not found"),
        }
        st.table(pd.DataFrame(list(pretty.items()), columns=["Field", "Excel Column"]))
        st.write(f"**Excel data rows:** {len(df)} | **PPT slides:** {len(Presentation(io.BytesIO(ppt_file.getvalue())).slides)}")

        if st.button("Analyse & Transfer SAP Codes", type="primary"):
            with st.spinner("Matching Name → Contact → Size and updating PPT..."):
                result_bytes, report = process_transfer(ppt_file.getvalue(), df, mapping)

            matched = int((report["Status"] == "MATCHED").sum())
            ambiguous = int((report["Status"] == "AMBIGUOUS").sum())
            no_match = int((report["Status"] == "NO_MATCH").sum())

            st.success(f"Completed: {matched} matched | {ambiguous} ambiguous | {no_match} not matched")
            st.dataframe(report, use_container_width=True)

            st.download_button(
                "Download Final PPT",
                data=result_bytes,
                file_name="Final_Sapcode_PPT.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.download_button(
                "Download Matching Report (Excel)",
                data=to_excel_bytes(report),
                file_name="PPT_Matching_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Error: {e}")
