import streamlit as st
from docx import Document
import pandas as pd
import re
import io

st.set_page_config(page_title="Form D Extractor", layout="wide")
st.title("📄 SEC Form D Extractor & Comment Generator")

uploaded_file = st.file_uploader("Upload a Form D (.docx) file", type=["docx"])

if uploaded_file:
    doc = Document(uploaded_file)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    full_text = "\n".join(lines)

    # --- FLEXIBLE PARSING FUNCTIONS ---
    def find_value_by_keyword(keyword, pattern=r"\d{8,10}"):
        for line in lines:
            if keyword.lower() in line.lower():
                match = re.search(pattern, line)
                if match:
                    return match.group()
        return "Not found"

    def find_text_after_label(label):
        for i, line in enumerate(lines):
            if label.lower() in line.lower():
                parts = line.split(label)
                if len(parts) > 1 and parts[1].strip():
                    return parts[1].strip()
                elif i + 1 < len(lines):
                    return lines[i + 1].strip()
        return "Not found"

    # --- FIELD EXTRACTION ---
    cik = find_value_by_keyword("CIK")
    issuer = find_text_after_label("Name of Issuer")

    # Year of Incorporation logic
    year = "Not found"
    if "Within Last Five Years (Specify Year)" in full_text and "2018" in full_text:
        year = "Within Last Five Years (Checked) with year 2018"
    elif "Within Last Five Years (Specify Year)" in full_text:
        year = "Within Last Five Years (Checked), year not found"
    elif "Over Five Years Ago" in full_text:
        year = "Over Five Years Ago (Checked)"
    elif "Yet to Be Formed" in full_text:
        year = "Yet to Be Formed (Checked)"

    # Entity Type detection
    entity_type = "Not found"
    entity_keywords = [
        "Corporation", "Limited Partnership", "Limited Liability Company",
        "General Partnership", "Business Trust", "Other"
    ]
    for line in lines:
        for keyword in entity_keywords:
            if keyword.lower() in line.lower():
                entity_type = keyword
                break

    # Section 13 — Offering Amounts
    offering = "Not found"
    sold = "Not found"
    remaining = "Not found"
    for i, line in enumerate(lines):
        if "Total Offering Amount" in line:
            match = re.search(r"\$?([\d,]+)", line)
            if match: offering = f"${match.group(1)}"
        if "Total Amount Sold" in line:
            match = re.search(r"\$?([\d,]+)", line)
            if match: sold = f"${match.group(1)}"
        if "Total Remaining" in line:
            match = re.search(r"\$?([\d,]+)", line)
            if match: remaining = f"${match.group(1)}"

    # Section 16 — Use of Proceeds
    use_of_proceeds = "Not found"
    for i, line in enumerate(lines):
        if "Use of Proceeds" in line:
            for j in range(i, i + 3):
                if j < len(lines):
                    use_of_proceeds = lines[j + 1].strip()
                    break
            break

    is_valid = "Yes" if cik != "Not found" and offering != "Not found" else "No"
    deal_type = "Tranche" if "Tranche" in full_text else "New"

    comment = f"{issuer} has filed a Form D indicating a {deal_type.lower()} deal. " \
              f"The offering amount is {offering}, of which {sold} has been sold. " \
              f"Proceeds are expected {use_of_proceeds}."

    # --- DISPLAY DATA ---
    st.subheader("📌 Extracted Data:")
    data = {
        "CIK": cik,
        "Issuer": issuer,
        "Year of Incorporation": year,
        "Entity Type": entity_type,
        "Total Offering": offering,
        "Amount Sold": sold,
        "Remaining": remaining,
        "Use of Proceeds": use_of_proceeds,
        "Valid Filing?": is_valid,
        "Deal Type": deal_type
    }
    st.table(pd.DataFrame(data.items(), columns=["Field", "Value"]))

    st.subheader("📝 Auto-Generated Comment:")
    st.code(comment, language="markdown")

    # --- EXCEL EXPORT ---
    df_out = pd.DataFrame([data])
    df_out["Comment"] = comment

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_out.to_excel(writer, index=False)
    output.seek(0)

    st.download_button(
        label="📥 Download Excel",
        data=output,
        file_name="form_d_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
