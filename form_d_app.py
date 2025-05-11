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

    def find_next_value_after(label):
        try:
            idx = lines.index(label)
            for offset in range(1, 5):
                value = lines[idx + offset].strip()
                if value and not value.startswith(label):
                    return value
        except ValueError:
            return "Not found"
        return "Not found"

    cik = find_next_value_after("CIK (Filer ID Number)")
    issuer = find_next_value_after("Name of Issuer")
    year = "Not found"
    if "Within Last Five Years (Specify Year)" in full_text and "2018" in full_text:
        year = "Within Last Five Years (Checked) with year 2018"
    elif "Within Last Five Years (Specify Year)" in full_text:
        year = "Within Last Five Years (Checked), year not found"
    elif "Over Five Years Ago" in full_text:
        year = "Over Five Years Ago (Checked)"
    elif "Yet to Be Formed" in full_text:
        year = "Yet to Be Formed (Checked)"

    # Entity Type detection (look for keywords after label)
    entity_type = "Not found"
    entity_keywords = ["Corporation", "Limited Partnership", "Limited Liability Company",
                       "General Partnership", "Business Trust", "Other"]
    for i, line in enumerate(lines):
        if "Entity Type" in line:
            for j in range(i, i+10):
                if j < len(lines):
                    for keyword in entity_keywords:
                        if keyword in lines[j]:
                            entity_type = keyword
                            break

    # Section 13 — Offering Amounts
    offering = "Not found"
    sold = "Not found"
    remaining = "Not found"
    for i, line in enumerate(lines):
        if "Total Offering Amount" in line:
            match = re.search(r"\$?([0-9,]+)", lines[i+1])
            if match: offering = f"${match.group(1)}"
        if "Total Amount Sold" in line:
            match = re.search(r"\$?([0-9,]+)", lines[i+1])
            if match: sold = f"${match.group(1)}"
        if "Total Remaining" in line:
            match = re.search(r"\$?([0-9,]+)", lines[i+1])
            if match: remaining = f"${match.group(1)}"

    # Section 16 — Use of Proceeds
    use_of_proceeds = "Not found"
    for i, line in enumerate(lines):
        if "Use of Proceeds" in line:
            for j in range(i, i+5):
                if "$" in lines[j]:
                    use_of_proceeds = lines[j].strip()
                    break
            break

    is_valid = "Yes" if cik != "Not found" and offering != "Not found" else "No"
    deal_type = "Tranche" if "Tranche" in full_text else "New"

    comment = f"{issuer} has filed a Form D indicating a {deal_type.lower()} deal. " \
              f"The offering amount is {offering}, of which {sold} has been sold. " \
              f"Proceeds are expected {use_of_proceeds}."

    # Display results
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

    # Excel export
    df_out = pd.DataFrame([data])
    df_out["Comment"] = comment
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False)
    st.download_button(
        label="📥 Download Excel",
        data=buffer.getvalue(),
        file_name="form_d_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
