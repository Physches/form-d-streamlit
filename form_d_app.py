# -------------------------------
# 📦 Step 0: Install Requirements
# -------------------------------
# Run this in terminal before using: pip install streamlit python-docx pandas

import streamlit as st
from docx import Document
import pandas as pd
import re

st.set_page_config(page_title="Form D Extractor", layout="wide")
st.title("📄 SEC Form D Extractor & Comment Generator")

# -------------------------------
# 📥 Step 1: Upload .docx file
# -------------------------------
uploaded_file = st.file_uploader("Upload a Form D (.docx) file", type=["docx"])

if uploaded_file:
    doc = Document(uploaded_file)
    all_lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    full_text = "\n".join(all_lines)

    # -------------------------------
    # 🔍 Step 2: Extract Core Data
    # -------------------------------
    def extract_field(pattern, context, fallback="Not found", flags=re.IGNORECASE):
        match = re.search(pattern, context, flags)
        return match.group(1).strip() if match else fallback

    # CIK
    cik = extract_field(r"CIK.*?(\d{8,10})", full_text)
    
    # Issuer Name
    issuer = extract_field(r"Name of Issuer\s*\n*([A-Za-z0-9 ,.&\-']+)", full_text)

    # Year of Incorporation
    if "Within Last Five" in full_text and "2018" in full_text:
        year_info = "Within Last Five Years (Checked) with year 2018"
    elif "Within Last Five" in full_text:
        year_info = "Within Last Five Years (Checked), year not found"
    elif "Over Five Years" in full_text:
        year_info = "Over Five Years Ago (Checked)"
    elif "Yet to Be" in full_text:
        year_info = "Yet to Be Formed (Checked)"
    else:
        year_info = "Not found"

    # Entity Type
    entity_type = "Not found"
    for etype in ["Corporation", "Limited Partnership", "Limited Liability Company",
                  "General Partnership", "Business Trust", "Other"]:
        if re.search(r"[ΠC✓Xx]\s*" + re.escape(etype), full_text, re.IGNORECASE):
            entity_type = etype
            break

    # Section 13: Financials
    offering, sold, remaining = "Not found", "Not found", "Not found"
    for i, line in enumerate(all_lines):
        if "total offering" in line.lower():
            for j in [i-2, i-1, i+1, i+2]:
                match = re.search(r"\$?\s*([0-9]{4,})", all_lines[j])
                if match: offering = f"${match.group(1)}"; break
        if re.fullmatch(r"Sold", line.strip(), re.IGNORECASE):
            for j in [i-2, i-1, i+1, i+2]:
                match = re.search(r"\$?\s*([0-9]{4,})", all_lines[j])
                if match: sold = f"${match.group(1)}"; break
        if "total remaining" in line.lower():
            for j in [i-2, i-1, i+1, i+2]:
                match = re.search(r"\$?\s*([0-9]{4,})", all_lines[j])
                if match: remaining = f"${match.group(1)}"; break

    # Section 16: Use of Proceeds
    use_of_proceeds = "Not found"
    for i, line in enumerate(all_lines):
        if "use of proceeds" in line.lower():
            match = re.search(r"Use of Proceeds.*?(to\s+[a-zA-Z0-9 ,\-]+)", line, re.IGNORECASE)
            if match:
                use_of_proceeds = match.group(1)
                break
            for j in [i+1, i+2]:
                match = re.search(r"(to\s+[a-zA-Z0-9 ,\-]+)", all_lines[j])
                if match:
                    use_of_proceeds = match.group(1)
                    break
            break

    # -------------------------------
    # 🧠 Step 3: Basic Logic/Validation
    # -------------------------------
    is_valid = "Yes" if cik != "Not found" and offering != "Not found" else "No"
    deal_type = "Tranche" if "Tranche" in full_text else "New"

    # -------------------------------
    # 🧾 Step 4: Comment Template
    # -------------------------------
    comment = f"{issuer} has filed a Form D indicating a {deal_type.lower()} deal. " \
              f"The offering amount is {offering}, of which {sold} has been sold. " \
              f"Proceeds are expected {use_of_proceeds}."

    # -------------------------------
    # 📊 Step 5: Display Results
    # -------------------------------
    st.subheader("📌 Extracted Data:")
    data = {
        "CIK": cik,
        "Issuer": issuer,
        "Year of Incorporation": year_info,
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

    # -------------------------------
    # 💾 Step 6: Download Option
    # -------------------------------
    df_out = pd.DataFrame([data])
    df_out["Comment"] = comment
    st.download_button("📥 Download Excel", df_out.to_excel(index=False), "form_d_output.xlsx")
