import streamlit as st
import pandas as pd
from deep_translator import GoogleTranslator
from datetime import datetime

st.set_page_config(page_title="Flyone Report Tool", layout="wide")

translator = GoogleTranslator(source='auto', target='hy')

st.title("🛫 Flyone Report Tool (Web Version)")

uploaded_excel = st.file_uploader("📁 Upload Excel File", type=["xlsx"])
start_date = st.date_input("📅 Start Date")
end_date = st.date_input("📅 End Date")

if uploaded_excel:
    df = pd.read_excel(uploaded_excel, sheet_name=None)
    sheet_names = list(df.keys())
    selected_sheet = st.selectbox("📑 Choose Excel sheet", sheet_names)
    data = df[selected_sheet]

    # Clean column names
    data.columns = data.columns.str.strip()
    
    # Debug: Show actual column names (optional)
    # st.write("Columns found:", data.columns.tolist())

    if 'Date & Time of Event (UTC)' in data.columns:
        data['Date & Time of Event (UTC)'] = pd.to_datetime(data['Date & Time of Event (UTC)'], errors='coerce')

        # Filter by selected dates
        filtered = data[
            (data['Date & Time of Event (UTC)'] >= pd.to_datetime(start_date)) &
            (data['Date & Time of Event (UTC)'] <= pd.to_datetime(end_date))
        ]

        if not filtered.empty:
            st.success(f"✅ Found {len(filtered)} reports between selected dates.")
            translations = []

            # Check required columns exist
            if 'Aircraft Registration' in filtered.columns and 'Type of report' in filtered.columns:
                grouped = filtered.groupby(['Aircraft Registration', 'Type of report'])

                for (aircraft, report_type), group in grouped:
                    st.markdown(f"### ✈️ {aircraft} — 🗂️ {report_type}")
                    for idx, row in group.iterrows():
                        original = str(row.get("Details", ""))
                        st.markdown(f"**📄 Original:** {original}")

                        try:
                            translated = translator.translate(original)
                        except Exception as e:
                            translated = "[Թարգմանությունը ձախողվեց]"

                        new_text = st.text_area(f"✏️ Edit Translation [{idx}]", translated, key=f"edit_{idx}")
                        translations.append({
                            "Aircraft": aircraft,
                            "Type": report_type,
                            "Date": row.get("Date & Time of Event (UTC)"),
                            "Flight Number": row.get("Flight Number"),
                            "Original": original,
                            "Translation": new_text
                        })

                if st.button("📥 Export Translated Reports"):
                    export_df = pd.DataFrame(translations)
                    export_df.to_excel("translated_reports.xlsx", index=False)
                    st.success("✅ Translations saved as translated_reports.xlsx")
                    with open("translated_reports.xlsx", "rb") as file:
                        st.download_button("⬇️ Download File", file, file_name="translated_reports.xlsx")
            else:
                st.error("❌ Required columns missing: 'Aircraft Registration' or 'Type of report'")
        else:
            st.warning("No reports found in selected date range.")
    else:
        st.error("❌ 'Date & Time of Event (UTC)' column not found.")
from docx import Document
import io
import os

def insert_into_word(translations, template_path, start_date, end_date):
    doc = Document(template_path)
    section_map = {
        "Տեխնիկական": "Տեխնիկական զեկույցներ՝",
        "Թռիչք": "Թռիչքային զեկույցներ՝",
        "Վերգետնյա": "Վերգետնյա սպասարկում/Նստեցման հետ կապված խնդիրներ՝",
        "Բողոք": "Ուղևորների բողոքներ",
        "աղտոտ": "Օդանավի աղտոտվածության վերաբերյալ զեկույցներ՝",
        "այլ": "Այլ խնդիրներ"
    }

    for paragraph in doc.paragraphs:
        for key, title in section_map.items():
            if title.strip() in paragraph.text.strip():
                idx = doc.paragraphs.index(paragraph)
                for i, table in enumerate(doc.tables):
                    if doc.paragraphs.index(doc.paragraphs[idx + 1]) < doc.paragraphs.index(table._element.getparent()):
                        target_table = table
                        break
                else:
                    continue

                for entry in translations:
                    text = entry["Translation"]
                    if key in text.lower():
                        row_cells = target_table.add_row().cells
                        row_cells[0].text = entry["Aircraft"]
                        row_cells[1].text = entry["Flight Number"] or ""
                        row_cells[2].text = entry["Date"].strftime("%Y-%m-%d %H:%M") if entry["Date"] else ""
                        row_cells[3].text = text.strip()
                break

    # Save file in memory
    output = io.BytesIO()
    output_name = f"Translated_Report_{start_date.strftime('%d.%m.%y')}-{end_date.strftime('%d.%m.%y')}.docx"
    doc.save(output)
    output.seek(0)
    return output, output_name


# Add this button where others are
if translations and st.button("📝 Export to Word Template"):
    try:
        template_path = "template/Զեկույցների ցանկ.docx"
        word_file, word_filename = insert_into_word(translations, template_path, start_date, end_date)
        st.download_button("⬇️ Download Word Report", word_file, file_name=word_filename)
        st.success("✅ Word document generated successfully.")
    except Exception as e:
        st.error(f"❌ Failed to create Word file: {e}")
