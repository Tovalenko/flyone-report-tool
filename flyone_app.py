import streamlit as st
import pandas as pd
from googletrans import Translator
from datetime import datetime

st.set_page_config(page_title="Flyone Report Tool", layout="wide")
translator = Translator()

st.title("🛫 Flyone Report Tool (Web Version)")

uploaded_excel = st.file_uploader("📁 Upload Excel File", type=["xlsx"])
start_date = st.date_input("📅 Start Date")
end_date = st.date_input("📅 End Date")

if uploaded_excel:
    df = pd.read_excel(uploaded_excel, sheet_name=None)
    sheet_names = list(df.keys())
    selected_sheet = st.selectbox("Choose sheet", sheet_names)
    data = df[selected_sheet]
    
    if 'Date & Time of Event (UTC)' in data.columns:
        data['Date & Time of Event (UTC)'] = pd.to_datetime(data['Date & Time of Event (UTC)'], errors='coerce')
        filtered = data[
            (data['Date & Time of Event (UTC)'] >= pd.to_datetime(start_date)) &
            (data['Date & Time of Event (UTC)'] <= pd.to_datetime(end_date))
        ]
        
        if not filtered.empty:
            st.success(f"✅ Found {len(filtered)} reports between selected dates.")
            translated_texts = []

            for idx, row in filtered.iterrows():
                original = str(row.get("Details", ""))
                st.markdown("---")
                st.markdown(f"**✈️ Aircraft:** {row.get('Aircraft Registration', '')}")
                st.markdown(f"**📅 Date:** {row.get('Date & Time of Event (UTC)')}")

                try:
                    translated = translator.translate(original, dest="hy").text
                except Exception:
                    translated = "[Թարգմանությունը ձախողվեց]"
                
                new_text = st.text_area(f"📝 Translation for report {idx}", translated, key=f"trans_{idx}")
                translated_texts.append(new_text)
            
            if st.button("📥 Export Translations"):
                output_df = filtered.copy()
                output_df["Translated Details (HY)"] = translated_texts
                output_df.to_excel("translated_reports.xlsx", index=False)
                st.success("✅ File saved as translated_reports.xlsx")
        else:
            st.warning("No reports found in selected date range.")
    else:
        st.error("❌ 'Date & Time of Event (UTC)' column not found.")
