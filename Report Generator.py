import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Faculty Report Generator", layout="wide")
st.title("🎓 Faculty Report Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    excel_data = pd.ExcelFile(uploaded_file)
    sheet_name = "Students by Faculty"
    
    if sheet_name not in excel_data.sheet_names:
        st.error(f"'{sheet_name}' sheet not found in the uploaded file.")
    else:
        df = excel_data.parse(sheet_name)
        faculties = df['Faculty'].dropna().unique().tolist()
        selected_faculties = st.multiselect("Select Faculty/Faculties", faculties, default=faculties)

        for faculty in selected_faculties:
            faculty_df = df[df['Faculty'] == faculty]
            st.subheader(f"📊 Report for {faculty}")

            col1, col2 = st.columns(2)
            with col1:
                fig, ax = plt.subplots()
                faculty_df['Flag'].value_counts().plot(kind='bar', ax=ax, color='teal')
                ax.set_title("Flag Distribution")
                st.pyplot(fig)

            with col2:
                fig2, ax2 = plt.subplots()
                faculty_df['Entrance Group'].value_counts().plot(kind='bar', ax=ax2, color='coral')
                ax2.set_title("Entrance Group Distribution")
                st.pyplot(fig2)

            # Generate Word report in memory
            doc = Document()
            doc.add_heading(f"Report for {faculty}", level=1)

            doc.add_paragraph(f"Total Students: {len(faculty_df)}")
            flag_counts = faculty_df['Flag'].value_counts()
            doc.add_heading("Flag Distribution", level=2)
            for flag, count in flag_counts.items():
                doc.add_paragraph(f"{flag}: {count}")

            entrance_counts = faculty_df['Entrance Group'].value_counts()
            doc.add_heading("Entrance Group Distribution", level=2)
            for group, count in entrance_counts.items():
                doc.add_paragraph(f"{group}: {count}")

            # Save report to memory for download
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            st.download_button(
                label=f"📥 Download Word Report for {faculty}",
                data=buffer,
                file_name=f"{faculty}_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
