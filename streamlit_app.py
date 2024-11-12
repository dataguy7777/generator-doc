import streamlit as st
from docx import Document
from docx.shared import Inches
import pandas as pd
import io

# Initialize session state for storing content
if 'paragraphs' not in st.session_state:
    st.session_state['paragraphs'] = []
if 'tables' not in st.session_state:
    st.session_state['tables'] = []

def add_paragraph(paragraph_text):
    st.session_state['paragraphs'].append(paragraph_text)

def add_table(table_df):
    st.session_state['tables'].append(table_df)

def generate_preview():
    previews = []
    for idx, para in enumerate(st.session_state['paragraphs']):
        previews.append(f"**Paragraph {idx + 1}:** {para}")
    for idx, table in enumerate(st.session_state['tables']):
        previews.append(f"**Table {idx + 1}:**")
        previews.append(table)
    return previews

def create_word_document():
    doc = Document()
    for para in st.session_state['paragraphs']:
        doc.add_paragraph(para)
    for table in st.session_state['tables']:
        if isinstance(table, pd.DataFrame):
            table_doc = doc.add_table(rows=1, cols=len(table.columns))
            hdr_cells = table_doc.rows[0].cells
            for i, column in enumerate(table.columns):
                hdr_cells[i].text = str(column)
            for _, row in table.iterrows():
                row_cells = table_doc.add_row().cells
                for i, item in enumerate(row):
                    row_cells[i].text = str(item)
            doc.add_paragraph("")  # Add a space after table
    # Save to a BytesIO buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.title("Word Document Generator")

st.header("Add Paragraphs")
paragraph_input = st.text_area("Enter your paragraph here:")
if st.button("Add Paragraph"):
    if paragraph_input.strip() != "":
        add_paragraph(paragraph_input)
        st.success("Paragraph added!")
    else:
        st.error("Paragraph cannot be empty.")

st.header("Add Tables")
with st.form("table_form"):
    num_rows = st.number_input("Number of rows", min_value=1, value=2)
    num_cols = st.number_input("Number of columns", min_value=1, value=2)
    table_data = []
    for row in range(int(num_rows)):
        row_data = []
        for col in range(int(num_cols)):
            cell = st.text_input(f"Row {row + 1}, Column {col + 1}", key=f"cell_{row}_{col}")
            row_data.append(cell)
        table_data.append(row_data)
    submitted = st.form_submit_button("Add Table")
    if submitted:
        df = pd.DataFrame(table_data, columns=[f"Column {i+1}" for i in range(int(num_cols))])
        add_table(df)
        st.success("Table added!")

st.header("Preview Document")
previews = generate_preview()
for preview in previews:
    if isinstance(preview, str):
        st.markdown(preview)
    elif isinstance(preview, pd.DataFrame):
        st.table(preview)

if st.button("Generate and Download Word Document"):
    if not st.session_state['paragraphs'] and not st.session_state['tables']:
        st.error("No content to generate. Please add paragraphs or tables.")
    else:
        doc_buffer = create_word_document()
        st.download_button(
            label="Download Word Document",
            data=doc_buffer,
            file_name="generated_document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
