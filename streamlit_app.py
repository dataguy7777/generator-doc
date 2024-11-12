import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
import pandas as pd
import io
import os
from PIL import Image
import networkx as nx
from pyvis.network import Network
import tempfile

# ---------------------------
# Session State Initialization
# ---------------------------
if 'cover_page' not in st.session_state:
    st.session_state['cover_page'] = None
if 'paragraphs' not in st.session_state:
    st.session_state['paragraphs'] = []
if 'tables' not in st.session_state:
    st.session_state['tables'] = []
if 'images' not in st.session_state:
    st.session_state['images'] = []
if 'templates' not in st.session_state:
    st.session_state['templates'] = {}
if 'document_graph' not in st.session_state:
    st.session_state['document_graph'] = nx.DiGraph()

# ---------------------------
# Helper Functions
# ---------------------------

def load_templates(folder_path):
    """
    Load .docx templates and corresponding images from the specified folder.
    Assumes that each .docx file has a corresponding image with the same base name.
    Supported image formats: .jpg, .jpeg, .png
    """
    templates = {}
    supported_image_formats = ['.jpg', '.jpeg', '.png']
    
    if not os.path.isdir(folder_path):
        st.error(f"The folder path '{folder_path}' does not exist or is not a directory.")
        return templates

    files = os.listdir(folder_path)
    docx_files = [f for f in files if f.lower().endswith('.docx')]

    for docx in docx_files:
        base_name = os.path.splitext(docx)[0]
        # Look for corresponding image
        image_path = None
        for ext in supported_image_formats:
            img_file = f"{base_name}{ext}"
            if img_file in files:
                image_path = os.path.join(folder_path, img_file)
                break
        if image_path:
            templates[base_name] = {
                "template_path": os.path.join(folder_path, docx),
                "image_path": image_path
            }
        else:
            st.warning(f"No corresponding image found for template '{docx}'. Skipping this template.")
    
    if not templates:
        st.error("No valid templates found in the specified folder.")
    return templates

def add_paragraph():
    st.subheader("Add Paragraph")
    parent_paragraph = st.text_area("Enter main paragraph:", key="parent_paragraph")
    if st.button("Add Paragraph"):
        if parent_paragraph.strip() != "":
            para_id = len(st.session_state['paragraphs']) + 1
            st.session_state['paragraphs'].append({
                "id": para_id,
                "content": parent_paragraph,
                "sub_paragraphs": [],
                "comments": []
            })
            st.session_state['document_graph'].add_node(f"Paragraph {para_id}", label=parent_paragraph, type='paragraph')
            st.success("Paragraph added!")
        else:
            st.error("Paragraph cannot be empty.")
    
    if st.session_state['paragraphs']:
        st.markdown("---")
        st.subheader("Existing Paragraphs")
        for para in st.session_state['paragraphs']:
            st.markdown(f"**Paragraph {para['id']}:** {para['content']}")
            # Sub-paragraphs
            for sub_idx, sub in enumerate(para['sub_paragraphs']):
                st.markdown(f"&nbsp;&nbsp;&nbsp;**Sub-paragraph {para['id']}.{sub_idx + 1}:** {sub}")
            # Comments
            for com_idx, com in enumerate(para['comments']):
                st.markdown(f"&nbsp;&nbsp;&nbsp;**Comment {para['id']}.{com_idx + 1}:** {com}")

def add_sub_paragraph():
    st.subheader("Add Sub-Paragraph")
    if not st.session_state['paragraphs']:
        st.warning("Add a main paragraph first.")
    else:
        para_options = [f"Paragraph {para['id']}" for para in st.session_state['paragraphs']]
        selected_para = st.selectbox("Select Paragraph to Add Sub-Paragraph", para_options)
        sub_para = st.text_area("Enter sub-paragraph:", key="sub_paragraph")
        if st.button("Add Sub-Paragraph"):
            if sub_para.strip() != "":
                para_id = int(selected_para.split(" ")[1])
                para = next((p for p in st.session_state['paragraphs'] if p['id'] == para_id), None)
                if para:
                    para['sub_paragraphs'].append(sub_para)
                    child_id = f"{para_id}.{len(para['sub_paragraphs'])}"
                    st.session_state['document_graph'].add_node(f"Sub-paragraph {child_id}", label=sub_para, type='sub_paragraph')
                    st.session_state['document_graph'].add_edge(f"Paragraph {para_id}", f"Sub-paragraph {child_id}")
                    st.success("Sub-paragraph added!")
                else:
                    st.error("Selected paragraph not found.")
            else:
                st.error("Sub-paragraph cannot be empty.")

def add_comment():
    st.subheader("Add Comment")
    if not st.session_state['paragraphs']:
        st.warning("Add a main paragraph first.")
    else:
        para_options = [f"Paragraph {para['id']}" for para in st.session_state['paragraphs']]
        selected_para = st.selectbox("Select Paragraph to Add Comment", para_options)
        comment = st.text_area("Enter comment:", key="comment")
        if st.button("Add Comment"):
            if comment.strip() != "":
                para_id = int(selected_para.split(" ")[1])
                para = next((p for p in st.session_state['paragraphs'] if p['id'] == para_id), None)
                if para:
                    para['comments'].append(comment)
                    comment_id = f"{para_id}.{len(para['comments'])}"
                    st.session_state['document_graph'].add_node(f"Comment {comment_id}", label=comment, type='comment')
                    st.session_state['document_graph'].add_edge(f"Paragraph {para_id}", f"Comment {comment_id}")
                    st.success("Comment added!")
                else:
                    st.error("Selected paragraph not found.")
            else:
                st.error("Comment cannot be empty.")

def add_table():
    st.subheader("Add Table")
    with st.form("table_form"):
        num_rows = st.number_input("Number of rows", min_value=1, value=2, key="table_rows")
        num_cols = st.number_input("Number of columns", min_value=1, value=2, key="table_cols")
        table_data = []
        for row in range(int(num_rows)):
            row_data = []
            for col in range(int(num_cols)):
                cell = st.text_input(f"Row {row + 1}, Column {col + 1}", key=f"table_cell_{row}_{col}")
                row_data.append(cell)
            table_data.append(row_data)
        submitted = st.form_submit_button("Add Table")
        if submitted:
            df = pd.DataFrame(table_data, columns=[f"Column {i+1}" for i in range(int(num_cols))])
            table_id = len(st.session_state['tables']) + 1
            st.session_state['tables'].append({
                "id": table_id,
                "data": df
            })
            st.session_state['document_graph'].add_node(f"Table {table_id}", label=f"Table {table_id}", type='table')
            st.session_state['document_graph'].add_edge("Document", f"Table {table_id}")
            st.success("Table added!")

def add_image():
    st.subheader("Add Image")
    uploaded_image = st.file_uploader("Upload an image", type=["png", "jpg", "jpeg"])
    if uploaded_image is not None:
        image = Image.open(uploaded_image)
        st.image(image, caption="Uploaded Image", use_column_width=True)
        if st.button("Add Image"):
            image_id = len(st.session_state['images']) + 1
            st.session_state['images'].append({
                "id": image_id,
                "file": uploaded_image
            })
            st.session_state['document_graph'].add_node(f"Image {image_id}", label=f"Image {image_id}", type='image')
            st.session_state['document_graph'].add_edge("Document", f"Image {image_id}")
            st.success("Image added!")

def generate_document_graph():
    """
    Generates and returns a PyVis network graph based on the document structure.
    """
    net = Network(height='400px', width='100%', directed=True)
    net.from_nx(st.session_state['document_graph'])
    
    # Customize node appearance based on type
    for node in net.nodes:
        node_type = st.session_state['document_graph'].nodes[node['id']]['type']
        if node_type == 'paragraph':
            node['color'] = '#1f78b4'
        elif node_type == 'sub_paragraph':
            node['color'] = '#33a02c'
        elif node_type == 'comment':
            node['color'] = '#ff7f00'
        elif node_type == 'table':
            node[
