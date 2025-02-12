import os
import pythoncom
import zipfile
from datetime import datetime

import streamlit as st
from docx2pdf import convert
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path

# ========== STREAMLIT PAGE CONFIG & CUSTOM CSS ==========

st.set_page_config(
    page_title="Affidavit Generator",
    page_icon="📝",
    layout="wide"
)

def local_css(file_name):
    """Load local CSS file for additional styling."""
    with open(file_name) as css:
        st.markdown("<style>{}</style>".format(css.read()), unsafe_allow_html=True)

# Load your local CSS (make sure the path is correct)
local_css("styles/main.css")

# ========== HELPER FUNCTIONS ==========

def convert_with_com(input_path, output_path):
    """
    Safely convert DOCX to PDF using the COM interface (docx2pdf).
    This function ensures the COM interface is initialized and uninitialized.
    """
    pythoncom.CoInitialize()
    try:
        convert(input_path, output_path)
    finally:
        pythoncom.CoUninitialize()

def display_pdf_preview(pdf_path):
    """
    Convert a PDF to images for a quick visual preview inside Streamlit.
    """
    try:
        images = convert_from_path(pdf_path)
        for page_num, image in enumerate(images, start=1):
            st.image(image, caption=f'Page {page_num}', use_column_width=True)
    except Exception as e:
        st.error(f"Error displaying PDF preview: {e}")

def create_zip_file(files, zip_name):
    """
    Create a ZIP file containing the given files.
    """
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for file in files:
            zipf.write(file, os.path.basename(file))

# ========== PATHS TO TEMPLATES ==========

AFFIDAVIT_TEMPLATE_PATH = "Affidavit template to change.docx"
EXHIBIT_TEMPLATE_PATH   = "Exhibit Template.docx"
PART_2_TEMPLATE_PATH    = "Affidavit template to change - Part 2.docx"

# ========== STREAMLIT INTERFACE ==========

st.title("Affidavit Generator with Exhibits & Live Preview")
st.markdown(
    """
    **Fill in the details below** to automatically generate an affidavit (Word/PDF) 
    with any exhibits attached.
    """
)

# 1. GATHER USER INPUT
st.subheader("1. General Information")

name = st.text_input("Full Name").upper()
case_file = st.text_input("Case File Number")
party_name = st.text_input("Party Name")
lawyer_name = st.text_input("Lawyer's Name")
date = st.date_input("Date", value=datetime.today())
stat_declaration = st.selectbox("Statutory Declaration", ["Sworn", "Affirmed"])
address = st.text_area("Address")
email = st.text_input("Email")
phone = st.text_input("Phone")
party_role = st.selectbox(
    "Your Role in This Proceeding", 
    ["Witness", "Plaintiff", "Defendant"]
)

# 2. ENTER THE FACTS
st.subheader("2. Facts for the Affidavit")
facts = st.text_area("Enter the facts for the affidavit:")

# 3. UPLOAD EXHIBITS
st.subheader("3. Upload Exhibits (PDF)")
exhibit_files = st.file_uploader("Upload one or more PDF exhibits", type=["pdf"], accept_multiple_files=True)

# PREVIEW BUTTON (optional quick check of input)
if st.button("Preview Input"):
    st.write("**Preview of your inputs**:")
    st.write(f"**Name:** {name}")
    st.write(f"**Case File:** {case_file}")
    st.write(f"**Party Name:** {party_name}")
    st.write(f"**Lawyer Name:** {lawyer_name}")
    st.write(f"**Date:** {date}")
    st.write(f"**Statutory Declaration:** {stat_declaration}")
    st.write(f"**Address:** {address}")
    st.write(f"**Email:** {email}")
    st.write(f"**Phone:** {phone}")
    st.write(f"**Party Role:** {party_role}")
    st.write(f"**Facts:** {facts}")

# ========== AFFIDAVIT GENERATION (WORD) ==========

def generate_affidavit_with_exhibits(output_filename):
    """
    Render and assemble the affidavit template, exhibits, and part 2 into a single DOCX file.
    """
    # Prepare the context for affidavit rendering
    context = {
        'name': name,
        'case_file': case_file,
        'party_name': party_name,
        'date': date.strftime('%Y-%m-%d'),
        'stat_declaration': stat_declaration,
        'address': address,
        'email': email,
        'phone': phone,
        'lawyer_name': lawyer_name,
        'party_role': party_role.lower()
    }

    # Load & render the main affidavit template
    doc = DocxTemplate(AFFIDAVIT_TEMPLATE_PATH)
    doc.render(context)
    affidavit_file_path = "temp_affidavit.docx"
    doc.save(affidavit_file_path)

    # Load the rendered affidavit into a python-docx Document
    affidavit_doc = Document(affidavit_file_path)

    # EXHIBIT placeholders (optional logic if you want to show them in Word as well)
    exhibit_letter = 'A'
    for _ in exhibit_files:
        # If you want to do anything special in the Word doc about the exhibits
        # (like listing them), you can do that here.
        exhibit_letter = chr(ord(exhibit_letter) + 1)

    # Load and render Part 2 of the affidavit
    part_2_context = {
        'party_name': party_name,
        'party_role': party_role,
        'lawyer_name': lawyer_name,
        'date': date.strftime('%Y-%m-%d'),
        'name': name,
        'address': address,
        'email': email,
        'phone': phone,
        'stat_declaration': stat_declaration,
        'case_file': case_file
    }
    part_2_doc = DocxTemplate(PART_2_TEMPLATE_PATH)
    part_2_doc.render(part_2_context)
    temp_part_2_path = "temp_part_2.docx"
    part_2_doc.save(temp_part_2_path)

    # Append rendered Part 2 to the main affidavit
    rendered_part_2 = Document(temp_part_2_path)
    for element in rendered_part_2.element.body:
        affidavit_doc.element.body.append(element)

    # Apply 1-inch margins
    for section in affidavit_doc.sections:
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)

    # Save final DOCX
    affidavit_doc.save(output_filename)

# BUTTON: Generate Word Affidavit
if st.button("Generate Affidavit with Exhibits (Word)"):
    # Generate the Word document
    word_output_filename = f"{name}_affidavit_with_exhibits.docx"
    generate_affidavit_with_exhibits(word_output_filename)
    
    # Allow user to download it
    with open(word_output_filename, "rb") as f:
        st.download_button(
            label="Download Word Document",
            data=f,
            file_name=word_output_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ========== PDF GENERATION SECTION ==========

st.subheader("Convert a Word file to PDF and attach exhibits?")

uploaded_word_file = st.file_uploader("Upload a Word file (DOCX) to convert to PDF", type=["docx"])

def generate_pdf_with_exhibits(uploaded_word):
    """
    Convert the uploaded affidavit DOCX to PDF and then append exhibits and Part 2.
    """
    # 1. Save uploaded Word file temporarily
    temp_word_path = "uploaded_temp_affidavit.docx"
    with open(temp_word_path, "wb") as f:
        f.write(uploaded_word.getbuffer())

    # 2. Convert Word to PDF
    pdf_output_path = temp_word_path.replace(".docx", "_converted.pdf")
    convert_with_com(temp_word_path, pdf_output_path)

    # 3. Prepare a new PDF writer
    pdf_writer = PdfWriter()

    # 4. Convert the main affidavit template again (Part 1) to PDF
    affidavit_part_1_template = DocxTemplate(AFFIDAVIT_TEMPLATE_PATH)
    affidavit_part_1_context = {
        'name': name,
        'case_file': case_file,
        'party_name': party_name,
        'date': date.strftime('%Y-%m-%d'),
        'stat_declaration': stat_declaration,
        'address': address,
        'email': email,
        'phone': phone,
        'party_role': party_role.lower(),
        'lawyer_name': lawyer_name
    }
    affidavit_part_1_template.render(affidavit_part_1_context)
    affidavit_part_1_path = "temp_affidavit_part_1.docx"
    affidavit_part_1_template.save(affidavit_part_1_path)

    convert_with_com(affidavit_part_1_path, "temp_affidavit_part_1.pdf")
    part_1_pdf = PdfReader("temp_affidavit_part_1.pdf")
    for page in part_1_pdf.pages:
        pdf_writer.add_page(page)

    # 5. Insert exhibit templates + user-uploaded exhibits
    exhibit_letter = 'A'
    for exhibit_file in exhibit_files:
        # a) Exhibit Template PDF
        exhibit_template = DocxTemplate(EXHIBIT_TEMPLATE_PATH)
        exhibit_template.render({
            'letter': exhibit_letter,
            'party_name': party_name,
            'date': date.strftime('%Y-%m-%d')
        })

        temp_exhibit_docx = f"temp_exhibit_{exhibit_letter}.docx"
        exhibit_template.save(temp_exhibit_docx)

        temp_exhibit_pdf = temp_exhibit_docx.replace(".docx", ".pdf")
        convert_with_com(temp_exhibit_docx, temp_exhibit_pdf)

        # Add the Exhibit Template PDF pages
        template_pdf = PdfReader(temp_exhibit_pdf)
        for page in template_pdf.pages:
            pdf_writer.add_page(page)

        # b) The actual exhibit PDF (uploaded by user)
        #   NOTE: exhibit_file is a Streamlit UploadedFile, so we need to read it in memory
        #   or temporarily write it to disk.
        exhibit_bytes = exhibit_file.read()
        temp_exhibit_file_path = f"temp_exhibit_{exhibit_letter}.pdf"
        with open(temp_exhibit_file_path, "wb") as ef:
            ef.write(exhibit_bytes)
        
        exhibit_pdf = PdfReader(temp_exhibit_file_path)
        for page in exhibit_pdf.pages:
            pdf_writer.add_page(page)

        # Go to next letter
        exhibit_letter = chr(ord(exhibit_letter) + 1)

    # 6. Append Part 2 (if needed in PDF)
    part_2_template = DocxTemplate(PART_2_TEMPLATE_PATH)
    part_2_template.render({
        'party_name': party_name,
        'party_role': party_role,
        'lawyer_name': lawyer_name,
        'date': date.strftime('%Y-%m-%d'),
        'name': name,
        'address': address,
        'email': email,
        'phone': phone,
        'stat_declaration': stat_declaration,
        'case_file': case_file
    })
    temp_part_2_docx = "temp_part_2.docx"
    part_2_template.save(temp_part_2_docx)

    temp_part_2_pdf = temp_part_2_docx.replace(".docx", ".pdf")
    convert_with_com(temp_part_2_docx, temp_part_2_pdf)

    part_2_pdf = PdfReader(temp_part_2_pdf)
    for page in part_2_pdf.pages:
        pdf_writer.add_page(page)

    # 7. Write the final PDF
    final_pdf_output_path = "affidavit_with_exhibits.pdf"
    with open(final_pdf_output_path, "wb") as output_pdf:
        pdf_writer.write(output_pdf)

    # Optional: Display a preview
    st.subheader("PDF Preview")
    display_pdf_preview(final_pdf_output_path)

    # 8. Download Button
    with open(final_pdf_output_path, "rb") as f:
        st.download_button(
            label="Download Final PDF",
            data=f,
            file_name=os.path.basename(final_pdf_output_path),
            mime="application/pdf"
        )

if uploaded_word_file:
    if st.button("Generate PDF with Exhibits"):
        generate_pdf_with_exhibits(uploaded_word_file)

# ========== SIMPLE EXAMPLE USAGE (Optional) ==========

# Example usage if you want to generate a doc quickly based on context:
def generate_word_doc(context, exhibit_contexts, output_path):
    """
    Generates a Word document using the context (facts, personal info, etc.)
    and exhibits details.
    """
    doc = Document()
    doc.add_heading("Affidavit", 0)
    doc.add_paragraph(f"Name: {context['Name']}")
    doc.add_paragraph(f"Case File: {context['Case File']}")
    doc.add_paragraph(f"Party Name: {context['Party Name']}")
    doc.add_paragraph(f"Lawyer Name: {context['Lawyer Name']}")
    doc.add_paragraph(f"Date: {context['Date']}")
    doc.add_paragraph(f"Statutory Declaration: {context['Statutory Declaration']}")
    doc.add_paragraph(f"Address: {context['Address']}")
    doc.add_paragraph(f"Email: {context['Email']}")
    doc.add_paragraph(f"Phone: {context['Phone']}")
    doc.add_paragraph(f"Role: {context['Role']}")
    
    doc.add_heading("Facts", level=1)
    doc.add_paragraph(context['Facts'])
    
    if exhibit_contexts:
        doc.add_heading("Exhibits", level=1)
        for exhibit in exhibit_contexts:
            letter = exhibit['letter']
            notes = exhibit.get('notes', '')
            doc.add_paragraph(f"Exhibit {letter}: {notes}")
    doc.save(output_path)

def generate_pdf_from_word(word_path, pdf_path):
    """
    Converts a Word document to PDF.
    """
    convert(word_path, pdf_path)

st.subheader("Optional: Quick Affidavit Example")
if st.button("Generate Example Word/PDF"):
    example_context = {
        'Name': name or "[Name]",
        'Case File': case_file or "[Case File]",
        'Party Name': party_name or "[Party]",
        'Date': date.strftime('%Y-%m-%d') if date else "[Date]",
        'Facts': facts or "[Facts]"
    }
    word_output_path = f"{name}_example_affidavit.docx"
    pdf_output_path = f"{name}_example_affidavit.pdf"

    generate_word_doc(example_context, [], word_output_path)
    generate_pdf_from_word(word_output_path, pdf_output_path)

    # ZIP them together
    zip_name = f"{name}_documents.zip"
    create_zip_file([word_output_path, pdf_output_path], zip_name)

    with open(zip_name, "rb") as f:
        st.download_button(
            label="Download Example Affidavit (Word+PDF)",
            data=f,
            file_name=zip_name,
            mime="application/zip"
        )
