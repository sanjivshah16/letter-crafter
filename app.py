import streamlit as st
from docx import Document
from urllib.parse import unquote
from datetime import date
import io

st.title("📄 Format Your Recommendation Letter")

# Get query parameters
params = st.query_params

# Handle query parameters more safely
letter_text = ""
addressee = ""
salutation = ""
letter_date = date.today().strftime("%B %d, %Y")

if "text" in params:
    letter_text = unquote(params["text"])
if "addressee" in params:
    addressee = unquote(params["addressee"])
if "salutation" in params:
    salutation = unquote(params["salutation"])
if "date" in params:
    letter_date = unquote(params["date"])

# Display the parsed content for debugging
st.subheader("Parsed Content:")
st.write(f"**Date:** {letter_date}")
st.write(f"**Addressee:** {addressee if addressee else '(None provided)'}")
st.write(f"**Salutation:** {salutation}")
st.write(f"**Letter Text Length:** {len(letter_text)} characters")

# Show a preview of the letter text
if letter_text:
    with st.expander("Preview Letter Text"):
        st.write(letter_text)

# Let user optionally rename file
filename = st.text_input("Enter filename (without extension)", value="recommendation_letter")

# Upload template
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

# Updated condition - don't require addressee since it can be empty
if template_file and letter_text and salutation:
    try:
        template = Document(template_file)
        
        # Improved replace function
        def replace_text_in_doc(doc, replacements):
            # Replace in paragraphs
            for paragraph in doc.paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        # Handle paragraph-level replacement
                        inline = paragraph.runs
                        for run in inline:
                            if key in run.text:
                                run.text = run.text.replace(key, value)
            
            # Replace in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for key, value in replacements.items():
                                if key in paragraph.text:
                                    for run in paragraph.runs:
                                        if key in run.text:
                                            run.text = run.text.replace(key, value)
            
            return doc
        
        replacements = {
            "<<Date>>": letter_date,
            "<<Addressee>>": addressee if addressee else "",
            "<<Salutation>>": salutation,
            "<<Enter text here>>": letter_text
        }
        
        updated_doc = replace_text_in_doc(template, replacements)
        
        # Save DOCX
        docx_buffer = io.BytesIO()
        updated_doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        st.success("Document processed successfully!")
        
        st.download_button(
            label="📥 Download DOCX",
            data=docx_buffer.getvalue(),
            file_name=f"{filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        st.info("To convert the DOCX file to PDF, please use an external tool or service.")
        
    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
        
elif not template_file:
    st.info("Please upload a Word template to begin.")
elif not letter_text:
    st.info("No letter text found in URL parameters.")
elif not salutation:
    st.info("No salutation found in URL parameters.")
else:
    st.info("Awaiting letter text and a valid template to begin.")
