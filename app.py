import streamlit as st
from docx import Document
from urllib.parse import unquote
from datetime import date
import io

st.title("📄 Format Your Recommendation Letter")

# Get query parameters
params = st.query_params
letter_text = unquote(params.get("text", ""))
addressee = unquote(params.get("addressee", ""))
salutation = unquote(params.get("salutation", ""))
letter_date = unquote(params.get("date", date.today().strftime("%B %d, %Y")))

# Let user optionally rename file
filename = st.text_input("Enter filename (without extension)", value="recommendation_letter")

# Upload template
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if template_file and letter_text and addressee and salutation:
    template = Document(template_file)

    # Replace placeholders
    def replace(doc, replacements):
        for p in doc.paragraphs:
            for key, val in replacements.items():
                if key in p.text:
                    for run in p.runs:
                        run.text = run.text.replace(key, val)
        return doc

    replacements = {
        "<<Date>>": letter_date,
        "<<Addressee>>": addressee,
        "<<Salutation>>": salutation,
        "<<Enter text here>>": letter_text
    }

    updated_doc = replace(template, replacements)

    # Save DOCX
    docx_buffer = io.BytesIO()
    updated_doc.save(docx_buffer)
    docx_buffer.seek(0)

    st.download_button(
        label="📥 Download DOCX",
        data=docx_buffer,
        file_name=f"{filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    st.info("To convert the DOCX file to PDF, please use an external tool or service.")
else:
    st.info("Awaiting letter text and a valid template to begin.")
