import streamlit as st
from docx import Document
from urllib.parse import unquote
from datetime import date
import io
import requests
import json

# Configure page to avoid resets
st.set_page_config(page_title="Letter Formatter", layout="wide")

st.title("📄 Format Your Recommendation Letter")

# Get query parameters
params = st.query_params

# Handle both URL parameters and letter ID from POST redirect
letter_text = ""
addressee = ""
salutation = ""
letter_date = date.today().strftime("%B %d, %Y")

# Check if we have a letter ID (from POST method)
if "letter_id" in params:
    letter_id = params["letter_id"]
    try:
        # Fetch the stored letter data
        # You'll need to implement this endpoint on your API server
        response = requests.get(f"https://your-api-domain.com/get-letter/{letter_id}")
        if response.status_code == 200:
            data = response.json()
            letter_text = data.get("text", "")
            addressee = data.get("addressee", "")
            salutation = data.get("salutation", "")
            letter_date = data.get("date", letter_date)
        else:
            st.error("Could not retrieve letter content. Please try the manual input below.")
    except Exception as e:
        st.error(f"Error retrieving letter: {str(e)}. Please use manual input below.")
else:
    # Fall back to original URL parameter method
    if "text" in params:
        letter_text = unquote(params["text"])
    if "addressee" in params:
        addressee = unquote(params["addressee"])
    if "salutation" in params:
        salutation = unquote(params["salutation"])
    if "date" in params:
        letter_date = unquote(params["date"])

# Create two columns for layout
col1, col2 = st.columns([2, 1])

with col1:
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

with col2:
    # Template instructions and download
    st.subheader("📋 Template Requirements")
    st.info("""
    **Important:** Your Word template must:
    - Be a .docx file (Word document)
    - Include these exact placeholders:
      - `<<Date>>` for the date  
      - `<<Addressee>>` for recipient info (optional)
      - `<<Salutation>>` for greeting
      - `<<Enter text here>>` for main content
      
    **Note:** If no addressee is provided, the addressee lines will be automatically removed.
    """)
    
    # Link to the provided template
    st.markdown("📥 **[Download Example Template](https://tinyurl.com/yc5au6un)**")
    st.caption("Use this template as a starting point for your letters")

# Add manual input option
st.markdown("---")
st.subheader("🔄 Manual Input Option")
st.write("If the letter data didn't load automatically, you can input it manually:")

manual_mode = st.checkbox("Use manual input mode")

if manual_mode:
    letter_text = st.text_area("Letter Text", value=letter_text, height=300)
    addressee = st.text_input("Addressee (optional)", value=addressee)
    salutation = st.text_input("Salutation", value=salutation or "Dear admissions committee members,")
    letter_date = st.text_input("Date", value=letter_date)

# Let user optionally rename file
filename = st.text_input("Enter filename (without extension)", value="recommendation_letter")

# Upload template
template_file = st.file_uploader(
    "Upload Your Word Template (.docx)", 
    type=["docx"],
    help="Upload a Word document with the required placeholders"
)

# Process document if all requirements are met
if template_file and letter_text and salutation:
    try:
        # Store processed document in session state to prevent reprocessing
        cache_key = f"{template_file.name}_{len(letter_text)}_{addressee}_{salutation}"
        
        if 'processed_doc' not in st.session_state or st.session_state.get('cache_key') != cache_key:
            template = Document(template_file)
            
            # Enhanced replace function that handles addressee removal
            def replace_text_in_doc(doc, replacements, remove_addressee_if_empty=False):
                paragraphs_to_remove = []
                
                # Replace in paragraphs
                for i, paragraph in enumerate(doc.paragraphs):
                    paragraph_text = paragraph.text.strip()
                    
                    # Handle addressee removal logic
                    if remove_addressee_if_empty and not addressee:
                        # If this paragraph contains <<Addressee>> placeholder, mark for removal
                        if "<<Addressee>>" in paragraph_text:
                            paragraphs_to_remove.append(i)
                            continue
                        # Also remove the blank line after addressee (if it exists)
                        elif i > 0 and "<<Addressee>>" in doc.paragraphs[i-1].text and paragraph_text == "":
                            paragraphs_to_remove.append(i)
                            continue
                    
                    # Normal text replacement
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            # Handle paragraph-level replacement
                            for run in paragraph.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, value)
                
                # Remove paragraphs marked for removal (in reverse order to maintain indices)
                for idx in sorted(paragraphs_to_remove, reverse=True):
                    # Remove paragraph by clearing its content
                    p = doc.paragraphs[idx]
                    p.clear()
                
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
            
            updated_doc = replace_text_in_doc(
                template, 
                replacements, 
                remove_addressee_if_empty=True
            )
            
            # Store in session state
            st.session_state.processed_doc = updated_doc
            st.session_state.cache_key = cache_key
        
        st.success("Document processed successfully!")
        
        # Create download buttons in columns
        download_col1, download_col2 = st.columns(2)
        
        with download_col1:
            # DOCX Download
            docx_buffer = io.BytesIO()
            st.session_state.processed_doc.save(docx_buffer)
            docx_buffer.seek(0)
            
            st.download_button(
                label="📥 Download as DOCX",
                data=docx_buffer.getvalue(),
                file_name=f"{filename}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="docx_download"
            )
        
        with download_col2:
            # PDF Download instruction
            st.info("📄 **PDF Download**")
            st.write("After downloading the DOCX file:")
            st.write("1. Open it in Microsoft Word")
            st.write("2. Go to File → Save As")
            st.write("3. Choose PDF format")
            st.write("")
            st.write("Or use an online converter like:")
            st.markdown("• [SmallPDF](https://smallpdf.com/word-to-pdf)")
            st.markdown("• [ILovePDF](https://www.ilovepdf.com/word_to_pdf)")
        
    except Exception as e:
        st.error(f"Error processing document: {str(e)}")
        st.write("Debug info:")
        st.write(f"Template file name: {template_file.name}")
        st.write(f"Letter text length: {len(letter_text)}")
        
elif not template_file:
    st.info("Please upload a Word template to begin.")
elif not letter_text:
    st.info("No letter text found. Please use manual input above or check your source link.")
elif not salutation:
    st.info("No salutation found. Please use manual input above.")
else:
    st.info("Awaiting letter text and a valid template to begin.")

# Add footer with instructions
st.markdown("---")
st.markdown("""
**How to use this app:**
1. Download the example template above and customize it with your letterhead
2. Make sure your template includes the required placeholders (see format above)
3. Upload your customized template
4. The app will automatically fill in the content from the URL parameters
5. If no addressee is provided, those lines will be automatically removed
6. Download your formatted letter as DOCX, then convert to PDF if needed
""")
