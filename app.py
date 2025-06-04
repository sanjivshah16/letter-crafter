import streamlit as st
from docx import Document
from urllib.parse import unquote
from datetime import date
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import requests
import json

# Configure page to avoid resets
st.set_page_config(page_title="Letter Formatter", layout="wide")

st.title("📄 Format Your Recommendation Letter")

# Get query parameters
params = st.query_params

# Handle both URL parameters and pastebin ID
letter_text = ""
addressee = ""
salutation = ""
letter_date = date.today().strftime("%B %d, %Y")

# Check if we have a pastebin ID
if "paste_id" in params:
    paste_id = params["paste_id"]
    try:
        pastebin_url = f"https://pastebin.com/raw/{paste_id}"
        response = requests.get(pastebin_url)
        if response.status_code == 200:
            data = json.loads(response.text)
            letter_text = data.get("text", "")
            addressee = data.get("addressee", "")
            salutation = data.get("salutation", "")
            letter_date = data.get("date", letter_date)
            st.success("✅ Letter data loaded successfully from pastebin!")
        else:
            st.error("Could not retrieve letter content from pastebin.")
    except Exception as e:
        st.error(f"Error retrieving letter: {str(e)}")

# Override with individual parameters
if "addressee" in params:
    addressee = unquote(params["addressee"])
if "salutation" in params:
    salutation = unquote(params["salutation"])
if "date" in params:
    letter_date = unquote(params["date"])
if "text" in params:
    letter_text = unquote(params["text"])

# Layout
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("Parsed Content:")
    st.write(f"**Date:** {letter_date}")
    st.write(f"**Addressee:** {addressee if addressee else '(None provided)'}")
    st.write(f"**Salutation:** {salutation}")
    st.write(f"**Letter Text Length:** {len(letter_text)} characters")
    if letter_text:
        with st.expander("Preview Letter Text"):
            st.write(letter_text)

with col2:
    st.subheader("📋 Template Requirements")
    st.info("""
    **Important:** Your Word template must:
    - Be a .docx file
    - Include these placeholders:
      - `<<Date>>`  
      - `<<Addressee>>`  
      - `<<Salutation>>`  
      - `<<Enter text here>>`
    """)
    st.markdown("📥 **[Download Example Template](https://tinyurl.com/yc5au6un)**")
    st.caption("Use this template as a starting point.")

# Filename input
filename = st.text_input("Enter filename (without extension)", value="recommendation_letter")

# Font & size selection
st.markdown("### ✏️ Choose Formatting")
font_name = st.selectbox("Font", ["Arial", "Times New Roman", "Calibri", "Aptos"], index=0)
font_size = st.selectbox("Font size", [9, 10, 10.5, 11, 11.5, 12], index=3)

# Check if font settings have changed after document was processed
font_changed = False
if 'processed_doc' in st.session_state:
    if ('last_font_name' in st.session_state and st.session_state.last_font_name != font_name) or \
       ('last_font_size' in st.session_state and st.session_state.last_font_size != font_size):
        font_changed = True

# Template upload
template_file = st.file_uploader(
    "Upload Your Word Template (.docx)", 
    type=["docx"],
    help="Upload a Word document with the required placeholders"
)

# Show regenerate button if font changed and document exists
if font_changed and template_file:
    if st.button("🔄 Regenerate Letter with Updated Font Formatting", type="primary"):
        # Clear the cached document to force regeneration
        if 'processed_doc' in st.session_state:
            del st.session_state.processed_doc
        if 'cache_key' in st.session_state:
            del st.session_state.cache_key
        st.rerun()

# Main processing
if template_file and letter_text and salutation:
    try:
        cache_key = f"{template_file.name}_{hash(letter_text)}_{hash(addressee)}_{hash(salutation)}_{font_name}_{font_size}"
        
        if 'processed_doc' not in st.session_state or st.session_state.get('cache_key') != cache_key:
            template = Document(template_file)
            
            placeholders_found = []
            for paragraph in template.paragraphs:
                if "<<" in paragraph.text and ">>" in paragraph.text:
                    placeholders_found.append(paragraph.text.strip())
            if placeholders_found:
                st.write("**Placeholders found in template:**")
                for placeholder in placeholders_found:
                    st.write(f"- {placeholder}")
            else:
                st.warning("⚠️ No placeholders found in template!")

            # Replacement logic
            def replace_text_in_document(doc, replacements):
                replacements_made = {}
                paragraphs_to_remove = []
                letter_content_paragraph_index = None

                for i, paragraph in enumerate(doc.paragraphs):
                    original_text = paragraph.text

                    if not addressee and "<<Addressee>>" in original_text:
                        paragraphs_to_remove.append(i)
                        continue
                    if not addressee and i > 0 and "<<Addressee>>" in doc.paragraphs[i-1].text and original_text.strip() == "":
                        paragraphs_to_remove.append(i)
                        continue

                    for placeholder, replacement in replacements.items():
                        if placeholder in original_text:
                            paragraph.clear()
                            run = paragraph.add_run(original_text.replace(placeholder, replacement))
                            run.font.name = font_name
                            run.font.size = Pt(font_size)
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                            replacements_made[placeholder] = True
                            
                            # Track where the letter content was inserted
                            if placeholder == "<<Enter text here>>":
                                letter_content_paragraph_index = i
                            break

                # Apply font formatting to all paragraphs after the letter content
                if letter_content_paragraph_index is not None:
                    for i in range(letter_content_paragraph_index + 1, len(doc.paragraphs)):
                        paragraph = doc.paragraphs[i]
                        # Only format paragraphs that have text content
                        if paragraph.text.strip():
                            for run in paragraph.runs:
                                run.font.name = font_name
                                run.font.size = Pt(font_size)
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

                for idx in sorted(paragraphs_to_remove, reverse=True):
                    p = doc.paragraphs[idx]
                    p_element = p._element
                    p_element.getparent().remove(p_element)

                return doc, replacements_made

            replacements = {
                "<<Date>>": letter_date,
                "<<Addressee>>": addressee if addressee else "",
                "<<Salutation>>": salutation,
                "<<Enter text here>>": letter_text
            }

            updated_doc, replacements_made = replace_text_in_document(template, replacements)

            st.write("**Replacements made:**")
            for placeholder, replacement in replacements.items():
                if placeholder in replacements_made:
                    st.write(f"✅ {placeholder} → {replacement[:50]}{'...' if len(replacement) > 50 else ''}")
                else:
                    st.write(f"❌ {placeholder} (not found in template)")

            st.session_state.processed_doc = updated_doc
            st.session_state.cache_key = cache_key
            # Store current font settings
            st.session_state.last_font_name = font_name
            st.session_state.last_font_size = font_size

        st.success("🎉 Document processed successfully!")
        docx_buffer = io.BytesIO()
        st.session_state.processed_doc.save(docx_buffer)
        docx_buffer.seek(0)

        st.download_button(
            label="📥 Download Formatted Letter (DOCX)",
            data=docx_buffer.getvalue(),
            file_name=f"{filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="docx_download"
        )

    except Exception as e:
        st.error(f"❌ Error processing document: {str(e)}")
        st.write("**Debug info:**")
        st.write(f"Template file name: {template_file.name}")
        st.write(f"Letter text length: {len(letter_text)}")
        if letter_text:
            st.write(f"Letter text preview: {letter_text[:100]}...")
elif not template_file:
    st.info("📤 Please upload a Word template to begin.")
elif not letter_text:
    st.info("📝 No letter text found. Please check your source link.")
elif not salutation:
    st.info("👋 No salutation found. Please check your source link.")
else:
    st.info("⏳ Awaiting letter text and a valid template to begin.")

st.markdown("---")
st.markdown("""
**📋 How to use this app:**
1. Download the example template above and customize it with your letterhead  
2. Make sure your template includes: `<<Date>>`, `<<Addressee>>`, `<<Salutation>>`, `<<Enter text here>>`  
3. Upload your customized template  
4. Choose your desired font and size  
5. The app fills in the content and lets you download the final DOCX
""")
