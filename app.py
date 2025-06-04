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

# Handle both URL parameters and pastebin ID
letter_text = ""
addressee = ""
salutation = ""
letter_date = date.today().strftime("%B %d, %Y")

# Check if we have a pastebin ID
if "paste_id" in params:
    paste_id = params["paste_id"]
    try:
        # Fetch the letter from pastebin
        pastebin_url = f"https://pastebin.com/raw/{paste_id}"
        response = requests.get(pastebin_url)
        if response.status_code == 200:
            # Parse the JSON data from pastebin
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

# Also get individual parameters if provided (these override pastebin data)
if "addressee" in params:
    addressee = unquote(params["addressee"])
if "salutation" in params:
    salutation = unquote(params["salutation"])
if "date" in params:
    letter_date = unquote(params["date"])
if "text" in params:
    letter_text = unquote(params["text"])

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
        cache_key = f"{template_file.name}_{hash(letter_text)}_{hash(addressee)}_{hash(salutation)}"
        
        if 'processed_doc' not in st.session_state or st.session_state.get('cache_key') != cache_key:
            
            # Read the uploaded template
            template = Document(template_file)
            
            # Debug: Show what placeholders we found
            placeholders_found = []
            for paragraph in template.paragraphs:
                if "<<" in paragraph.text and ">>" in paragraph.text:
                    placeholders_found.append(paragraph.text.strip())
            
            if placeholders_found:
                st.write("**Placeholders found in template:**")
                for placeholder in placeholders_found:
                    st.write(f"- {placeholder}")
            else:
                st.warning("⚠️ No placeholders found in template! Make sure your template contains <<Date>>, <<Addressee>>, <<Salutation>>, and <<Enter text here>>")
            
            # Format-preserving replace function
            def replace_text_preserve_formatting(doc, replacements):
                replacements_made = {}
                paragraphs_to_remove = []
                
                # Process paragraphs
                for i, paragraph in enumerate(doc.paragraphs):
                    original_text = paragraph.text
                    
                    # Check if this paragraph should be removed (empty addressee case)
                    if not addressee and "<<Addressee>>" in original_text:
                        paragraphs_to_remove.append(i)
                        continue
                    
                    # Check if this is a blank line after addressee that should be removed
                    if (not addressee and i > 0 and 
                        "<<Addressee>>" in doc.paragraphs[i-1].text and 
                        original_text.strip() == ""):
                        paragraphs_to_remove.append(i)
                        continue
                    
                    # Replace text while preserving formatting
                    for placeholder, replacement in replacements.items():
                        if placeholder in original_text:
                            # Handle multi-paragraph replacement for long text
                            if placeholder == "<<Enter text here>>" and "\n\n" in replacement:
                                # Split replacement text into paragraphs
                                paragraphs = replacement.split('\n\n')
                                
                                # Replace the first paragraph in the existing paragraph
                                first_paragraph = paragraphs[0]
                                replace_text_in_paragraph(paragraph, placeholder, first_paragraph)
                                replacements_made[placeholder] = True
                                
                                # Add additional paragraphs after this one
                                if len(paragraphs) > 1:
                                    # Get the paragraph's style for consistency
                                    style = paragraph.style
                                    
                                    # Insert new paragraphs after the current one
                                    parent = paragraph._element.getparent()
                                    current_p = paragraph._element
                                    
                                    for extra_text in paragraphs[1:]:
                                        # Create new paragraph with same style
                                        new_p = doc.add_paragraph(extra_text, style)
                                        # Move it to the correct position
                                        parent.insert(parent.index(current_p) + 1, new_p._element)
                                        current_p = new_p._element
                            else:
                                # Simple single replacement
                                replace_text_in_paragraph(paragraph, placeholder, replacement)
                                replacements_made[placeholder] = True
                            break
                
                # Remove paragraphs marked for removal (in reverse order)
                for idx in sorted(paragraphs_to_remove, reverse=True):
                    p = doc.paragraphs[idx]
                    p_element = p._element
                    p_element.getparent().remove(p_element)
                
                # Process tables
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                for placeholder, replacement in replacements.items():
                                    if placeholder in paragraph.text:
                                        replace_text_in_paragraph(paragraph, placeholder, replacement)
                                        replacements_made[placeholder] = True
                
                return doc, replacements_made
            
            def replace_text_in_paragraph(paragraph, placeholder, replacement):
                """Replace text in a paragraph while preserving formatting of the first run that contains the placeholder"""
                
                # Find the run that contains the placeholder
                target_run = None
                for run in paragraph.runs:
                    if placeholder in run.text:
                        target_run = run
                        break
                
                if target_run is None:
                    return
                
                # Store the formatting from the target run
                font = target_run.font
                font_name = font.name
                font_size = font.size
                font_bold = font.bold
                font_italic = font.italic
                font_underline = font.underline
                font_color = font.color.rgb if font.color.rgb else None
                
                # Clear all runs in the paragraph
                for run in paragraph.runs:
                    run.clear()
                
                # Add new run with the replacement text and preserved formatting
                new_run = paragraph.runs[0]
                new_run.text = paragraph.text.replace(placeholder, replacement)
                
                # Apply the preserved formatting
                new_run.font.name = font_name
                new_run.font.size = font_size
                new_run.font.bold = font_bold
                new_run.font.italic = font_italic
                new_run.font.underline = font_underline
                if font_color:
                    new_run.font.color.rgb = font_color
            
            # Define replacements
            replacements = {
                "<<Date>>": letter_date,
                "<<Addressee>>": addressee if addressee else "",
                "<<Salutation>>": salutation,
                "<<Enter text here>>": letter_text
            }
            
            # Apply replacements
            updated_doc, replacements_made = replace_text_preserve_formatting(template, replacements)
            
            # Show what replacements were made
            st.write("**Replacements made:**")
            for placeholder, replacement in replacements.items():
                if placeholder in replacements_made:
                    st.write(f"✅ {placeholder} → {replacement[:50]}{'...' if len(replacement) > 50 else ''}")
                else:
                    st.write(f"❌ {placeholder} (not found in template)")
            
            # Store in session state
            st.session_state.processed_doc = updated_doc
            st.session_state.cache_key = cache_key
        
        st.success("🎉 Document processed successfully!")
        
        # DOCX Download
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
        
        # Show the first few characters of letter text for debugging
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

# Add footer with instructions
st.markdown("---")
st.markdown("""
**📋 How to use this app:**
1. Download the example template above and customize it with your letterhead
2. Make sure your template includes the required placeholders: `<<Date>>`, `<<Addressee>>`, `<<Salutation>>`, `<<Enter text here>>`
3. Upload your customized template
4. The app will automatically fill in the content from the URL parameters or pastebin
5. If no addressee is provided, those lines will be automatically removed
6. Download your formatted letter as a DOCX file

**✨ Font matching:** The replacement text will automatically match the font, size, and formatting of your template placeholders!
""")
