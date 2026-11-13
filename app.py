import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate
from PIL import Image
import json
import io

# Streamlit page configuration
st.set_page_config(page_title="Handwritten to Word Invoice", layout="centered", page_icon="📄")
st.title("📄 Handwritten Invoice to Word Document")
st.markdown("Upload a handwritten invoice image and a Word template to auto-generate a typed invoice.")

# Fetching API Key from Streamlit Secrets
api_key = st.secrets.get("GEMINI_API_KEY") 
if not api_key:
    st.error("⚠️ GEMINI_API_KEY not found! Please add it to Streamlit Secrets.")
    st.stop()

# Configure Google Gemini
genai.configure(api_key=api_key)
model = genai.GenerativeModel('gemini-1.5-flash')

def extract_data_from_image(image):
    prompt = """
    Extract the invoice data from this handwritten image and return ONLY a valid JSON object. 
    Do not use markdown formatting like ```json or anything else. Just the raw JSON.
    The JSON structure MUST perfectly match this:
    {
        "date": "extract date",
        "to_name": "extract company/client name",
        "address": "extract full address",
        "subject": "extract subject line",
        "items": [
            {"sno": "1", "desc": "item name", "unit": "unit", "qty": "qty", "rate": "rate", "amount": "total amount"}
        ],
        "taxable_amount": "extract taxable amount",
        "gst_amount": "extract GST amount",
        "round_off": "extract round off if any, else 0",
        "gross_amount": "extract gross amount"
    }
    """
    try:
        response = model.generate_content([prompt, image])
        # Clean up output to ensure valid JSON
        json_text = response.text.replace('```json', '').replace('```', '').strip()
        return json.loads(json_text)
    except Exception as e:
        st.error(f"Error during AI extraction: {e}")
        return None

# UI for File Uploads
st.subheader("1. Upload Handwritten Invoice (Image)")
img_file = st.file_uploader("Upload Image (JPG/PNG)", type=["jpg", "jpeg", "png"])

st.subheader("2. Upload Word Template (.docx)")
template_file = st.file_uploader("Upload Word Template", type=["docx"])

# Processing logic
if img_file and template_file:
    image = Image.open(img_file)
    st.image(image, caption="Uploaded Handwritten Invoice", use_container_width=True)
    
    if st.button("🔄 Convert to Word Invoice"):
        with st.spinner("Analyzing handwriting using Gemini AI..."):
            extracted_data = extract_data_from_image(image)
            
        if extracted_data:
            st.success("✅ Data extracted successfully!")
            
            with st.expander("View Extracted Data"):
                st.json(extracted_data)
            
            with st.spinner("Generating Word Document..."):
                try:
                    # Load template and render data
                    doc = DocxTemplate(template_file)
                    doc.render(extracted_data)
                    
                    # Save to memory
                    output = io.BytesIO()
                    doc.save(output)
                    output.seek(0)
                    
                    # Download button
                    st.download_button(
                        label="⬇️ Download Converted Word Invoice",
                        data=output,
                        file_name="Converted_Invoice.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Error generating Word file: {e}")
