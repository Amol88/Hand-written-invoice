import streamlit as st
import base64
import json
import io
from PIL import Image
from docxtpl import DocxTemplate
from google import genai
from pydantic import BaseModel, Field
from typing import List

# ---------------------------------------------------------
# 1. STRUCTURED OUTPUT SCHEMA (For 100% Accurate Formatting)
# ---------------------------------------------------------
class InvoiceItem(BaseModel):
    sno: str = Field(description="Serial number")
    desc: str = Field(description="Description of the item")
    unit: str = Field(description="Unit of measurement")
    qty: str = Field(description="Quantity")
    rate: str = Field(description="Rate per unit")
    amount: str = Field(description="Total amount for the item")

class InvoiceData(BaseModel):
    date: str = Field(description="Invoice date")
    to_name: str = Field(description="Client or Company Name")
    address: str = Field(description="Full address")
    subject: str = Field(description="Subject of the quotation or invoice")
    items: List[InvoiceItem] = Field(description="List of all items in the table")
    taxable_amount: str = Field(description="Total taxable amount before taxes")
    gst_amount: str = Field(description="GST or tax amount")
    round_off: str = Field(description="Round off amount if any, otherwise 0")
    gross_amount: str = Field(description="Final gross amount")

# ---------------------------------------------------------
# 2. STREAMLIT APP UI
# ---------------------------------------------------------
st.set_page_config(page_title="Handwritten to Word Invoice", layout="centered")
st.title("📄 Handwritten Invoice to Word Document")

# Secret Key Check
api_key = st.secrets.get("GEMINI_API_KEY") 
if not api_key:
    st.error("⚠️ GEMINI_API_KEY not found in Secrets.")
    st.stop()

# Initialize Client
client = genai.Client(api_key=api_key)

def extract_data_from_image(image_file):
    image_bytes = image_file.getvalue()
    image_b64 = base64.b64encode(image_bytes).decode("utf-8")
    mime_type = image_file.type 
    
    try:
        # Changed to the most stable and fastest production model
        interaction = client.interactions.create(
            model="gemini-1.5-flash", 
            input=[
                {"type": "text", "text": "Extract all data from this handwritten invoice exactly as per the schema."},
                {"type": "image", "data": image_b64, "mime_type": mime_type}
            ],
            response_format={
                "type": "text",
                "mime_type": "application/json",
                "schema": InvoiceData.model_json_schema()
            }
        )
        return json.loads(interaction.output_text)
    
    except Exception as e:
        st.error(f"Error during AI extraction: {e}")
        return None

# ---------------------------------------------------------
# 3. FILE UPLOADS & PROCESSING
# ---------------------------------------------------------
img_file = st.file_uploader("Upload Image (JPG/PNG)", type=["jpg", "jpeg", "png"])
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if img_file and template_file:
    st.image(img_file, caption="Uploaded Invoice", use_container_width=True)
    
    if st.button("🔄 Fast Convert to Word Invoice"):
        with st.spinner("Extracting data rapidly..."):
            extracted_data = extract_data_from_image(img_file)
            
        if extracted_data:
            st.success("✅ Data extracted successfully!")
            
            with st.expander("View AI Extracted Data"):
                st.json(extracted_data)
                
            with st.spinner("Generating Word Document..."):
                try:
                    doc = DocxTemplate(template_file)
                    doc.render(extracted_data)
                    output = io.BytesIO()
                    doc.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        label="⬇️ Download Converted Word Invoice",
                        data=output,
                        file_name="Converted_Invoice.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Template Error: {e}")
