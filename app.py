import streamlit as st
import base64
import json
import io
from PIL import Image
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from pydantic import BaseModel, Field
from typing import List

# ---------------------------------------------------------
# 1. STRUCTURED OUTPUT SCHEMA
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
# 2. WORD DOCUMENT GENERATOR (No Template Needed!)
# ---------------------------------------------------------
def create_word_invoice(data):
    doc = Document()
    
    # Header
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Prakash Fire Services\n")
    run.bold = True
    run.font.size = Pt(16)
    doc.add_paragraph("Fire & Safety Engineers & Contractors\n402, 4th floor, Shripad Vallabh Apartment, Nandivali Road, Near Ashok Vatika, Sunil Nagar, Dombivali (East), 421201.\nEmail : pfsfire2005@gmail.com Mob. : 91-7045078180 / 9167573647.\nGST No. : 27BGTPS6345E2ZS").alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Title
    p_title = doc.add_paragraph("\nQUOTATION")
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.runs[0].bold = True
    p_title.runs[0].underline = True

    # Details
    doc.add_paragraph(f"Ref. No.: PFS/Q/073/26-27\t\t\t\t\tDate: {data.get('date', '')}")
    doc.add_paragraph(f"To,\n{data.get('to_name', '')}\n{data.get('address', '')}\n\nSubject: {data.get('subject', '')}\n")

    # Table
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    headers = ['S.No', 'Description', 'Unit', 'Qty', 'Rate (₹)', 'Amount (₹)']
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].bold = True

    # Add Items
    for item in data.get('items', []):
        row_cells = table.add_row().cells
        row_cells[0].text = str(item.get('sno', ''))
        row_cells[1].text = str(item.get('desc', ''))
        row_cells[2].text = str(item.get('unit', ''))
        row_cells[3].text = str(item.get('qty', ''))
        row_cells[4].text = str(item.get('rate', ''))
        row_cells[5].text = str(item.get('amount', ''))

    doc.add_paragraph("\n")

    # Totals
    tot_table = doc.add_table(rows=4, cols=2)
    tot_table.rows[0].cells[0].text = 'Taxable Amount'
    tot_table.rows[0].cells[1].text = str(data.get('taxable_amount', ''))
    tot_table.rows[1].cells[0].text = '18% GST'
    tot_table.rows[1].cells[1].text = str(data.get('gst_amount', ''))
    tot_table.rows[2].cells[0].text = 'Round Off'
    tot_table.rows[2].cells[1].text = str(data.get('round_off', ''))
    tot_table.rows[3].cells[0].text = 'Gross Amount'
    tot_table.rows[3].cells[1].text = str(data.get('gross_amount', ''))

    # Terms and conditions
    doc.add_paragraph("\nTerms and Conditions:\n50% Advance as per work order\n30% against material purchase.\n20% after completion of Work.\n18% GST Separately applicable.\nExtra work will be charges extra.\nValidity of Estimates- 7 Days")
    doc.add_paragraph("\nThanks & Regards,\nFor Prakash Fire Services\n\n\nProprietor")
    
    return doc

# ---------------------------------------------------------
# 3. STREAMLIT APP UI
# ---------------------------------------------------------
st.set_page_config(page_title="Handwritten to Word Invoice", layout="centered")
st.title("📄 Handwritten Invoice to Word Document")
st.info("💡 Now fully automatic! Just upload the handwritten image and get the Word file directly.")

api_key = st.secrets.get("GEMINI_API_KEY") 
if not api_key:
    st.error("⚠️ GEMINI_API_KEY not found in Secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

def extract_data_from_image(image_file):
    image_bytes = image_file.getvalue()
    image_b64 = base64.b64encode(image_bytes).decode("utf-8")
    mime_type = image_file.type 
    
    try:
        interaction = client.interactions.create(
            model="gemini-3.5-flash", 
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

# ONLY ask for image now!
img_file = st.file_uploader("Upload Handwritten Image (JPG/PNG)", type=["jpg", "jpeg", "png"])

if img_file:
    st.image(img_file, caption="Uploaded Invoice", use_container_width=True)
    
    if st.button("🔄 Fast Convert to Word Invoice"):
        with st.spinner("Extracting data rapidly..."):
            extracted_data = extract_data_from_image(img_file)
            
        if extracted_data:
            st.success("✅ Data extracted successfully!")
            
            with st.expander("View AI Extracted Data"):
                st.json(extracted_data)
                
            with st.spinner("Generating Perfect Word Document..."):
                try:
                    # Naya tarika: Seedha doc generate karna bina template ke
                    doc = create_word_invoice(extracted_data)
                    output = io.BytesIO()
                    doc.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        label="⬇️ Download Converted Word Invoice",
                        data=output,
                        file_name="Converted_Invoice_Final.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Generation Error: {e}")
