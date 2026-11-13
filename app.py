import streamlit as st
import base64
import json
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from pydantic import BaseModel, Field

# 1. COMPACT SCHEMA
class Item(BaseModel):
    sno: str; desc: str; unit: str; qty: str; rate: str; amount: str
class Invoice(BaseModel):
    date: str; to_name: str; address: str; subject: str
    items: list[Item]
    taxable_amount: str; gst_amount: str; round_off: str; gross_amount: str

# 2. EXACT FORMATTING GENERATOR (Fast & Short)
def create_word_invoice(data):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'; style.font.size = Pt(11)
    
    # Header
    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = head.add_run("Prakash Fire Services\n")
    r.bold = True; r.font.size = Pt(16)
    head.add_run("Fire & Safety Engineers & Contractors\n402, 4th floor, Shripad Vallabh Apartment, Nandivali Road, Near Ashok Vatika, Sunil Nagar, Dombivali (East), 421201.\nEmail : pfsfire2005@gmail.com Mob. : 91-7045078180 / 9167573647.\nGST No. : 27BGTPS6345E2ZS")
    
    doc.add_paragraph("QUOTATION").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].runs[0].bold = True
    doc.paragraphs[-1].runs[0].underline = True

    # Details
    doc.add_paragraph(f"Ref. No.: PFS/Q/073/26-27\t\t\t\t\t\tDate: {data.get('date', '')}")
    doc.add_paragraph(f"To,\n{data.get('to_name', '')}\n{data.get('address', '')}\n\nSubject: {data.get('subject', '')}\n")

    # Table with Items & Totals Perfectly Aligned
    t = doc.add_table(rows=1, cols=6, style='Table Grid')
    for i, h in enumerate(['S.No', 'Description', 'Unit', 'Qty', 'Rate (₹)', 'Amount (₹)']):
        t.cell(0,i).text = h; t.cell(0,i).paragraphs[0].runs[0].bold = True
        
    for it in data.get('items', []):
        r = t.add_row().cells
        r[0].text, r[1].text, r[2].text = it.get('sno',''), it.get('desc',''), it.get('unit','')
        r[3].text, r[4].text, r[5].text = it.get('qty',''), it.get('rate',''), it.get('amount','')

    # Totals in the same table (Columns 4 & 5)
    for label, val in [('Taxable Amount', 'taxable_amount'), ('18% GST', 'gst_amount'), ('Round Off', 'round_off'), ('Gross Amount', 'gross_amount')]:
        r = t.add_row().cells
        r[4].text = label; r[4].paragraphs[0].runs[0].bold = True
        r[5].text = str(data.get(val, '')); r[5].paragraphs[0].runs[0].bold = True

    # Footer
    doc.add_paragraph("\nTerms and Conditions:\n50% Advance as per work order\n30% against material purchase.\n20% after completion of Work.\n18% GST Separately applicable.\nExtra work will be charges extra.\nValidity of Estimates- 7 Days")
    doc.add_paragraph("\nThanks & Regards,\nFor Prakash Fire Services\n\n\nProprietor")
    return doc

# 3. STREAMLIT APP
st.set_page_config(page_title="Fast Invoice Generator", layout="centered")
st.title("📄 Auto Invoice to Word")

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

img_file = st.file_uploader("Upload Handwritten Invoice", type=["jpg", "png", "jpeg"])
if img_file and st.button("🚀 Fast Convert"):
    with st.spinner("Extracting & Formatting..."):
        try:
            res = client.interactions.create(
                model="gemini-1.5-flash", # Fastest model for quick conversion
                input=[
                    {"type": "text", "text": "Extract invoice data exactly as per schema."},
                    {"type": "image", "data": base64.b64encode(img_file.getvalue()).decode("utf-8"), "mime_type": img_file.type}
                ],
                response_format={"type": "text", "mime_type": "application/json", "schema": Invoice.model_json_schema()}
            )
            
            doc = create_word_invoice(json.loads(res.output_text))
            out = io.BytesIO()
            doc.save(out)
            
            st.success("✅ Ready!")
            st.download_button("⬇️ Download Perfect Word File", data=out.getvalue(), file_name="Invoice_Formatted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Error: {e}")
