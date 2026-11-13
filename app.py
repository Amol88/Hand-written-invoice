import streamlit as st
import base64
import json
import io
from docx import Document
from docx.shared import Pt
from google import genai
from pydantic import BaseModel, Field

# 1. COMPACT SCHEMA
class Item(BaseModel):
    sno: str; desc: str; unit: str; qty: str; rate: str; amount: str
class Invoice(BaseModel):
    date: str; to_name: str; address: str; subject: str
    items: list[Item]
    taxable_amount: str; gst_amount: str; round_off: str; gross_amount: str

# 2. EXACT FORMATTING GENERATOR (Fixing Line Breaks)
def create_word_invoice(data):
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin, sec.right_margin = Pt(50), Pt(50)
    
    # Custom function to force exact line spacing and breaks
    def add_p(text, b=False, align=0, size=11):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(0) # Forces zero spacing between lines like original
        r = p.add_run(text)
        r.bold = b
        r.font.size = Pt(size)
        return p

    # Header perfectly structured
    add_p("Prakash Fire Services", b=True, align=1, size=16)
    add_p("Fire & Safety Engineers & Contractors", align=1)
    add_p("402, 4th floor, Shripad Vallabh Apartment, Nandivali Road, Near Ashok Vatika, Sunil Nagar, Dombivali (East), 421201.", align=1)
    add_p("Email : pfsfire2005@gmail.com Mob. : 91-7045078180 / 9167573647.", align=1)
    add_p("GST No. : 27BGTPS6345E2ZS", align=1)
    
    add_p("")
    pq = add_p("QUOTATION", b=True, align=1)
    pq.runs[0].underline = True
    add_p("")

    # Ref & Date separated
    tr = doc.add_table(rows=1, cols=2)
    tr.cell(0,0).text = "Ref. No.: PFS/Q/073/26-27"
    tr.cell(0,1).text = f"Date: {data.get('date', '')}"
    tr.cell(0,1).paragraphs[0].alignment = 2
    for c in tr.rows[0].cells: c.paragraphs[0].paragraph_format.space_after = Pt(0)

    # To Block (Fixing the squished text issue)
    add_p("\nTo,")
    add_p(data.get('to_name', ''), b=True)
    # Split address into multiple clean lines
    for line in data.get('address', '').replace(', ', ',\n').split('\n'):
        add_p(line)
    
    add_p("")
    psub = add_p("Subject: ")
    psub.runs[0].bold = True
    psub.add_run(data.get('subject', ''))
    add_p("")

    # Table perfectly aligned
    t = doc.add_table(rows=1, cols=6, style='Table Grid')
    for i, h in enumerate(['S.No', 'Description', 'Unit', 'Qty', 'Rate (₹)', 'Amount (₹)']):
        t.cell(0,i).text = h
        t.cell(0,i).paragraphs[0].runs[0].bold = True
        t.cell(0,i).paragraphs[0].paragraph_format.space_after = Pt(0)
        
    for it in data.get('items', []):
        row = t.add_row().cells
        row[0].text, row[1].text, row[2].text = it.get('sno',''), it.get('desc',''), it.get('unit','')
        row[3].text, row[4].text, row[5].text = it.get('qty',''), it.get('rate',''), it.get('amount','')
        for c in row: c.paragraphs[0].paragraph_format.space_after = Pt(0)

    # Totals in column 4 and 5
    for lbl, val in [('Taxable Amount', 'taxable_amount'), ('18% GST', 'gst_amount'), ('Round Off', 'round_off'), ('Gross Amount', 'gross_amount')]:
        row = t.add_row().cells
        row[4].text = lbl; row[4].paragraphs[0].runs[0].bold = True
        row[5].text = str(data.get(val, '')); row[5].paragraphs[0].runs[0].bold = True
        for c in row: c.paragraphs[0].paragraph_format.space_after = Pt(0)

    # T&C properly structured
    add_p("\nTerms and Conditions: ", b=True)
    add_p("50% Advance as per work order")
    add_p("30% against material purchase.")
    add_p("20% after completion of Work.")
    add_p("18% GST Separately applicable.")
    add_p("Extra work will be charges extra.")
    add_p("Validity of Estimates- 7 Days")
    
    add_p("\nThanks & Regards,")
    add_p("For Prakash Fire Services", b=True)
    add_p("\nProprietor")
    return doc

# 3. STREAMLIT APP
st.set_page_config(page_title="Fast Invoice Generator", layout="centered")
st.title("📄 Auto Invoice to Word")

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

img_file = st.file_uploader("Upload Handwritten Invoice", type=["jpg", "png", "jpeg"])
if img_file and st.button("🚀 Fast Convert"):
    with st.spinner("Extracting & Perfecting Format..."):
        try:
            res = client.interactions.create(
                model="gemini-3.5-flash",
                input=[
                    {"type": "text", "text": "Extract invoice data exactly as per schema. For empty fields, use blank string."},
                    {"type": "image", "data": base64.b64encode(img_file.getvalue()).decode("utf-8"), "mime_type": img_file.type}
                ],
                response_format={"type": "text", "mime_type": "application/json", "schema": Invoice.model_json_schema()}
            )
            
            doc = create_word_invoice(json.loads(res.output_text))
            out = io.BytesIO()
            doc.save(out)
            
            st.success("✅ Perfect Formatting Applied!")
            st.download_button("⬇️ Download Perfect Word File", data=out.getvalue(), file_name="Invoice_Perfectly_Formatted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Error: {e}")
