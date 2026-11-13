import streamlit as st
import base64, json, io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from pydantic import BaseModel

# 1. COMPACT SCHEMA
class Item(BaseModel): sno:str; desc:str; unit:str; qty:str; rate:str; amount:str
class Invoice(BaseModel):
    date:str; to_name:str; address:str; subject:str; items:list[Item]
    taxable_amount:str; gst_amount:str; round_off:str; gross_amount:str

# 2. EXACT COLORED GENERATOR (Ultra Short)
def generate_doc(data):
    doc = Document()
    doc.styles['Normal'].font.name = 'Times New Roman' # Professional Serif font
    
    def add_p(txt, b=False, sz=11, color=(0,0,0), align=0):
        p = doc.add_paragraph(); p.alignment = align; p.paragraph_format.space_after = Pt(0)
        r = p.add_run(txt); r.bold = b; r.font.size = Pt(sz); r.font.color.rgb = RGBColor(*color)
        return p

    # Colored Header (Orange Title, Dark Blue Subtext)
    add_p("Prakash Fire Services", True, 26, (235, 107, 24), 1) 
    blue = (0, 32, 96) 
    add_p("Fire & Safety Engineers & Contractors", True, 12, blue, 1)
    add_p("402, 4th floor, Shripad Vallabh Apartment, Nandivali Road, Near Ashok Vatika, Sunil Nagar, Dombivali (East), 421201.", True, 10, blue, 1)
    add_p("Email : pfsfire2005@gmail.com Mob. : 91-7045078180 / 9167573647.", True, 10, blue, 1)
    add_p("GST No. : 27BGTPS6345E2ZS", True, 10, blue, 1)
    add_p("════════════════════════════════════════════════════════════════════════", False, 12, (0,0,0), 1)
    
    add_p("QUOTATION", True, 14, align=1).runs[0].underline = True

    # Details
    doc.add_paragraph(f"Ref. No.: PFS/Q/073/26-27\t\t\t\t\t\tDate: {data.get('date','')}")
    add_p(f"To,\n{data.get('to_name','')}", True)
    add_p(data.get('address',''))
    add_p(f"\nSubject: {data.get('subject','')}", True)

    # Table
    t = doc.add_table(rows=1, cols=6, style='Table Grid')
    for i, h in enumerate(['S.No', 'Description', 'Unit', 'Qty', 'Rate (₹)', 'Amount (₹)']):
        t.cell(0,i).text = h; t.cell(0,i).paragraphs[0].runs[0].bold = True
        
    for it in data.get('items', []):
        r = t.add_row().cells
        r[0].text, r[1].text, r[2].text = it.get('sno',''), it.get('desc',''), it.get('unit','')
        r[3].text, r[4].text, r[5].text = it.get('qty',''), it.get('rate',''), it.get('amount','')

    for lbl, val in [('Taxable Amount', 'taxable_amount'), ('18% GST', 'gst_amount'), ('Round Off', 'round_off'), ('Gross Amount', 'gross_amount')]:
        r = t.add_row().cells
        r[4].text = lbl; r[4].paragraphs[0].runs[0].bold = True
        r[5].text = str(data.get(val, '')); r[5].paragraphs[0].runs[0].bold = True

    # Footer
    add_p("\nTerms and Conditions: ", True)
    add_p("50% Advance as per work order\n30% against material purchase.\n20% after completion of Work.\n18% GST Separately applicable.\nExtra work will be charges extra.\nValidity of Estimates- 7 Days")
    add_p("\nThanks & Regards,\nFor Prakash Fire Services", True)
    add_p("Proprietor")
    return doc

# 3. STREAMLIT APP
st.set_page_config(page_title="Fast Auto Invoice", layout="centered")
st.title("📄 Auto Invoice to Word")

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
img_file = st.file_uploader("Upload Handwritten Invoice", type=["jpg", "png", "jpeg"])

if img_file and st.button("🚀 Fast Convert"):
    with st.spinner("Extracting & Formatting Colors..."):
        try:
            res = client.interactions.create(
                model="gemini-3.5-flash",
                input=[
                    {"type": "text", "text": "Extract invoice exactly as per schema. Empty fields as blank."},
                    {"type": "image", "data": base64.b64encode(img_file.getvalue()).decode("utf-8"), "mime_type": img_file.type}
                ],
                response_format={"type": "text", "mime_type": "application/json", "schema": Invoice.model_json_schema()}
            )
            out = io.BytesIO()
            generate_doc(json.loads(res.output_text)).save(out)
            
            st.success("✅ Colored Formatting Applied!")
            st.download_button("⬇️ Download Perfect Word File", out.getvalue(), "Invoice_Colored.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Error: {e}")
