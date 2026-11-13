import streamlit as st
import google.generativeai as genai
from docxtpl import DocxTemplate
from PIL import Image
import json
import io

st.set_page_config(page_title="Handwritten to Word Invoice", layout="centered")
st.title("📄 Handwritten Invoice to Word Document")

api_key = st.secrets.get("GEMINI_API_KEY") 
if not api_key:
    st.error("⚠️ GEMINI_API_KEY not found in Secrets.")
    st.stop()

genai.configure(api_key=api_key)

# -----------------------------------------------------------------
# AUTO-DETECT MODEL: Ye khud check karega ki aapke paas kaunsa model hai
# -----------------------------------------------------------------
def get_best_model():
    available_models = []
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                available_models.append(m.name)
    except Exception as e:
        st.error(f"API Check Error: {e}")
        return "gemini-1.5-pro"

    if 'models/gemini-1.5-flash' in available_models:
        return 'gemini-1.5-flash'
    elif 'models/gemini-1.5-pro' in available_models:
        return 'gemini-1.5-pro'
    elif 'models/gemini-1.5-flash-latest' in available_models:
        return 'gemini-1.5-flash-latest'
    elif 'models/gemini-pro-vision' in available_models:
        return 'gemini-pro-vision'
    elif len(available_models) > 0:
        return available_models[0].replace('models/', '') 
    else:
        return 'gemini-1.5-flash' 

best_model_name = get_best_model()
model = genai.GenerativeModel(best_model_name)
st.success(f"🤖 Connected to Available Model: {best_model_name}")
# -----------------------------------------------------------------

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
        json_text = response.text.replace('```json', '').replace('```', '').strip()
        return json.loads(json_text)
    except Exception as e:
        st.error(f"Error during AI extraction: {e}")
        return None

img_file = st.file_uploader("Upload Image (JPG/PNG)", type=["jpg", "jpeg", "png"])
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if img_file and template_file:
    image = Image.open(img_file)
    st.image(image, caption="Uploaded Invoice", use_container_width=True)
    
    if st.button("🔄 Convert to Word Invoice"):
        with st.spinner("Extracting data..."):
            extracted_data = extract_data_from_image(image)
            
        if extracted_data:
            st.success("✅ Data extracted!")
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
