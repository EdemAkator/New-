import streamlit as st
import pandas as pd
import openai
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from io import BytesIO
from PIL import Image
import base64
import tempfile

openai.api_key = st.secrets["OPENAI_API_KEY"] if "OPENAI_API_KEY" in st.secrets else st.text_input("Enter OpenAI API Key", type="password")

st.title("📊 GPT-4 Excel Item Analyzer (Images in Column D)")

uploaded_file = st.file_uploader("Upload your Excel file with embedded images in Column D", type=["xlsx"])

if uploaded_file and openai.api_key:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    df = pd.read_excel(tmp_path)
    wb = load_workbook(tmp_path)
    ws = wb.active
    image_loader = SheetImageLoader(ws)

    df["Description"] = ""
    df["Brand"] = ""
    df["Original Price"] = ""
    df["Age"] = ""

    def image_to_base64(img):
        with BytesIO() as output:
            img.save(output, format="JPEG")
            return base64.b64encode(output.getvalue()).decode()

    def call_gpt(image_b64):
        prompt = (
            "You are an insurance expert. From the photo, extract:\n"
            "- Description\n- Brand (if visible)\n- Original price (USD)\n- Approximate age (years)\n\n"
            "Reply format:\nDescription: ...\nBrand: ...\nOriginal Price: ...\nAge: ..."
        )
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4-vision-preview",
                messages=[
                    {"role": "user", "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}
                    ]}
                ],
                max_tokens=500
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"Error: {str(e)}"

    if st.button("⚙️ Process File"):
        with st.spinner("Analyzing images and extracting data..."):
            for idx in df.index:
                row_num = idx + 2
                cell = f"D{row_num}"
                if image_loader.image_in(cell):
                    img = image_loader.get(cell)
                    img_b64 = image_to_base64(img)
                    gpt_result = call_gpt(img_b64)
                    for line in gpt_result.split("\n"):
                        if line.lower().startswith("description:"):
                            df.at[idx, "Description"] = line.partition(":")[2].strip()
                        elif line.lower().startswith("brand:"):
                            df.at[idx, "Brand"] = line.partition(":")[2].strip()
                        elif line.lower().startswith("original price:"):
                            df.at[idx, "Original Price"] = line.partition(":")[2].strip()
                        elif line.lower().startswith("age:"):
                            df.at[idx, "Age"] = line.partition(":")[2].strip()
                else:
                    df.at[idx, "Description"] = "No image"

        st.success("✅ Processing complete!")
        output = BytesIO()
        df.to_excel(output, index=False)
        st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="output_gpt_items.xlsx")
