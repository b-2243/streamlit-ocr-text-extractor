import streamlit as st
import pytesseract
from PIL import Image
import io
import os
from pdf2image import convert_from_bytes
from docx import Document

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
poppler_path = r'C:\Users\BHAUMIK\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin'    

st.title("📄 OCR Text Extractor")
st.write("Upload an **image**, **PDF**, or **Word (.docx)** file to extract text.")

uploaded_file = st.file_uploader("Upload file", type=["png", "jpg", "jpeg", "pdf", "docx"])

if uploaded_file:
    file_type = uploaded_file.type

    extracted_text = ""

    # Image (PNG, JPG)
    if "image" in file_type:
        image = Image.open(uploaded_file)
        st.image(image, caption="Uploaded Image", use_column_width=True)
        extracted_text = pytesseract.image_to_string(image)

    # PDF
    elif file_type == "application/pdf":
        try:
            images = convert_from_bytes(uploaded_file.read(), dpi=300, poppler_path=poppler_path)
            for i, img in enumerate(images):
                st.image(img, caption=f"Page {i+1}", use_column_width=True)
                extracted_text += f"\n--- Page {i+1} ---\n"
                extracted_text += pytesseract.image_to_string(img)
        except Exception as e:
            st.error(f"PDF processing failed: {e}")

    # Word document (.docx)
    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc = Document(uploaded_file)
        for para in doc.paragraphs:
            extracted_text += para.text + "\n"

        for rel in doc.part._rels:
            rel = doc.part._rels[rel]
            if "image" in rel.target_ref:
                img_data = rel.target_part.blob
                image = Image.open(io.BytesIO(img_data))
                st.image(image, caption="Embedded Image", use_column_width=True)
                extracted_text += "\n" + pytesseract.image_to_string(image)

    st.subheader("📝 Extracted Text:")
    st.text_area("Text Output", extracted_text, height=300)

    st.download_button("📥 Download Text", extracted_text, file_name="extracted_text.txt")
