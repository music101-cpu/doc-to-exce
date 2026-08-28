import streamlit as st
import pandas as pd
import json
import io
import pypdfium2 as pdfium
from PIL import Image
from google import genai

# 1. ตั้งค่าหน้าตา Web App
st.set_page_config(page_title="Document to Excel Converter", layout="wide")
st.title("📄 ระบบแปลงเอกสาร ท.ร. 14/1 เป็น Excel")

# 2. ช่องใส่ API Key และ อัปโหลดไฟล์
api_key = st.sidebar.text_input("กรอก Gemini API Key", type="password")
uploaded_files = st.file_uploader(
    "อัปโหลดเอกสาร (PDF, JPG, PNG)", 
    type=["pdf", "jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

PROMPT = """
คุณคือระบบ OCR สำหรับสกัดข้อมูลจากเอกสารทางทะเบียนราษฎร ท.ร.14/1 ของประเทศไทย

คำสั่ง: 
อ่านข้อมูลบุคคลทุกคนที่ปรากฏในเอกสารนี้ แล้วสร้างเป็น JSON Array ของ Object
โดยสกัดข้อความจริงที่ปรากฏในเอกสารตามหัวข้อต่อไปนี้ (อย่าละเว้นข้อมูล):

- id_card: เลขประจำตัวประชาชน (13 หลัก)
- house_code: เลขรหัสประจำบ้าน (11 หลัก)
- title: คำนำหน้านาม (เช่น นาย / นาง / นางสาว / เด็กชาย / เด็กหญิง)
- first_name: ชื่อตัว
- last_name: ชื่อสกุล
- gender: เพศ (ชาย/หญิง)
- dob: วันเดือนปีเกิด (ตามที่ระบุในเอกสาร)
- status: สถานภาพในบ้าน (เช่น เจ้าบ้าน / ผู้อาศัย)
- mother_name: ชื่อ-นามสกุลมารดา
- father_name: ชื่อ-นามสกุลบิดา
- address: ที่อยู่ตามทะเบียนบ้าน
- registrar_office: สำนักทะเบียน
- move_in_date: วันที่ย้ายเข้า
- note: รายการบันทึก/หมายเหตุเพิ่มเติม (ถ้ามี)

**กฎสำคัญ:**
1. หากมีหลายคนในเอกสาร ให้สร้าง Object ของทุกคนแยกกันเป็นรายการใน JSON Array
2. อ่านข้อความภาษาไทยให้ถูกต้องแม่นยำที่สุด
3. ให้ตอบกลับเป็นรูปแบบข้อความ JSON เพียวๆ เท่านั้น ห้ามใส่ข้อความเกริ่นนำหรือ Markdown```json ใดๆ ทั้งสิ้น
"""

if st.button("เริ่มแปลงข้อมูล"):
    if not api_key:
        st.error("กรุณากรอก Gemini API Key ในแถบด้านซ้ายก่อนครับ")
    elif not uploaded_files:
        st.warning("กรุณาอัปโหลดไฟล์เอกสารอย่างน้อย 1 ไฟล์")
    else:
        client = genai.Client(api_key=api_key)
        all_results = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()

        total_files = len(uploaded_files)
        for i, file in enumerate(uploaded_files):
            status_text.text(f"กำลังประมวลผลไฟล์ที่ {i+1}/{total_files}: {file.name}")
            
            images_to_process = []
            
            # แปลง PDF เป็นรูปภาพคมชัดสูง (scale=3)
            if file.name.lower().endswith(".pdf"):
                pdf = pdfium.PdfDocument(file.read())
                for page in pdf:
                    image = page.render(scale=3).to_pil()
                    images_to_process.append(image)
            else:
                images_to_process.append(Image.open(file))

            # ประมวลผลภาพทีละหน้า
            for img in images_to_process:
                try:
                    # เปลี่ยนโมเดลเป็น gemini-3.6-flash ตามที่ API แจ้งเตือน
                    response = client.models.generate_content(
                        model='gemini-3.6-flash',
                        contents=[img, PROMPT]
                    )
                    
                    cleaned_json = response.text.strip()
                    if cleaned_json.startswith("```json"):
                        cleaned_json = cleaned_json.removeprefix("```json").removesuffix("```").strip()
                    elif cleaned_json.startswith("```"):
                        cleaned_json = cleaned_json.removeprefix("```").removesuffix("```").strip()
                        
                    data = json.loads(cleaned_json)
                    
                    if isinstance(data, list):
                        all_results.extend(data)
                    elif isinstance(data, dict):
                        all_results.append(data)
                        
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดกับไฟล์ {file.name}: {e}")
            
            progress_bar.progress((i + 1) / total_files)

        status_text.success("ประมวลผลเสร็จสิ้น!")

        # 3. แสดงตารางและดาวน์โหลด Excel
        if all_results:
            df = pd.DataFrame(all_results)
            st.subheader("ตัวอย่างข้อมูลที่สกัดได้")
            st.dataframe(df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='รายการทะเบียนราษฎร')
            output.seek(0)

            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
                data=output,
                file_name="รายการทะเบียนราษฎร.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("ไม่พบข้อมูลที่สกัดได้จากเอกสาร")