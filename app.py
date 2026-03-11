import streamlit as st
import pandas as pd
import io
import zipfile
from Code import process_call_plan_logic

# 1. ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="Call Plan Sales-team Generator", 
    page_icon="🚚",
    layout="centered"
)

# 2. ปรับแต่ง UI ด้วย CSS (เน้นโทนสีน้ำเงินเป็นหลัก)
st.markdown("""
    <style>
    /* ปรับฟอนต์และพื้นหลัง */
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700&display=swap');
    html, body, [class*="css"]  {
        font-family: 'Sarabun', sans-serif;
    }
    
    /* ปรับแต่งปุ่มเริ่มทำงาน */
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3.5em;
        background-color: #004AAD; /* สีน้ำเงินเข้ม */
        color: white;
        font-weight: bold;
        border: none;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #003070;
        border: none;
        color: white;
    }
    
    /* ปรับแต่งปุ่ม Download */
    .stDownloadButton>button {
        width: 100%;
        border-radius: 8px;
        background-color: #28a745;
        color: white;
    }
    
    /* ซ่อน Footer ของ Streamlit */
    footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# 3. ส่วน Sidebar (เครดิตและข้อมูลแอป)
with st.sidebar:
    # ใช้ Emoji แทนรูปภาพที่โหลดไม่ติด เพื่อความเสถียร
    st.markdown("<h1 style='text-align: center; font-size: 100px;'>🚛</h1>", unsafe_allow_html=True)
    st.title("Call Plan System")
    st.info("""
    **Version 1.1.0** ระบบแยกแผนการออกเยี่ยมอัตโนมัติ  
    ช่วยลดระยะเวลาการทำงาน RPA/Manual
    """)
    st.markdown("---")
    st.write("👨‍💻 **Developer:** PJS SPIRIT R4")
    st.write("🏢 **Department:** บริษัทนำรุ่งโรจน์")
    st.caption("© 2026 RPA Salesteam Support")

# 4. หน้าหลักของแอป
st.title("🚚 Call Plan Salesteam Generator")
st.markdown("##### ระบบจัดการไฟล์แผนงานรถขายและเยี่ยมร้านค้ารายบุคคล")

# ส่วนคำแนะนำการใช้งานแบบสวยงาม
with st.expander("💡 วิธีการใช้งาน (คลิกเพื่ออ่าน)"):
    st.write("""
    1. **เลือกไฟล์:** คลิกปุ่ม 'Browse files' เพื่อเลือกไฟล์ `R4_Visit plan.xlsx`
    2. **ระบุชื่องาน:** พิมพ์เดือนและปีในช่อง 'ชื่อ Job' (เช่น 10_2026)
    3. **ประมวลผล:** กดปุ่ม 'เริ่มสร้างไฟล์' ระบบจะทำการแยกไฟล์ตามรหัสพนักงานให้ทันที
    4. **ดาวน์โหลด:** เมื่อเสร็จสิ้น ให้กดปุ่ม 'ดาวน์โหลดไฟล์ ZIP' เพื่อรับไฟล์ทั้งหมด
    """)

st.write("") # เว้นวรรค

# ส่วนรับ Input (จัดวางแบบคอลัมน์)
col1, col2 = st.columns([3, 2])

with col1:
    uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ R4_Visit plan (Excel)", type=['xlsx'])

with col2:
    job_name = st.text_input("📝 ระบุชื่อ Job", placeholder="เช่น 10_2026")

st.markdown("---")

# 5. ส่วนประมวลผล Logic
if st.button("🚀 เริ่มสร้างไฟล์แผนงาน"):
    if uploaded_file and job_name:
        with st.spinner("⏳ กำลังคัดแยกข้อมูลพนักงาน... กรุณารอสักครู่"):
            try:
                # อ่านชีต Template
                df_source = pd.read_excel(uploaded_file, sheet_name='Template', engine='openpyxl')
                
                # เรียกใช้ Logic จากไฟล์ Code.py
                success, result = process_call_plan_logic(df_source, job_name)
                
                if success and result:
                    # สร้างไฟล์ ZIP ในหน่วยความจำ
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for filename, file_data in result.items():
                            zip_file.writestr(filename, file_data)
                    
                    st.balloons() # เอฟเฟกต์ลูกโป่งเมื่อสำเร็จ
                    st.success(f"✅ ประมวลผลสำเร็จ! แยกไฟล์พนักงานได้ทั้งหมด {len(result)} คน")
                    
                    # ปุ่มดาวน์โหลด
                    st.download_button(
                        label="📥 คลิกเพื่อดาวน์โหลดไฟล์ ZIP ทั้งหมด",
                        data=zip_buffer.getvalue(),
                        file_name=f"Call_Plan_{job_name}.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("❌ ไม่พบข้อมูลพนักงานในชีต 'Template' กรุณาตรวจสอบไฟล์ต้นฉบับ")
            
            except Exception as e:
                st.error(f"❌ ระบบขัดข้อง: {str(e)}")
    else:
        st.warning("⚠️ โปรดอัปโหลดไฟล์และระบุชื่อ Job ให้เรียบร้อยก่อนกดเริ่ม")

# ฟอนต์ด้านล่างสุด

st.markdown("<br><p style='text-align: center; color: gray; font-size: 12px;'>RPA Solutions for All-Channel Sales Team</p>", unsafe_allow_html=True)
