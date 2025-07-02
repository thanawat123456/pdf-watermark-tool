import streamlit as st
import pandas as pd
import io
from datetime import datetime
import zipfile
import math
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PyPDF2 import PdfReader, PdfWriter
import tempfile
import os

# ลงทะเบียนฟอนต์ THSarabunNew.ttf
try:
    pdfmetrics.registerFont(TTFont('THSarabunNew', 'THSarabunNew.ttf'))
    FONT_AVAILABLE = True
except:
    FONT_AVAILABLE = False

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="PDF Individual Watermark Tool",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

def read_file_names(uploaded_file, column_index=0, sheet_name=None):
    """อ่านรายชื่อจากไฟล์ CSV หรือ Excel"""
    try:
        file_extension = uploaded_file.name.lower().split('.')[-1]
        
        # อ่านไฟล์ตามประเภท
        if file_extension == 'csv':
            # ลองหลาย encoding สำหรับ CSV
            try:
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            except:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
                except:
                    try:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, encoding='cp874')
                    except:
                        try:
                            uploaded_file.seek(0)
                            df = pd.read_csv(uploaded_file, encoding='tis-620')
                        except:
                            uploaded_file.seek(0)
                            df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')
        
        elif file_extension == 'xlsx':
            # อ่าน Excel
            if sheet_name:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        else:
            return [], f"ไม่รองรับไฟล์ประเภท: {file_extension}"
        
        if df.empty or len(df.columns) == 0:
            return [], f"ไฟล์ {file_extension.upper()} ว่างเปล่า"
        
        if column_index >= len(df.columns):
            return [], f"ไม่พบคอลัมน์ที่ {column_index + 1}"
        
        # ดึงรายชื่อจากคอลัมน์ที่เลือก
        names_column = df.iloc[:, column_index]
        names = [str(name).strip() for name in names_column.dropna() if str(name).strip()]
        
        return names, None
    except Exception as e:
        return [], str(e)

def get_excel_sheets(uploaded_file):
    """ดึงรายชื่อ Sheet จากไฟล์ Excel"""
    try:
        if uploaded_file.name.lower().endswith('.xlsx'):
            excel_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
            return excel_file.sheet_names
        return []
    except:
        return []


def create_watermark_overlay(name, page_width, page_height):
    """สร้างลายน้ำสำหรับชื่อหนึ่งคน - แบบทแยงมุม 9 จุดตามตำแหน่งใหม่"""
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    # ตั้งค่าลายน้ำ
    if FONT_AVAILABLE:
        can.setFont("THSarabunNew", 45)  # ลดขนาดลงมาจาก 60
    else:
        can.setFont("Helvetica", 30)      # ลดขนาดลงมาจาก 40
    
    can.setFillColorRGB(0.6, 0.6, 0.6, alpha=0.20)  # ลดความเข้มเล็กน้อย
    
    # กำหนดตำแหน่ง 9 จุดตามรูปแบบใหม่ (ตำแหน่งเฉพาะ)
    positions = [
        # แถวบนซ้าย-บน-บนขวา
        (page_width * 0.15, page_height * 0.85),  # บนซ้าย
        (page_width * 0.5, page_height * 0.9),    # บนกลาง (สูงขึ้น)
        (page_width * 0.85, page_height * 0.85),  # บนขวา
        
        # แถวกลางซ้าย-กลาง-กลางขวา
        (page_width * 0.1, page_height * 0.55),   # กลางซ้าย (ซ้ายมาก)
        (page_width * 0.5, page_height * 0.5),    # กลางกลาง
        (page_width * 0.9, page_height * 0.55),   # กลางขวา (ขวามาก)
        
        # แถวล่างซ้าย-ล่าง-ล่างขวา
        (page_width * 0.15, page_height * 0.2),   # ล่างซ้าย
        (page_width * 0.5, page_height * 0.15),   # ล่างกลาง (ต่ำลง)
        (page_width * 0.85, page_height * 0.2),   # ล่างขวา
    ]
    
    # วางลายน้ำใน 9 ตำแหน่งตามรูปแบบใหม่
    for i, (x, y) in enumerate(positions):
        can.saveState()
        can.translate(x, y)
        
        # หมุนตามแนวทแยงมุม (+45 องศา สำหรับทุกตำแหน่ง)
        rotation_angle = 45
        can.rotate(rotation_angle)
        
        # คำนวณความกว้างข้อความ
        font_name = "THSarabunNew" if FONT_AVAILABLE else "Helvetica"
        font_size = 45 if FONT_AVAILABLE else 30
        text_width = can.stringWidth(name, font_name, font_size)
        
        # วาดข้อความ (จัดกึ่งกลาง)
        can.drawString(-text_width / 2, 0, name)
        can.restoreState()
    
    can.save()
    packet.seek(0)
    return packet

def add_watermark_to_pdf(pdf_file, name, original_filename):
    """เพิ่มลายน้ำให้กับ PDF ทุกหน้า"""
    try:
        # อ่าน PDF ต้นฉบับ
        pdf_reader = PdfReader(pdf_file)
        pdf_writer = PdfWriter()
        
        for page_num, page in enumerate(pdf_reader.pages):
            # ดึงขนาดหน้า
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)
            
            # สร้างลายน้ำสำหรับหน้านี้
            watermark_packet = create_watermark_overlay(name, page_width, page_height)
            watermark_reader = PdfReader(watermark_packet)
            watermark_page = watermark_reader.pages[0]
            
            # รวมลายน้ำกับหน้าเดิม
            page.merge_page(watermark_page)
            pdf_writer.add_page(page)
        
        # สร้างไฟล์ PDF ใหม่
        output_buffer = io.BytesIO()
        pdf_writer.write(output_buffer)
        output_buffer.seek(0)
        
        # สร้างชื่อไฟล์ใหม่
        base_name = original_filename.rsplit('.', 1)[0] if '.' in original_filename else original_filename
        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '_', '-')).strip()
        new_filename = f"{base_name}_{safe_name}.pdf"
        
        return output_buffer, new_filename, None
        
    except Exception as e:
        return None, None, str(e)

def process_pdf_with_names(pdf_file, names_list, original_filename):
    """ประมวลผล PDF กับรายชื่อทั้งหมด"""
    results = []
    
    for name in names_list:
        pdf_file.seek(0)  # Reset file pointer
        pdf_buffer, filename, error = add_watermark_to_pdf(pdf_file, name, original_filename)
        
        if pdf_buffer:
            results.append({
                'name': name,
                'filename': filename,
                'buffer': pdf_buffer,
                'error': None
            })
        else:
            results.append({
                'name': name,
                'filename': None,
                'buffer': None,
                'error': error
            })
    
    return results

def create_zip_file(pdf_results):
    """สร้างไฟล์ ZIP จาก PDF หลายไฟล์"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for result in pdf_results:
            if result['buffer'] and result['filename']:
                zip_file.writestr(result['filename'], result['buffer'].getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

def main():
    st.title("📄 PDF Individual Watermark Tool")
    st.markdown("### อัพโหลด CSV (รายชื่อ) และ PDF (ต้นฉบับ) เพื่อสร้าง PDF แยกไฟล์พร้อมลายน้ำรายบุคคล")
    
    with st.sidebar:
        st.header("⚙️ การตั้งค่า")
        
        st.subheader("📋 ตั้งค่าข้อมูล")
        column_option = st.selectbox(
            "เลือกคอลัมน์ที่มีรายชื่อ:",
            ["คอลัมน์ A (แรก)", "คอลัมน์ B (ที่สอง)", "คอลัมน์ C (ที่สาม)", "คอลัมน์ D (ที่สี่)", "คอลัมน์ E (ที่ห้า)"],
            help="เลือกคอลัมน์ใน CSV ที่มีรายชื่อ"
        )
        
        column_index = {
            "คอลัมน์ A (แรก)": 0, 
            "คอลัมน์ B (ที่สอง)": 1, 
            "คอลัมน์ C (ที่สาม)": 2, 
            "คอลัมน์ D (ที่สี่)": 3, 
            "คอลัมน์ E (ที่ห้า)": 4
        }[column_option]
        
        st.markdown("---")
        st.markdown("### 📋 วิธีใช้งาน")
        st.markdown("""
        1. อัพโหลดไฟล์ CSV หรือ Excel ที่มีรายชื่อ
        2. อัพโหลดไฟล์ PDF ต้นฉบับ
        3. เลือกคอลัมน์ที่มีรายชื่อ
        4. สำหรับ Excel: เลือก Sheet ที่มีข้อมูล
        5. กดปุ่มสร้างไฟล์
        6. ดาวน์โหลด ZIP ที่มี PDF แยกไฟล์ตามจำนวนชื่อ
        """)
        
        st.markdown("### 🎯 ผลลัพธ์")
        st.markdown("""
        - **1 ชื่อ = 1 PDF**
        - ลายน้ำจะปรากฏทุกหน้าใน PDF
        - ชื่อไฟล์: `ชื่อไฟล์เดิม_รายชื่อ.pdf`
        - ไฟล์ทั้งหมดจะถูกรวมใน ZIP
        """)
        
        if not FONT_AVAILABLE:
            st.warning("⚠️ ไม่พบฟอนต์ THSarabunNew - จะใช้ฟอนต์ Helvetica แทน")
        
        # ไฟล์ตัวอย่าง CSV
        sample_csv = """ชื่อ,ตำแหน่ง,แผนก
สมชาย ใจดี,ผู้จัดการ,การตลาด
สุดา รักงาน,โปรแกรมเมอร์,IT
วิชัย ขยัน,นักบัญชี,การเงิน
นิดา เก่งงาน,นักวิเคราะห์,วิจัยและพัฒนา
ธนา ฉลาด,หัวหน้าขาย,ขาย"""
        
        if st.button("📥 ดาวน์โหลดไฟล์ตัวอย่าง CSV"):
            st.download_button(
                label="💾 ดาวน์โหลด sample.csv",
                data=sample_csv,
                file_name="sample_names.csv",
                mime="text/csv"
            )
    
    # ส่วนหลัก
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📁 อัพโหลดไฟล์")
        
        # อัพโหลด CSV/Excel
        data_file = st.file_uploader(
            "📊 เลือกไฟล์ CSV หรือ Excel ที่มีรายชื่อ",
            type=['csv', 'xlsx'],
            help="ไฟล์ CSV หรือ Excel ที่มีรายชื่อในคอลัมน์ที่เลือก"
        )
        
        # อัพโหลด PDF
        pdf_file = st.file_uploader(
            "📄 เลือกไฟล์ PDF ต้นฉบับ",
            type=['pdf'],
            help="ไฟล์ PDF ที่ต้องการเพิ่มลายน้ำ"
        )
        
        # แสดงข้อมูลไฟล์ที่อัพโหลด
        selected_sheet = None
        if data_file:
            file_extension = data_file.name.lower().split('.')[-1]
            
            # ถ้าเป็น Excel ให้เลือก Sheet
            if file_extension == 'xlsx':
                sheets = get_excel_sheets(data_file)
                if sheets:
                    st.subheader("📑 เลือก Sheet")
                    selected_sheet = st.selectbox(
                        "เลือก Sheet ที่มีรายชื่อ:",
                        sheets,
                        help="เลือก Sheet ใน Excel ที่มีข้อมูลรายชื่อ"
                    )
            
            with st.expander(f"📊 ตัวอย่างข้อมูล {file_extension.upper()}"):
                names, error = read_file_names(data_file, column_index, selected_sheet)
                
                if names:
                    # แสดง DataFrame
                    data_file.seek(0)
                    try:
                        if file_extension == 'csv':
                            df = pd.read_csv(data_file, encoding='utf-8')
                        else:
                            if selected_sheet:
                                df = pd.read_excel(data_file, sheet_name=selected_sheet, engine='openpyxl')
                            else:
                                df = pd.read_excel(data_file, engine='openpyxl')
                        
                        st.dataframe(df.head(), use_container_width=True)
                        
                        col_info1, col_info2 = st.columns(2)
                        with col_info1:
                            st.info(f"📊 จำนวนแถว: {len(df)}")
                        with col_info2:
                            st.info(f"📋 จำนวนคอลัมน์: {len(df.columns)}")
                        
                        # แสดงชื่อคอลัมน์
                        st.write("**ชื่อคอลัมน์ที่พบ:**")
                        for i, col in enumerate(df.columns):
                            st.write(f"- คอลัมน์ {chr(65+i)}: `{col}`")
                        
                        if file_extension == 'xlsx' and selected_sheet:
                            st.info(f"📑 Sheet ที่เลือก: {selected_sheet}")
                        
                    except Exception as e:
                        st.warning(f"ไม่สามารถแสดงตัวอย่างได้: {str(e)}")
                    
                    # แสดงรายชื่อ
                    st.success(f"👥 พบรายชื่อ {len(names)} คน:")
                    
                    # แสดงรายชื่อใน columns
                    name_cols = st.columns(3)
                    for i, name in enumerate(names):
                        with name_cols[i % 3]:
                            st.write(f"• {name}")
                else:
                    st.error(f"❌ {error}")
        
        if pdf_file:
            with st.expander("📄 ข้อมูล PDF"):
                try:
                    pdf_reader = PdfReader(pdf_file)
                    num_pages = len(pdf_reader.pages)
                    
                    col_pdf1, col_pdf2 = st.columns(2)
                    with col_pdf1:
                        st.info(f"📄 จำนวนหน้า: {num_pages}")
                    with col_pdf2:
                        st.info(f"📁 ชื่อไฟล์: {pdf_file.name}")
                    
                    # แสดงข้อมูลหน้าแรก
                    first_page = pdf_reader.pages[0]
                    width = float(first_page.mediabox.width)
                    height = float(first_page.mediabox.height)
                    st.write(f"**ขนาดหน้า**: {width:.0f} x {height:.0f} points")
                    
                except Exception as e:
                    st.error(f"❌ ไม่สามารถอ่าน PDF ได้: {str(e)}")
        
        # ปุ่มประมวลผล
        if data_file and pdf_file:
            st.markdown("---")
            
            names, error = read_file_names(data_file, column_index, selected_sheet)
            
            if names and not error:
                st.info(f"🎯 จะสร้าง PDF จำนวน {len(names)} ไฟล์")
                
                if st.button("🚀 สร้างไฟล์ PDF พร้อมลายน้ำ", type="primary"):
                    with st.spinner("กำลังประมวลผล..."):
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # ประมวลผล PDF
                        results = process_pdf_with_names(pdf_file, names, pdf_file.name)
                        
                        # แสดงผลลัพธ์
                        success_count = sum(1 for r in results if r['buffer'])
                        error_count = len(results) - success_count
                        
                        progress_bar.progress(100)
                        status_text.success(f"✅ สำเร็จ {success_count} ไฟล์, ผิดพลาด {error_count} ไฟล์")
                        
                        if success_count > 0:
                            # สร้าง ZIP
                            zip_buffer = create_zip_file(results)
                            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                            zip_filename = f"pdf_watermarked_{current_time}.zip"
                            
                            st.download_button(
                                label=f"📦 ดาวน์โหลด ZIP ({success_count} ไฟล์)",
                                data=zip_buffer.getvalue(),
                                file_name=zip_filename,
                                mime="application/zip"
                            )
                            
                            # แสดงรายการไฟล์ที่สร้างสำเร็จ
                            with st.expander("📋 รายการไฟล์ที่สร้างแล้ว"):
                                for result in results:
                                    if result['buffer']:
                                        st.success(f"✅ {result['filename']} - {result['name']}")
                                    else:
                                        st.error(f"❌ {result['name']} - {result['error']}")
                        
                        if error_count > 0:
                            st.error(f"❌ มีไฟล์ที่ประมวลผลไม่สำเร็จ {error_count} ไฟล์")
            else:
                st.error("❌ กรุณาตรวจสอบไฟล์ข้อมูลและเลือกคอลัมน์ที่ถูกต้อง")
    
    with col2:
        st.subheader("📊 สถิติ")
        
        if data_file:
            names, _ = read_file_names(data_file, column_index, selected_sheet)
            file_type = data_file.name.split('.')[-1].upper()
            st.metric("รายชื่อทั้งหมด", len(names))
            st.metric(f"ไฟล์ {file_type}", "1 ไฟล์")
        
        if pdf_file:
            try:
                pdf_reader = PdfReader(pdf_file)
                st.metric("หน้า PDF", f"{len(pdf_reader.pages)} หน้า")
                st.metric("ไฟล์ PDF ต้นฉบับ", "1 ไฟล์")
            except:
                st.metric("ไฟล์ PDF ต้นฉบับ", "ไม่สามารถอ่านได้")
        
        if data_file and pdf_file:
            names, _ = read_file_names(data_file, column_index, selected_sheet)
            st.metric("ไฟล์ PDF ที่จะได้", f"{len(names)} ไฟล์")
        
        st.subheader("💡 ขั้นตอนการใช้งาน")
        st.markdown("""
        - อัพโหลดทั้ง CSV/Excel และ PDF
        - CSV/Excel ต้องมีรายชื่อในคอลัมน์ที่เลือก
        - สำหรับ Excel: เลือก Sheet ที่มีข้อมูล
        - PDF สามารถมีหลายหน้าได้
        - ลายน้ำจะปรากฏทุกหน้าของ PDF
        - ชื่อไฟล์จะเป็น: `ชื่อไฟล์เดิม_รายชื่อ.pdf`
        - ไฟล์ ZIP จะมี PDF แยกไฟล์ตามจำนวนชื่อ
        """)
        
        st.subheader("🎨 ตัวอย่างลายน้ำ")
        if data_file:
            names, _ = read_file_names(data_file, column_index, selected_sheet)
            if names:
                st.write("**ตัวอย่างชื่อที่จะเป็นลายน้ำ:**")
                for name in names[:5]:  # แสดง 5 ชื่อแรก
                    st.write(f"🏷️ {name}")
                if len(names) > 5:
                    st.write(f"และอีก {len(names)-5} คน...")

if __name__ == "__main__":
    main()