# pdf-watermark-tool# PDF Individual Watermark Tool

เครื่องมือสำหรับสร้างลายน้ำรายบุคคลในไฟล์ PDF

## คุณสมบัติ
- อัพโหลด CSV หรือ Excel ที่มีรายชื่อ
- อัพโหลด PDF ต้นฉบับ
- สร้าง PDF แยกไฟล์ตามจำนวนรายชื่อ (1 ชื่อ = 1 PDF)
- ลายน้ำปรากฏทุกหน้าใน PDF
- ดาวน์โหลดเป็นไฟล์ ZIP

## วิธีใช้งาน
1. อัพโหลดไฟล์ CSV/Excel ที่มีรายชื่อ
2. อัพโหลดไฟล์ PDF ต้นฉบับ
3. เลือกคอลัมน์ที่มีรายชื่อ
4. กดสร้างไฟล์
5. ดาวน์โหลด ZIP ที่มี PDF แยกไฟล์

## Demo
[ลิงก์ของแอพ](https://your-app-url.streamlit.app)

## การติดตั้งในเครื่อง
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py