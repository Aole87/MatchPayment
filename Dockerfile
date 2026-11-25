FROM python:3.10-slim

WORKDIR /app

# คัดลอกและติดตั้ง requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# คัดลอกไฟล์โปรเจกต์ทั้งหมด
COPY . .

# กำหนด Port เริ่มต้นเป็น 8080 (Cloud Run จะ inject ค่า PORT เข้ามาอัตโนมัติ)
ENV PORT=8080

# รัน gunicorn สำหรับรองรับ Webhook
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app:app
