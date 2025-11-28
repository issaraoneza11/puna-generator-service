FROM node:20.11.1-slim

# ไม่ถาม interactive + ตั้ง timezone
ENV DEBIAN_FRONTEND=noninteractive
ENV TZ=Asia/Bangkok

# ติดตั้ง LibreOffice + ฟอนต์ไทยพื้นฐาน + fontconfig (ไว้ใช้ fc-cache)
RUN apt-get update && \
    apt-get install -y \
        libreoffice \
        libreoffice-writer \
        libreoffice-calc \
        libreoffice-impress \
        fonts-thai-tlwg \
        fontconfig \
        tzdata \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# โฟลเดอร์โปรเจกต์
RUN mkdir -p /usr/src/app
WORKDIR /usr/src/app

# คัดลอกโปรเจกต์ทั้งหมดเข้า image
COPY . .

# ✅ ติดตั้งฟอนต์ TH Sarabun จากไฟล์ในโปรเจกต์
# font/THSarabun.ttf -> ตำแหน่งฟอนต์บนระบบ
RUN mkdir -p /usr/share/fonts/truetype/thai && \
    cp app/font/THSarabun.ttf /usr/share/fonts/truetype/thai/THSarabun.ttf && \
    fc-cache -f -v

# ติดตั้ง dependencies
RUN npm install --legacy-peer-deps --include=dev --unsafe-perm

# env อื่น ๆ
ENV START_PROJECT=development

# เริ่มโปรเจกต์
CMD ["npm", "start"]