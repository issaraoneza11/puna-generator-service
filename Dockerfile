FROM node:20.11.1-slim

# ป้องกัน interactive mode + ตั้ง timezone
ENV DEBIAN_FRONTEND=noninteractive
ENV TZ=Asia/Bangkok

# ติดตั้ง LibreOffice + ฟอนต์ไทย (TLWG + TH Sarabun) + timezone
RUN apt-get update && \
    apt-get install -y \
        libreoffice \
        libreoffice-writer \
        libreoffice-calc \
        libreoffice-impress \
        fonts-thai-tlwg \
        fonts-th-sarabun-new \
        tzdata \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# โฟลเดอร์โปรเจค
RUN mkdir -p /usr/src/app
WORKDIR /usr/src/app

# คัดลอกโปรเจคทั้งหมด
COPY . .

# ติดตั้ง dependencies
RUN npm install --legacy-peer-deps --include=dev --unsafe-perm

# ตั้งค่า environment
ENV START_PROJECT=development

# เริ่มโปรเจค
CMD ["npm", "start"]