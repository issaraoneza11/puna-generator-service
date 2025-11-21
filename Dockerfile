FROM node:20.11.1-alpine

RUN mkdir -p /usr/src/app
WORKDIR /usr/src/app

# timezone (Alpine)
RUN apk add --no-cache tzdata
ENV TZ=Asia/Bangkok

# คัดลอกซอร์ส
COPY . .

# ติดตั้ง deps รวม devDeps + เปิดสิทธิ์เฉพาะที่คำสั่ง (npm v10 ไม่มี config unsafe-perm)
RUN npm install --legacy-peer-deps --include=dev --unsafe-perm

# (ถ้าใช้ webpack/nodemon ใน npm start ต้องอยู่ใน devDeps แล้ว)
ENV START_PROJECT=development

CMD ["npm", "start"]