# PDFCONVERTOR

Web app for converting Office files (`.docx`, `.xlsx`, `.pptx`) to PDF, adding text annotations, and saving the result as PDF or DOCX.

## Frontend

```bash
npm install
npm run dev
```

Open locally:

```text
http://localhost:3000/PDFCONVERTOR/
```

Open from another device on the same network:

```text
http://<SERVER_IP>:3000/PDFCONVERTOR/
```

## Backend

LibreOffice must be installed on the machine that runs the backend.

```bash
cd backend
npm install
npm run dev
```

The backend listens on:

```text
http://0.0.0.0:4000
```

API:

```text
POST /api/convert/office-to-pdf
multipart/form-data field: file

POST /api/pdf/save
multipart/form-data fields: file, annotations

POST /api/pdf/export/docx
multipart/form-data fields: file, annotations, pageImages
```

## Editor Notes

This first version edits PDF files by adding a text annotation layer. It does not rewrite existing PDF text.

Supported annotation styling:

- Thai text
- font family
- font size
- color
- bold
- italic

DOCX export extracts the PDF text layer into editable Word text and appends added annotations as editable text. Complex PDF layouts, scanned files, multi-column documents, and advanced tables may not round-trip perfectly. Scanned PDFs need OCR before text can be extracted.

## Adding Custom Fonts (วิธีเพิ่มฟอนต์ใหม่)

หากต้องการเพิ่มฟอนต์ภาษาไทยใหม่ในระบบ ให้ทำตาม 3 ขั้นตอนนี้:

1. **วางไฟล์ฟอนต์ `.ttf`:**
   - คัดลอกไฟล์ฟอนต์นามสกุล `.ttf` (ทั้งตัวปกติและตัวหนา) ไปวางไว้ที่โฟลเดอร์:
     ```text
     backend/fonts/
     ```
     *(เช่น `MyFont-Regular.ttf` และ `MyFont-Bold.ttf`)*

2. **แก้ไข `backend/server.js`:**
   - เพิ่ม path ไฟล์ใน `thaiFontCandidates` และ `thaiBoldFontCandidates` (บรรทัด ~31-60):
     ```javascript
     const thaiFontCandidates = [
       path.join(FONTS_DIR, 'MyFont-Regular.ttf'),
       ...
     ]
     const thaiBoldFontCandidates = [
       path.join(FONTS_DIR, 'MyFont-Bold.ttf'),
       ...
     ]
     ```
   - เพิ่มเงื่อนไขในฟังก์ชัน `resolveFont` (บรรทัด ~345):
     ```javascript
     if (norm.includes('myfont')) {
       candidates = isBold
         ? [path.join(FONTS_DIR, 'MyFont-Bold.ttf'), path.join(FONTS_DIR, 'Sarabun-Bold.ttf')]
         : [path.join(FONTS_DIR, 'MyFont-Regular.ttf'), path.join(FONTS_DIR, 'Sarabun-Regular.ttf')]
     }
     ```

3. **เพิ่มตัวเลือกใน Frontend (`src/App.tsx`):**
   - เพิ่มชื่อฟอนต์ในอาร์เรย์ `fontOptions` (บรรทัด ~55) เพื่อให้แสดงใน Dropdown บนหน้าเว็บ:
     ```typescript
     const fontOptions = ['MyFont', 'Angsana New', 'Sarabun', 'Noto Sans Thai', 'Tahoma', 'Arial']
     ```

## Linux Deployment (การติดตั้งบน Linux)

1. ติดตั้ง Node.js, LibreOffice และฟอนต์ระบบ:
   ```bash
   sudo apt-get update
   sudo apt-get install -y nodejs npm libreoffice fonts-thai-tlwg fonts-noto-core fonts-noto-ui-core fontconfig
   sudo fc-cache -f -v
   ```

2. คัดลอกโฟลเดอร์ `backend/` (รวมถึง `backend/fonts/`) ไปที่เซิร์ฟเวอร์แล้วติดตั้ง dependencies:
   ```bash
   cd backend
   npm install
   ```

3. รัน backend ด้วย PM2:
   ```bash
   sudo npm install -g pm2
   pm2 start server.js --name "pdf-backend"
   pm2 startup
   pm2 save
   ```

## Web Server / IIS Notes

- หาก deploy Frontend บน **IIS (Windows Server)** ตรวจสอบให้แน่ใจว่าได้คัดลอกไฟล์ `dist/web.config` ไปด้วย เพื่อให้ IIS รองรับ MIME Type ของไฟล์ `.mjs` (PDF.js Worker)

