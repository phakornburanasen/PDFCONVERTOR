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

## Thai Fonts on Linux/Ubuntu

Install LibreOffice and common Thai fonts:

```bash
sudo apt-get update
sudo apt-get install -y libreoffice fonts-thai-tlwg fonts-noto-core fonts-noto-ui-core fontconfig
sudo fc-cache -f -v
```

For a specific `Sarabun` or `TH Sarabun PSK` font, copy the `.ttf` files to:

```text
/usr/share/fonts/truetype/thai/
```

Then refresh the font cache:

```bash
sudo fc-cache -f -v
```

For best Office-to-PDF layout accuracy, the font names used inside the Office file should exist on the backend machine.

For saving Thai text into edited PDF files, install at least one Thai-capable font such as Sarabun, Noto Sans Thai, Garuda, Tahoma, or TH Sarabun New on the backend machine.
