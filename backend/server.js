import fontkit from '@pdf-lib/fontkit'
import cors from 'cors'
import { Document, Packer, PageBreak, Paragraph, TextRun } from 'docx'
import express from 'express'
import libre from 'libreoffice-convert'
import multer from 'multer'
import { existsSync, readFileSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { degrees, PDFDocument, rgb, StandardFonts } from 'pdf-lib'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const FONTS_DIR = path.join(__dirname, 'fonts')

const app = express()
const convertAsync = promisify(libre.convert)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
    fieldSize: 80 * 1024 * 1024,
  },
})

const PORT = Number(process.env.PORT || 4000)
const HOST = process.env.HOST || '0.0.0.0'
const allowedOfficeExtensions = new Set(['docx', 'xlsx', 'pptx'])
const thaiFontCandidates = [
  path.join(FONTS_DIR, 'AngsanaNew.ttf'),
  path.join(FONTS_DIR, 'AngsanaUPC.ttf'),
  path.join(FONTS_DIR, 'Sarabun-Regular.ttf'),
  path.join(FONTS_DIR, 'tahoma.ttf'),
  path.join(FONTS_DIR, 'segoeui.ttf'),
  'C:/Windows/Fonts/Sarabun-Regular.ttf',
  'C:/Windows/Fonts/THSarabunNew.ttf',
  'C:/Windows/Fonts/tahoma.ttf',
  'C:/Windows/Fonts/segoeui.ttf',
  '/usr/share/fonts/truetype/sarabun/Sarabun-Regular.ttf',
  '/usr/share/fonts/truetype/thai/Sarabun-Regular.ttf',
  '/usr/share/fonts/truetype/tlwg/Garuda.ttf',
  '/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf',
]
const thaiBoldFontCandidates = [
  path.join(FONTS_DIR, 'AngsanaNew-Bold.ttf'),
  path.join(FONTS_DIR, 'AngsanaUPC-Bold.ttf'),
  path.join(FONTS_DIR, 'Sarabun-Bold.ttf'),
  path.join(FONTS_DIR, 'tahomabd.ttf'),
  path.join(FONTS_DIR, 'segoeuib.ttf'),
  'C:/Windows/Fonts/Sarabun-Bold.ttf',
  'C:/Windows/Fonts/THSarabunNew Bold.ttf',
  'C:/Windows/Fonts/tahomabd.ttf',
  'C:/Windows/Fonts/segoeuib.ttf',
  '/usr/share/fonts/truetype/sarabun/Sarabun-Bold.ttf',
  '/usr/share/fonts/truetype/thai/Sarabun-Bold.ttf',
  '/usr/share/fonts/truetype/tlwg/Garuda-Bold.ttf',
  '/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf',
]

app.use(cors({ origin: '*' }))

app.get('/health', (_req, res) => {
  res.json({ ok: true, service: 'PDFCONVERTOR backend' })
})

app.post('/api/convert/office-to-pdf', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ Office' })
    return
  }

  const extension = req.file.originalname.split('.').pop()?.toLowerCase()

  if (!extension || !allowedOfficeExtensions.has(extension)) {
    res.status(400).json({
      error: 'รองรับเฉพาะไฟล์ .docx, .xlsx, .pptx เท่านั้น',
    })
    return
  }

  try {
    const pdfBuffer = await convertAsync(req.file.buffer, '.pdf', undefined)
    const outputName = req.file.originalname.replace(/\.[^.]+$/, '.pdf')

    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `inline; filename="${encodeURIComponent(outputName)}"`)
    res.send(pdfBuffer)
  } catch (error) {
    console.error('Office to PDF conversion failed:', error)
    res.status(500).json({
      error:
        'แปลงไฟล์ไม่สำเร็จ กรุณาตรวจสอบว่าเครื่อง backend ติดตั้ง LibreOffice และฟอนต์ไทยแล้ว',
    })
  }
})

app.post('/api/pdf/save', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' })
    return
  }

  try {
    const annotations = parseAnnotations(req.body.annotations)
    const masks = parseMasks(req.body.masks)
    const signatures = parseSignatures(req.body.signatures)
    const pdfDoc = await PDFDocument.load(req.file.buffer)
    pdfDoc.registerFontkit(fontkit)

    const fontCache = new Map()
    async function getFontForAnnotation(family, isBold) {
      const key = `${family || 'default'}_${Boolean(isBold)}`
      if (fontCache.has(key)) {
        return fontCache.get(key)
      }
      const font = await resolveFont(pdfDoc, family, isBold)
      fontCache.set(key, font)
      return font
    }

    // Draw masks first so text annotations are always visible on top.
    for (const mask of masks) {
      if (mask.page < 1 || mask.page > pdfDoc.getPageCount()) continue
      const page = pdfDoc.getPage(mask.page - 1)
      const { width, height } = page.getSize()
      page.drawRectangle({
        x: clamp(Number(mask.x) || 0, 0, 1) * width,
        y: height - (clamp(Number(mask.y) || 0, 0, 1) + clamp(Number(mask.height) || 0, 0, 1)) * height,
        width: clamp(Number(mask.width) || 0, 0, 1) * width,
        height: clamp(Number(mask.height) || 0, 0, 1) * height,
        color: hexToRgb(mask.color),
      })
    }

    // Draw text annotations
    for (const annotation of annotations) {
      if (annotation.page < 1 || annotation.page > pdfDoc.getPageCount()) continue
      const page = pdfDoc.getPage(annotation.page - 1)
      const { width, height } = page.getSize()
      const fontSize = clamp(Number(annotation.fontSize) || 24, 8, 96)
      const font = await getFontForAnnotation(annotation.fontFamily, annotation.bold)
      const rot = Number(annotation.rotation) || 0

      page.drawText(String(annotation.text || ''), {
        x: clamp(Number(annotation.x) || 0, 0, 1) * width,
        y: height - clamp(Number(annotation.y) || 0, 0, 1) * height - fontSize * 0.45,
        size: fontSize,
        font,
        color: hexToRgb(annotation.color),
        lineHeight: fontSize * 1.25,
        rotate: degrees(-rot),
      })
    }

    // Draw signatures
    for (const sig of signatures) {
      if (sig.page < 1 || sig.page > pdfDoc.getPageCount()) continue
      const page = pdfDoc.getPage(sig.page - 1)
      const { width, height } = page.getSize()
      const sigWidth = (Number(sig.width) || 0.15) * width
      const sigHeight = (Number(sig.height) || 0.1) * width
      const centerX = clamp(Number(sig.x) || 0, 0, 1) * width
      const centerY = height - clamp(Number(sig.y) || 0, 0, 1) * height
      const rot = Number(sig.rotation) || 0

      try {
        const base64Data = sig.dataUrl.replace(/^data:image\/\w+;base64,/, '')
        const imageBytes = Buffer.from(base64Data, 'base64')
        const img = sig.dataUrl.startsWith('data:image/jpeg') || sig.dataUrl.startsWith('data:image/jpg')
          ? await pdfDoc.embedJpg(imageBytes)
          : await pdfDoc.embedPng(imageBytes)

        page.drawImage(img, {
          x: centerX - sigWidth / 2,
          y: centerY - sigHeight / 2,
          width: sigWidth,
          height: sigHeight,
          rotate: degrees(-rot),
          opacity: clamp(Number(sig.opacity ?? 1), 0, 1),
        })
      } catch (imgError) {
        console.warn('Failed to embed signature image:', imgError)
      }
    }

    const pdfBytes = await pdfDoc.save()
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', 'attachment; filename="edited.pdf"')
    res.send(Buffer.from(pdfBytes))
  } catch (error) {
    console.error('PDF save failed:', error)
    res.status(500).json({
      error: 'บันทึก PDF ไม่สำเร็จ กรุณาตรวจสอบว่าไฟล์ PDF และ annotation ถูกต้อง',
    })
  }
})

app.post('/api/pdf/export/docx', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'กรุณาอัปโหลดไฟล์ PDF' })
    return
  }

  try {
    const annotations = parseAnnotations(req.body.annotations)
    const extractedPages = await extractPdfText(req.file.buffer, annotations)
    const children = []

    for (const page of extractedPages) {
      children.push(
        new Paragraph({
          spacing: { before: 80, after: 120 },
          children: [
            new TextRun({
              text: sanitizeXmlText(`หน้า ${page.pageNumber}`),
              bold: true,
              font: 'Sarabun',
            }),
          ],
        }),
      )

      if (!page.lines.length) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: sanitizeXmlText(
                  'ไม่พบ text layer ในหน้านี้ อาจเป็น PDF แบบสแกนหรือข้อความถูกแปลงเป็นรูปภาพ',
                ),
                italics: true,
                font: 'Sarabun',
              }),
            ],
          }),
        )
      }

      for (const line of page.lines) {
        children.push(
          new Paragraph({
            children: line.items.map(
              (item) =>
                new TextRun({
                  text: sanitizeXmlText(item.text),
                  bold: Boolean(item.bold),
                  italics: Boolean(item.italic),
                  color: normalizeHex(item.color),
                  size: clamp(Number(item.fontSize) || 12, 8, 96) * 2,
                  font: item.fontFamily || 'Sarabun',
                }),
            ),
          }),
        )
      }

      children.push(new Paragraph({ children: [new PageBreak()] }))
    }

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              font: 'Sarabun',
            },
          },
        },
      },
      sections: [
        {
          children,
        },
      ],
    })
    const docxBuffer = await Packer.toBuffer(doc)

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    res.setHeader('Content-Disposition', 'attachment; filename="edited.docx"')
    res.send(docxBuffer)
  } catch (error) {
    console.error('DOCX export failed:', error)
    res.status(500).json({
      error:
        'สร้าง DOCX ไม่สำเร็จ กรุณาลองใหม่ หากเป็นไฟล์สแกนต้องใช้ OCR ก่อนจึงจะดึงข้อความเป็น DOCX ได้',
    })
  }
})

app.listen(PORT, HOST, () => {
  console.log(`PDFCONVERTOR backend listening on http://${HOST}:${PORT}`)
})

async function extractPdfText(pdfBuffer, annotations) {
  const loadingTask = getDocument({
    data: new Uint8Array(pdfBuffer),
    disableFontFace: true,
    useSystemFonts: true,
  })
  const pdf = await loadingTask.promise
  const pages = []

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber)
    const viewport = page.getViewport({ scale: 1 })
    const content = await page.getTextContent()
    const items = content.items
      .filter((item) => 'str' in item && item.str.trim())
      .map((item) => ({
        text: item.str,
        x: item.transform[4],
        y: item.transform[5],
        fontSize: Math.max(10, Math.round(Math.abs(item.transform[3]) || 12)),
        fontFamily: 'Sarabun',
        color: '111827',
        bold: false,
        italic: false,
      }))

    for (const annotation of annotations.filter((item) => item.page === pageNumber)) {
      items.push({
        text: String(annotation.text || ''),
        x: clamp(Number(annotation.x) || 0, 0, 1) * viewport.width,
        y: (1 - clamp(Number(annotation.y) || 0, 0, 1)) * viewport.height,
        fontSize: clamp(Number(annotation.fontSize) || 12, 8, 96),
        fontFamily: annotation.fontFamily || 'Sarabun',
        color: normalizeHex(annotation.color),
        bold: Boolean(annotation.bold),
        italic: Boolean(annotation.italic),
      })
    }

    pages.push({
      pageNumber,
      lines: groupTextItemsIntoLines(items),
    })
  }

  await loadingTask.destroy()
  return pages
}

function groupTextItemsIntoLines(items) {
  const sortedItems = [...items].sort((a, b) => b.y - a.y || a.x - b.x)
  const lines = []

  for (const item of sortedItems) {
    const line = lines.find((candidate) => Math.abs(candidate.y - item.y) <= 5)

    if (line) {
      line.items.push(item)
      line.items.sort((a, b) => a.x - b.x)
      continue
    }

    lines.push({ y: item.y, items: [item] })
  }

  return lines.map((line) => ({
    ...line,
    items: joinLineItems(line.items),
  }))
}

function joinLineItems(items) {
  return items.map((item, index) => {
    const previous = items[index - 1]
    const needsSpace = previous && item.x - previous.x > Math.max(previous.text.length * previous.fontSize * 0.45, 8)

    return {
      ...item,
      text: `${needsSpace ? ' ' : ''}${item.text}`,
    }
  })
}

function parseAnnotations(value) {
  if (!value) {
    return []
  }

  const parsed = JSON.parse(value)

  if (!Array.isArray(parsed)) {
    throw new Error('Invalid annotations JSON')
  }

  return parsed.filter((annotation) => Number.isInteger(annotation.page) && annotation.page > 0)
}

function parseMasks(value) {
  if (!value) return []
  try {
    const parsed = JSON.parse(value)
    if (!Array.isArray(parsed)) return []
    return parsed.filter((mask) => Number.isInteger(mask.page) && mask.page > 0)
  } catch {
    return []
  }
}

function parseSignatures(value) {
  if (!value) {
    return []
  }
  try {
    const parsed = JSON.parse(value)
    if (!Array.isArray(parsed)) {
      return []
    }
    return parsed.filter(
      (sig) => Number.isInteger(sig.page) && sig.page > 0 && typeof sig.dataUrl === 'string',
    )
  } catch {
    return []
  }
}

async function resolveFont(pdfDoc, familyName, isBold) {
  const norm = String(familyName || '').toLowerCase().trim()
  const fontMap = [
    { names: ['angsa'], regular: ['AngsanaNew.ttf', 'AngsanaUPC.ttf'], bold: ['AngsanaNew-Bold.ttf', 'AngsanaUPC-Bold.ttf'] },
    { names: ['sarabun'], regular: ['Sarabun-Regular.ttf'], bold: ['Sarabun-Bold.ttf'] },
    { names: ['noto sans thai', 'noto'], regular: ['NotoSansThai-Regular.ttf'], bold: ['NotoSansThai-Bold.ttf'] },
    { names: ['tahoma'], regular: ['tahoma.ttf'], bold: ['tahomabd.ttf'] },
    { names: ['arial'], regular: ['arial.ttf'], bold: ['arialbd.ttf'] },
    { names: ['calibri'], regular: ['calibri.ttf'], bold: ['calibrib.ttf'] },
    { names: ['courier new'], regular: ['cour.ttf'], bold: ['courbd.ttf'] },
    { names: ['georgia'], regular: ['georgia.ttf'], bold: ['georgiab.ttf'] },
    { names: ['times new roman'], regular: ['times.ttf'], bold: ['timesbd.ttf'] },
    { names: ['verdana'], regular: ['verdana.ttf'], bold: ['verdanab.ttf'] },
    { names: ['segoe ui'], regular: ['segoeui.ttf'], bold: ['segoeuib.ttf'] },
    { names: ['cordia new'], regular: ['cordia.ttf'], bold: ['cordiab.ttf'] },
    { names: ['browallia new'], regular: ['browa.ttf'], bold: ['browab.ttf'] },
    { names: ['leelawadee ui'], regular: ['leelawad.ttf'], bold: ['leelawdb.ttf'] },
  ]
  const selected = fontMap.find((font) => font.names.some((name) => norm === name || norm.includes(name)))
  const candidates = selected
    ? [
        ...(isBold ? selected.bold : selected.regular).map((fileName) => path.join(FONTS_DIR, fileName)),
        ...(isBold ? thaiBoldFontCandidates : thaiFontCandidates),
      ]
    : isBold
      ? thaiBoldFontCandidates
      : thaiFontCandidates

  return embedBestFont(pdfDoc, candidates)
}

async function embedBestFont(pdfDoc, candidates) {
  const fontPath = candidates.find((candidate) => existsSync(candidate))

  if (!fontPath) {
    return pdfDoc.embedFont(StandardFonts.Helvetica)
  }

  return pdfDoc.embedFont(readFileSync(fontPath), { subset: true })
}

function hexToRgb(hex) {
  const normalized = normalizeHex(hex)
  const red = Number.parseInt(normalized.slice(0, 2), 16) / 255
  const green = Number.parseInt(normalized.slice(2, 4), 16) / 255
  const blue = Number.parseInt(normalized.slice(4, 6), 16) / 255

  return rgb(red, green, blue)
}

function normalizeHex(hex) {
  const normalized = String(hex || '#111827').replace('#', '')
  return /^[0-9a-fA-F]{6}$/.test(normalized) ? normalized : '111827'
}

function sanitizeXmlText(value) {
  return Array.from(String(value ?? ''))
    .filter((character) => {
      const codePoint = character.codePointAt(0) ?? 0
      return (
        codePoint === 0x9 ||
        codePoint === 0xa ||
        codePoint === 0xd ||
        (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
        (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
        (codePoint >= 0x10000 && codePoint <= 0x10ffff)
      )
    })
    .join('')
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max)
}
