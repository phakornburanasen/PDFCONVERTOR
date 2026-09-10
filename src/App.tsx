import {

  useEffect,
  useRef,
  useState,
  type CSSProperties,
  type FormEvent,
  type MouseEvent,
  type PointerEvent,
  type ChangeEvent,
} from 'react'
import * as pdfjsLib from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import './App.css'

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl

type Tool = 'move' | 'text' | 'mask' | 'delete'

type Annotation = {
  id: string
  page: number
  x: number
  y: number
  text: string
  fontFamily: string
  fontSize: number
  color: string
  bold: boolean
  italic: boolean
  rotation: number
}

type Mask = {
  id: string
  page: number
  x: number
  y: number
  width: number
  height: number
  color: string
}

type GuideLine = {
  id: string
  pageNumber: number
  type: 'vertical' | 'horizontal'
  position: number
}

type Signature = {
  id: string
  page: number
  x: number
  y: number
  width: number
  height: number
  rotation: number
  opacity: number
  dataUrl: string
  fileName: string
}

type RenderedPage = {
  pageNumber: number
  width: number
  height: number
  dataUrl: string
}

const acceptedTypes = '.pdf,.docx,.xlsx,.pptx'
const signatureAcceptedTypes = '.png,.jpg,.jpeg,.pdf'
const fontOptions = [
  'Angsana New',
  'Sarabun',
  'Noto Sans Thai',
  'Tahoma',
  'Arial',
  'Calibri',
  'Courier New',
  'Georgia',
  'Times New Roman',
  'Verdana',
  'Segoe UI',
  'Cordia New',
  'Browallia New',
  'Leelawadee UI',
]
const baseRenderScale = 1.35

const toolConfig: Record<Tool, { label: string; icon: string; cursor: string }> = {
  move: { label: 'เลื่อนดู', icon: '✋', cursor: 'grab' },
  text: { label: 'เพิ่มข้อความ', icon: 'T', cursor: 'crosshair' },
  mask: { label: 'แปะพื้นที่', icon: '■', cursor: 'crosshair' },
  delete: { label: 'ลบ', icon: '✕', cursor: 'pointer' },
}

function App() {
  const [sourceFile, setSourceFile] = useState<File | null>(null)
  const [pdfBlob, setPdfBlob] = useState<Blob | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pages, setPages] = useState<RenderedPage[]>([])
  const [annotations, setAnnotations] = useState<Annotation[]>([])
  const [masks, setMasks] = useState<Mask[]>([])
  const [signatures, setSignatures] = useState<Signature[]>([])
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [selectedSignatureId, setSelectedSignatureId] = useState<string | null>(null)
  const [selectedMaskId, setSelectedMaskId] = useState<string | null>(null)
  const [activeTool, setActiveTool] = useState<Tool>('text')
  const [draftText, setDraftText] = useState('ทดสอบ ข้อความ')
  const [fontFamily, setFontFamily] = useState(fontOptions[0])
  const [fontSize, setFontSize] = useState(24)
  const [color, setColor] = useState('#111827')
  const [maskColor, setMaskColor] = useState('#ffffff')
  const [bold, setBold] = useState(false)
  const [italic, setItalic] = useState(false)
  const [rotation, setRotation] = useState(0)
  const [busyLabel, setBusyLabel] = useState('')
  const [error, setError] = useState('')
  const [warning, setWarning] = useState('')
  const [draggingId, setDraggingId] = useState<string | null>(null)
  const [draggingSignatureId, setDraggingSignatureId] = useState<string | null>(null)
  const [draggingMaskId, setDraggingMaskId] = useState<string | null>(null)
  const [maskDraft, setMaskDraft] = useState<{
    page: number
    startX: number
    startY: number
    endX: number
    endY: number
  } | null>(null)
  const [zoomPercent, setZoomPercent] = useState(100)
  const [renderZoomPercent, setRenderZoomPercent] = useState(100)
  const [showRuler, setShowRuler] = useState(true)
  const [guideLines, setGuideLines] = useState<GuideLine[]>([])
  const [draggingGuideId, setDraggingGuideId] = useState<string | null>(null)
  const [hoverPos, setHoverPos] = useState<{ pageNumber: number; x: number; y: number } | null>(null)
  const [isPanning, setIsPanning] = useState(false)
  const [panStart, setPanStart] = useState({ x: 0, y: 0, scrollLeft: 0, scrollTop: 0 })
  const pageRefs = useRef<Record<number, HTMLDivElement | null>>({})
  const documentPanelRef = useRef<HTMLDivElement | null>(null)

  const backendUrl = `${window.location.protocol}//${window.location.hostname}:4000`
  const zoomScale = zoomPercent / 100

  useEffect(() => {
    const timeoutId = window.setTimeout(() => {
      setRenderZoomPercent(zoomPercent)
    }, 140)

    return () => window.clearTimeout(timeoutId)
  }, [zoomPercent])

  useEffect(() => {
    if (!pdfBlob) {
      return
    }

    let isCancelled = false
    const activePdfBlob = pdfBlob

    async function renderPdf() {
      setBusyLabel(
        renderZoomPercent > 100 ? `กำลังปรับความคมชัด ${renderZoomPercent}%...` : 'กำลังโหลด PDF...',
      )
      setError('')

      try {
        const bytes = await activePdfBlob.arrayBuffer()
        const pdf = await pdfjsLib.getDocument({ data: bytes.slice(0) }).promise
        const nextPages: RenderedPage[] = []
        const outputScale = baseRenderScale * Math.max(1, renderZoomPercent / 100)

        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
          const page = await pdf.getPage(pageNumber)
          const viewport = page.getViewport({ scale: outputScale })
          const canvas = document.createElement('canvas')
          const context = canvas.getContext('2d')

          if (!context) {
            throw new Error('ไม่สามารถสร้าง canvas สำหรับแสดง PDF ได้')
          }

          canvas.width = Math.floor(viewport.width)
          canvas.height = Math.floor(viewport.height)
          await page.render({ canvas, canvasContext: context, viewport }).promise

          nextPages.push({
            pageNumber,
            width: canvas.width,
            height: canvas.height,
            dataUrl: canvas.toDataURL('image/png'),
          })
        }

        if (!isCancelled) {
          setPages(nextPages)
        }
      } catch (renderError) {
        if (!isCancelled) {
          setError(renderError instanceof Error ? renderError.message : 'เปิด PDF ไม่สำเร็จ')
        }
      } finally {
        if (!isCancelled) {
          setBusyLabel('')
        }
      }
    }

    void renderPdf()

    return () => {
      isCancelled = true
    }
  }, [pdfBlob, renderZoomPercent])

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    if (!sourceFile) {
      setError('กรุณาเลือกไฟล์ PDF หรือ Office ก่อน')
      return
    }

    setBusyLabel(sourceFile.name.toLowerCase().endsWith('.pdf') ? 'กำลังเปิด PDF...' : 'กำลังแปลงไฟล์...')
    setError('')
    setWarning('')
    setAnnotations([])
    setMasks([])
    setSignatures([])
    setSelectedId(null)
    setSelectedSignatureId(null)
    setSelectedMaskId(null)

    try {
      if (sourceFile.name.toLowerCase().endsWith('.pdf')) {
        setPdfBlob(sourceFile)
        setPdfName(sourceFile.name)
        setWarning('โหมดนี้เป็นการเพิ่มข้อความทับบน PDF ไม่ได้แก้ข้อความเดิมในไฟล์โดยตรง')
        return
      }

      const formData = new FormData()
      formData.append('file', sourceFile)

      const response = await fetch(`${backendUrl}/api/convert/office-to-pdf`, {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        const message = await response.json().catch(() => null)
        throw new Error(message?.error || 'แปลงไฟล์ไม่สำเร็จ')
      }

      const blob = await response.blob()
      setPdfBlob(blob)
      setPdfName(sourceFile.name.replace(/\.[^.]+$/, '.pdf'))
      setWarning('หากเอกสารมีตารางซับซ้อน หลายคอลัมน์ หรือฟอนต์ไม่มีในเครื่อง เลย์เอาต์อาจคลาดเคลื่อน')
    } catch (conversionError) {
      setError(conversionError instanceof Error ? conversionError.message : 'แปลงไฟล์ไม่สำเร็จ')
    } finally {
      setBusyLabel('')
    }
  }

  // --- Tool-based page click handler ---
  function handlePageClick(page: RenderedPage, event: MouseEvent<HTMLDivElement>) {
    const target = event.target as HTMLElement

    // If clicking on annotation or signature, let their own handlers manage it
    if (target.closest('.annotation-box') || target.closest('.signature-overlay') || target.closest('.mask-overlay')) {
      return
    }

    const rect = event.currentTarget.getBoundingClientRect()
    const relX = (event.clientX - rect.left) / rect.width
    const relY = (event.clientY - rect.top) / rect.height

    if (activeTool === 'text') {
      addAnnotation(page.pageNumber, relX, relY)
    } else if (activeTool === 'mask') {
      finishMask(page.pageNumber, relX, relY)
    } else if (activeTool === 'delete') {
      // Delete mode: clicking empty area does nothing
    }
    // Move mode: clicking empty area does nothing (panning is handled by pointer events)
  }

  function handlePagePointerDown(_page: RenderedPage, event: PointerEvent<HTMLDivElement>) {
    if (activeTool === 'mask') {
      const rect = event.currentTarget.getBoundingClientRect()
      const x = clamp((event.clientX - rect.left) / rect.width, 0, 1)
      const y = clamp((event.clientY - rect.top) / rect.height, 0, 1)
      setMaskDraft({ page: _page.pageNumber, startX: x, startY: y, endX: x, endY: y })
      event.currentTarget.setPointerCapture(event.pointerId)
      return
    }

    if (activeTool === 'move') {
      const panel = documentPanelRef.current
      if (!panel) return
      setIsPanning(true)
      setPanStart({
        x: event.clientX,
        y: event.clientY,
        scrollLeft: panel.scrollLeft,
        scrollTop: panel.scrollTop,
      })
        ; (event.currentTarget as HTMLElement).setPointerCapture(event.pointerId)
    }
  }

  function handlePagePointerMove(pageNumber: number, event: PointerEvent<HTMLDivElement>) {
    if (draggingMaskId) {
      const mask = masks.find((item) => item.id === draggingMaskId)
      const pageElement = pageRefs.current[pageNumber]
      if (mask?.page === pageNumber && pageElement) {
        const rect = pageElement.getBoundingClientRect()
        updateMask(mask.id, {
          x: clamp((event.clientX - rect.left) / rect.width - mask.width / 2, 0, 1 - mask.width),
          y: clamp((event.clientY - rect.top) / rect.height - mask.height / 2, 0, 1 - mask.height),
        })
      }
      return
    }

    if (maskDraft?.page === pageNumber) {
      const rect = event.currentTarget.getBoundingClientRect()
      setMaskDraft((current) =>
        current
          ? {
              ...current,
              endX: clamp((event.clientX - rect.left) / rect.width, 0, 1),
              endY: clamp((event.clientY - rect.top) / rect.height, 0, 1),
            }
          : current,
      )
      return
    }

    if (activeTool === 'move' && isPanning) {
      const panel = documentPanelRef.current
      if (!panel) return
      const dx = event.clientX - panStart.x
      const dy = event.clientY - panStart.y
      panel.scrollLeft = panStart.scrollLeft - dx
      panel.scrollTop = panStart.scrollTop - dy
    }

    const rect = event.currentTarget.getBoundingClientRect()
    setHoverPos({
      pageNumber,
      x: clamp((event.clientX - rect.left) / rect.width, 0, 1),
      y: clamp((event.clientY - rect.top) / rect.height, 0, 1),
    })
  }

  function handlePagePointerUp() {
    if (draggingMaskId) {
      setDraggingMaskId(null)
    }
    if (maskDraft) {
      finishMask(maskDraft.page, maskDraft.endX, maskDraft.endY)
    }
    if (isPanning) {
      setIsPanning(false)
    }
  }

  function finishMask(pageNumber: number, endX: number, endY: number) {
    if (!maskDraft || maskDraft.page !== pageNumber) return
    const x = Math.min(maskDraft.startX, endX)
    const y = Math.min(maskDraft.startY, endY)
    const width = Math.abs(endX - maskDraft.startX)
    const height = Math.abs(endY - maskDraft.startY)
    setMaskDraft(null)

    if (width < 0.005 || height < 0.005) return
    const mask: Mask = { id: createId(), page: pageNumber, x, y, width, height, color: maskColor }
    setMasks((current) => [...current, mask])
    setSelectedMaskId(mask.id)
    setSelectedId(null)
    setSelectedSignatureId(null)
  }

  function handleMaskPointerDown(id: string, event: PointerEvent<HTMLDivElement>) {
    event.stopPropagation()
    if (activeTool === 'delete') {
      removeMask(id)
      return
    }
    setSelectedMaskId(id)
    setSelectedId(null)
    setSelectedSignatureId(null)
    event.currentTarget.setPointerCapture(event.pointerId)
    setDraggingMaskId(id)
  }

  function updateMask(id: string, next: Partial<Mask>) {
    setMasks((current) => current.map((mask) => (mask.id === id ? { ...mask, ...next } : mask)))
  }

  function removeMask(id: string) {
    setMasks((current) => current.filter((mask) => mask.id !== id))
    if (selectedMaskId === id) setSelectedMaskId(null)
  }

  function removeSelectedMask() {
    if (selectedMaskId) removeMask(selectedMaskId)
  }

  function handleRulerClick(
    type: 'vertical' | 'horizontal',
    pageNumber: number,
    event: MouseEvent<HTMLDivElement>,
  ) {
    event.stopPropagation()
    const rect = event.currentTarget.getBoundingClientRect()
    const position =
      type === 'vertical'
        ? clamp((event.clientX - rect.left) / rect.width, 0, 1)
        : clamp((event.clientY - rect.top) / rect.height, 0, 1)

    const newGuide: GuideLine = {
      id: createId(),
      pageNumber,
      type,
      position,
    }
    setGuideLines((current) => [...current, newGuide])
  }

  function handleGuidePointerDown(id: string, event: PointerEvent<HTMLDivElement>) {
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)
    setDraggingGuideId(id)
  }

  function handleGuidePointerMove(
    id: string,
    pageNumber: number,
    event: PointerEvent<HTMLDivElement>,
  ) {
    if (draggingGuideId !== id) return
    const pageEl = pageRefs.current[pageNumber]
    if (!pageEl) return
    const rect = pageEl.getBoundingClientRect()
    const guide = guideLines.find((g) => g.id === id)
    if (!guide) return

    const position =
      guide.type === 'vertical'
        ? clamp((event.clientX - rect.left) / rect.width, 0, 1)
        : clamp((event.clientY - rect.top) / rect.height, 0, 1)

    setGuideLines((current) =>
      current.map((g) => (g.id === id ? { ...g, position } : g)),
    )
  }

  function removeGuideLine(id: string) {
    setGuideLines((current) => current.filter((g) => g.id !== id))
  }

  function clearAllGuideLines() {
    setGuideLines([])
  }

  function addAnnotation(pageNumber = pages[0]?.pageNumber ?? 1, x = 0.18, y = 0.18) {
    const text = draftText.trim()

    if (!pages.length) {
      setError('กรุณาเปิด PDF ก่อนเพิ่มข้อความ')
      return
    }

    if (!text) {
      setError('กรุณาพิมพ์ข้อความก่อนวางลงบน PDF')
      return
    }

    const annotation: Annotation = {
      id: createId(),
      page: pageNumber,
      x: clamp(x, 0, 1),
      y: clamp(y, 0, 1),
      text,
      fontFamily,
      fontSize,
      color,
      bold,
      italic,
      rotation,
    }

    setAnnotations((current) => [...current, annotation])
    setSelectedId(annotation.id)
    setSelectedSignatureId(null)
    setError('')
  }

  function handleAnnotationPointerDown(id: string, event: PointerEvent<HTMLButtonElement>) {
    event.stopPropagation()

    if (activeTool === 'delete') {
      removeAnnotation(id)
      return
    }

    event.currentTarget.setPointerCapture(event.pointerId)
    selectAnnotation(id)
    setDraggingId(id)
  }

  function handleAnnotationPointerMove(id: string, event: PointerEvent<HTMLButtonElement>) {
    if (draggingId !== id) {
      return
    }

    const annotation = annotations.find((item) => item.id === id)
    const pageElement = annotation ? pageRefs.current[annotation.page] : null

    if (!annotation || !pageElement) {
      return
    }

    const rect = pageElement.getBoundingClientRect()
    updateAnnotation(id, {
      x: clamp((event.clientX - rect.left) / rect.width, 0, 1),
      y: clamp((event.clientY - rect.top) / rect.height, 0, 1),
    })
  }

  function applyStyleToSelected(next: Partial<Annotation>) {
    if (!selectedId) {
      return
    }

    updateAnnotation(selectedId, next)
  }

  function updateAnnotation(id: string, next: Partial<Annotation>) {
    setAnnotations((current) =>
      current.map((annotation) => (annotation.id === id ? { ...annotation, ...next } : annotation)),
    )
  }

  function selectAnnotation(id: string) {
    const annotation = annotations.find((item) => item.id === id)
    setSelectedId(id)
    setSelectedSignatureId(null)

    if (!annotation) {
      return
    }

    setDraftText(annotation.text)
    setFontFamily(annotation.fontFamily)
    setFontSize(annotation.fontSize)
    setColor(annotation.color)
    setBold(annotation.bold)
    setItalic(annotation.italic)
    setRotation(annotation.rotation ?? 0)
  }

  function removeAnnotation(id: string) {
    setAnnotations((current) => current.filter((annotation) => annotation.id !== id))
    if (selectedId === id) {
      setSelectedId(null)
    }
  }

  function removeSelectedAnnotation() {
    if (!selectedId) {
      return
    }
    removeAnnotation(selectedId)
  }

  // --- Signature handlers ---
  async function handleSignatureUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return

    const ext = file.name.toLowerCase().match(/\.(png|jpg|jpeg|pdf)$/)
    if (!ext) {
      setError('รองรับเฉพาะไฟล์ .png, .jpg, .pdf')
      return
    }

    try {
      let dataUrl = ''

      if (ext[1] === 'pdf') {
        // Render first page of PDF signature to image
        const bytes = await file.arrayBuffer()
        const pdf = await pdfjsLib.getDocument({ data: bytes.slice(0) }).promise
        const page = await pdf.getPage(1)
        const viewport = page.getViewport({ scale: 2 })
        const canvas = document.createElement('canvas')
        const context = canvas.getContext('2d')!
        canvas.width = Math.floor(viewport.width)
        canvas.height = Math.floor(viewport.height)
        await page.render({ canvas, canvasContext: context, viewport }).promise
        dataUrl = canvas.toDataURL('image/png')
        await pdf.cleanup()
      } else {
        // Read image file as data URL
        dataUrl = await readFileAsDataUrl(file)
      }

      // Get image dimensions to calculate aspect ratio
      const img = await loadImage(dataUrl)
      const defaultWidth = 0.15 // 15% of page width
      const aspectRatio = img.height / img.width
      const defaultHeight = defaultWidth * aspectRatio

      const signature: Signature = {
        id: createId(),
        page: pages[0]?.pageNumber ?? 1,
        x: 0.4,
        y: 0.7,
        width: defaultWidth,
        height: defaultHeight,
        rotation: 0,
        opacity: 1,
        dataUrl,
        fileName: file.name,
      }

      setSignatures((current) => [...current, signature])
      setSelectedSignatureId(signature.id)
      setSelectedId(null)
      setError('')
    } catch (uploadError) {
      setError(uploadError instanceof Error ? uploadError.message : 'อัพโหลดลายเซ็นไม่สำเร็จ')
    }

    // Reset file input
    event.target.value = ''
  }

  function handleSignaturePointerDown(id: string, event: PointerEvent<HTMLDivElement>) {
    event.stopPropagation()

    if (activeTool === 'delete') {
      removeSignature(id)
      return
    }

    event.currentTarget.setPointerCapture(event.pointerId)
    setSelectedSignatureId(id)
    setSelectedId(null)
    setDraggingSignatureId(id)
  }

  function handleSignaturePointerMove(id: string, event: PointerEvent<HTMLDivElement>) {
    if (draggingSignatureId !== id) {
      return
    }

    const signature = signatures.find((item) => item.id === id)
    const pageElement = signature ? pageRefs.current[signature.page] : null

    if (!signature || !pageElement) {
      return
    }

    const rect = pageElement.getBoundingClientRect()
    updateSignature(id, {
      x: clamp((event.clientX - rect.left) / rect.width, 0, 1),
      y: clamp((event.clientY - rect.top) / rect.height, 0, 1),
    })
  }

  function updateSignature(id: string, next: Partial<Signature>) {
    setSignatures((current) =>
      current.map((sig) => (sig.id === id ? { ...sig, ...next } : sig)),
    )
  }

  function removeSignature(id: string) {
    setSignatures((current) => current.filter((sig) => sig.id !== id))
    if (selectedSignatureId === id) {
      setSelectedSignatureId(null)
    }
  }

  function removeSelectedSignature() {
    if (!selectedSignatureId) return
    removeSignature(selectedSignatureId)
  }

  const selectedSignature = signatures.find((s) => s.id === selectedSignatureId)

  function changeZoom(nextZoom: number) {
    setZoomPercent(clamp(Math.round(nextZoom), 50, 500))
  }

  async function downloadEditedPdf() {
    await postForDownload('/api/pdf/save', 'edited.pdf', 'กำลังบันทึก PDF...')
  }

  async function postForDownload(endpoint: string, fallbackName: string, label: string) {
    if (!pdfBlob) {
      setError('กรุณาเปิด PDF ก่อนบันทึกไฟล์')
      return
    }

    const formData = new FormData()
    const baseName = pdfName.replace(/\.[^.]+$/, '') || 'edited'
    formData.append('file', pdfBlob, pdfName || 'document.pdf')
    formData.append('annotations', JSON.stringify(annotations))
    formData.append('masks', JSON.stringify(masks))
    formData.append('signatures', JSON.stringify(signatures))

    setBusyLabel(label)
    setError('')

    try {
      const response = await fetch(`${backendUrl}${endpoint}`, {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        const message = await response.json().catch(() => null)
        throw new Error(message?.error || 'บันทึกไฟล์ไม่สำเร็จ')
      }

      const blob = await response.blob()
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = `${baseName}.pdf`
      link.click()
      URL.revokeObjectURL(url)
    } catch (saveError) {
      setError(saveError instanceof Error ? saveError.message : fallbackName)
    } finally {
      setBusyLabel('')
    }
  }

  return (
    <div className="app-frame">
      <header className="top-navbar">
        <a className="brand-mark" href="#tools" aria-label="PDFCONVERTOR home">
          <img src="/PDFCONVERTOR/favicon.svg" alt="" />
          <span>
            <strong>PDFCONVERTOR</strong>
            <small>Editor workspace</small>
          </span>
        </a>

        <nav className="nav-links" aria-label="Primary">
          <a href="#tools">Tools</a>
          <a href="#preview">Preview</a>
          <a href="#export">Export</a>
        </nav>

        <div className="nav-status" aria-live="polite">
          <span className={busyLabel ? 'status-dot busy' : 'status-dot'} />
          <span>{busyLabel || (pdfBlob ? 'Ready to edit' : 'Waiting for file')}</span>
        </div>
      </header>

      <main className="app-shell">
        <aside id="tools" className="tool-panel" aria-labelledby="app-title">
          <div className="intro">
            <p className="eyebrow">PDFCONVERTOR</p>
            <h1 id="app-title">PDF Editor</h1>
            <p>เพิ่มข้อความและลายเซ็นบน PDF</p>
          </div>

          {/* --- File Upload --- */}
          <form className="upload-form" onSubmit={handleSubmit}>
            <label className="file-picker">
              <span>เลือกไฟล์ PDF หรือ Office</span>
              <input
                type="file"
                accept={acceptedTypes}
                onChange={(event) => {
                  setSourceFile(event.target.files?.[0] ?? null)
                  setError('')
                }}
              />
            </label>

            <div className="file-summary">
              {sourceFile ? (
                <>
                  <strong>{sourceFile.name}</strong>
                  <span>{(sourceFile.size / 1024 / 1024).toFixed(2)} MB</span>
                </>
              ) : (
                <span>รองรับ .pdf, .docx, .xlsx, .pptx</span>
              )}
            </div>

            <button type="submit" disabled={Boolean(busyLabel)}>
              {busyLabel || 'เปิดเอกสาร'}
            </button>
          </form>

          {/* --- Tool Selector --- */}
          <section className="tool-selector" aria-label="Tool selection">
            <h3 className="section-title">เครื่องมือ</h3>
            <div className="tool-buttons">
              {(Object.keys(toolConfig) as Tool[]).map((tool) => (
                <button
                  key={tool}
                  type="button"
                  className={`tool-btn ${activeTool === tool ? 'active' : ''}`}
                  onClick={() => {
                    setActiveTool(tool)
                    setSelectedId(null)
                    setSelectedSignatureId(null)
                    setSelectedMaskId(null)
                  }}
                  title={toolConfig[tool].label}
                >
                  <span className="tool-icon">{toolConfig[tool].icon}</span>
                  <span className="tool-label">{toolConfig[tool].label}</span>
                </button>
              ))}
            </div>
            <p className="tool-hint">{toolConfig[activeTool].label}: {activeTool === 'move' ? 'ลากบนเอกสารเพื่อเลื่อนดู' : activeTool === 'text' ? 'คลิกบนเอกสารเพื่อวางข้อความ' : activeTool === 'mask' ? 'ลากเพื่อเลือกพื้นที่และปิดข้อความเดิม' : 'คลิกที่ข้อความ ลายเซ็น หรือพื้นที่เพื่อ ลบ'}</p>
          </section>

          {/* --- Text Tools (only visible when text tool is active or text is selected) --- */}
          {(activeTool === 'text' || selectedId) && (
            <section className="editor-tools" aria-label="Text tools">
              <h3 className="section-title">ข้อความ</h3>

              <label>
                ข้อความ
                <textarea
                  value={draftText}
                  onChange={(event) => {
                    setDraftText(event.target.value)
                    applyStyleToSelected({ text: event.target.value })
                  }}
                />
              </label>

              <label>
                ฟอนต์
                <select
                  value={fontFamily}
                  onChange={(event) => {
                    setFontFamily(event.target.value)
                    applyStyleToSelected({ fontFamily: event.target.value })
                  }}
                >
                  {fontOptions.map((font) => (
                    <option key={font} value={font}>
                      {font}
                    </option>
                  ))}
                </select>
              </label>

              <div className="tool-row">
                <label>
                  ขนาด
                  <input
                    type="number"
                    min="8"
                    max="96"
                    value={fontSize}
                    onChange={(event) => {
                      const nextSize = Number(event.target.value)
                      setFontSize(nextSize)
                      applyStyleToSelected({ fontSize: nextSize })
                    }}
                  />
                </label>
                <label>
                  สี
                  <input
                    type="color"
                    value={color}
                    onChange={(event) => {
                      setColor(event.target.value)
                      applyStyleToSelected({ color: event.target.value })
                    }}
                  />
                </label>
              </div>

              <div className="rotation-control">
                <div className="rotation-header">
                  <span>การหมุน: {rotation}°</span>
                  <div className="rotation-actions">
                    <button
                      type="button"
                      title="หมุนซ้าย 90°"
                      onClick={() => {
                        const nextRot = rotation - 90 < -180 ? rotation - 90 + 360 : rotation - 90
                        setRotation(nextRot)
                        applyStyleToSelected({ rotation: nextRot })
                      }}
                    >
                      ↺ -90°
                    </button>
                    <button
                      type="button"
                      title="รีเซ็ต 0°"
                      disabled={rotation === 0}
                      onClick={() => {
                        setRotation(0)
                        applyStyleToSelected({ rotation: 0 })
                      }}
                    >
                      0°
                    </button>
                    <button
                      type="button"
                      title="หมุนขวา 90°"
                      onClick={() => {
                        const nextRot = rotation + 90 > 180 ? rotation + 90 - 360 : rotation + 90
                        setRotation(nextRot)
                        applyStyleToSelected({ rotation: nextRot })
                      }}
                    >
                      ↻ +90°
                    </button>
                  </div>
                </div>
                <input
                  type="range"
                  min="-180"
                  max="180"
                  step="1"
                  value={rotation}
                  onChange={(event) => {
                    const nextRot = Number(event.target.value)
                    setRotation(nextRot)
                    applyStyleToSelected({ rotation: nextRot })
                  }}
                />
              </div>

              <div className="segmented-controls">
                <button type="button" className="add-text-button" disabled={!pages.length} onClick={() => addAnnotation()}>
                  เพิ่มข้อความ
                </button>
                <button
                  type="button"
                  className={bold ? 'active' : ''}
                  onClick={() => {
                    setBold((current) => !current)
                    applyStyleToSelected({ bold: !bold })
                  }}
                >
                  B
                </button>
                <button
                  type="button"
                  className={italic ? 'active' : ''}
                  onClick={() => {
                    setItalic((current) => !current)
                    applyStyleToSelected({ italic: !italic })
                  }}
                >
                  I
                </button>
                <button type="button" disabled={!selectedId} onClick={removeSelectedAnnotation}>
                  ลบ
                </button>
              </div>
            </section>
          )}

          {(activeTool === 'mask' || selectedMaskId) && (
            <section className="mask-tools" aria-label="Mask tools">
              <h3 className="section-title">แปะพื้นที่</h3>
              <label>
                สีพื้นที่ปิดข้อความ
                <input
                  type="color"
                  value={maskColor}
                  onChange={(event) => {
                    const nextColor = event.target.value
                    setMaskColor(nextColor)
                    if (selectedMaskId) {
                      setMasks((current) =>
                        current.map((mask) =>
                          mask.id === selectedMaskId ? { ...mask, color: nextColor } : mask,
                        ),
                      )
                    }
                  }}
                />
              </label>
              <p className="tool-hint">ลากบนเอกสารเพื่อสร้างพื้นที่ปิดทับ ข้อความใหม่จะอยู่ด้านบน</p>
              <button type="button" disabled={!selectedMaskId} onClick={removeSelectedMask}>
                ลบพื้นที่ที่เลือก
              </button>
            </section>
          )}

          {/* --- Signature Tools --- */}
          <section className="signature-tools" aria-label="Signature tools">
            <h3 className="section-title">ลายเซ็น</h3>

            <label className="file-picker signature-picker">
              <span>อัพโหลดลายเซ็น (.png, .jpg, .pdf)</span>
              <input
                type="file"
                accept={signatureAcceptedTypes}
                onChange={handleSignatureUpload}
              />
            </label>

            {signatures.length > 0 && (
              <div className="signature-list">
                <p className="signature-count">{signatures.length} ลายเซ็น</p>
                {signatures.map((sig) => (
                  <button
                    key={sig.id}
                    type="button"
                    className={`signature-list-item ${sig.id === selectedSignatureId ? 'selected' : ''}`}
                    onClick={() => {
                      setSelectedSignatureId(sig.id)
                      setSelectedId(null)
                    }}
                  >
                    <img src={sig.dataUrl} alt={sig.fileName} />
                    <span>{sig.fileName}</span>
                  </button>
                ))}
              </div>
            )}

            {/* Signature controls */}
            {selectedSignature && (
              <div className="signature-controls">
                <label>
                  ขนาด: {Math.round(selectedSignature.width * 100)}%
                  <input
                    type="range"
                    min="3"
                    max="60"
                    step="1"
                    value={Math.round(selectedSignature.width * 100)}
                    onChange={(event) => {
                      const newWidth = Number(event.target.value) / 100
                      const aspectRatio = selectedSignature.height / selectedSignature.width
                      updateSignature(selectedSignature.id, {
                        width: newWidth,
                        height: newWidth * aspectRatio,
                      })
                    }}
                  />
                </label>

                <label>
                  หมุน: {selectedSignature.rotation}°
                  <input
                    type="range"
                    min="-180"
                    max="180"
                    step="1"
                    value={selectedSignature.rotation}
                    onChange={(event) => {
                      updateSignature(selectedSignature.id, {
                        rotation: Number(event.target.value),
                      })
                    }}
                  />
                </label>

                <label>
                  ความเข้ม: {Math.round(selectedSignature.opacity * 100)}%
                  <input
                    type="range"
                    min="10"
                    max="100"
                    step="5"
                    value={Math.round(selectedSignature.opacity * 100)}
                    onChange={(event) => {
                      updateSignature(selectedSignature.id, {
                        opacity: Number(event.target.value) / 100,
                      })
                    }}
                  />
                </label>

                <label>
                  หน้า
                  <select
                    value={selectedSignature.page}
                    onChange={(event) => {
                      updateSignature(selectedSignature.id, {
                        page: Number(event.target.value),
                      })
                    }}
                  >
                    {pages.map((p) => (
                      <option key={p.pageNumber} value={p.pageNumber}>
                        หน้า {p.pageNumber}
                      </option>
                    ))}
                  </select>
                </label>

                <button
                  type="button"
                  className="delete-signature-btn"
                  onClick={() => removeSelectedSignature()}
                >
                  ✕ ลบลายเซ็นนี้
                </button>
              </div>
            )}
          </section>

          {/* --- Save Actions --- */}
          <section id="export" className="save-actions" aria-label="Save actions">
            <button type="button" disabled={!pdfBlob || Boolean(busyLabel)} onClick={downloadEditedPdf}>
              Save PDF
            </button>
          </section>

          {warning && <p className="warning-message">{warning}</p>}
          {error && <p className="error-message">{error}</p>}
        </aside>

        <section
          id="preview"
          className="document-panel"
          aria-label="PDF editor canvas"
          ref={documentPanelRef}
        >
          <div className="preview-toolbar">
            <div className="preview-toolbar-info">
              <strong>Preview</strong>
              <span>{pages.length ? `${pages.length} page${pages.length > 1 ? 's' : ''}` : 'No document'}</span>
            </div>

            <div className="preview-toolbar-actions">
              <button
                type="button"
                className={`ruler-btn ${showRuler ? 'active' : ''}`}
                title={showRuler ? 'ปิดไม้บรรทัด' : 'เปิดไม้บรรทัด'}
                onClick={() => setShowRuler((current) => !current)}
              >
                📐 ไม้บรรทัด
              </button>
              {guideLines.length > 0 && (
                <button
                  type="button"
                  className="clear-guides-btn"
                  title="ล้างเส้นบอกตำแหน่งทั้งหมด"
                  onClick={clearAllGuideLines}
                >
                  ✕ ล้างเส้น ({guideLines.length})
                </button>
              )}
            </div>

            <div className="zoom-controls" aria-label="Document zoom controls">
              <button
                type="button"
                aria-label="Zoom out"
                disabled={zoomPercent <= 50}
                onClick={() => changeZoom(zoomPercent - 25)}
              >
                -
              </button>
              <input
                type="range"
                min="50"
                max="500"
                step="25"
                value={zoomPercent}
                aria-label="Document zoom"
                onChange={(event) => changeZoom(Number(event.target.value))}
              />
              <button
                type="button"
                aria-label="Zoom in"
                disabled={zoomPercent >= 500}
                onClick={() => changeZoom(zoomPercent + 25)}
              >
                +
              </button>
              <button type="button" className="zoom-reset" onClick={() => changeZoom(100)}>
                {zoomPercent}%
              </button>
            </div>
          </div>

          {pages.length > 0 ? (
            <div className="pages" style={{ '--zoom-scale': zoomScale } as CSSProperties}>
              {pages.map((page) => (
                <div key={page.pageNumber} className={`pdf-page-wrapper ${showRuler ? 'with-ruler' : ''}`}>
                  {showRuler && (
                    <>
                      <div className="ruler-corner" title="ไม้บรรทัด (เซนติเมตร)">
                        cm
                      </div>
                      <div
                        className="ruler-top"
                        title="คลิกเพื่อปักเส้นบอกตำแหน่งแนวตั้ง"
                        onClick={(event) => handleRulerClick('vertical', page.pageNumber, event)}
                      >
                        <RulerHorizontal maxCm={21} />
                        {hoverPos?.pageNumber === page.pageNumber && (
                          <div className="ruler-tracker-v" style={{ left: `${hoverPos.x * 100}%` }} />
                        )}
                      </div>
                      <div
                        className="ruler-left"
                        title="คลิกเพื่อปักเส้นบอกตำแหน่งแนวนอน"
                        onClick={(event) => handleRulerClick('horizontal', page.pageNumber, event)}
                      >
                        <RulerVertical maxCm={29.7} />
                        {hoverPos?.pageNumber === page.pageNumber && (
                          <div className="ruler-tracker-h" style={{ top: `${hoverPos.y * 100}%` }} />
                        )}
                      </div>
                    </>
                  )}

                  <div
                    ref={(node) => {
                      pageRefs.current[page.pageNumber] = node
                    }}
                    className="pdf-page"
                    style={{
                      aspectRatio: `${page.width} / ${page.height}`,
                      cursor:
                        activeTool === 'move'
                          ? isPanning
                            ? 'grabbing'
                            : 'grab'
                          : toolConfig[activeTool].cursor,
                    }}
                    onClick={(event) => handlePageClick(page, event)}
                    onPointerDown={(event) => handlePagePointerDown(page, event)}
                    onPointerMove={(event) => handlePagePointerMove(page.pageNumber, event)}
                    onPointerUp={handlePagePointerUp}
                    onMouseLeave={() => setHoverPos(null)}
                  >
                    <img src={page.dataUrl} alt={`หน้า ${page.pageNumber}`} />

                    {/* Guidelines */}
                    {showRuler &&
                      guideLines
                        .filter((guide) => guide.pageNumber === page.pageNumber)
                        .map((guide) => (
                          <div
                            key={guide.id}
                            className={`guideline guideline-${guide.type === 'vertical' ? 'v' : 'h'}`}
                            style={
                              guide.type === 'vertical'
                                ? { left: `${guide.position * 100}%` }
                                : { top: `${guide.position * 100}%` }
                            }
                            onPointerDown={(event) => handleGuidePointerDown(guide.id, event)}
                            onPointerMove={(event) =>
                              handleGuidePointerMove(guide.id, page.pageNumber, event)
                            }
                            onPointerUp={() => setDraggingGuideId(null)}
                          >
                            <span
                              className="guideline-badge"
                              title="คลิกเพื่อลบเส้นนี้"
                              onClick={(event) => {
                                event.stopPropagation()
                                removeGuideLine(guide.id)
                              }}
                            >
                              {guide.type === 'vertical'
                                ? `X: ${(guide.position * 21).toFixed(1)} cm`
                                : `Y: ${(guide.position * 29.7).toFixed(1)} cm`} ✕
                            </span>
                          </div>
                        ))}

                    {/* Masks: rendered before text so new text stays on top */}
                    {masks
                      .filter((mask) => mask.page === page.pageNumber)
                      .map((mask) => (
                        <div
                          key={mask.id}
                          className={`mask-overlay ${mask.id === selectedMaskId ? 'selected' : ''}`}
                          style={{
                            left: `${mask.x * 100}%`,
                            top: `${mask.y * 100}%`,
                            width: `${mask.width * 100}%`,
                            height: `${mask.height * 100}%`,
                            backgroundColor: mask.color,
                          }}
                          onPointerDown={(event) => handleMaskPointerDown(mask.id, event)}
                          onPointerUp={() => setDraggingMaskId(null)}
                        />
                      ))}

                    {maskDraft?.page === page.pageNumber && (
                      <div
                        className="mask-overlay mask-draft"
                        style={{
                          left: `${Math.min(maskDraft.startX, maskDraft.endX) * 100}%`,
                          top: `${Math.min(maskDraft.startY, maskDraft.endY) * 100}%`,
                          width: `${Math.abs(maskDraft.endX - maskDraft.startX) * 100}%`,
                          height: `${Math.abs(maskDraft.endY - maskDraft.startY) * 100}%`,
                          backgroundColor: maskColor,
                        }}
                      />
                    )}

                    {/* Text Annotations */}
                    {annotations
                      .filter((annotation) => annotation.page === page.pageNumber)
                      .map((annotation) => (
                        <button
                          key={annotation.id}
                          type="button"
                          className={`annotation-box ${annotation.id === selectedId ? 'selected' : ''}`}
                          style={{
                            left: `${annotation.x * 100}%`,
                            top: `${annotation.y * 100}%`,
                            color: annotation.color,
                            fontFamily: annotation.fontFamily,
                            fontSize: annotation.fontSize * zoomScale,
                            fontWeight: annotation.bold ? 700 : 400,
                            fontStyle: annotation.italic ? 'italic' : 'normal',
                            transform: `translate(-2px, -50%) rotate(${annotation.rotation || 0}deg)`,
                            transformOrigin: 'left center',
                          }}
                          onPointerDown={(event) => handleAnnotationPointerDown(annotation.id, event)}
                          onPointerMove={(event) => handleAnnotationPointerMove(annotation.id, event)}
                          onPointerUp={() => setDraggingId(null)}
                        >
                          {annotation.text}
                        </button>
                      ))}

                    {/* Signatures */}
                    {signatures
                      .filter((sig) => sig.page === page.pageNumber)
                      .map((sig) => (
                        <div
                          key={sig.id}
                          className={`signature-overlay ${sig.id === selectedSignatureId ? 'selected' : ''}`}
                          style={{
                            left: `${sig.x * 100}%`,
                            top: `${sig.y * 100}%`,
                            width: `${sig.width * 100}%`,
                            height: `${sig.height * 100}%`,
                            transform: `translate(-50%, -50%) rotate(${sig.rotation}deg)`,
                            opacity: sig.opacity,
                          }}
                          onPointerDown={(event) => handleSignaturePointerDown(sig.id, event)}
                          onPointerMove={(event) => handleSignaturePointerMove(sig.id, event)}
                          onPointerUp={() => setDraggingSignatureId(null)}
                        >
                          <img src={sig.dataUrl} alt={sig.fileName} draggable={false} />
                        </div>
                      ))}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="empty-preview">
              <h2>PDF Preview</h2>
              <p>เปิด PDF หรือแปลง Office เป็น PDF แล้วใช้เครื่องมือเพื่อเพิ่มข้อความหรือลายเซ็น</p>
            </div>
          )}
        </section>
      </main>
    </div>
  )
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max)
}

function createId() {
  if ('randomUUID' in crypto) {
    return crypto.randomUUID()
  }

  return `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(new Error('อ่านไฟล์ไม่สำเร็จ'))
    reader.readAsDataURL(file)
  })
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error('โหลดรูปไม่สำเร็จ'))
    img.src = src
  })
}

function RulerHorizontal({ maxCm = 21 }: { maxCm?: number }) {
  const mm = Math.round(maxCm * 10)
  const ticks = []
  for (let i = 0; i <= mm; i++) {
    const isCm = i % 10 === 0
    const isHalfCm = i % 5 === 0 && !isCm
    const h = isCm ? 12 : isHalfCm ? 7 : 4
    ticks.push(
      <line
        key={`th-${i}`}
        x1={i}
        y1={20 - h}
        x2={i}
        y2={20}
        stroke="currentColor"
        strokeWidth={isCm ? 0.6 : 0.35}
      />,
    )
    if (isCm && i < mm) {
      ticks.push(
        <text
          key={`lh-${i}`}
          x={i + 1.2}
          y={8}
          fontSize="5.5"
          fill="currentColor"
          fontFamily="system-ui, sans-serif"
        >
          {i / 10}
        </text>,
      )
    }
  }

  return (
    <svg
      viewBox={`0 0 ${mm} 20`}
      preserveAspectRatio="none"
      className="ruler-svg"
      aria-hidden="true"
    >
      <rect width={mm} height={20} fill="#f8fafc" />
      <line x1="0" y1="19.5" x2={mm} y2="19.5" stroke="#cbd5e1" strokeWidth="0.5" />
      {ticks}
    </svg>
  )
}

function RulerVertical({ maxCm = 29.7 }: { maxCm?: number }) {
  const mm = Math.round(maxCm * 10)
  const ticks = []
  for (let i = 0; i <= mm; i++) {
    const isCm = i % 10 === 0
    const isHalfCm = i % 5 === 0 && !isCm
    const w = isCm ? 12 : isHalfCm ? 7 : 4
    ticks.push(
      <line
        key={`tv-${i}`}
        x1={20 - w}
        y1={i}
        x2={20}
        y2={i}
        stroke="currentColor"
        strokeWidth={isCm ? 0.6 : 0.35}
      />,
    )
    if (isCm && i < mm) {
      ticks.push(
        <text
          key={`lv-${i}`}
          x={3}
          y={i + 6.5}
          fontSize="5.5"
          fill="currentColor"
          fontFamily="system-ui, sans-serif"
        >
          {i / 10}
        </text>,
      )
    }
  }

  return (
    <svg
      viewBox={`0 0 20 ${mm}`}
      preserveAspectRatio="none"
      className="ruler-svg"
      aria-hidden="true"
    >
      <rect width={20} height={mm} fill="#f8fafc" />
      <line x1="19.5" y1="0" x2="19.5" y2={mm} stroke="#cbd5e1" strokeWidth="0.5" />
      {ticks}
    </svg>
  )
}

export default App
