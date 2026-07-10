import { useCallback, useRef, useState } from 'react'
import { useLocale } from '../LocaleContext'
import JSZip from 'jszip'

interface Props {
  file: File | null
  onFile: (f: File) => void
  disabled: boolean
  onTwoColumnDetected?: (v: boolean) => void
}

export default function DropZone({ file, onFile, disabled, onTwoColumnDetected }: Props) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const { messages } = useLocale()
  const msg = messages.dropZone

  const detectTwoColumn = useCallback(async (f: File) => {
    if (!onTwoColumnDetected) return
    try {
      const buf = await f.arrayBuffer()
      const zip = await JSZip.loadAsync(buf)
      const docXml = await zip.file('word/document.xml')?.async('string')
      if (!docXml) { onTwoColumnDetected(false); return }
      const hasWordCols = /<w:cols[^>]*w:num\s*=\s*"2"/.test(docXml)
      const gridMatch = docXml.match(/<w:tblGrid>(.*?)<\/w:tblGrid>/)
      const has2ColTable = gridMatch ? (gridMatch[1].match(/<w:gridCol/g) || []).length === 2 : false
      onTwoColumnDetected(hasWordCols || has2ColTable)
    } catch {
      onTwoColumnDetected(false)
    }
  }, [onTwoColumnDetected])

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault()
      setDragging(false)
      if (disabled) return
      const f = e.dataTransfer.files[0]
      if (f && f.name.endsWith('.docx')) {
        onFile(f)
        detectTwoColumn(f)
      }
    },
    [disabled, onFile, detectTwoColumn]
  )

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0]
      if (f) {
        onFile(f)
        detectTwoColumn(f)
      }
    },
    [onFile, detectTwoColumn]
  )

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); if (!disabled) setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors ${
        dragging
          ? 'border-brand-500 bg-brand-50'
          : 'border-brand-200 hover:border-brand-300 bg-brand-50/50'
      } ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".docx"
        className="hidden"
        onChange={handleChange}
      />
      {file ? (
        <p className="text-brand-500 font-medium">{file.name}</p>
      ) : (
        <div>
          <p className="text-brand-400">{msg.dragDrop}</p>
          <p className="text-brand-300 text-sm mt-1">{msg.clickToBrowse}</p>
        </div>
      )}
    </div>
  )
}
