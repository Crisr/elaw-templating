import { useCallback, useRef, useState } from 'react'
import { messages } from '../messages'

interface Props {
  file: File | null
  onFile: (f: File) => void
  disabled: boolean
}

export default function DropZone({ file, onFile, disabled }: Props) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const { dropZone: msg } = messages

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault()
      setDragging(false)
      if (disabled) return
      const f = e.dataTransfer.files[0]
      if (f && f.name.endsWith('.docx')) onFile(f)
    },
    [disabled, onFile]
  )

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0]
      if (f) onFile(f)
    },
    [onFile]
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
