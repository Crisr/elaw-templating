import { useCallback, useRef, useState } from 'react'

interface Props {
  file: File | null
  onFile: (f: File) => void
  disabled: boolean
}

export default function DropZone({ file, onFile, disabled }: Props) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

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
          ? 'border-indigo-500 bg-indigo-50'
          : 'border-gray-300 hover:border-gray-400 bg-gray-50'
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
        <p className="text-gray-700 font-medium">{file.name}</p>
      ) : (
        <div>
          <p className="text-gray-500">Drag & drop a .docx file here</p>
          <p className="text-gray-400 text-sm mt-1">or click to browse</p>
        </div>
      )}
    </div>
  )
}
