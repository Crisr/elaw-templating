import { useState, useCallback } from 'react'
import DropZone from './components/DropZone'
import OptionsForm from './components/OptionsForm'
import ProgressBar from './components/ProgressBar'
import DownloadLink from './components/DownloadLink'

type Status = 'idle' | 'uploading' | 'translating' | 'done' | 'error'

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [lang, setLang] = useState('ro')
  const [mode, setMode] = useState('inline')
  const [providerName, setProviderName] = useState('')
  const [modelName, setModelName] = useState('')
  const [status, setStatus] = useState<Status>('idle')
  const [jobId, setJobId] = useState<string | null>(null)
  const [progress, setProgress] = useState(0)
  const [total, setTotal] = useState(0)
  const [error, setError] = useState('')

  const handleSubmit = useCallback(async () => {
    if (!file) return
    setStatus('uploading')
    setError('')

    const formData = new FormData()
    formData.append('file', file)
    formData.append('lang', lang)
    formData.append('mode', mode)
    if (providerName) formData.append('provider', providerName)
    if (modelName) formData.append('model', modelName)

    try {
      const resp = await fetch('/api/translate', { method: 'POST', body: formData })
      if (!resp.ok) throw new Error(`Upload failed: ${resp.statusText}`)
      const data = await resp.json()
      setJobId(data.job_id)
      setStatus('translating')

      const poll = setInterval(async () => {
        const sResp = await fetch(`/api/translate/${data.job_id}/status`)
        if (!sResp.ok) {
          clearInterval(poll)
          setStatus('error')
          setError('Failed to fetch status')
          return
        }
        const sData = await sResp.json()
        setProgress(sData.progress)
        setTotal(sData.total)

        if (sData.status === 'done') {
          clearInterval(poll)
          setStatus('done')
        } else if (sData.status === 'failed') {
          clearInterval(poll)
          setStatus('error')
          setError(sData.error || 'Translation failed')
        }
      }, 500)
    } catch (e) {
      setStatus('error')
      setError(e instanceof Error ? e.message : 'Unknown error')
    }
  }, [file, lang, mode, providerName, modelName])

  const handleReset = () => {
    setFile(null)
    setJobId(null)
    setStatus('idle')
    setProgress(0)
    setTotal(0)
    setError('')
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-start justify-center p-4 sm:p-8">
      <div className="w-full max-w-xl bg-white rounded-2xl shadow-lg p-6 sm:p-8 space-y-6">
        <h1 className="text-2xl font-bold text-gray-800">DOCX Translator</h1>

        {status === 'idle' || status === 'uploading' ? (
          <>
            <DropZone file={file} onFile={setFile} disabled={status === 'uploading'} />
            <OptionsForm
              lang={lang}
              mode={mode}
              providerName={providerName}
              modelName={modelName}
              onLangChange={setLang}
              onModeChange={setMode}
              onProviderChange={setProviderName}
              onModelChange={setModelName}
            />
            <button
              onClick={handleSubmit}
              disabled={!file || status === 'uploading'}
              className="w-full py-3 px-4 bg-indigo-600 text-white font-medium rounded-xl disabled:opacity-40 disabled:cursor-not-allowed hover:bg-indigo-700 transition-colors"
            >
              {status === 'uploading' ? 'Uploading...' : 'Convert'}
            </button>
          </>
        ) : status === 'translating' ? (
          <ProgressBar progress={progress} total={total} />
        ) : status === 'done' && jobId ? (
          <DownloadLink jobId={jobId} onReset={handleReset} />
        ) : status === 'error' ? (
          <div className="space-y-4">
            <div className="bg-red-50 border border-red-200 text-red-700 rounded-xl p-4">
              {error || 'An error occurred'}
            </div>
            <button
              onClick={handleReset}
              className="w-full py-3 px-4 bg-gray-600 text-white font-medium rounded-xl hover:bg-gray-700 transition-colors"
            >
              Try Again
            </button>
          </div>
        ) : null}
      </div>
    </div>
  )
}

export default App
