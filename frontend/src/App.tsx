import { useState, useCallback, useRef, useEffect } from 'react'
import DropZone from './components/DropZone'
import OptionsForm from './components/OptionsForm'
import ProgressBar from './components/ProgressBar'
import DownloadLink from './components/DownloadLink'
import { useLocale } from './LocaleContext'

type Status = 'idle' | 'uploading' | 'translating' | 'done' | 'error'

function App() {
  const { messages, locale, setLocale } = useLocale()
  const [file, setFile] = useState<File | null>(null)
  const [lang, setLang] = useState('ro')
  const [mode, setMode] = useState('inline')
  const [providerName, setProviderName] = useState('')
  const [modelName, setModelName] = useState('')
  const [isTwoColumn, setIsTwoColumn] = useState(false)
  const [status, setStatus] = useState<Status>('idle')
  const [jobId, setJobId] = useState<string | null>(null)
  const [progress, setProgress] = useState(0)
  const [total, setTotal] = useState(0)
  const [error, setError] = useState('')
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null)

  useEffect(() => {
    return () => {
      if (pollRef.current) clearInterval(pollRef.current)
    }
  }, [])

  // Reset isTwoColumn when file changes
  useEffect(() => {
    setIsTwoColumn(false)
  }, [file])

  const { app: msg } = messages

  const handleSubmit = useCallback(async () => {
    if (!file) return
    setStatus('uploading')
    setError('')

    const formData = new FormData()
    formData.append('file', file)
    if (isTwoColumn) {
      formData.append('transform2cell', 'true')
    } else {
      formData.append('lang', lang)
      formData.append('mode', mode)
    }
    if (providerName) formData.append('provider', providerName)
    if (modelName) formData.append('model', modelName)

    try {
      const resp = await fetch('/api/translate', { method: 'POST', body: formData })
      if (!resp.ok) throw new Error(`${msg.uploadFailed}: ${resp.statusText}`)
      const data = await resp.json()
      setJobId(data.job_id)
      setStatus('translating')

      pollRef.current = setInterval(async () => {
        const sResp = await fetch(`/api/translate/${data.job_id}/status`)
        if (!sResp.ok) {
          if (pollRef.current) clearInterval(pollRef.current)
          pollRef.current = null
          setStatus('error')
          setError(msg.fetchStatusFailed)
          return
        }
        const sData = await sResp.json()
        setProgress(sData.progress)
        setTotal(sData.total)

        if (sData.status === 'done') {
          if (pollRef.current) clearInterval(pollRef.current)
          pollRef.current = null
          setStatus('done')
        } else if (sData.status === 'failed') {
          if (pollRef.current) clearInterval(pollRef.current)
          pollRef.current = null
          setStatus('error')
          setError(sData.error || msg.translationFailed)
        }
      }, 500)
    } catch (e) {
      setStatus('error')
      setError(e instanceof Error ? e.message : msg.errorOccurred)
    }
  }, [file, lang, mode, providerName, modelName, isTwoColumn, msg])

  const handleReset = () => {
    setFile(null)
    setJobId(null)
    setStatus('idle')
    setProgress(0)
    setTotal(0)
    setError('')
  }

  return (
    <div className="min-h-screen bg-brand-50 flex items-start justify-center p-4 sm:p-8">
      <div className="w-full max-w-xl bg-white rounded-2xl shadow-lg p-6 sm:p-8 space-y-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <img src="/logo.png" alt="Emplawra" className="h-8 w-auto" />
            <h1 className="text-2xl font-bold text-brand-500">{msg.title}</h1>
          </div>
          <button
            onClick={() => setLocale(locale === 'en' ? 'ro' : 'en')}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-brand-200 text-sm font-medium text-brand-500 hover:bg-brand-100 transition-colors"
          >
            <span className={locale === 'en' ? 'text-brand-500' : 'text-brand-300'}>EN</span>
            <span className="text-brand-200">/</span>
            <span className={locale === 'ro' ? 'text-brand-500' : 'text-brand-300'}>RO</span>
          </button>
        </div>

        {status === 'idle' || status === 'uploading' ? (
          <>
            <DropZone file={file} onFile={setFile} disabled={status === 'uploading'} onTwoColumnDetected={setIsTwoColumn} />
            <OptionsForm
              lang={lang}
              mode={mode}
              providerName={providerName}
              modelName={modelName}
              onLangChange={setLang}
              onModeChange={setMode}
              onProviderChange={setProviderName}
              onModelChange={setModelName}
              disabled={isTwoColumn}
            />
            <button
              onClick={handleSubmit}
              disabled={!file || status === 'uploading'}
              className="w-full py-3 px-4 bg-brand-500 text-white font-medium rounded-xl disabled:opacity-40 disabled:cursor-not-allowed hover:bg-brand-400 transition-colors"
            >
              {status === 'uploading' ? msg.uploading : (isTwoColumn ? msg.convertToCellColumns : msg.convert)}
            </button>
          </>
        ) : status === 'translating' ? (
          <ProgressBar progress={progress} total={total} />
        ) : status === 'done' && jobId ? (
          <DownloadLink jobId={jobId} onReset={handleReset} />
        ) : status === 'error' ? (
          <div className="space-y-4">
            <div className="bg-red-50 border border-red-200 text-red-700 rounded-xl p-4">
              {error || msg.errorOccurred}
            </div>
            <button
              onClick={handleReset}
              className="w-full py-3 px-4 bg-brand-300 text-white font-medium rounded-xl hover:bg-brand-400 transition-colors"
            >
              {msg.tryAgain}
            </button>
          </div>
        ) : null}
      </div>
    </div>
  )
}

export default App
