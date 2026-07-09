import { useState, useEffect } from 'react'

interface Props {
  lang: string
  mode: string
  providerName: string
  modelName: string
  onLangChange: (v: string) => void
  onModeChange: (v: string) => void
  onProviderChange: (v: string) => void
  onModelChange: (v: string) => void
}

interface Provider {
  name: string
  model: string
}

export default function OptionsForm({
  lang, mode, providerName, modelName,
  onLangChange, onModeChange, onProviderChange, onModelChange,
}: Props) {
  const [providers, setProviders] = useState<Provider[]>([])
  const [defaultProvider, setDefaultProvider] = useState('')

  useEffect(() => {
    fetch('/api/providers')
      .then((r) => r.json())
      .then((data) => {
        setProviders(data.providers || [])
        setDefaultProvider(data.default_provider || '')
        if (!providerName && data.default_provider) {
          onProviderChange(data.default_provider)
        }
      })
      .catch(() => {})
  }, [])

  const selectedProvider = providers.find((p) => p.name === (providerName || defaultProvider))

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Language</label>
          <select
            value={lang}
            onChange={(e) => onLangChange(e.target.value)}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          >
            <option value="ro">Romanian</option>
            <option value="en">English</option>
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Mode</label>
          <select
            value={mode}
            onChange={(e) => onModeChange(e.target.value)}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          >
            <option value="inline">Inline</option>
            <option value="side-by-side">Side-by-Side</option>
          </select>
        </div>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Provider</label>
          <select
            value={providerName || defaultProvider}
            onChange={(e) => {
              onProviderChange(e.target.value)
              const p = providers.find((pr) => pr.name === e.target.value)
              if (p) onModelChange(p.model)
            }}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          >
            {providers.map((p) => (
              <option key={p.name} value={p.name}>{p.name}</option>
            ))}
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Model</label>
          <input
            type="text"
            value={modelName || selectedProvider?.model || ''}
            onChange={(e) => onModelChange(e.target.value)}
            placeholder={selectedProvider?.model || 'Model name'}
            className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
          />
        </div>
      </div>
    </div>
  )
}
