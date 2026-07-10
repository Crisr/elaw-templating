import { useState, useEffect } from 'react'
import { useLocale } from '../LocaleContext'

interface Props {
  lang: string
  mode: string
  providerName: string
  modelName: string
  onLangChange: (v: string) => void
  onModeChange: (v: string) => void
  onProviderChange: (v: string) => void
  onModelChange: (v: string) => void
  disabled?: boolean
}

interface Provider {
  name: string
  model: string
  models?: string[]
}

export default function OptionsForm({
  lang, mode, providerName, modelName,
  onLangChange, onModeChange, onProviderChange, onModelChange,
  disabled = false,
}: Props) {
  const [providers, setProviders] = useState<Provider[]>([])
  const [defaultProvider, setDefaultProvider] = useState('')
  const { messages } = useLocale()
  const msg = messages.optionsForm

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
      .catch((err) => console.error('Failed to load providers:', err))
  }, [])

  const selectedProvider = providers.find((p) => p.name === (providerName || defaultProvider))
  const availableModels = selectedProvider?.models

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-brand-500 mb-1">{msg.translateTo}</label>
          <select
            value={lang}
            onChange={(e) => onLangChange(e.target.value)}
            disabled={disabled}
            className={`w-full border border-brand-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-brand-300 focus:border-brand-300 text-brand-500 ${
              disabled ? 'opacity-50 cursor-not-allowed' : ''
            }`}
          >
            <option value="ro">{msg.romanian}</option>
            <option value="en">{msg.english}</option>
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-brand-500 mb-1">{msg.mode}</label>
          <select
            value={mode}
            onChange={(e) => onModeChange(e.target.value)}
            disabled={disabled}
            className={`w-full border border-brand-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-brand-300 focus:border-brand-300 text-brand-500 ${
              disabled ? 'opacity-50 cursor-not-allowed' : ''
            }`}
          >
            <option value="inline">{msg.inline}</option>
            <option value="side-by-side">{msg.sideBySide}</option>
          </select>
        </div>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-brand-500 mb-1">{msg.provider}</label>
          <select
            value={providerName || defaultProvider}
            onChange={(e) => {
              onProviderChange(e.target.value)
              const p = providers.find((pr) => pr.name === e.target.value)
              if (p) onModelChange(p.model)
            }}
            className="w-full border border-brand-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-brand-300 focus:border-brand-300 text-brand-500"
          >
            {providers.map((p) => (
              <option key={p.name} value={p.name}>{p.name}</option>
            ))}
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-brand-500 mb-1">{msg.model}</label>
          {availableModels ? (
            <select
              value={modelName || selectedProvider?.model || ''}
              onChange={(e) => onModelChange(e.target.value)}
              className="w-full border border-brand-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-brand-300 focus:border-brand-300 text-brand-500"
            >
              {availableModels.map((m) => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>
          ) : (
            <input
              type="text"
              value={modelName || selectedProvider?.model || ''}
              onChange={(e) => onModelChange(e.target.value)}
              placeholder={selectedProvider?.model || msg.modelName}
              className="w-full border border-brand-200 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-brand-300 focus:border-brand-300 text-brand-500 placeholder-brand-300"
            />
          )}
        </div>
      </div>
    </div>
  )
}
