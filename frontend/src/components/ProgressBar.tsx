import { useLocale } from '../LocaleContext'

interface Props {
  progress: number
  total: number
}

export default function ProgressBar({ progress, total }: Props) {
  const pct = total > 0 ? Math.round((progress / total) * 100) : 0
  const { messages } = useLocale()
  const label = messages.progressBar.translating
    .replace('{progress}', String(progress))
    .replace('{total}', String(total))
    .replace('{pct}', String(pct))

  return (
    <div className="space-y-3">
      <p className="text-sm text-brand-400 text-center">{label}</p>
      <div className="w-full bg-brand-100 rounded-full h-4 overflow-hidden">
        <div
          className="h-full bg-brand-500 rounded-full transition-all duration-300 ease-out"
          style={{ width: `${pct}%` }}
        />
      </div>
    </div>
  )
}
