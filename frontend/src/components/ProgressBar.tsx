interface Props {
  progress: number
  total: number
}

export default function ProgressBar({ progress, total }: Props) {
  const pct = total > 0 ? Math.round((progress / total) * 100) : 0

  return (
    <div className="space-y-3">
      <p className="text-sm text-gray-600 text-center">
        Translating... {progress} / {total} chunks ({pct}%)
      </p>
      <div className="w-full bg-gray-200 rounded-full h-4 overflow-hidden">
        <div
          className="h-full bg-indigo-600 rounded-full transition-all duration-300 ease-out"
          style={{ width: `${pct}%` }}
        />
      </div>
    </div>
  )
}
