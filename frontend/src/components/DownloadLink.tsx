import { messages } from '../messages'

interface Props {
  jobId: string
  onReset: () => void
}

export default function DownloadLink({ jobId, onReset }: Props) {
  const { downloadLink: msg } = messages

  return (
    <div className="text-center space-y-4">
      <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl p-4">
        {msg.complete}
      </div>
      <a
        href={`/api/translate/${jobId}/download`}
        className="inline-block w-full py-3 px-4 bg-brand-500 text-white font-medium rounded-xl text-center hover:bg-brand-400 transition-colors"
      >
        {msg.download}
      </a>
      <button
        onClick={onReset}
        className="w-full py-2 px-4 text-brand-300 hover:text-brand-500 text-sm transition-colors"
      >
        {msg.translateAnother}
      </button>
    </div>
  )
}
