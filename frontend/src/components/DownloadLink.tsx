interface Props {
  jobId: string
  onReset: () => void
}

export default function DownloadLink({ jobId, onReset }: Props) {
  return (
    <div className="text-center space-y-4">
      <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl p-4">
        Translation complete!
      </div>
      <a
        href={`/api/translate/${jobId}/download`}
        className="inline-block w-full py-3 px-4 bg-indigo-600 text-white font-medium rounded-xl text-center hover:bg-indigo-700 transition-colors"
      >
        Download Translated File
      </a>
      <button
        onClick={onReset}
        className="w-full py-2 px-4 text-gray-600 hover:text-gray-800 text-sm transition-colors"
      >
        Translate Another
      </button>
    </div>
  )
}
