export interface Messages {
  app: {
    title: string
    convert: string
    convertToCellColumns: string
    uploading: string
    tryAgain: string
    errorOccurred: string
    uploadFailed: string
    fetchStatusFailed: string
    translationFailed: string
  }
  dropZone: {
    dragDrop: string
    clickToBrowse: string
  }
  optionsForm: {
    translateTo: string
    romanian: string
    english: string
    none: string
    mode: string
    inline: string
    sideBySide: string
    provider: string
    model: string
    modelName: string
  }
  progressBar: {
    translating: string
  }
  downloadLink: {
    complete: string
    download: string
    translateAnother: string
  }
}

export const en: Messages = {
  app: {
    title: 'DOCX Translator',
    convert: 'Convert',
    convertToCellColumns: 'Convert to cell columns',
    uploading: 'Uploading...',
    tryAgain: 'Try Again',
    errorOccurred: 'An error occurred',
    uploadFailed: 'Upload failed',
    fetchStatusFailed: 'Failed to fetch status',
    translationFailed: 'Translation failed',
  },
  dropZone: {
    dragDrop: 'Drag & drop a .docx file here',
    clickToBrowse: 'or click to browse',
  },
  optionsForm: {
    translateTo: 'Translate to',
    romanian: 'Romanian',
    english: 'English',
    none: 'None (same language)',
    mode: 'Mode',
    inline: 'Inline',
    sideBySide: 'Side-by-Side',
    provider: 'Provider',
    model: 'Model',
    modelName: 'Model name',
  },
  progressBar: {
    translating: 'Translating... {progress} / {total} chunks ({pct}%)',
  },
  downloadLink: {
    complete: 'Translation complete!',
    download: 'Download Translated File',
    translateAnother: 'Translate Another',
  },
}

export const ro: Messages = {
  app: {
    title: 'Traducător DOCX',
    convert: 'Convertește',
    convertToCellColumns: 'Convertește în coloane',
    uploading: 'Se încarcă...',
    tryAgain: 'Încercați din nou',
    errorOccurred: 'A apărut o eroare',
    uploadFailed: 'Încărcarea a eșuat',
    fetchStatusFailed: 'Nu s-a putut obține starea',
    translationFailed: 'Traducerea a eșuat',
  },
  dropZone: {
    dragDrop: 'Trageți un fișier .docx aici',
    clickToBrowse: 'sau faceți clic pentru a răsfoi',
  },
  optionsForm: {
    translateTo: 'Traduceți în',
    romanian: 'Română',
    english: 'Engleză',
    none: 'Niciuna (aceeași limbă)',
    mode: 'Mod',
    inline: 'Direct',
    sideBySide: 'Față în față',
    provider: 'Furnizor',
    model: 'Model',
    modelName: 'Numele modelului',
  },
  progressBar: {
    translating: 'Se traduce... {progress} / {total} bucăți ({pct}%)',
  },
  downloadLink: {
    complete: 'Traducere finalizată!',
    download: 'Descărcați fișierul tradus',
    translateAnother: 'Traduceți altul',
  },
}

export type Locale = 'en' | 'ro'

export function getMessages(locale: Locale): Messages {
  return locale === 'ro' ? ro : en
}
