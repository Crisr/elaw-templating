import { createContext, useContext, useState, type ReactNode } from 'react'
import { getMessages, type Messages, type Locale } from './messages'

interface LocaleContextValue {
  locale: Locale
  setLocale: (l: Locale) => void
  messages: Messages
}

const LocaleContext = createContext<LocaleContextValue | null>(null)

export function LocaleProvider({ children }: { children: ReactNode }) {
  const [locale, setLocale] = useState<Locale>('en')
  const messages = getMessages(locale)

  return (
    <LocaleContext.Provider value={{ locale, setLocale, messages }}>
      {children}
    </LocaleContext.Provider>
  )
}

export function useLocale(): LocaleContextValue {
  const ctx = useContext(LocaleContext)
  if (!ctx) throw new Error('useLocale must be used within LocaleProvider')
  return ctx
}
