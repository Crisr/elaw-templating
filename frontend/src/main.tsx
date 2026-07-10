import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import { LocaleProvider } from './LocaleContext'
import App from './App'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <LocaleProvider>
      <App />
    </LocaleProvider>
  </StrictMode>,
)
