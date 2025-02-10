import React from 'react'
import { createRoot } from 'react-dom/client'
import { App } from './app/App'
import { registerValidationNotice } from './core/validation'
import { showNotice } from './app/notice'
import './app/theme.css'
import './app/style.css'

registerValidationNotice(showNotice)

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
