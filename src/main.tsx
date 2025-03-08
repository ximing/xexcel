import React from 'react'
import { createRoot } from 'react-dom/client'
import { App } from './app/App'
import { registerValidationNotice } from './core/validation'
import { registerPluginNotice } from './plugins/notify'
import { showNotice } from './app/notice'
import './app/theme.css'

registerValidationNotice(showNotice)
registerPluginNotice(showNotice)

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
