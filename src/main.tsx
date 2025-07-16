import React from 'react'
import { createRoot } from 'react-dom/client'
import { App } from './app/App'
import { registerValidationNotice } from '@gmi/excel-core'
import { registerPluginNotice } from '@gmi/excel-view'
import { showNotice } from './app/notice'
import './app/theme.css'

registerValidationNotice(showNotice)
registerPluginNotice(showNotice)

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
