import React from 'react'
import { createRoot } from 'react-dom/client'
import { registerValidationNotice } from '@gmi/excel-core'
import { registerPluginNotice } from '@gmi/excel-view'
import { showNotice } from '@gmi/excel-react'
import '@gmi/excel-react/theme.css'
import { App } from './app/App'

registerValidationNotice(showNotice)
registerPluginNotice(showNotice)

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
