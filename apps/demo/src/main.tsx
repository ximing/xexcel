import React from 'react'
import { createRoot } from 'react-dom/client'
import { registerValidationNotice } from '@xexcel/core'
import { registerPluginNotice } from '@xexcel/view'
import { showNotice } from '@xexcel/react'
import '@xexcel/react/theme.css'
import { App } from './App'

registerValidationNotice(showNotice)
registerPluginNotice(showNotice)

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
