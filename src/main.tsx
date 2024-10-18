import React from 'react'
import { createRoot } from 'react-dom/client'
import { App } from './app/App'
import { setupExcelJS } from './app/exceljs'
import './app/style.css'

setupExcelJS()

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
