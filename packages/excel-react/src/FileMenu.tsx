// src/react/FileMenu.tsx
// 「文件」下拉：打开 CSV（新建 sheet 导入，可撤销）/ 打开 xlsx（整簿替换，不可撤销）/ 导出 CSV / 导出 xlsx / 清除浏览器存档。
import { FileDown, FilePlus2, FileSpreadsheet, FileText, FolderOpen, Trash2, type LucideIcon } from 'lucide-react'
import { createStateFromWorkbook } from './createState'
import { downloadBlob, pickFile, readFileArrayBuffer, readFileText } from './fileio'
import { showNotice } from './notice'
import { workbookStorage } from './storage'
import { csvToGrid, sheetToCSV } from '@xexcel/core'
import { parseXlsx, workbookToExcelJS } from '@xexcel/core'
import type { Workbook } from '@xexcel/core'
import type { EditorView } from '@xexcel/view'
import { buildImportTr } from './csvImport'
import {
  CSV_MAX_BYTES,
  CSV_MAX_ROWS,
  XLSX_MAX_BYTES,
  csvBaseName,
  FileMenuId,
  fileMenuItems,
  isGridEmpty,
} from './fileMenuCore'
import { t, useLocale } from './i18n'
import { askConfirm } from './ui/confirmStore'
import { Dropdown } from './ui/Dropdown'
import { IconButton } from './ui/IconButton'
import type { MenuEntry } from './ui/Menu'

interface Props {
  view: EditorView
}

const ITEM_ICON: Record<FileMenuId, LucideIcon> = {
  openCsv: FilePlus2,
  openXlsx: FolderOpen,
  exportCsv: FileText,
  exportXlsx: FileDown,
  clearStorage: Trash2,
}

export function FileMenu({ view }: Props) {
  const locale = useLocale()
  const openCsv = async (): Promise<void> => {
    const file = await pickFile('.csv,text/csv')
    if (!file) return
    if (
      file.size > CSV_MAX_BYTES &&
      !(await askConfirm({ title: t(locale, 'file.importCsvTitle'), body: t(locale, 'file.sizeWarn', { size: Math.round(file.size / 1024 / 1024) }) }))
    )
      return
    let grid
    try {
      grid = csvToGrid(await readFileText(file))
    } catch {
      showNotice(t(locale, 'file.readFail'))
      return
    }
    if (isGridEmpty(grid)) {
      showNotice(t(locale, 'file.csvEmpty'))
      return
    }
    if (
      grid.length > CSV_MAX_ROWS &&
      !(await askConfirm({ title: t(locale, 'file.importCsvTitle'), body: t(locale, 'file.csvRowsWarn', { n: grid.length }) }))
    )
      return
    view.dispatch(buildImportTr(view.state, grid, csvBaseName(file.name)))
    view.focus()
  }

  const exportCsv = (): void => {
    const name = view.state.doc.names.get(view.state.doc.active) ?? 'Sheet1'
    const csv = sheetToCSV(view.state.activeSheet)
    downloadBlob(`${name}.csv`, new Blob([csv], { type: 'text/csv;charset=utf-8' }))
    view.focus()
  }

  const openXlsx = async (): Promise<void> => {
    if (
      !(await askConfirm({
        title: t(locale, 'file.openXlsxTitle'),
        body: t(locale, 'file.openXlsxBody'),
        danger: true,
        confirmLabel: t(locale, 'file.openXlsxConfirm'),
      }))
    )
      return
    const file = await pickFile('.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if (!file) return
    if (
      file.size > XLSX_MAX_BYTES &&
      !(await askConfirm({ title: t(locale, 'file.openXlsxTitle'), body: t(locale, 'file.sizeWarn', { size: Math.round(file.size / 1024 / 1024) }) }))
    )
      return
    // 先暂停自动保存：导入期间（含 updateState 触发的订阅）不写存档；不清存档，失败时现场与存档都不动
    workbookStorage.pause()
    let opened: Workbook | null = null
    try {
      const wb = await parseXlsx(new Uint8Array(await readFileArrayBuffer(file)))
      view.updateState(createStateFromWorkbook(wb)) // 同启动恢复路径，不可撤销
      opened = wb
      showNotice(t(locale, 'file.opened', { name: file.name }))
    } catch {
      showNotice(t(locale, 'file.parseFail'))
    } finally {
      workbookStorage.resume()
    }
    // 落档均须在 resume 之后（suspended 时 saveNow 是 no-op）：
    // 成功 → 立即写入新 workbook；失败 → 补写当前 doc（pause 清掉的 pending 防抖可能含最后 <1s 的在途编辑）
    if (opened) workbookStorage.saveNow(opened)
    else workbookStorage.saveNow(view.state.doc)
    view.focus()
  }

  const exportXlsx = async (): Promise<void> => {
    try {
      const ewb = workbookToExcelJS(view.state.doc)
      const buf = await ewb.xlsx.writeBuffer()
      const name = view.state.doc.names.get(view.state.doc.active) ?? 'workbook'
      downloadBlob(
        `${name}.xlsx`,
        new Blob([buf as BlobPart], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
      )
    } catch {
      showNotice(t(locale, 'file.exportFail'))
    }
    view.focus()
  }

  const clearStorage = async (): Promise<void> => {
    if (
      !(await askConfirm({
        title: t(locale, 'file.clearTitle'),
        body: t(locale, 'file.clearBody'),
        danger: true,
        confirmLabel: t(locale, 'file.clearConfirm'),
      }))
    )
      return
    workbookStorage.suspend()
    showNotice(t(locale, 'file.clearDone'))
    view.focus()
  }

  const run = (id: FileMenuId): void => {
    if (id === 'openCsv') void openCsv()
    else if (id === 'openXlsx') void openXlsx()
    else if (id === 'exportCsv') exportCsv()
    else if (id === 'exportXlsx') void exportXlsx()
    else void clearStorage()
  }

  const entries: MenuEntry[] = fileMenuItems(locale).flatMap((it): MenuEntry[] => {
    const entry: MenuEntry = {
      id: it.id,
      label: it.label,
      icon: ITEM_ICON[it.id],
      danger: it.danger,
      onSelect: () => run(it.id),
    }
    // 危险项（清除存档）前分隔
    return it.danger ? [{ sep: true }, entry] : [entry]
  })

  return (
    <Dropdown
      trigger={(open, toggle) => <IconButton icon={FileSpreadsheet} tip={t(locale, 'file.tip')} active={open} onClick={toggle} />}
      entries={entries}
    />
  )
}
