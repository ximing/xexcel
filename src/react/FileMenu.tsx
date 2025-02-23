// src/react/FileMenu.tsx
// 「文件」下拉：打开 CSV（新建 sheet 导入，可撤销）/ 打开 xlsx（整簿替换，不可撤销）/ 导出 CSV / 导出 xlsx / 清除浏览器存档。
import { FileDown, FilePlus2, FileSpreadsheet, FileText, FolderOpen, Trash2, type LucideIcon } from 'lucide-react'
import { createStateFromWorkbook } from '../app/demo'
import { downloadBlob, pickFile, readFileArrayBuffer, readFileText } from '../app/fileio'
import { showNotice } from '../app/notice'
import { workbookStorage } from '../app/storage'
import { csvToGrid, sheetToCSV } from '../core/io/csv'
import { parseXlsx, workbookToExcelJS } from '../core/io/xlsx'
import type { Workbook } from '../core/model'
import type { EditorView } from '../view/editorview'
import { buildImportTr } from './csvImport'
import {
  CSV_MAX_BYTES,
  CSV_MAX_ROWS,
  XLSX_MAX_BYTES,
  csvBaseName,
  FileMenuId,
  fileMenuItems,
  isGridEmpty,
} from './filemenu'
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
  const openCsv = async (): Promise<void> => {
    const file = await pickFile('.csv,text/csv')
    if (!file) return
    if (
      file.size > CSV_MAX_BYTES &&
      !(await askConfirm({ title: '导入 CSV', body: `文件 ${Math.round(file.size / 1024 / 1024)}MB，较大，确定导入？` }))
    )
      return
    let grid
    try {
      grid = csvToGrid(await readFileText(file))
    } catch {
      showNotice('文件读取失败')
      return
    }
    if (isGridEmpty(grid)) {
      showNotice('CSV 文件为空')
      return
    }
    if (
      grid.length > CSV_MAX_ROWS &&
      !(await askConfirm({ title: '导入 CSV', body: `共 ${grid.length} 行，较多，确定导入？` }))
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
        title: '打开 xlsx',
        body: '将替换当前表格内容，浏览器存档将被覆盖。',
        danger: true,
        confirmLabel: '打开',
      }))
    )
      return
    const file = await pickFile('.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if (!file) return
    if (
      file.size > XLSX_MAX_BYTES &&
      !(await askConfirm({ title: '打开 xlsx', body: `文件 ${Math.round(file.size / 1024 / 1024)}MB，较大，确定导入？` }))
    )
      return
    // 先暂停自动保存：导入期间（含 updateState 触发的订阅）不写存档；不清存档，失败时现场与存档都不动
    workbookStorage.pause()
    let opened: Workbook | null = null
    try {
      const wb = await parseXlsx(new Uint8Array(await readFileArrayBuffer(file)))
      view.updateState(createStateFromWorkbook(wb)) // 同启动恢复路径，不可撤销
      opened = wb
      showNotice(`已打开 ${file.name}`)
    } catch {
      showNotice('文件无法解析')
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
      showNotice('导出失败')
    }
    view.focus()
  }

  const clearStorage = async (): Promise<void> => {
    if (
      !(await askConfirm({
        title: '清除浏览器存档',
        body: '当前内容刷新后将不再恢复。',
        danger: true,
        confirmLabel: '清除',
      }))
    )
      return
    workbookStorage.suspend()
    showNotice('已清除浏览器存档，刷新后生效')
    view.focus()
  }

  const run = (id: FileMenuId): void => {
    if (id === 'openCsv') void openCsv()
    else if (id === 'openXlsx') void openXlsx()
    else if (id === 'exportCsv') exportCsv()
    else if (id === 'exportXlsx') void exportXlsx()
    else void clearStorage()
  }

  const entries: MenuEntry[] = fileMenuItems().flatMap((it): MenuEntry[] => {
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
      trigger={(open, toggle) => <IconButton icon={FileSpreadsheet} tip="文件" active={open} onClick={toggle} />}
      entries={entries}
    />
  )
}
