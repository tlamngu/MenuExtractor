import React from 'react'
import type { SheetJson } from './excelParser'
import { parseExcelFile } from './excelParser'

export interface Props {
  onData: (data: SheetJson[]) => void
}

export function ExcelImporter({ onData }: Props) {
  const handleChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    try {
      const sheets = await parseExcelFile(file)
      onData(sheets)
    } catch (err) {
      console.error('Failed to parse Excel:', err)
    }
  }

  return (
    <input
      type="file"
      accept=".xls,.xlsx"
      onChange={handleChange}
    />
  )
}