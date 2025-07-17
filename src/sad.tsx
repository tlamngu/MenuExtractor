import React from 'react'
import { ExcelImporter } from './utils/ExcelImporter'
import type { SheetJson } from './utils/excelParser'

function App() {
  const handleData = (sheets: SheetJson[]) => {
    console.log('Parsed sheets:', sheets)
    // send to your backend or store in state...
  }

  return (
    <div>
      <h1>Upload Flight Menus</h1>
      <ExcelImporter onData={handleData} />
    </div>
  )
}

export default App