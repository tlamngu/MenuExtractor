import * as XLSX from 'xlsx'

/** Core types for your menu JSON */
export interface Item {
  upliftRatio:      string
  componentDescription: string
  qty?:             string | null
  remark?:          string | null
}

export interface Menu {
  menuCode?:        string
  cycle?:           number
  applicableFor?:   string
  items:            Item[]
}

export interface FareClass {
  className:        string
  menus:            Menu[]
}

export interface BeverageGroup {
  groupName:        string
  items:            Item[]
}

export interface SheetJson {
  sheetName:        string
  metadata:         Record<string,string | null>
  fareClasses:      FareClass[]
  beverageGroups:   BeverageGroup[]
  notes:            string[]
}

/**
 * Reads an Excel file and returns structured JSON for each sheet.
 */
export async function parseExcelFile(file: File): Promise<SheetJson[]> {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array' })

  return workbook.SheetNames.map(sheetName => {
    const sheet = workbook.Sheets[sheetName]
    // get rows as array of arrays
    const raw: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false
    })
    console.log(raw)
    return transformSheetToJson(sheetName, raw)
  })
}

/**
 * Transform raw row data into your JSON structure.
 * You may need to tweak parsing logic to match your exact Excel layout.
 */
function transformSheetToJson(
  sheetName: string,
  rows: any[][]
): SheetJson {
  // 1. Extract metadata from top rows (key : value)
  const metadata: Record<string,string|null> = {}
  let rowIdx = 0
  while (rows[rowIdx] && rows[rowIdx].some(cell => typeof cell === 'string' && cell.includes(':'))) {
    const [key, val] = rows[rowIdx].join('').split(':').map(s => s.trim())
    metadata[key] = val || null
    rowIdx++
  }

  // 2. Skip to header row ( ["Uplift Ratio", "Component Description", "Qty", "Remark"] )
  while (rows[rowIdx] && !rows[rowIdx][0]?.toString().startsWith('Uplift Ratio')) {
    rowIdx++
  }
  const headerRow = rows[rowIdx++] || []

  // 3. Parse the body into fareClasses / beverageGroups
  const fareClasses: FareClass[] = []
  const beverageGroups: BeverageGroup[] = []
  let currentClass: FareClass | null = null
  let currentMenu: Menu | null = null

  for (; rowIdx < rows.length; rowIdx++) {
    const [col1, col2, col3, col4] = rows[rowIdx].map(c => c?.toString().trim() || '')

    // detect a new fare‐class header
    if (/^[A-Z0-9 &\+]+$/.test(col1) && !col2) {
      currentClass = { className: col1, menus: [] }
      fareClasses.push(currentClass)
      currentMenu = null
      continue
    }

    // detect a new menu line within a class (e.g. "Menu A", "CYCLE 5")
    if (col2?.startsWith('Menu') || col2?.startsWith('STD')) {
      const [menuCode, maybeCycle] = col2.split(';').map(s => s.trim())
      const cycleMatch = maybeCycle?.match(/CYCLE\s*(\d+)/i)
      currentMenu = {
        menuCode: menuCode,
        cycle: cycleMatch ? parseInt(cycleMatch[1], 10) : undefined,
        items: []
      }
      currentClass?.menus.push(currentMenu)
      continue
    }

    // if still inside a class & menu, push item
    const item: Item = {
      upliftRatio: col1,
      componentDescription: col2,
      qty:           col3 || null,
      remark:        col4 || null
    }
    if (currentMenu) {
      currentMenu.items.push(item)
      continue
    }

    // else treat as beverage group
    if (col1 && !currentMenu) {
      const group = beverageGroups.find(g => g.groupName === col1)
                   || (() => {
                         const g = { groupName: col1, items: [] }
                         beverageGroups.push(g)
                         return g
                       })()
      group.items.push(item)
      continue
    }
  }

  // 4. Tail notes (anything not parsed)
  const notes: string[] = []
  // (you can append leftover rows as notes)
  return { sheetName, metadata, fareClasses, beverageGroups, notes }
}