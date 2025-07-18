/**

Converted by VerseAI Toolkit
Source language: Python (Google Colab)
Target language: Typescript
Version: 1.0.2

Note: File dữ liệu nhận số 2

**/
import * as xlsx from "xlsx";
/**
 * Represents a processed food item with its associated metadata.
 */
interface FoodItem {
  UpliftRatio: string | number | null;
  Name: string | null;
  Qty: string | number | null;
  Remark: string | null;
  Class: string | null;
  MenuID: string;
  Cycle: string | number | null;
}

// Define a type for the raw sheet data for clarity.
type RawRow = any[];
type RawSheetData = RawRow[];

/**
 * Checks if a cell's value is effectively empty (null, undefined, or an empty string).
 * @param cell The cell value to check.
 * @returns True if the cell is empty, false otherwise.
 */
const isCellEmpty = (cell: any): boolean => {
  return cell === null || cell === undefined || String(cell).trim() === '';
};

/**
 * Processes a worksheet array to extract structured data based on specific keywords and layout,
 * and returns a flattened list of food items.
 *
 * @param sheetData The worksheet data, expected as an array of arrays (from xlsx.js: sheet_to_json(ws, { header: 1 })).
 * @returns A flattened list of FoodItem objects.
 */
export function procDfDln2(sheet: xlsx.Sheet): FoodItem[] {
  let sheetData: any[][] = xlsx.utils.sheet_to_json(sheet, {
      header: 1,
      defval: null,
    });
  if (!sheetData || sheetData.length === 0) {
    console.error("Error: Input sheetData is empty or invalid.");
    return [];
  }

  // 1. Find the boundaries of the main data table
  let startRow = -1;
  let endRow = -1;
  const kw_start = "Uplift Ratio".trim().toLowerCase();
  const kw_end = "Ghi chú".trim().toLowerCase();
  let emptyRowCount = 0;

  const isRowEmpty = (row: RawRow): boolean => {
    // A row is considered empty if all its cells are empty.
    return row.every(isCellEmpty);
  };

  for (let i = 0; i < sheetData.length; i++) {
    const row = sheetData[i];
    if (!row || row.length === 0) {
      if (startRow !== -1) emptyRowCount++; // Count empty rows only after table starts
      continue;
    }

    const firstCell = String(row[0] ?? '').trim().toLowerCase();

    if (startRow === -1 && firstCell.includes(kw_start)) {
      // The table officially starts one row *before* the header row with "Uplift Ratio"
      // because a MENUID might be on that row.
      startRow = i - 1;
    } else if (startRow !== -1) {
      if (firstCell.includes(kw_end)) {
        endRow = i;
        break;
      }
      if (isRowEmpty(row)) {
        emptyRowCount++;
        if (emptyRowCount >= 2) {
          // The end is the first of the two consecutive empty rows
          endRow = i - 1; 
          break;
        }
      } else {
        // Reset counter if a non-empty row is found
        emptyRowCount = 0;
      }
    }
  }

  // If end was not found by keywords, assume it's the end of the sheet.
  if (startRow !== -1 && endRow === -1) {
    endRow = sheetData.length;
  }
  
  if (startRow === -1 || endRow === -1) {
    console.error("Error: Could not find start ('Uplift Ratio') or end ('Ghi chú'/empty rows) keywords.");
    return [];
  }

  // 2. Extract and clean the sub-table
  const subTable = sheetData.slice(startRow, endRow);

  // The header is the second row of our initial slice (the one with "Uplift Ratio")
  const headerRow = subTable[1];
  if (!headerRow) {
    console.error("Error: Header row not found in the identified table boundaries.");
    return [];
  }

  // Create a map from header name to column index for easy data access
  const headerMap: { [key: string]: number } = {};
  headerRow.forEach((header, index) => {
    if (typeof header === 'string' && header.trim() !== '') {
      headerMap[header.trim()] = index;
    }
  });

  // 3. Segment the DataFrame using a state machine approach and flatten
  const foodItemsList: FoodItem[] = [];

  // State variables that will be updated as we iterate through rows
  let currentMenuID: string = "%STANDALONE"; // Default for items before the first MENUID
  let currentCycle: string | number | null = null;
  let currentClass: string | null = null;

  const kw_cycle = "cycle";
  const kw_menu = "menu";

  for (const row of subTable) {
    // Helper conditions to identify row type
    const isMenuIdHeader = !isCellEmpty(row[0]) &&
                           !isCellEmpty(row[2]) &&
                           isCellEmpty(row[1]) &&
                           String(row[0]).toLowerCase().includes(kw_menu) &&
                           String(row[2]).toLowerCase().includes(kw_cycle);

    const isClassHeader = !isCellEmpty(row[0]) &&
                          isCellEmpty(row[1]) &&
                          row.slice(2).every(isCellEmpty);
    
    // A data row must have a value in the "Uplift Ratio" column (or first column)
    const isDataRow = !isCellEmpty(row[headerMap['Uplift Ratio'] ?? 0]);

    if (isMenuIdHeader) {
      currentMenuID = String(row[0]);
      // Cycle value is in the 4th column (index 3)
      currentCycle = row[3] ?? null;
      currentClass = null; // Reset class when a new menu starts
      continue; // This row is a header, not data
    }

    if (isClassHeader) {
      currentClass = String(row[0]);
      continue; // This row is a header, not data
    }
    
    // If it's a data row, create the FoodItem object
    if(isDataRow) {
      const foodItem: FoodItem = {
        UpliftRatio: row[headerMap['Uplift Ratio']] ?? null,
        Name: row[headerMap['Component Description']] ?? null,
        Qty: row[headerMap['Qty']] ?? null,
        Remark: row[headerMap['Remark']] ?? null,
        Class: currentClass,
        MenuID: currentMenuID,
        Cycle: currentCycle,
      };
      foodItemsList.push(foodItem);
    }
  }
  foodItemsList.shift()
  return  foodItemsList;
}