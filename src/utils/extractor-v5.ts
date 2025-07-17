/**

Converted by VerseAI Toolkit
Source language: Python (Google Colab)
Target language: Typescript
Version: 1.0.2

**/
import * as xlsx from 'xlsx';

// Define the structure of the final output object, similar to the Python dictionary
interface FoodItem {
  UpliftRatio: string | number;
  Name: string | null;
  Unit: string | null;
  Qty: string | number | null;
  Remark: string | null;
  Class: string | null;
  MenuID: string | null;
  Cycle: string;
  TimeStart: string | null;
  TimeEnd: string | null;
}

// Equivalent to the Python Checkpoint class
interface Checkpoint {
  x: number; // column index
  y: number; // row index
  typ: string;
}

/**
 * Parses a date range string like "1-7 APR.2025" into ISO 8601 format.
 * This function is a direct port of the Python version, designed to handle
 * variations in spacing, capitalization, and the optional period after the month.
 *
 * @param dateString - A string representing the date range (e.g., "1-7 APR.2025").
 * @returns A tuple containing the start and end dates as ISO 8601 strings
 *   (e.g., ['2025-04-01', '2025-04-07']).
 *   Returns -1 if the string format is invalid or represents an impossible date.
 */
function parseDateRangeToISO(dateString: string): [string, string] | -1 {
  // Dictionary to map three-letter month abbreviations to month numbers (0-indexed for JS Date).
  const monthMap: { [key: string]: number } = {
    'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'MAY': 4, 'JUN': 5,
    'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11
  };

  // Regex to capture the components, equivalent to the Python re.compile.
  // The 'i' flag makes it case-insensitive.
  const pattern = /(\d{1,2})-(\d{1,2})\s*([A-Za-z]{3})\.?\s*(\d{4})/;
  
  // Strip leading/trailing whitespace and try to match.
  const match = dateString.trim().match(pattern);

  if (!match) {
    // If the string doesn't match the expected format, return -1.
    return -1;
  }

  try {
    // Extract the captured groups from the match array.
    // match[0] is the full string, match[1] is the first group, etc.
    const [, startDayStr, endDayStr, monthStr, yearStr] = match;

    // Convert string components to integers.
    const year = parseInt(yearStr, 10);
    const startDay = parseInt(startDayStr, 10);
    const endDay = parseInt(endDayStr, 10);
    
    // Look up the month number from the map (case-insensitively).
    const monthAbbr = monthStr.toUpperCase();
    if (!(monthAbbr in monthMap)) {
      // Invalid month abbreviation.
      return -1;
    }
    const month = monthMap[monthAbbr];

    // Create Date objects for the start and end dates.
    // JavaScript's Date constructor will handle validation. We must check it
    // because it can "roll over" invalid dates (e.g., April 31 becomes May 1).
    const startDate = new Date(Date.UTC(year, month, startDay));
    const endDate = new Date(Date.UTC(year, month, endDay));

    // Validate that the dates were not rolled over.
    if (startDate.getUTCFullYear() !== year || startDate.getUTCMonth() !== month || startDate.getUTCDate() !== startDay) {
        return -1; // Invalid start date
    }
    if (endDate.getUTCFullYear() !== year || endDate.getUTCMonth() !== month || endDate.getUTCDate() !== endDay) {
        return -1; // Invalid end date
    }

    // Return a tuple of the dates formatted as ISO 8601 strings (date part only).
    return [
      startDate.toISOString().split('T')[0],
      endDate.toISOString().split('T')[0]
    ];
  } catch (error) {
    // Catches errors from parseInt or other unexpected issues.
    return -1;
  }
}

/**
 * Finds a keyword in a specific row of a 2D data array.
 * @param data - The 2D array representing the sheet data.
 * @param rowIndex - The index of the row to search.
 * @param kw - The keyword to find (case-insensitive).
 * @returns A tuple [found: boolean, columnIndex: number | null].
 */
function findKwInRow(data: any[][], rowIndex: number, kw: string): [boolean, number | null] {
  const row = data[rowIndex];
  if (!row) return [false, null];

  for (let i = 0; i < row.length; i++) {
    const cellValue = String(row[i] ?? '').trim().toLowerCase();
    if (cellValue.includes(kw)) {
      return [true, i];
    }
  }
  return [false, null];
}

/**
 * Finds a keyword in a specific column of a 2D data array.
 * @param data - The 2D array representing the sheet data.
 * @param colIndex - The index of the column to search.
 * @param kw - The keyword to find (case-insensitive).
 * @returns A tuple [found: boolean, rowIndex: number | null].
 */
function findKwInCol(data: any[][], colIndex: number, kw: string): [boolean, number | null] {
  for (let i = 0; i < data.length; i++) {
    const cellValue = String(data[i]?.[colIndex] ?? '').trim().toLowerCase();
    if (cellValue.includes(kw)) {
      return [true, i];
    }
  }
  return [false, null];
}


export function procSheetDLN3(sheet: xlsx.Sheet): FoodItem[] {
  // Convert sheet to a 2D array of values, which is the JS/TS equivalent of a basic DataFrame.
  // `header: 1` creates an array of arrays. `defval: null` ensures empty cells are null.
  const df: any[][] = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

  const shape = [df.length, df.length > 0 ? df[0].length : 0];

  let start: [number, number] | null = null;
  let end: [number, number] | null = null;
  const kw_start = "Uplift Ratio".trim().toLowerCase();
  const kw_end = "Lưu ý".trim().toLowerCase();

  for (let i = 0; i < shape[0]; i++) {
    // Check the first column for start and end keywords
    const cellValue = String(df[i]?.[0] ?? '').trim().toLowerCase();
    if (cellValue.includes(kw_start)) {
      start = [0, i];
    } else if (cellValue.includes(kw_end)) {
      // In Python, end col was `len(cols)-1`. We'll find max row length for safety.
      const rowLength = df[i]?.length ?? 0;
      end = [rowLength > 0 ? rowLength - 1 : 0, i];
    }
  }

  // console.log("Start:", start);
  // console.log("End:", end);

  if (!start || !end) {
    console.error("Could not determine table boundaries. 'Uplift Ratio' or 'Lưu ý' not found.");
    return [];
  }

  // Create sub data frame from start and end, equivalent to df.iloc[...].
  const sub_df = df
    .slice(start[1], end[1] + 1)
    .map(row => row.slice(start![0], end![0] + 1));
  
  // make very first row of sub_df as heading for it.
  // const new_header = sub_df[0]; // Header is not strictly used later, but we follow the logic
  // The next line in Python removes the header row and the last row ("Lưu ý" row)
  const tableBody = sub_df.slice(1, sub_df.length - 1);

  // Extracts data from prev row of start of table for CLASS definition
  const row_of_class_data = df[start[1] - 1];

  // The Python code uses find_kw_in_col on a reshaped DataFrame.
  // Here we can simply search the array `row_of_class_data` directly.
  let col_class: number | null = null;
  let col_menu: number | null = null;
  row_of_class_data.forEach((cell, index) => {
    const cellStr = String(cell ?? '').trim().toLowerCase();
    if (cellStr.includes("class")) {
      col_class = index;
    }
    if (cellStr.includes("menu")) {
      col_menu = index;
    }
  });

  // console.log("Found Class:", col_class !== null, col_class);
  // console.log("Found Menu:", col_menu !== null, col_menu);
  
  const MENU_VAL = col_menu !== null ? row_of_class_data[col_menu] : null;
  const CLASS_VAL = col_class !== null ? row_of_class_data[col_class] : null;
  
  // console.log("MENU_VAL:", MENU_VAL);
  // console.log("CLASS_VAL:", CLASS_VAL);

  // Segmentation to find Checkpoints (CYCLE rows)
  const CP: Checkpoint[] = [];
  tableBody.forEach((row, i) => {
    // `pd.isna(sub_df[sub_heads[0]][i])` checks if the first cell is empty.
    if (row[0] === null || row[0] === '') {
        // Search for "cycle" in this row
        const [found, col] = findKwInRow(tableBody, i, "cycle");
        if (found && col !== null) {
            CP.push({ x: col, y: i, typ: "CYCLE" });
            // console.log(`(1) Detected a cycle at position: col=${col} | row=${i} | KEYWORD: ${tableBody[i][col]}`);
        }
    }
  });
  
  /*
  Finally segment it to usable data, FoodItem consist of:
  food_item = {
      "UpliftRatio": at column 0,
      "Name": at column 1,
      "Qty": at column 2,
      "Remark": at column 3,
      "Class": None, // No specific class for this segment
      "MenuID": variable MENU_VAL,
      "Cycle": at checkpoint, column 1
      "TimeStart": at checkpoint, column 4(parsed by parse_date_range),
      "TimeEnd" : at checkpoint, column 4 (parsed by parse_data_range)
  }
  */
  
  const food_items: FoodItem[] = [];
  let time_start: string | null = null;
  let time_end: string | null = null;
  let cycle_state: string = "ALL";
  let prev_uplift_ratio: string | number = ""; // Fallback

  tableBody.forEach((row, index_row) => {
    // Check if the current row is a checkpoint row by its index
    if (CP.some(cp => cp.y === index_row)) {
      // console.log("Iterate and found checkpoint, updating state");
      // console.log(row);
      cycle_state = row[1] ?? 'UNKNOWN'; // Cycle is in the second column of the checkpoint row
      const dateRangeResult = parseDateRangeToISO(String(row[4] ?? ''));
      if (dateRangeResult !== -1) {
          [time_start, time_end] = dateRangeResult;
      }
      return; // Continue to the next row
    }

    // Handle merged cells for 'UpliftRatio'
    let currentUpliftRatio = row[0];
    if (currentUpliftRatio === null || currentUpliftRatio === '') {
      currentUpliftRatio = prev_uplift_ratio;
    } else {
      prev_uplift_ratio = currentUpliftRatio;
    }

    const foodItem: FoodItem = {
      UpliftRatio: currentUpliftRatio,
      Name: row[1],
      Unit: row[2],
      Qty: row[3],
      Remark: row[4],
      Class: CLASS_VAL,
      MenuID: MENU_VAL,
      Cycle: cycle_state,
      TimeStart: time_start,
      TimeEnd: time_end
    };
    
    food_items.push(foodItem);
  });

  return food_items;
}