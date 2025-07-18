/**

Note: Menu do uong

**/

import * as XLSX from 'xlsx';

// --- Type Definitions for Type Safety ---

/**
 * Represents the row and column boundaries of a data table within the sheet.
 */
interface TableBoundaries {
    item_start_row: number;
    item_end_row: number;
    flight_start_col: number;
    flight_end_col: number;
    flight_header_start_row: number;
    flight_header_end_row: number;
}

/**
 * Represents a found flight header with its name and column index.
 */
interface FlightHeader {
    name: string;
    col: number;
}

/**
 * Represents the extracted content of a single table block.
 * Format: { [itemName: string]: { [flightName: string]: any } }
 */
type ExtractedTableContent = Record<string, Record<string, any>>;

/**
 * Represents the final structured data from the entire sheet.
 * Format: { [classification: string]: ExtractedTableContent }
 */
type AllExtractedData = Record<string, ExtractedTableContent>;


// --- Helper Class for xlsx.WorkSheet ---

/**
 * A wrapper around an xlsx.WorkSheet to provide pandas.DataFrame-like
 * integer-based access (e.g., iloc[row, col]).
 */
class SheetHelper {
    private sheet: XLSX.WorkSheet;
    public rowCount: number;
    public colCount: number;

    constructor(sheet: XLSX.WorkSheet) {
        this.sheet = sheet;
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        this.rowCount = range.e.r + 1;
        this.colCount = range.e.c + 1;
    }

    /**
     * Gets the value of a cell using 0-based row and column indices.
     * @param row The row index.
     * @param col The column index.
     * @returns The cell's value, or undefined if the cell is empty or out of bounds.
     */
    public getValue(row: number, col: number): any {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = this.sheet[cellAddress];
        return cell ? cell.v : undefined;
    }
}


// --- Helper Functions (ported from Python) ---

/**
 * Case-insensitive keyword matching in a string.
 */
function match_kw(str: any, kw: string): boolean {
    // Use `?? ''` to handle null or undefined inputs gracefully
    return String(str ?? '').trim().toLowerCase().includes(kw.trim().toLowerCase());
}

/**
 * Check if a value can be converted to a float.
 */
function is_numeric(value: any): boolean {
    if (value === null || value === undefined || value === '') {
        return false;
    }
    // `!isNaN(parseFloat(value))` checks if it's a number representation
    // `isFinite(value)` excludes Infinity, -Infinity
    return !isNaN(parseFloat(value)) && isFinite(value);
}

// --- Constants for Keywords ---
const STT_KW = "STT";
const FLIGHT_INFO_KW = "Thông tin chuyến bay";
const FLIGHT_LEG_KW = "Thông tin chặng bay";
const UNIT_KW = "ĐVT";
const END_DATA_KW = "Số lượng xe giao đi";
const NOTE_KW = "ghi chú";
const REEXPORT_KW = "Tái xuất";
const NM_KW = "PHIẾU GIAO NHẬN ĐỒ UỐNG"

/**
 * Finds the row and column boundaries of a data table, given the row index of its 'STT' header.
 * This version strictly respects block boundaries defined by the next 'STT' occurrence.
 *
 * @param sheetHelper - The helper object for the worksheet.
 * @param stt_row_idx - The row index where 'STT' was found.
 * @returns An object with boundaries, or null if essential markers are not found.
 */
function find_table_boundaries(sheetHelper: SheetHelper, stt_row_idx: number): TableBoundaries | null {
    // 1. Find the hard end of the current block by finding the *next* STT or end-marker
    let block_end_row = sheetHelper.rowCount; // Default to the end of the sheet
    for (let i = stt_row_idx + 1; i < sheetHelper.rowCount; i++) {
        const cell_value = sheetHelper.getValue(i, 0);
        if (match_kw(cell_value, STT_KW) || match_kw(cell_value, END_DATA_KW)) {
            block_end_row = i;
            break;
        }
    }
    console.log(`| | Block boundary determined to be row ${block_end_row} (next STT or end-marker).`);

    // 2. Find Vertical Boundaries (Item Rows) *within the determined block*
    let item_start_row = -1;
    for (let i = stt_row_idx + 1; i < block_end_row; i++) {
        if (is_numeric(sheetHelper.getValue(i, 0))) {
            item_start_row = i;
            break;
        }
    }

    if (item_start_row === -1) {
        console.log("| | No numeric items found in this block. Treating as a header-only section.");
        item_start_row = block_end_row;
    }
    
    const item_end_row = block_end_row;

    // 3. Find Horizontal Boundaries (Flight Columns)
    let flight_start_col = -1;
    for (let i = 0; i < sheetHelper.colCount; i++) {
        const cell = sheetHelper.getValue(stt_row_idx, i);
        if (match_kw(cell, UNIT_KW) || match_kw(cell, FLIGHT_LEG_KW)) {
            flight_start_col = i;
            break;
        }
    }

    if (flight_start_col === -1) {
        console.log(`| Warning: Could not find '${UNIT_KW}' or '${FLIGHT_LEG_KW}' in STT row ${stt_row_idx}.`);
        return null;
    }

    let flight_end_col = -1;
    // Loop backwards from the last column
    for (let i = sheetHelper.colCount - 1; i > flight_start_col; i--) {
        const note_cell_1 = sheetHelper.getValue(stt_row_idx, i);
        const note_cell_2 = sheetHelper.getValue(stt_row_idx + 1, i);
        if (match_kw(note_cell_1, NOTE_KW) || match_kw(note_cell_2, NOTE_KW)) {
            flight_end_col = i;
            break;
        }
    }
    
    if (flight_end_col === -1) {
        console.log(`| Warning: Could not find '${NOTE_KW}'. Using last column as fallback.`);
        flight_end_col = sheetHelper.colCount;
    }

    const flight_header_end_row = item_start_row;

    return {
        item_start_row,
        item_end_row,
        flight_start_col,
        flight_end_col,
        flight_header_start_row: stt_row_idx + 1,
        flight_header_end_row,
    };
}

/**
 * Extracts item data for each flight from a table with defined boundaries.
 * Handles duplicate flight names by appending a counter.
 *
 * @param sheetHelper - The helper object for the worksheet.
 * @param boundaries - An object containing the row/column boundaries of the table.
 * @returns A dictionary where keys are item names and values are dicts of {flight: value}.
 */
function extract_table_data(sheetHelper: SheetHelper, boundaries: TableBoundaries): ExtractedTableContent {
    // Unpack boundaries for clarity
    const {
        item_start_row,
        item_end_row,
        flight_start_col,
        flight_end_col,
        flight_header_start_row,
        flight_header_end_row
    } = boundaries;

    // 1. Find all flight headers and their column indices
    const flight_headers: FlightHeader[] = [];
    for (let r = flight_header_start_row; r < flight_header_end_row; r++) {
        for (let c = flight_start_col; c < flight_end_col; c++) {
            const cell_value = sheetHelper.getValue(r, c);
            if (cell_value !== undefined && cell_value !== null && !match_kw(cell_value, REEXPORT_KW)) {
                flight_headers.push({ name: String(cell_value), col: c });
            }
        }
    }

    // 2. Extract data for each item against each flight
    const table_content: ExtractedTableContent = {};

    for (let item_row_idx = item_start_row; item_row_idx < item_end_row; item_row_idx++) {
        const itemNameRaw = sheetHelper.getValue(item_row_idx, 1);
        if (itemNameRaw === undefined || itemNameRaw === null) {
            continue;
        }

        const item_name = String(itemNameRaw).trim();
        table_content[item_name] = {};
        
        const item_flight_name_tracker: Record<string, number> = {};

        for (const header of flight_headers) {
            const raw_flight_name = header.name;
            const flight_col_idx = header.col;

            let processed_flight_name: string;
            if (raw_flight_name in item_flight_name_tracker) {
                item_flight_name_tracker[raw_flight_name]++;
                processed_flight_name = `${raw_flight_name}_${item_flight_name_tracker[raw_flight_name]}`;
            } else {
                item_flight_name_tracker[raw_flight_name] = 0;
                processed_flight_name = raw_flight_name;
            }
            
            const value = sheetHelper.getValue(item_row_idx, flight_col_idx);
            table_content[item_name][processed_flight_name] = value;
        }
    }

    return table_content;
}

/**
 * Processes a complex worksheet to extract structured data based on keywords.
 * It identifies data blocks starting with 'STT', determines their boundaries,
 * and extracts item data associated with flight information.
 *
 * @param sheet - The raw worksheet object from the xlsx library.
 * @returns An object containing the structured data, classified by sections.
 */
export function procDoUong(sheet: XLSX.WorkSheet): {} {
    console.log("Start processing.");
    
    // Create the helper to provide iloc-like functionality
    const sheetHelper = new SheetHelper(sheet);
    console.log("Shape of table (rows, cols):", sheetHelper.rowCount, sheetHelper.colCount);
    const all_extracted_data: AllExtractedData = {};
    
    let packed : {spill_id: string, data: AllExtractedData} = {
        spill_id: '',
        data: all_extracted_data
    };

    for (let i = 0; i < sheetHelper.rowCount; i++) {
        const cell_value = sheetHelper.getValue(i, 0);
        console.log(cell_value)
        if(match_kw(cell_value, NM_KW)){
            packed["spill_id"] = cell_value;
        }
        if (match_kw(cell_value, STT_KW)) {
            console.log(`\nFound '${STT_KW}' at row ${i}. Starting block analysis.`);

            if (i === 0) {
                console.log("| Warning: Found 'STT' on the first row. Cannot determine classification title.");
                continue;
            }

            const classification_title = String(sheetHelper.getValue(i - 1, 0) ?? 'Untitled').trim();
            console.log(`| Classification: '${classification_title}'`);

            const boundaries = find_table_boundaries(sheetHelper, i);
            if (!boundaries) {
                console.log(`| Skipping block under '${classification_title}' due to missing boundaries.`);
                continue;
            }
            console.log(`| Detected Boundaries:`, boundaries);

            const table_data = extract_table_data(sheetHelper, boundaries);

            if (Object.keys(table_data).length === 0) {
                console.log(`| No data extracted for block '${classification_title}'.`);
                continue;
            }

            if (!(classification_title in all_extracted_data)) {
                all_extracted_data[classification_title] = table_data;
            } else {
                console.log(`| Warning: Duplicate classification '${classification_title}'. Merging data.`);
                // Deep merge the data
                for (const item in table_data) {
                    if (item in all_extracted_data[classification_title]) {
                        // Merge flight data for an existing item
                        Object.assign(all_extracted_data[classification_title][item], table_data[item]);
                    } else {
                        // Add new item to the classification
                        all_extracted_data[classification_title][item] = table_data[item];
                    }
                }
            }
        }
    }
    packed["data"] = all_extracted_data
    console.log("\nProcessing finished.");
    return packed;
}

// --- Example Usage ---
// This part shows how you would use the function in a real scenario.
/*
import { readFileSync } from 'fs';

try {
    // 1. Read the Excel file
    const filePath = './your_excel_file.xlsx';
    const workbook = XLSX.read(readFileSync(filePath));
    
    // 2. Get the first sheet
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // 3. Process the sheet
    const extractedData = procDoUong(worksheet);

    // 4. Do something with the result
    console.log("\n--- Final Extracted Data ---");
    console.log(JSON.stringify(extractedData, null, 2));

} catch (error) {
    console.error("An error occurred:", error);
}
*/