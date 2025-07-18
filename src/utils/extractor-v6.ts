/**

Converted by VerseAI Toolkit
Source language: Python (Google Colab)
Target language: Typescript
Version: 1.0.2
Agent: Qwen3-CVT-IT-12B

Note: Lịch nhân sự / work plan

**/
import * as XLSX from 'xlsx';

// --- Type Definitions (Interfaces) ---
// These interfaces provide strong typing for the function's output, a best practice in TypeScript.

/**
 * Describes the metadata extracted from the roster header.
 */
interface PlanDetails {
    valid_from?: string;
    valid_to?: string;
    issued_date?: string;
    version?: string;
}

/**
 * Represents a single day's work assignment for an employee.
 */
interface WorkDay {
    date: string;
    shift_code: string | null; // null represents an empty cell (e.g., day off)
}

/**
 * Represents a single employee and their complete work schedule.
 */
interface Employee {
    id: string;
    name: string;
    group: string;
    work_schedule: WorkDay[];
}

/**
 * Defines the structure of the final output object.
 */
interface RosterOutput {
    plan_details: PlanDetails;
    employees: Employee[];
}


/**
 * Extracts employee work plan information from a roster worksheet.
 *
 * This function dynamically parses a worksheet object from xlsx.js,
 * extracting metadata, employee details, and daily schedules. It is designed to
 * be robust against structural variations in the input file.
 *
 * @param sheet A worksheet object from the xlsx.js library. It is recommended to
 *              generate this from a file where employee codes are formatted as text
 *              to preserve leading zeros.
 * @returns An object containing the extracted plan details and a list of employees,
 *          where each employee is an object with their ID, name, group, and work schedule.
 */
export function procEmployees(sheet: XLSX.WorkSheet): RosterOutput {
    // In TypeScript with xlsx.js, we convert the sheet to an array of arrays.
    // This is the closest equivalent to a raw pandas DataFrame loaded with header=None.
    // `defval: null` ensures empty cells become `null` for consistent checks.
    const data: (string | number | null)[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    // Helper function to safely get a string value from a cell.
    const getCellString = (row: number, col: number): string | null => {
        if (row < 0 || row >= data.length || col < 0 || !data[row] || col >= data[row].length) {
            return null;
        }
        const value = data[row][col];
        return value !== null && value !== undefined ? String(value).trim() : null;
    };


    // --- 1. Extract Header Information ---
    const header_info: PlanDetails = {};
    const headerSearchLimit = Math.min(5, data.length);

    try {
        for (let r = 0; r < headerSearchLimit; r++) {
            for (let c = 0; c < (data[r] ? data[r].length : 0); c++) {
                const cellValue = getCellString(r, c);
                if (!cellValue) continue;

                if (cellValue.includes('Từ :')) {
                    header_info.valid_from = getCellString(r, c + 1) || undefined;
                }
                if (cellValue.includes('Đến :')) {
                    header_info.valid_to = getCellString(r, c + 1) || undefined;
                }
                if (cellValue.includes('Phát hành :')) {
                    header_info.issued_date = getCellString(r, c + 2) || undefined;
                }
                if (cellValue.startsWith('Phát hành lần')) {
                    const match = cellValue.match(/\d+/);
                    header_info.version = match ? match[0] : cellValue.split(' ').pop() || undefined;
                }
            }
        }
    } catch (e) {
        console.warn(`Warning: Could not parse all header details. Error: ${e instanceof Error ? e.message : String(e)}`);
    }


    // --- 2. Locate Data Table and Columns ---
    let headerRowIdx = -1;
    for (let i = 0; i < data.length; i++) {
        const row = data[i].map(cell => String(cell).trim());
        if (row.includes('Code') && row.includes('Họ và Tên')) {
            headerRowIdx = i;
            break;
        }
    }

    if (headerRowIdx === -1) {
        throw new Error("Could not find the main data header row (containing 'Code' and 'Họ và Tên').");
    }

    const dateRowIdx = headerRowIdx + 1;
    const dataStartIdx = headerRowIdx + 2;

    const mainHeaderList = data[headerRowIdx].map(s => String(s).trim());
    const summaryColStartIdx = mainHeaderList.indexOf('CÔNG QL');

    if (summaryColStartIdx === -1) {
        throw new Error("Could not determine the end of the schedule columns. 'CÔNG QL' marker not found.");
    }

    // Define column indices based on the located header
    const codeColIdx = 1;
    const nameColIdx = 2;
    const dateColStartIdx = 3;
    const dateColEndIdx = summaryColStartIdx; // Exclusive index

    const dates: string[] = [];
    if (dateRowIdx < data.length) {
        for (let i = dateColStartIdx; i < dateColEndIdx; i++) {
            const dateStr = getCellString(dateRowIdx, i);
            dates.push(dateStr || `Day_${i}`); // Fallback if date is missing
        }
    }


    // --- 3. Extract Employee Data ---
    const all_employees: Employee[] = [];
    let current_group = "Nhóm 1";
    let group_counter = 1;
    // This flag helps differentiate a '1' that starts a new group from the very first employee
    let isNewGroupContext = true;

    for (let i = dataStartIdx; i < data.length; i++) {
        const row = data[i];
        if (!row) continue; // Skip empty rows

        const employeeCodeRaw = getCellString(i, codeColIdx);
        const employeeNameRaw = getCellString(i, nameColIdx);

        // Check if the row contains valid employee data
        // Equivalent to `pd.notna(employee_code_raw) and pd.notna(employee_name_raw)`
        if (employeeCodeRaw && employeeNameRaw) {
            // Clean up employee code (handles cases like '294.0' -> '0294')
            const employeeCode = String(employeeCodeRaw).split('.')[0].padStart(4, '0');
            const employeeName = String(employeeNameRaw).trim();

            // Check for group change based on 'Số'/'TT' column resetting to '1'
            const soTt = getCellString(i, 0)?.split('.')[0];
            if (soTt === '1' && !isNewGroupContext) {
                group_counter++;
                current_group = `Nhóm ${group_counter}`;
            }

            isNewGroupContext = false;

            // Extract work schedule for the employee
            const work_schedule: WorkDay[] = [];
            for (let j = 0; j < dates.length; j++) {
                const colIdx = dateColStartIdx + j;
                const shift_code = getCellString(i, colIdx);

                work_schedule.push({
                    date: dates[j],
                    shift_code: shift_code // Will be null if cell is empty
                });
            }

            const employeeData: Employee = {
                id: employeeCode,
                name: employeeName,
                group: current_group,
                work_schedule: work_schedule
            };
            all_employees.push(employeeData);

        // Check if the row is a named group header
        // Equivalent to `pd.isna(employee_code_raw) and pd.isna(employee_name_raw)`
        } else if (!employeeCodeRaw && !employeeNameRaw) {
            const potentialHeader = getCellString(i, 0);
            const isNumeric = (str: string) => /^\d+$/.test(str);
            
            // Check for a meaningful, non-numeric header in the first column
            if (potentialHeader && !isNumeric(potentialHeader)) {
                const ignoreList = ['anh/em', 'giao non-air', 'tùy thuộc', 'prepared by', 'write python code'];
                const potentialHeaderLower = potentialHeader.toLowerCase();
                
                // If it's not in the ignore list, treat it as a new group name
                if (!ignoreList.some(keyword => potentialHeaderLower.includes(keyword))) {
                    current_group = potentialHeader; // Assign the new named group
                    isNewGroupContext = true; // Set flag for the start of a new group
                }
            }
        }
    }

    // --- 4. Assemble Final Result ---
    const result: RosterOutput = {
        plan_details: header_info,
        employees: all_employees
    };

    return result;
}