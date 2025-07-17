export interface MenuItem {
    class: string | null;
    AircraftType: 'normal' | 'NEO' | null;
    Name: string | null;
    StartTime: string | null;
    EndTime: string | null;
    Quantity: string | number | null;
    Remark: string | null;
    Note: string | null;
    "Uplift Ratio": string | number | null;
    MenuId: string | null;
}

type ExcelData = { [key: string]: (string | number | null)[][] };

const cleanCell = (cell: any): string => {
    if (cell === null || cell === undefined) return '';
    return String(cell).trim();
};

const CLASS_HEADER_REGEX = /\b(CLASS|CREW|CAPTAIN|COPILOT)\b/i;
const CONDITION_HEADER_REGEX = /(dành cho tàu|<etd<)/i;
const TIME_EXTRACTION_REGEX = /(\d{2}:\d{2})\s*<\s*ETD\s*<\s*(\d{2}:\d{2})/;

const isClassHeader = (cell: string): boolean => {
    return CLASS_HEADER_REGEX.test(cell);
};

const isConditionHeader = (cell: string): boolean => {
    return CONDITION_HEADER_REGEX.test(cell);
};

const extractMenuId = (cell: string): string | null => {
    if (!cell) return null;

    // Tìm có từ khóa MENU
    let match = cell.match(/MENU\s*([A-Z0-9]+)/i);
    if (match && match[1]) {
        return match[1].trim().toUpperCase();
    }

    //fallback tìm ID độc lập (ko có MENU)
    match = cell.match(/\b([A-Z]{1,2}\d{0,2})\b/);
    if (match && cell.trim() === match[0]) {
        return match[1].trim().toUpperCase();
    }
    
    return null;
};

export function processDetailedFlightMenuData(sheetData: ExcelData):MenuItem[]{
    
}

export function processFlightMenuDataV2(sheetData: ExcelData): MenuItem[] {
    const allMenuItems: MenuItem[] = [];

    const HEADER_ROW_UPLIFT_REGEX = /uplift ratio/i;
    const HEADER_ROW_COMPONENT_REGEX = /component description/i;
    const NOTE_ROW_REGEX = /^Ghi chú:/i;
    const LOADED_BY_REGEX = /^Loaded By/i;

    for (const sheetName in sheetData) {
        if (!sheetData.hasOwnProperty(sheetName)) continue;

        const rows = sheetData[sheetName];
        if (!rows || rows.length === 0) continue;

        // Tìm ghi chú chung của cả sheet
        const noteRow = rows.find(row => row && NOTE_ROW_REGEX.test(cleanCell(row[0])));
        const sheetNote = noteRow ? cleanCell(noteRow[0]) : null;

        // Tìm dòng tiêu đề chính của bảng 
        const headerRowIndex = rows.findIndex(row =>
            row &&
            HEADER_ROW_UPLIFT_REGEX.test(cleanCell(row[0])) &&
            HEADER_ROW_COMPONENT_REGEX.test(cleanCell(row[1]))
        );

        if (headerRowIndex === -1) {
            console.warn(`[Cảnh báo] Bỏ qua sheet "${sheetName}" vì không tìm thấy dòng tiêu đề (header).`);
            continue;
        }

        // State machine
        let currentState = {
            class: null as string | null,
            aircraftType: null as 'normal' | 'NEO' | null,
            menuId: null as string | null,
            startTime: null as string | null,
            endTime: null as string | null,
        };

        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            
            if (!row || row.length === 0 || LOADED_BY_REGEX.test(cleanCell(row[0]))) {
                continue;
            }

            const [rawUplift, rawName, rawQty, rawRemark] = row;
            const upliftCell = cleanCell(rawUplift);
            const nameCell = cleanCell(rawName);
            const qtyCell = cleanCell(rawQty);
            const remarkCell = cleanCell(rawRemark);

            if (!upliftCell && !nameCell && !qtyCell && !remarkCell) {
                continue;
            }

            // --- Logic nhận dạng loại dòng && cập nhật trạng thái ---

            // Trường hợp 1: Dòng là TIÊU ĐỀ NHÓM CHÍNH (CLASS, CREW, etc....)
            if (isClassHeader(upliftCell)) {
                currentState.class = upliftCell;
                // Đổi hạng ghế reset trạng thái
                currentState.aircraftType = 'normal';
                currentState.startTime = null;
                currentState.endTime = null;
                // Tiêu đề có thể đi kèm Menu ID ở cột Quantity :v
                currentState.menuId = extractMenuId(qtyCell) || currentState.menuId;
                continue; 
            }

            // Trường hợp 2: Dòng là TIÊU ĐỀ ĐIỀU KIỆN (Tàu bay, thời gian)
            if (isConditionHeader(upliftCell)) {
                currentState.aircraftType = upliftCell.includes('NEO') ? 'NEO' : 'normal';
                
                const timeMatch = upliftCell.match(TIME_EXTRACTION_REGEX);
                if (timeMatch) {
                    currentState.startTime = timeMatch[1];
                    currentState.endTime = timeMatch[2];
                } else {
                    // Nếu dòng điều kiện không có giờ, reset
                    currentState.startTime = null;
                    currentState.endTime = null;
                }
                
                // Tiêu đề điều kiện cũng có thể đi kèm Menu ID :v
                currentState.menuId = extractMenuId(qtyCell) || currentState.menuId;
                continue; 
            }
            
            // Trường hợp 3: Dòng chỉ chứa MENU ID (thường cột 1, 2 trống)
            const potentialMenuId = extractMenuId(qtyCell);
            if (!upliftCell && !nameCell && potentialMenuId) {
                currentState.menuId = potentialMenuId;
                //reset qua menu mới
                currentState.aircraftType = 'normal'; 
                currentState.startTime = null;
                currentState.endTime = null;
                continue; 
            }

            // Trường hợp 4 (Fallback): Món ăn
            if (nameCell) {
                let itemStartTime = currentState.startTime;
                let itemEndTime = currentState.endTime;
                let itemRemark = remarkCell;
                let itemMenuId = currentState.menuId;

                // Kt có ghi đè điều kiện bên remark ko (<ETD<...)
                const remarkTimeMatch = itemRemark.match(TIME_EXTRACTION_REGEX);
                if (remarkTimeMatch) {
                    itemStartTime = remarkTimeMatch[1];
                    itemEndTime = remarkTimeMatch[2];
                }
                const remarkMenuId = extractMenuId(itemRemark);
                if (remarkMenuId) {
                    itemMenuId = remarkMenuId;
                }

                let upliftRatio: string | number | null = null;
                if (typeof rawUplift === "number" && !isNaN(rawUplift)) {
                    upliftRatio = `${rawUplift * 100}%`;
                } else if (typeof rawUplift === "string" && rawUplift.trim() !== "") {
                    upliftRatio = rawUplift;
                } else {
                    upliftRatio = null;
                }

                const menuItem: MenuItem = {
                    class: currentState.class,
                    AircraftType: currentState.aircraftType,
                    Name: nameCell,
                    StartTime: itemStartTime,
                    EndTime: itemEndTime,
                    Quantity: qtyCell,
                    Remark: itemRemark,
                    Note: sheetNote,
                    "Uplift Ratio": upliftRatio,
                    MenuId: itemMenuId,
                };

                allMenuItems.push(menuItem);
            }
        }
    }

    return allMenuItems;
}