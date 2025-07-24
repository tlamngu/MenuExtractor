import * as xlsx from "xlsx";

interface FlightMeta {
  documentNumber: string;
  revision: string;
  attachment: string;
}

interface DayFlightsMeta {
  date: string;
  shift: string;
  section: string;
  acsupervisor: string;
  tpoSupervisor: string;
}

interface ExtractedData {
  meta: FlightMeta;
  dayFlightsMeta: DayFlightsMeta;
  nightFlightsMeta: DayFlightsMeta;
  dayFlights: any;
  nightFlights: any;
}

interface bounderies {
  startX: number;
  endX: number;
  startY: number;
  endY: number;
}

interface pviot_cnt {
  name: String;
  count: number;
}

function fillMerges(merges: xlsx.Range[] | undefined, data: any[][]): any[][] {
  merges?.forEach((merge) => {
    let scan_col_start = merge.s.c;
    let scan_col_end = merge.e.c;
    let scan_row_start = merge.s.r;
    let scan_row_end = merge.e.r;

    let dat = null;
    for (let y = scan_row_start; y <= scan_row_end; y++) {
      for (let x = scan_col_start; x <= scan_col_end; x++) {
        if (data[y] && data[y][x] != null) {
          dat = data[y][x];
        }
      }
    }

    if (dat != null) {
      for (let y = scan_row_start; y <= scan_row_end; y++) {
        for (let x = scan_col_start; x <= scan_col_end; x++) {
          if (data[y] && data[y][x] == null) {
            data[y][x] = dat;
          }
        }
      }
    }
  });
  return data;
}

function match_kw(str: any, kw: string): boolean {
  if (!str) return false;
  return String(str)
    .toLowerCase()
    .trim()
    .includes(String(kw).trim().toLowerCase());
}

function match_strict(str: any, kw: string): boolean {
  if (!str) return false;
  return String(str).toLowerCase().trim() == String(kw).trim().toLowerCase();
}

function extractMetadata(data: any[][]): FlightMeta {
  let documentNumber = "";
  let revision = "";
  let attachment = "";
  for (let y = 0; y < Math.min(5, data.length); y++) {
    for (let x = 0; x < (data[y]?.length || 0); x++) {
      const cell = data[y]?.[x];
      if (!cell) continue;
      const cellStr = String(cell);
      if (match_kw(cellStr, "Document Number:")) {
        const match = cellStr.match(/Document Number:\s*(.+)/i);
        if (match) documentNumber = match[1].trim();
      }
      if (match_kw(cellStr, "Revision:")) {
        const match = cellStr.match(/Revision:\s*(.+)/i);
        if (match) revision = match[1].trim();
      }
      if (match_kw(cellStr, "Attachment")) {
        attachment = cellStr.trim();
      }
    }
  }

  return { documentNumber, revision, attachment };
}

function extractTableMeta(data: any[][], bounds: bounderies): DayFlightsMeta {
  let date = "";
  let shift = "";
  let section = "";
  let acsupervisor = "";
  let tpoSupervisor = "";

  const searchStartY = Math.max(0, bounds.startY - 10);
  const searchEndY = bounds.startY + 3;
  const searchStartX = bounds.startX;
  const searchEndX = bounds.endX + 5;

  for (let y = searchStartY; y <= searchEndY; y++) {
    for (let x = searchStartX; x <= searchEndX; x++) {
      const cell = data[y]?.[x];
      if (!cell) continue;

      const cellStr = String(cell);

      if (match_kw(cellStr, "Ngày-Date:") || match_kw(cellStr, "Date:")) {
        const match = cellStr.match(/(?:Ngày-)?Date:\s*(.+)/i);
        if (match) date = match[1].trim();
      }

      if (match_kw(cellStr, "Ca- shift:") || match_kw(cellStr, "shift:")) {
        const match = cellStr.match(/(?:Ca-?\s*)?shift:\s*(.+)/i);
        if (match) shift = match[1].trim();
      }

      if (match_kw(cellStr, "Section:")) {
        const match = cellStr.match(/Section:\s*(.+)/i);
        if (match) section = match[1].trim();
      }

      if (match_kw(cellStr, "ACS Supervisor:")) {
        const match = cellStr.match(/ACS Supervisor:\s*(.+)/i);
        if (match) acsupervisor = match[1].trim();
      }

      if (
        match_kw(cellStr, "TPO Supervisor") ||
        match_kw(cellStr, "Giám sát TPO")
      ) {
        const lines = cellStr.split(/\r?\n/);
        for (const line of lines) {
          if (match_kw(line, "TPO Supervisor")) {
            tpoSupervisor = line.replace(/TPO Supervisor/i, "").trim();
            break;
          }
        }
        if (!tpoSupervisor && lines.length > 1) {
          tpoSupervisor = lines[lines.length - 1].trim();
        }
      }
    }
  }

  return { date, shift, section, acsupervisor, tpoSupervisor };
}

function extractFlightData(
  data: any[][],
  OrgData: any[][],
  boundTB1: bounderies,
  sts_dict: string[]
): any {
  let flightData: any = {};
  let pviot_count: pviot_cnt[] = [];

  for (let x = boundTB1.startX; x <= boundTB1.endX; x++) {
    let Piviot = data[boundTB1.startY][x];

    if (match_kw(data[boundTB1.startY][x], "ETD/ETA")) {
      Piviot = "ETD/ETA";
    } else if (match_kw(data[boundTB1.startY][x], "Flt No")) {
      Piviot = "FlightNo";
    }

    sts_dict.map((sts) => {
      if (match_strict(data[boundTB1.startY][x], sts)) {
        let found = false;

        if (match_kw(data[boundTB1.startY - 1][x], "TPO")) {
          Piviot = "TPO." + String(data[boundTB1.startY][x] || "").trim();
        } else {
          const parentCell = data[boundTB1.startY - 1][x];
          Piviot =
            String(parentCell || "").trim() +
            "." +
            String(data[boundTB1.startY][x] || "").trim();
        }

        pviot_count.map((p: pviot_cnt, index: number) => {
          if (match_strict(p.name, Piviot)) {
            found = true;
            Piviot = Piviot + String(pviot_count[index].count);
            pviot_count[index].count += 1;
            return;
          }
        });

        if (!found) {
          pviot_count.push({
            name: Piviot,
            count: 1,
          });
        }
      }
    });

    for (let y = boundTB1.startY + 1; y < data.length; y++) {
      if (data[y] && data[y][x] != null) {
        if (OrgData[y] && OrgData[y][boundTB1.startX + 1] == null) continue;

        if (x !== boundTB1.startX) {
          if (data[y][boundTB1.startX] != null) {
            const aircraftKey = String(data[y][boundTB1.startX]).trim();
            if (!flightData[aircraftKey]) flightData[aircraftKey] = {};
            flightData[aircraftKey][Piviot] = data[y][x];
          }
        } else {
          const aircraftKey = String(data[y][boundTB1.startX]).trim();
          if (aircraftKey && aircraftKey !== "") {
            flightData[aircraftKey] = {};
          }
        }
      }
    }
  }

  return flightData;
}

export function extractTimeTable(worksheet: xlsx.WorkSheet): ExtractedData {
  const merges = worksheet["!merges"];
  let data: any[][] = xlsx.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: null,
  });
  const OrgData = data;
  data = fillMerges(merges, data);
  const meta = extractMetadata(data);
  const sts_dict = ["FWD", "MID", "AFT", "No"];
  let boundTB1: bounderies = {
    startX: 0,
    endX: 0,
    startY: 0,
    endY: 0,
  };
  for (let y = 0; y < data.length; y++) {
    if (match_kw(data[y]?.[0], "M/bay")) {
      for (let x = 0; x < (data[y]?.length || 0); x++) {
        if (
          match_strict(data[y][x], "No") &&
          match_kw(data[y]?.[x + 1], "M/bay")
        ) {
          boundTB1.startX = 0;
          boundTB1.endX = x;
          boundTB1.startY = y;
          boundTB1.endY = data.length - 1;
          break;
        }
      }
      if (boundTB1.endX > 0) break;
    }
  }
  let boundTB2: bounderies = {
    startX: boundTB1.endX + 1,
    endX: (data[boundTB1.startY]?.length || 1) - 1,
    startY: boundTB1.startY,
    endY: data.length - 1,
  };

  const TB1flightData = extractFlightData(data, OrgData, boundTB1, sts_dict);
  const TB2flightData = extractFlightData(data, OrgData, boundTB2, sts_dict);
  const dayFlightsMeta = extractTableMeta(data, boundTB1);
  const nightFlightsMeta = extractTableMeta(data, boundTB2);

  return {
    meta,
    dayFlightsMeta,
    nightFlightsMeta,
    dayFlights: TB1flightData,
    nightFlights: TB2flightData,
  };
}
