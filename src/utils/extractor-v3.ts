/*
Note: File dữ liệu nhận số 1
*/
import * as xlsx from "xlsx";

function isNA(value: any): boolean {
  return value === null || value === undefined;
}
export interface MenuItem {
  class: string | null;
  AircraftType: "normal" | "NEO" | null;
  Name: string | null;
  StartTime: string | null;
  EndTime: string | null;
  Quantity: string | number | null;
  Remark: string | null;
  Note: string | null;
  "Uplift Ratio": string | number | null;
  MenuId: string | null;
  Cycle: string | null;
}
export class Checkpoint {
  x: string | number;
  y: number;
  typ: "CLASS" | "AIRCRAFT";

  constructor(
    x: string | number,
    y: number,
    typ: "CLASS" | "AIRCRAFT" = "CLASS"
  ) {
    this.x = x;
    this.y = y;
    this.typ = typ;
  }

  toString(): string {
    return `Checkpoint(x=${this.x}, y=${this.y}, typ=${this.typ})`;
  }
}

/**
 * Processes a worksheet to extract structured data based on keywords and layout.
 * This is a direct port of the provided Python pandas script.
 * @param worksheet The xlsx.WorkSheet object from the xlsx library.
 */
export function procSheet(worksheet: xlsx.WorkSheet): MenuItem[] {
  const data: any[][] = xlsx.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: null,
  });

  const shape = { rows: data.length, cols: data[0]?.length || 0 };
  let start: [number, number] | [] = [];
  let end: [number, number] | [] = [];
  const kw_start = "Uplift Ratio".trim().toLowerCase();
  const kw_end = "Ghi chú".trim().toLowerCase();
  console.log(worksheet);

  for (let i = 0; i < shape.rows; i++) {
    const firstColCell = data[i][0];
    if (
      String(firstColCell || "")
        .trim()
        .toLowerCase()
        .includes(kw_start)
    ) {
      start = [0, i];
    } else if (
      String(firstColCell || "")
        .trim()
        .toLowerCase()
        .includes(kw_end)
    ) {
      const lastColIndex = data[i].length - 1;
      end = [lastColIndex, i];
    }
  }

  console.log("Found start coordinates:", start);
  console.log("Found end coordinates:", end);

  if (start.length === 0 || end.length === 0) {
    console.error(
      "Could not determine the data table range. Check for 'Uplift Ratio' and 'Ghi chú' keywords."
    );
    return [];
  }

  const sub_data_rows = data.slice(start[1], end[1] + 1);
  const new_header: string[] = sub_data_rows[0].map((h) => String(h || ""));
  let body_rows = sub_data_rows.slice(1);
  body_rows = body_rows.slice(0, body_rows.length - 1);
  const sub_df = body_rows.map((row) => {
    const obj: { [key: string]: any } = {};
    new_header.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
  const Classes: string[] = [];
  const ForAircraftExtract: string[] = [];
  const CP: Checkpoint[] = [];
  const sub_heads = new_header;

  let justDetectedAclass = false;
  let isDetectingAircraft = false;

  console.log("\nStarting segmentation loop...");

  for (let i = 0; i < sub_df.length; i++) {
    const row = sub_df[i];
    if (isNA(row[sub_heads[0]])) {
      continue;
    }

    const descriptionValue = row[sub_heads[1]];
    const qtyValue = String(row[sub_heads[2]] || "")
      .trim()
      .toLowerCase();
    const remarkValue = String(row[sub_heads[3]] || "")
      .trim()
      .toLowerCase();
    const firstColValue = row[sub_heads[0]];

    if (
      isNA(descriptionValue) &&
      qtyValue.includes("menu") &&
      remarkValue.includes("cycle")
    ) {
      if (!isDetectingAircraft && !justDetectedAclass) {
        if (typeof firstColValue !== "string") continue;
        Classes.push(firstColValue);
        CP.push(new Checkpoint(sub_heads[0], i, "CLASS"));
        console.log(
          `(1) Detected a CLASS at row ${i} | KEYWORD: ${firstColValue}`
        );
        justDetectedAclass = true;
      } else {
        if (typeof firstColValue !== "string") continue;
        ForAircraftExtract.push(firstColValue);
        CP.push(new Checkpoint(sub_heads[0], i, "AIRCRAFT"));
        console.log(
          `(1) Detected an AIRCRAFT at row ${i} | KEYWORD: ${firstColValue}`
        );
        isDetectingAircraft = justDetectedAclass;
      }
      continue;
    } else if (isNA(descriptionValue)) {
      if (justDetectedAclass) {
        if (typeof firstColValue !== "string") continue;
        ForAircraftExtract.push(firstColValue);
        CP.push(new Checkpoint(sub_heads[0], i, "AIRCRAFT"));
        console.log(
          `(2) Detected an AIRCRAFT at row ${i} | KEYWORD: ${firstColValue}`
        );
        isDetectingAircraft = true;
      } else {
        if (typeof firstColValue !== "string") continue;
        Classes.push(firstColValue);
        CP.push(new Checkpoint(sub_heads[0], i, "CLASS"));
        console.log(
          `(2) Detected a CLASS at row ${i} | KEYWORD: ${firstColValue}`
        );
        justDetectedAclass = true;
      }
      continue;
    } else {
      if (justDetectedAclass) {
        console.log(`Reset 'justDetectedAclass' at row ${i}`);
        justDetectedAclass = false;
      }
    }
  }

  console.log("\n--- Initial Data Table (`sub_df`) ---");
  console.table(sub_df);
  console.log("\nDetected Classes:", Classes);
  console.log("Detected Aircraft:", ForAircraftExtract);
  console.log("\nCheckpoints found:");
  CP.forEach((c) => console.log(c.toString()));

  // --- Python: Perform sub-data segmentation ---
  console.log("\n--- Performing sub-data segmentation ---");
  type Segment = { type: "CLASS" | "AIRCRAFT"; df: any[] };

  const segmented_dataframes: Segment[] = [];
  let current_class_start_index: number | null = null;
  let last_cp_index = -1;

  for (let i = 0; i < CP.length; i++) {
    if (CP[i].typ === "CLASS") {
      if (current_class_start_index !== null) {
        // End the previous class segment
        const end_row = CP[i].y;
        const segmented_df = sub_df.slice(current_class_start_index, end_row);
        segmented_dataframes.push({ type: "CLASS", df: segmented_df });
      }
      // Start a new class segment
      current_class_start_index = CP[i].y;
      last_cp_index = i;
    }
  }

  // Add the last class segment
  if (current_class_start_index !== null) {
    const segmented_df = sub_df.slice(current_class_start_index);
    segmented_dataframes.push({ type: "CLASS", df: segmented_df });
  }

  // --- Python: Now, within each class segment, identify aircraft sub-segments ---
  type AircraftSegment = {
    type: "AIRCRAFT";
    df: any[];
  };
  type ClassData = {
    class_df: any[];
    aircraft_segments: AircraftSegment[];
  };
  type NestedStructure = {
    className: string;
    classData: ClassData;
  };
  const nested_structured_dataframes: NestedStructure[] = [];

  // This requires mapping the sliced array indices back to the original `sub_df` indices
  // by finding the corresponding checkpoints.
  for (const class_segment of segmented_dataframes) {
    if (class_segment.type === "CLASS") {
      const class_df = class_segment.df;
      const class_name = class_df[0][sub_heads[0]];

      // Find the original start index of this class_df in sub_df
      const original_start_index = sub_df.indexOf(class_df[0]);

      const class_segment_data: ClassData = {
        class_df: class_df,
        aircraft_segments: [],
      };
      let current_aircraft_start_index_in_class_df: number | null = null;

      for (let i = 0; i < class_df.length; i++) {
        const original_sub_df_index = original_start_index + i;

        // Find if there is a checkpoint at this original index
        const current_cp = CP.find((cp) => cp.y === original_sub_df_index);

        if (current_cp && current_cp.typ === "AIRCRAFT") {
          if (current_aircraft_start_index_in_class_df !== null) {
            const aircraft_end_index_in_class_df = i;
            const aircraft_sub_df = class_df.slice(
              current_aircraft_start_index_in_class_df,
              aircraft_end_index_in_class_df
            );
            class_segment_data.aircraft_segments.push({
              type: "AIRCRAFT",
              df: aircraft_sub_df,
            });
          }
          current_aircraft_start_index_in_class_df = i;
        }
      }

      // Add the last aircraft sub-segment in the class segment
      if (current_aircraft_start_index_in_class_df !== null) {
        const aircraft_sub_df = class_df.slice(
          current_aircraft_start_index_in_class_df
        );
        class_segment_data.aircraft_segments.push({
          type: "AIRCRAFT",
          df: aircraft_sub_df,
        });
      }

      nested_structured_dataframes.push({
        className: class_name,
        classData: class_segment_data,
      });
    }
  }
  let food_items: MenuItem[] = [];
  // --- Python: Display the nested segmented dataframes ---

  console.log("\n\n--- FINAL NESTED STRUCTURE ---");
  for (const { className, classData } of nested_structured_dataframes) {
    console.log(`\n--- Class Segment: ${className} ---`);

    if (classData.aircraft_segments.length > 0) {
      for (const aircraft_segment of classData.aircraft_segments) {
        const aircraft_name =
          aircraft_segment.df[0]?.[sub_heads[0]] || "Unknown Aircraft";
        console.log(`----- Aircraft Sub-segment for: ${aircraft_name} -----`);

        let items: MenuItem[] = [];

        const Class = className;
        let Menu = aircraft_segment.df[0]["Qty"];
        let Cycle = aircraft_segment.df[0]["Remark"];
        console.log("Class:", Class, "Menu:", Menu, "Cycle:", Cycle);

        for (let i = 1; i < aircraft_segment.df.length; i++) {
          let newdat: MenuItem = {
            "Uplift Ratio":
              typeof aircraft_segment.df[i]["Uplift Ratio"] == "number"
                ? `${aircraft_segment.df[i]["Uplift Ratio"] * 100}%`
                : aircraft_segment.df[i]["Uplift Ratio"],
            Name: aircraft_segment.df[i]["Component Description"],
            MenuId: Menu,
            Remark: aircraft_segment.df[i]["Remark"],
            class: Class, // Use the corrected 'Class' variable
            Cycle: Cycle,
            AircraftType: aircraft_name
              .toLowerCase()
              .includes("NEO".toLowerCase())
              ? "NEO"
              : "normal",
            EndTime: null,
            StartTime: null,
            Note: null,
            Quantity: aircraft_segment.df[i]["Qty"],
          };
          items.push(newdat);
        }
        console.table(items);
        items.forEach((e) => {
          food_items.push(e);
        });
      }
    } else {
      let items: MenuItem[] = [];

      const Class = className;
      let Menu = classData.class_df[0]["Qty"];
      let Cycle = classData.class_df[0]["Remark"];
      console.log("Class:", Class, "Menu:", Menu, "Cycle:", Cycle);

      for (let i = 1; i < classData.class_df.length; i++) {
        let newdat: MenuItem = {
          "Uplift Ratio":
            typeof classData.class_df[i]["Uplift Ratio"] == "number"
              ? `${classData.class_df[i]["Uplift Ratio"] * 100}%`
              : classData.class_df[i]["Uplift Ratio"],
          Name: classData.class_df[i]["Component Description"],
          MenuId: Menu,
          Remark: classData.class_df[i]["Remark"],
          class: Class, // Use the corrected 'Class' variable
          Cycle: Cycle,
          AircraftType: "normal", // No specific aircraft, so default to "normal"
          EndTime: null,
          StartTime: null,
          Note: null,
          Quantity: classData.class_df[i]["Qty"],
        };
        items.push(newdat);
      }
      console.table(items);
      items.forEach((e) => {
        food_items.push(e);
      });
    }
  }
  console.log("Res ----------------------");
  console.table(food_items);
  return food_items;
}
