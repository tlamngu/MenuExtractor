import { useRef, useEffect, useState } from "react";
import * as XLSX from "xlsx";
// Make sure this path is correct for your project structure
import { procSheetDLN3 } from "./utils/extractor-v5"; 

import "./App.css";
import type { MenuItem } from "./utils/ExtractMenuFromSheet";

function App() {
  const uploadRef = useRef<null | HTMLInputElement>(null);
  const [extracted, setExtracted] = useState<MenuItem[]>([]);
  
  async function extractDataFromExcel(file: File) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    // workbook.SheetNames.forEach((e)=>{
      // procSheet(workbook.Sheets[e])
    // })
      console.table(procSheetDLN3(workbook.Sheets[workbook.SheetNames[0]]))

  }

  useEffect(() => {
    // This effect hook for the file input is well-written and can remain as is.
    if (uploadRef.current != null) {
      const handler = (e: Event) => {
        const files = (e.target as HTMLInputElement).files;
        if (files && files[0]) {
          extractDataFromExcel(files[0]);
        }
      };
      const input = uploadRef.current as HTMLInputElement;
      input.addEventListener("change", handler);
      return () => input.removeEventListener("change", handler);
    }
  }, [uploadRef]);

  return (
    <>
      <h1>Vietnam Airlines Menu Extractor</h1>
      <div className="import">
        <p>Select an Excel file (.xls, .xlsx) to parse.</p>
        <input type="file" ref={uploadRef} accept=".xls,.xlsx" />
      </div>
      
      {/* Show table only if there is data */}
      {extracted.length > 0 && (
        <div style={{ overflowX: "auto", marginTop: 24 }}>
          <p>Found {extracted.length} menu items.</p>
          <table style={{ borderCollapse: "collapse", width: "100%" }}>
            <thead>
              <tr>
                {/* Updated headers to be more descriptive */}
                <th style={{ border: "1px solid #ccc", padding: 4 }}>#</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Class</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Aircraft</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Name</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Start Time</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>End Time</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Qty</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Remark</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Note</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Uplift Ratio</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Menu ID</th>
                <th style={{ border: "1px solid #ccc", padding: 4 }}>Source Sheet</th>
              </tr>
            </thead>
            <tbody>
              {extracted.map((item, idx) => (
                <tr key={idx}>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{idx + 1}</td>
                  {/* Ensure property names match the MenuItem interface */}
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.class}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.AircraftType}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4, whiteSpace: "pre-wrap" }}>{item.Name}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.StartTime}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.EndTime}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.Quantity}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.Remark}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.Note}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item["Uplift Ratio"]}</td>
                  <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.MenuId}</td>
                  {/* <td style={{ border: "1px solid #ccc", padding: 4 }}>{item.sheetName}</td> */}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  );
}

export default App;