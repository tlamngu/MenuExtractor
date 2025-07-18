/**

Converted by VerseAI Toolkit
Source language: Python (Google Colab)
Target language: Typescript
Version: 1.0.6
Agent: Qwen3-CVT-IT-12B-VI

Note: Khăn

**/

import * as XLSX from 'xlsx';

// Định nghĩa kiểu dữ liệu cho output để code rõ ràng hơn
type DataObject = { [key: string]: any };

/**
 * Xử lý dữ liệu từ một sheet của file Excel, tương đương với logic của hàm Python ProcTissueData.
 * 
 * @param sheet - Một đối tượng worksheet từ thư viện XLSX.js.
 * @returns Một mảng các đối tượng JSON, mỗi đối tượng đại diện cho một dòng dữ liệu.
 */
export function procTissueData(sheet: XLSX.WorkSheet): DataObject[] {
  // --- BƯỚC CHUYỂN ĐỔI BAN ĐẦU ---
  // Để thao tác với dữ liệu theo từng dòng như Pandas, ta chuyển sheet thành một mảng của các mảng (array of arrays).
  // { header: 1 } đảm bảo mỗi dòng là một mảng.
  // { defval: null } để các ô trống có giá trị null, tương tự như NaN trong Pandas.
  const dataAsArrays: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  // --- BẮT ĐẦU CHUYỂN ĐỔI LOGIC TỪ PYTHON ---

  // Python: df.dropna(how='all', inplace=True)
  // Lọc ra các dòng mà tất cả các ô đều là null hoặc rỗng.
  let processedData = dataAsArrays.filter(row => 
    !row.every(cell => cell === null || cell === undefined || cell === '')
  );

  // Nếu không có dữ liệu sau khi lọc, trả về mảng rỗng.
  if (processedData.length === 0) {
    return [];
  }
  
  // Python: df.columns = df.iloc[0]
  // Lấy dòng đầu tiên làm header.
  // Dùng .map(String) để đảm bảo tất cả header đều là chuỗi.
  const headers: string[] = processedData[1].map(String);
  
  // Python: df.drop(df.index[0], inplace=True)
  // Bỏ dòng đầu tiên (vì nó đã trở thành header).
  let dataRows = processedData.slice(0);
  dataRows = dataRows.slice(1);
//   dataRows = dataRows.slice(2);
  console.table(dataRows)
  // Python: df.reset_index(drop=True, inplace=True)
  // Trong JavaScript, các hàm như .filter() và .slice() tự động trả về mảng mới với index được "reset" (bắt đầu từ 0).
  // Vì vậy, không cần bước xử lý tương đương.

  // Python: df.dropna(subset=['STT'], inplace=True)
  // Tiền xử lí dữ liệu: drop các dòng có cột 'STT' là NaN/null/rỗng.
  
  // 1. Tìm vị trí (index) của cột 'STT'.
  const sttColumnIndex = headers.indexOf('STT');

  if (sttColumnIndex !== -1) {
    // 2. Nếu tìm thấy cột 'STT', lọc các dòng có giá trị trong cột đó không phải là null/undefined/rỗng.
    dataRows = dataRows.filter(row => {
      const sttValue = row[sttColumnIndex];
      return sttValue !== null && sttValue !== undefined && sttValue !== '';
    });
  } else {
    // Nếu không tìm thấy cột STT, đưa ra cảnh báo để người dùng biết.
    // Logic này tương đương với giả định "không có STT -> hỏng data" trong code Python.
    console.warn("Cảnh báo: Cột 'STT' không được tìm thấy. Bỏ qua bước lọc theo STT.");
  }
  
  // Python: df.reset_index(drop=True, inplace=True)
  // Tương tự như trên, không cần bước này trong JavaScript.

  // Python: return df.to_dict(orient='records')
  // Map các dòng dữ liệu còn lại thành một mảng các object, với key là header.
  let result: DataObject[] = dataRows.map(row => {
    const rowObject: DataObject = {};
    headers.forEach((header, index) => {
      // Gán giá trị của ô (row[index]) cho key tương ứng (header)
      rowObject[header] = row[index] !== null ? row[index] : undefined; // Có thể đổi null thành undefined cho hợp chuẩn JSON
    });
    return rowObject;
  });

  // Python: display(df)
  // Lệnh tương đương để xem dữ liệu dạng bảng trong console.
  console.log("Dữ liệu đã xử lý:");
  console.table(result);
  

  return result;
}