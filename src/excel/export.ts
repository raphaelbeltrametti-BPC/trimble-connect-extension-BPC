import * as XLSX from "xlsx";

export interface ExportSheet {
  name: string;
  teams: string[];
  rows: Array<{ depth: number; folderName: string; values: string[] }>;
}

/**
 * Writes the current Trimble state in the same layout the importer reads: one sheet per phase,
 * 2 leading spaces per hierarchy level in column A, one column per team with V / L / empty.
 */
export function downloadCurrentMatrix(sheets: ExportSheet[], fileName: string): void {
  const workbook = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    const aoa: string[][] = [
      ["Phase / Ordner", ...sheet.teams],
      ...sheet.rows.map((row) => [`${"  ".repeat(Math.max(0, row.depth))}${row.folderName}`, ...row.values]),
    ];
    // Excel sheet names: max. 31 chars, none of : \ / ? * [ ]
    const safeName = sheet.name.replace(/[:\\/?*[\]]/g, "-").slice(0, 31) || "Phase";
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(aoa), safeName);
  });

  XLSX.writeFile(workbook, fileName);
}
