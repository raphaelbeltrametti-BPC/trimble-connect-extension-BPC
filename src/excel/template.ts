import * as XLSX from "xlsx";

const EXAMPLE_SHEET: (string | number)[][] = [
  ["Ordner", "Team A", "Team B", "Team C"],
  ["Hauptordner", "V", "L", "K"],
  ["  Unterordner 1", "V", "L", "K"],
  ["  Unterordner 2", "L", "L", "K"],
  ["    Unter-Unterordner", "K", "V", "L"],
];

const INSTRUCTIONS_SHEET: (string | number)[][] = [
  ["CDE-Berechtigungsmatrix - Vorlage"],
  [""],
  ["Jedes Tabellenblatt ist eine Phase; der Blattname wird als oberster Ordner verwendet."],
  ["Spalte A = Ordnername, weitere Spalten = ein Team pro Spalte (Kopfzeile = Teamname)."],
  ["Einrueckung in Spalte A mit 2 Leerzeichen pro Ebene bildet die Ordner-Verschachtelung ab."],
  [""],
  ["Erlaubte Werte pro Zelle:"],
  ["  Vollzugriff: V oder Vollzugriff"],
  ["  Lesezugriff: L oder Lesezugriff"],
  ["  Kein Zugriff: K, Kein Zugriff oder leer"],
  [""],
  ["Siehe Beispielblatt fuer eine 3-stufige Ordnerstruktur mit 3 Teams."],
];

export function buildMatrixTemplateWorkbook(): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(INSTRUCTIONS_SHEET), "Anleitung");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(EXAMPLE_SHEET), "Beispielphase");
  return workbook;
}

export function downloadMatrixTemplate(fileName = "Berechtigungsmatrix-Vorlage.xlsx"): void {
  XLSX.writeFile(buildMatrixTemplateWorkbook(), fileName);
}
