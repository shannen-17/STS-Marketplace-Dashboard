import * as XLSX from 'xlsx';
import { ParsedWorkbook, SheetInfo, WorkbookData, WorkbookMetadata, Row } from '@/lib/types';

const normalizeHeader = (header: string): string => {
  return header
    .toLowerCase()
    .trim()
    .replace(/[\s\-_]+/g, ' ')
    .replace(/^\s+|\s+$/g, '');
};

const parseCell = (value: any): any => {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const num = Number(value.replace(/[$,%]/g, ''));
    return !Number.isNaN(num) ? num : value;
  }
  return value;
};

export async function parseWorkbook(buffer: ArrayBuffer, fileName: string): Promise<ParsedWorkbook> {
  const errors: string[] = [];
  const data: WorkbookData = {};
  const sheets: SheetInfo[] = [];

  try {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

    workbook.SheetNames.forEach((sheetName) => {
      try {
        let rows: Row[] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
          defval: '',
          blankrows: false
        });

        rows = rows.map((row) => {
          const normalized: Row = {};
          Object.entries(row).forEach(([key, value]) => {
            normalized[normalizeHeader(key)] = parseCell(value);
          });
          return normalized;
        });

        rows = rows.filter((row) => Object.values(row).some((v) => v !== null && v !== '' && v !== undefined));

        if (rows.length > 0) {
          const headers = Object.keys(rows[0]);
          data[sheetName] = rows;
          sheets.push({
            name: sheetName,
            rowCount: rows.length,
            headers,
            sample: rows.slice(0, 3)
          });
        }
      } catch (err) {
        errors.push(`Error parsing sheet "${sheetName}": ${err instanceof Error ? err.message : 'Unknown error'}`);
      }
    });
  } catch (err) {
    errors.push(`Critical parsing error: ${err instanceof Error ? err.message : 'Unknown error'}`);
  }

  const metadata: WorkbookMetadata = {
    parsedAt: new Date(),
    fileName,
    totalSheets: sheets.length,
    validSheets: sheets,
    errors
  };

  return { sheets, data, metadata };
}

export function getSheet(data: WorkbookData, sheetName: string): Row[] {
  return data[sheetName] || [];
}

export function findColumn(rows: Row[], columnNames: string[]): string | null {
  if (rows.length === 0) return null;
  const headers = Object.keys(rows[0]);
  return columnNames.find((name) => headers.some((h) => h.includes(normalizeHeader(name)))) || null;
}
