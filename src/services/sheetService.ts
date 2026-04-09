
import Papa from 'papaparse';
import { DashboardRow } from '../types';

export const fetchSheetData = async (url: string, sheetName: string, query?: string): Promise<{ data: DashboardRow[]; headers: string[] }> => {
  try {
    // Extract Spreadsheet ID from typical URLs
    const match = url.match(/\/d\/(.*?)(\/|$)/);
    if (!match) throw new Error("Invalid Google Sheet URL. Please ensure it follows the standard format.");
    const spreadsheetId = match[1];
    
    // Construct CSV export URL
    let csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
    if (query) {
      csvUrl += `&tq=${encodeURIComponent(query)}`;
    }
    
    // Add timeout to fetch call
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 20000); // 20s timeout
    
    try {
      const response = await fetch(csvUrl, { signal: controller.signal });
      clearTimeout(timeoutId);
      
      if (!response.ok) throw new Error("Failed to fetch sheet. Ensure it is shared as 'Anyone with the link can view'.");
      
      const csvText = await response.text();
      
      return new Promise((resolve, reject) => {
        Papa.parse(csvText, {
          header: true,
          skipEmptyLines: true,
          transformHeader: (header) => header.trim(),
          complete: (results) => {
            if (results.errors.length > 0 && results.data.length === 0) {
              reject(new Error("CSV parsing failed: " + results.errors[0].message));
            } else {
              resolve({
                data: results.data as DashboardRow[],
                headers: results.meta.fields || []
              });
            }
          },
          error: (error: Error) => reject(error)
        });
      });
    } catch (error: any) {
      clearTimeout(timeoutId);
      console.error("fetchSheetData error:", error);
      if (error.name === 'AbortError') {
        throw new Error("Request timed out. Please check your connection.");
      }
      throw new Error(error.message || "An error occurred while fetching data.");
    }
  } catch (error: any) {
    throw new Error(error.message || "An error occurred while fetching data.");
  }
};

export const parseDate = (dateStr: string): Date | null => {
  // 1. Null Check
  if (!dateStr || typeof dateStr !== 'string' || dateStr.trim() === '') return null;
  const trimmed = dateStr.trim();
  
  let date: Date | null = null;

  // 2. Sheet Native: Handle Date(y,m,d) regex
  const gsheetMatch = trimmed.match(/Date\((\d+),\s*(\d+),\s*(\d+)/);
  if (gsheetMatch) {
    const y = parseInt(gsheetMatch[1], 10);
    const m = parseInt(gsheetMatch[2], 10); // GSheets Date() is 0-indexed for months
    const d = parseInt(gsheetMatch[3], 10);
    date = new Date(y, m, d);
  } 
  // 3. DD-MM-YYYY Format: Use regex ^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$
  else {
    const dmyMatch = trimmed.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (dmyMatch) {
      const day = parseInt(dmyMatch[1], 10);
      const month = parseInt(dmyMatch[2], 10);
      const year = parseInt(dmyMatch[3], 10);
      date = new Date(year, month - 1, day);
    }
    // 4. YYYY-MM-DD Format: Use regex ^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$
    else {
      const ymdMatch = trimmed.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
      if (ymdMatch) {
        const year = parseInt(ymdMatch[1], 10);
        const month = parseInt(ymdMatch[2], 10);
        const day = parseInt(ymdMatch[3], 10);
        date = new Date(year, month - 1, day);
      }
    }
  }

  // Final fallback for any other format that might work (e.g. ISO)
  if (!date || isNaN(date.getTime())) {
    const fallbackDate = new Date(trimmed);
    if (!isNaN(fallbackDate.getTime())) {
      date = fallbackDate;
    }
  }

  // CRITICAL: Before returning the parsed Date, force it to local midnight using .setHours(0, 0, 0, 0)
  if (date && !isNaN(date.getTime())) {
    date.setHours(0, 0, 0, 0);
    return date;
  }
  
  return null;
};

export const getColLetter = (index: number): string => {
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 65) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
};
