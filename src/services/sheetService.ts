
import Papa from 'papaparse';
import { DashboardRow } from '../types';

export const fetchSheetData = async (url: string, sheetName: string, query?: string): Promise<{ data: DashboardRow[]; headers: string[] }> => {
  try {
    // 1. Robust Spreadsheet ID Extraction
    let spreadsheetId = '';
    const isPublished = url.includes('/d/e/');
    
    if (isPublished) {
      const match = url.match(/\/d\/e\/(.*?)(?:\/|#|\?|$)/);
      if (match) spreadsheetId = match[1];
    } else {
      const match = url.match(/\/d\/(.*?)(?:\/|#|\?|$)/);
      if (match) spreadsheetId = match[1];
    }
    
    if (!spreadsheetId) {
      const fallbackMatch = url.match(/([a-zA-Z0-9-_]{20,})/);
      if (fallbackMatch) spreadsheetId = fallbackMatch[0];
    }

    if (!spreadsheetId) throw new Error("Invalid Google Sheet URL. Please ensure it follows the standard format.");
    
    // 2. Construct target URL
    // Use gviz/tq for queries, otherwise export?format=csv is more reliable for standard fetches
    let csvUrl = '';
    if (isPublished) {
      csvUrl = `https://docs.google.com/spreadsheets/d/e/${spreadsheetId}/pub?output=csv&sheet=${encodeURIComponent(sheetName)}`;
    } else if (query) {
      csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}&tq=${encodeURIComponent(query)}`;
    } else {
      csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&sheet=${encodeURIComponent(sheetName)}`;
    }
    
    // 3. Fetch with Timeout and Multiple CORS Proxy Fallbacks
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 20000); // 20s timeout
    
    const fetchOptions: RequestInit = {
      signal: controller.signal,
      mode: 'cors',
      credentials: 'omit',
    };

    const tryFetch = async (targetUrl: string): Promise<Response> => {
      const response = await fetch(targetUrl, fetchOptions);
      if (!response.ok) {
        if (response.status === 404) {
          throw new Error(`Sheet "${sheetName}" not found. Please check the sheet name.`);
        }
        if (response.status === 403 || response.status === 401) {
          throw new Error("Access Denied: Ensure the sheet is shared as 'Anyone with the link can view'.");
        }
        throw new Error(`Server returned ${response.status}: ${response.statusText}`);
      }
      return response;
    };

    try {
      let response: Response;
      try {
        // Try 1: Direct Fetch
        response = await tryFetch(csvUrl);
      } catch (fetchError: any) {
        console.warn("Direct fetch failed, trying Proxy 1 (corsproxy.io)...", fetchError.message);
        
        try {
          // Try 2: Proxy 1 (corsproxy.io)
          const proxyUrl1 = `https://corsproxy.io/?${encodeURIComponent(csvUrl)}`;
          response = await tryFetch(proxyUrl1);
        } catch (proxy1Error: any) {
          console.warn("Proxy 1 failed, trying Proxy 2 (allorigins)...", proxy1Error.message);
          
          try {
            // Try 3: Proxy 2 (allorigins.win)
            const proxyUrl2 = `https://api.allorigins.win/raw?url=${encodeURIComponent(csvUrl)}`;
            response = await tryFetch(proxyUrl2);
          } catch (proxy2Error: any) {
            // If all fail, throw the most descriptive error
            if (proxy2Error.message.includes('not found')) throw proxy2Error;
            if (proxy2Error.message.includes('Access Denied')) throw proxy2Error;
            
            throw new Error("CONNECTION FAILED: Unable to reach Google Sheets directly or via proxies. This may be due to network restrictions or a private sheet.");
          }
        }
      }

      clearTimeout(timeoutId);
      
      const csvText = await response.text();
      
      // Check if we got HTML instead of CSV (happens if sheet is private)
      if (csvText.trim().startsWith('<!DOCTYPE html>') || csvText.includes('google-signin')) {
        throw new Error("Access Denied: Please ensure the Google Sheet is shared as 'Anyone with the link can view'.");
      }

      return new Promise((resolve, reject) => {
        Papa.parse(csvText, {
          header: true,
          skipEmptyLines: true,
          transformHeader: (header) => header.trim(),
          complete: (results) => {
            const rawHeaders = results.meta.fields || [];
            const cleanHeaders = rawHeaders
              .filter(h => typeof h === 'string' && h.trim() !== '')
              .map(h => h.trim());

            if (results.errors.length > 0 && results.data.length === 0) {
              reject(new Error("CSV parsing failed: " + results.errors[0].message));
            } else {
              resolve({
                data: results.data as DashboardRow[],
                headers: cleanHeaders
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
