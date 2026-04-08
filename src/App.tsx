import React, { useState, useEffect, useMemo, useCallback, useRef, memo } from 'react';
import { Menu, AlertCircle, ExternalLink, RefreshCw, Database, Copy, Check } from 'lucide-react';
import { INITIAL_CONFIG, DEFAULT_MAPPING } from './constants';
import { SheetConfig, DashboardRow, KPIStats, SKUDetail, RemainingQtyItem } from './types';
import { fetchSheetData, parseDate } from './services/sheetService';
import KPIGrid from './components/KPIGrid';
import FilterSection from './components/FilterSection';
import SKUDetailsSection from './components/SKUDetailsSection';
import WipDrilldownModal from './components/WipDrilldownModal';
import RejectionDetailsSection from './components/RejectionDetailsSection';
import RejectionDrilldownModal from './components/RejectionDrilldownModal';
import AcceptedDrilldownModal from './components/AcceptedDrilldownModal';
import SettingsMenu from './components/SettingsMenu';
import ErrorBoundary from './components/ErrorBoundary';
import { isWithinInterval, startOfDay, endOfDay } from 'date-fns';

// Memoize heavy components
const MemoizedKPIGrid = memo(KPIGrid);
const MemoizedFilterSection = memo(FilterSection);
const MemoizedSKUDetailsSection = memo(SKUDetailsSection);
const MemoizedRejectionDetailsSection = memo(RejectionDetailsSection);
const MemoizedWipDrilldownModal = memo(WipDrilldownModal);
const MemoizedRejectionDrilldownModal = memo(RejectionDrilldownModal);
const MemoizedAcceptedDrilldownModal = memo(AcceptedDrilldownModal);

const MemoizedSettingsMenu = memo(SettingsMenu);

const App: React.FC = () => {
  const [config, setConfig] = useState<SheetConfig>(() => {
    const saved = localStorage.getItem('qc_dashboard_config');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          return {
            ...INITIAL_CONFIG,
            ...parsed,
            mapping: { ...INITIAL_CONFIG.mapping, ...(parsed.mapping || {}) }
          };
        }
      } catch (e) {
        console.error("Failed to parse config from localStorage", e);
      }
    }
    return INITIAL_CONFIG;
  });
  
  const [data, setData] = useState<DashboardRow[]>(() => {
    const saved = localStorage.getItem('qc_dashboard_cached_data');
    if (!saved) return [];
    try {
      const parsed = JSON.parse(saved) as DashboardRow[];
      // Re-parse dates because JSON.parse turns them into strings
      return parsed.map(row => ({
        ...row,
        date: row.date ? new Date(row.date) : null,
        _parsedDate: row._parsedDate ? new Date(row._parsedDate) : null
      }));
    } catch (e) {
      console.error("Failed to parse cached data", e);
      return [];
    }
  });
  const [headers, setHeaders] = useState<string[]>(() => {
    const saved = localStorage.getItem('qc_dashboard_cached_headers');
    return saved ? JSON.parse(saved) : [];
  });
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [showCopyToast, setShowCopyToast] = useState(false);
  const [syncStatus, setSyncStatus] = useState<{ type: 'success' | 'error' | 'syncing' | null; message: string | null }>({ type: null, message: null });
  const syncTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  const latestDataRef = useRef<DashboardRow[] | null>(null);
  const latestHeadersRef = useRef<string[] | null>(null);
  const latestMappingRef = useRef<any>(null);

  const syncLatestData = useCallback(() => {
    if (latestDataRef.current) {
      setData(latestDataRef.current);
      if (latestHeadersRef.current) setHeaders(latestHeadersRef.current);
      if (latestMappingRef.current) setConfig(prev => ({ ...prev, mapping: latestMappingRef.current }));
      latestDataRef.current = null;
      latestHeadersRef.current = null;
      latestMappingRef.current = null;
    }
  }, []);

  const setSyncMessage = (type: 'success' | 'error' | 'syncing', message: string) => {
    setSyncStatus({ type, message });
    if (syncTimeoutRef.current) clearTimeout(syncTimeoutRef.current);
    if (type !== 'syncing') {
      syncTimeoutRef.current = setTimeout(() => setSyncStatus({ type: null, message: null }), 5000);
    }
  };

  const [selectedBatches, setSelectedBatches] = useState<string[]>(() => {
    const saved = localStorage.getItem('qc_dashboard_selected_batches');
    return saved ? JSON.parse(saved) : [];
  });
  const [dateRange, setDateRange] = useState<{ start: Date | null; end: Date | null }>(() => {
    const saved = localStorage.getItem('qc_dashboard_date_range');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return {
          start: parsed.start ? new Date(parsed.start) : null,
          end: parsed.end ? new Date(parsed.end) : null
        };
      } catch (e) {
        return { start: null, end: null };
      }
    }
    return { start: null, end: null };
  });
  const [uidSearch, setUidSearch] = useState(() => {
    return localStorage.getItem('qc_dashboard_uid_search') || '';
  });
  const [debouncedUidSearch, setDebouncedUidSearch] = useState(uidSearch);
  const [isRejectionModalOpen, setIsRejectionModalOpen] = useState(false);
  const [isAcceptedModalOpen, setIsAcceptedModalOpen] = useState(false);
  const [isWipModalOpen, setIsWipModalOpen] = useState(false);

  const handleOpenRejectionModal = useCallback(() => setIsRejectionModalOpen(true), []);
  const handleCloseRejectionModal = useCallback(() => setIsRejectionModalOpen(false), []);
  const handleOpenAcceptedModal = useCallback(() => setIsAcceptedModalOpen(true), []);
  const handleCloseAcceptedModal = useCallback(() => setIsAcceptedModalOpen(false), []);
  const handleOpenWipModal = useCallback(() => setIsWipModalOpen(true), []);
  const handleCloseWipModal = useCallback(() => setIsWipModalOpen(false), []);

  // Debounce UID search
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedUidSearch(uidSearch);
    }, 300);
    return () => clearTimeout(timer);
  }, [uidSearch]);

  const handleSetSelectedBatches = useCallback((batches: string[]) => {
    syncLatestData();
    setSelectedBatches(batches);
  }, [syncLatestData]);

  const handleSetDateRange = useCallback((range: { start: Date | null; end: Date | null }) => {
    syncLatestData();
    setDateRange(range);
  }, [syncLatestData]);

  const handleSetUidSearch = useCallback((search: string) => {
    syncLatestData();
    setUidSearch(search);
  }, [syncLatestData]);

  const handleConfigUpdate = useCallback((newConfig: SheetConfig) => {
    syncLatestData();
    setConfig(newConfig);
    localStorage.setItem('qc_dashboard_config', JSON.stringify(newConfig));
  }, [syncLatestData]);

  // Persistence effects
  useEffect(() => {
    localStorage.setItem('qc_dashboard_selected_batches', JSON.stringify(selectedBatches));
  }, [selectedBatches]);

  useEffect(() => {
    localStorage.setItem('qc_dashboard_date_range', JSON.stringify(dateRange));
  }, [dateRange]);

  useEffect(() => {
    localStorage.setItem('qc_dashboard_uid_search', uidSearch);
  }, [uidSearch]);

  useEffect(() => {
    if (data.length > 0) {
      localStorage.setItem('qc_dashboard_cached_data', JSON.stringify(data));
    }
  }, [data]);

  useEffect(() => {
    if (headers.length > 0) {
      localStorage.setItem('qc_dashboard_cached_headers', JSON.stringify(headers));
    }
  }, [headers]);

  const lastRawData = useRef<string>('');
  const lastSyncTime = useRef<Date>(new Date());

  // Robust column detection with priority on SKU
  const findHeaderMatch = useCallback((availableHeaders: string[], searchTerms: string[]): string | undefined => {
    if (!availableHeaders.length) return undefined;
    
    const lowerHeaders = availableHeaders.map(h => h.trim().toLowerCase());
    
    for (const term of searchTerms) {
      const idx = lowerHeaders.indexOf(term.toLowerCase());
      if (idx !== -1) return availableHeaders[idx];
    }

    const contains = availableHeaders.find(h => 
      searchTerms.some(term => h.toLowerCase().includes(term.toLowerCase()))
    );
    if (contains) return contains;

    return availableHeaders.find(h => {
      const normalized = h.toLowerCase().replace(/[^a-z0-9]/g, '');
      return searchTerms.some(term => {
        const termNorm = term.toLowerCase().replace(/[^a-z0-9]/g, '');
        return normalized.includes(termNorm) || termNorm.includes(normalized);
      });
    });
  }, []);

  const autoDetectMapping = useCallback((availableHeaders: string[]) => {
    const mapping = { ...(config.mapping || DEFAULT_MAPPING) };
    
    const aliases = {
      sku: ['sku', 'item', 'part', 'article', 'model', 'product', 'code'],
      batchNo: ['batch', 'lot', 'serial', 'batch no', 'batch number', 'batch id', 'lot no', 'lot id'],
      ringStatus: ['status', 'result', 'outcome', 'quality'],
      uid: ['uid', 'id', 'barcode', 'serial'],
      inward: ['inward', 'qty', 'count', 'total'],
      reason: ['reason', 'rejection', 'defect', 'cause', 'fault'],
      movedToInventory: ['moved to inventory', 'inventory', 'moved'],
    };

    // 1. Explicitly look for "DATE" column first (case-insensitive)
    const dateHeader = availableHeaders.find(h => h.trim().toUpperCase() === 'DATE');
    if (dateHeader) {
      mapping.date = dateHeader;
    } else {
      // Fallback only if "DATE" is not found, using a very restricted set of aliases
      const fallbackDate = findHeaderMatch(availableHeaders, ['date', 'timestamp', 'day']);
      if (fallbackDate) mapping.date = fallbackDate;
    }

    (Object.entries(aliases) as [keyof typeof aliases, string[]][]).forEach(([key, terms]) => {
      const currentVal = mapping[key as keyof typeof mapping];
      if (!currentVal || !availableHeaders.includes(currentVal)) {
        const detected = findHeaderMatch(availableHeaders, terms);
        if (detected) {
          (mapping as any)[key] = detected;
        }
      }
    });

    return mapping;
  }, [config.mapping, findHeaderMatch]);

  const loadData = useCallback(async (silent = false) => {
    if (!config.url) {
      if (!silent) setError("CONFIGURATION REQUIRED: Please link a valid public Google Sheet.");
      return;
    }
    
    if (!silent) {
      setLoading(true);
      setError(null);
      setSyncMessage('syncing', 'Syncing data...');
    }
    
    try {
      const { data: rawData, headers: sheetHeaders } = await fetchSheetData(config.url, config.sheetName);
      
      // Performance: Avoid re-processing if data hasn't changed
      const currentDataStr = JSON.stringify(rawData);
      if (silent && currentDataStr === lastRawData.current) {
        lastSyncTime.current = new Date();
        return;
      }
      lastRawData.current = currentDataStr;
      lastSyncTime.current = new Date();

      setHeaders(sheetHeaders);
      
      const updatedMapping = autoDetectMapping(sheetHeaders);
      if (JSON.stringify(updatedMapping) !== JSON.stringify(config.mapping)) {
        setConfig(prev => ({ ...prev, mapping: updatedMapping }));
      }

      // Pre-parse dates and extract unique batches in a single pass
      const batchCol = updatedMapping.batchNo;
      const hasBatchCol = batchCol && sheetHeaders.includes(batchCol);
      const uniqueBatchesSet = new Set<string>();

      const updatedData = rawData.map(row => {
        const parsedDate = parseDate(String(row[updatedMapping.date] || ''));
        if (hasBatchCol) {
          const batchVal = String(row[batchCol] || '').trim();
          if (batchVal) uniqueBatchesSet.add(batchVal);
        }
        return {
          ...row,
          date: parsedDate,
          _parsedDate: parsedDate
        };
      });

      // Background sync: update internal dataset ONLY, do not trigger UI refresh
      if (silent) {
        latestDataRef.current = updatedData;
        latestHeadersRef.current = sheetHeaders;
        latestMappingRef.current = updatedMapping;
        
        // Still update cache for persistence
        localStorage.setItem('qc_dashboard_cached_data', JSON.stringify(updatedData));
        localStorage.setItem('qc_dashboard_cached_headers', JSON.stringify(sheetHeaders));
        
        setError(null);
        return;
      }

      // Use requestAnimationFrame for smooth UI update
      requestAnimationFrame(() => {
        setData(updatedData);
        lastDateMapping.current = updatedMapping.date;

        if (hasBatchCol) {
          const uniqueBatches = Array.from(uniqueBatchesSet)
            .sort((a: string, b: string) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
          
          if (selectedBatches.length === 0) {
            setSelectedBatches(uniqueBatches);
          }
        }
        
        setError(null);
        if (!silent) setSyncMessage('success', 'Data synced successfully');
      });
    } catch (err: any) {
      console.error("Data sync failed:", err);
      const isNetworkError = err.message === 'Failed to fetch' || 
                             err.message.includes('timed out') || 
                             !navigator.onLine;
      
      if (!silent) {
        const userMessage = isNetworkError 
          ? "NETWORK ERROR: Unable to reach Google Sheets. Please check your connection." 
          : `SYNC ERROR: ${err.message || 'Unknown error occurred during data fetch.'}`;
        setError(userMessage);
        setSyncMessage('error', 'Sync failed');
      } else {
        // Silent error: just log it and keep previous data
        if (!isNetworkError) {
          console.error("Background sync failed:", err.message);
        } else {
          console.warn("Background sync skipped: Connection issue or timeout.");
        }
      }
    } finally {
      if (!silent) setLoading(false);
    }
  }, [config.url, config.sheetName, autoDetectMapping, selectedBatches.length]);

  const lastDateMapping = useRef(config.mapping?.date);

  useEffect(() => {
    if (data.length > 0 && config.mapping?.date && lastDateMapping.current !== config.mapping.date) {
      lastDateMapping.current = config.mapping.date;
      setData(prev => prev.map(row => {
        const parsedDate = parseDate(String(row[config.mapping.date] || ''));
        return {
          ...row,
          date: parsedDate,
          _parsedDate: parsedDate
        };
      }));
    }
  }, [config.mapping?.date, data.length]);

  const lastSyncAttempt = useRef<number>(Date.now());

  // Initial load: Use silent sync if we have cached data to avoid loading spinner
  useEffect(() => { 
    if (config.url) {
      if (data.length > 0) {
        loadData(true); // Background sync
      } else {
        loadData(false); // Initial full sync
      }
    }
  }, [config.url, config.sheetName]);

  // Handle visibility change: Resume instantly and sync if needed
  useEffect(() => {
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        // Tab became active again - check if we should sync
        const now = Date.now();
        // Only sync if it's been more than 5 minutes since last attempt
        if (config.url && !loading && (now - lastSyncAttempt.current > 5 * 60 * 1000)) {
          lastSyncAttempt.current = now;
          loadData(true);
        }
      }
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);
    window.addEventListener('focus', handleVisibilityChange);
    
    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      window.removeEventListener('focus', handleVisibilityChange);
    };
  }, [config.url, config.sheetName, loadData, loading, syncLatestData]);

  // Auto-sync every 10 minutes, ONLY when tab is active and online
  useEffect(() => {
    const interval = setInterval(() => {
      if (document.visibilityState === 'visible' && config.url && !loading && navigator.onLine) {
        lastSyncAttempt.current = Date.now();
        loadData(true);
      }
    }, 10 * 60 * 1000);
    return () => clearInterval(interval);
  }, [config.url, config.sheetName, loadData, loading]);

  const allUniqueBatches = useMemo(() => {
    const batchCol = config.mapping?.batchNo;
    if (!batchCol || !headers.includes(batchCol)) {
      console.log('App: Batch column not found in headers:', batchCol, headers);
      return [];
    }
    const unique = Array.from(new Set(data.map(r => String(r[batchCol] || '').trim())))
      .filter(Boolean)
      .sort((a: string, b: string) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
    console.log('App: Found unique batches:', unique.length);
    return unique;
  }, [data, config.mapping?.batchNo, headers]);

  const filteredData = useMemo(() => {
    const mapping = config.mapping || DEFAULT_MAPPING;
    if (data.length === 0) return [];

    const parsedStartDate = dateRange.start ? new Date(dateRange.start) : null;
    if (parsedStartDate) parsedStartDate.setHours(0, 0, 0, 0);
    
    const parsedEndDate = dateRange.end ? new Date(dateRange.end) : null;
    if (parsedEndDate) parsedEndDate.setHours(0, 0, 0, 0);

    const batchCol = mapping.batchNo;
    const hasBatchFilter = batchCol && headers.includes(batchCol) && selectedBatches.length > 0 && selectedBatches.length < allUniqueBatches.length;
    const searchTerm = debouncedUidSearch.trim().toLowerCase();
    const uidCol = mapping.uid;

    const dateCol = mapping.date;
    const hasDateMapping = dateCol && headers.includes(dateCol);

    return data.filter(row => {
      // Date Filter
      if (hasDateMapping && (parsedStartDate || parsedEndDate)) {
        const rawRowDate = row.date;
        const rowDate = rawRowDate instanceof Date ? rawRowDate : (rawRowDate ? new Date(rawRowDate) : null);
        
        // If a row has a blank or invalid date, exclude it when a filter is active
        if (!rowDate || isNaN(rowDate.getTime())) return false;
        
        // Ensure rowDate is also forced to local midnight for comparison if not already
        // (Though parseDate should have already handled this)
        const rowTime = new Date(rowDate.getTime());
        rowTime.setHours(0, 0, 0, 0);

        const isAfterStart = !parsedStartDate || rowTime.getTime() >= parsedStartDate.getTime();
        const isBeforeEnd = !parsedEndDate || rowTime.getTime() <= parsedEndDate.getTime();

        if (!isAfterStart || !isBeforeEnd) return false;
      }
      
      // Batch Filter
      if (hasBatchFilter) {
        const rowBatch = String(row[batchCol] || '').trim();
        if (!selectedBatches.includes(rowBatch)) return false;
      }

      // UID Search Filter
      if (searchTerm) {
        const rowUid = String(row[uidCol] || '').toLowerCase();
        if (!rowUid.includes(searchTerm)) return false;
      }

      return true;
    });
  }, [data, config, dateRange, selectedBatches, debouncedUidSearch, headers, allUniqueBatches]);

  const stats: KPIStats = useMemo(() => {
    const mapping = config.mapping || DEFAULT_MAPPING;
    let total = 0;
    let accepted = 0;
    let rejected = 0;
    let inwardCount = 0;
    let movedToInventory = 0;

    filteredData.forEach(r => {
      const uid = String(r[mapping.uid] || '').trim();
      const sku = String(r[mapping.sku] || '').trim();
      
      // Only count rows that have either a UID or an SKU
      if (uid !== '' || sku !== '') {
        total++;
        const status = String(r[mapping.ringStatus] || '').trim().toLowerCase();
        if (['accepted', 'ok', 'pass', '1', 'true', 'yes'].includes(status)) {
          accepted++;
        } else if (['rejected', 'nok', 'fail', '0', 'false', 'no'].includes(status)) {
          rejected++;
        }
      }

      if (String(r[mapping.inward] || '').trim() !== '') {
        inwardCount++;
      }

      const movedVal = String(r[mapping.movedToInventory] || '').trim();
      if (movedVal !== '') {
        movedToInventory++;
      }
    });

    const wip = Math.max(0, inwardCount - total);
    const yieldVal = total > 0 ? (accepted / total) * 100 : 0;
    
    return { total, accepted, rejected, wip, yield: yieldVal, movedToInventory };
  }, [filteredData, config]);

  const skuDetails = useMemo(() => {
    const mapping = config.mapping || DEFAULT_MAPPING;
    const skuMap: Record<string, { total: number; accepted: number; rejected: number }> = {};
    
    let skuKey = mapping.sku;
    if (!headers.includes(skuKey)) {
      skuKey = findHeaderMatch(headers, ['sku', 'item', 'part', 'model']) || (headers.length > 0 ? headers[0] : '');
    }

    if (filteredData.length === 0 || !skuKey) {
      return [];
    }

    filteredData.forEach(r => {
      const val = r[skuKey];
      if (val !== undefined && val !== null) {
        const sku = String(val).trim().replace(/[\u0000-\u001F\u007F-\u009F]/g, ""); 
        if (sku !== '') {
          if (!skuMap[sku]) {
            skuMap[sku] = { total: 0, accepted: 0, rejected: 0 };
          }
          skuMap[sku].total += 1;
          
          const status = String(r[mapping.ringStatus] || '').trim().toLowerCase();
          if (['accepted', 'ok', 'pass', '1', 'true', 'yes'].includes(status)) {
            skuMap[sku].accepted += 1;
          } else if (['rejected', 'nok', 'fail', '0', 'false', 'no'].includes(status)) {
            skuMap[sku].rejected += 1;
          }
        }
      }
    });

    return Object.entries(skuMap).map(([sku, s]) => ({
      sku,
      total: s.total,
      accepted: s.accepted,
      rejected: s.rejected,
      yield: s.total > 0 ? (s.accepted / s.total) * 100 : 0
    }));
  }, [filteredData, config, headers, findHeaderMatch]);

  const handleCopyReport = () => {
    syncLatestData();
    if (stats.total === 0) {
      alert("No data available");
      return;
    }

    const mapping = config.mapping || DEFAULT_MAPPING;
    
    const totalData = filteredData.filter(r => 
      String(r[mapping.uid] || '').trim() !== '' || 
      String(r[mapping.sku] || '').trim() !== ''
    );

    const acceptedRows = totalData.filter(r => {
      const status = String(r[mapping.ringStatus] || '').trim().toLowerCase();
      return ['accepted', 'ok', 'pass', '1', 'true', 'yes'].includes(status);
    });
    
    const acceptedGroups: Record<string, number> = {};
    acceptedRows.forEach(r => {
      const sku = String(r[mapping.sku] || 'Unknown SKU').trim();
      acceptedGroups[sku] = (acceptedGroups[sku] || 0) + 1;
    });
    
    const acceptedDetailsStr = Object.entries(acceptedGroups)
      .map(([sku, count]) => `${sku}: ${count}`)
      .join('\n');

    const rejectedRows = totalData.filter(r => {
      const status = String(r[mapping.ringStatus] || '').trim().toLowerCase();
      return ['rejected', 'nok', 'fail', '0', 'false', 'no'].includes(status);
    });
    
    const rejectedGroups: Record<string, number> = {};
    rejectedRows.forEach(r => {
      const reason = String(r[mapping.reason] || 'No Reason Specified').trim();
      rejectedGroups[reason] = (rejectedGroups[reason] || 0) + 1;
    });
    
    const rejectedDetailsStr = Object.entries(rejectedGroups)
      .map(([reason, count]) => `${reason}: ${count}`)
      .join('\n');

    const reportText = `------------------------------------
BATCH REPORT

TOTAL : ${stats.total}
ACCEPTED : ${stats.accepted}
REJECTED : ${stats.rejected}
YIELD : ${stats.yield.toFixed(1)}%

ACCEPTED DETAILS
${acceptedDetailsStr || 'None'}

REJECTION DETAILS
${rejectedDetailsStr || 'None'}
------------------------------------`;

    navigator.clipboard.writeText(reportText).then(() => {
      setShowCopyToast(true);
      // Fixed: Change setShowResetToast to setShowCopyToast to fix reference error
      setTimeout(() => setShowCopyToast(false), 3000);
    }).catch(err => {
      console.error('Failed to copy: ', err);
    });
  };

  return (
    <ErrorBoundary>
      <div className="min-h-screen pb-12 w-full max-w-[100vw] bg-[#0f1117]">
        <header className="sticky top-0 z-40 bg-[#161a23]/90 backdrop-blur-xl border-b border-white/5 shadow-2xl w-full">
        <div className="w-full px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-20">
            <div className="flex items-center gap-4">
              <div className="flex items-center justify-center">
                <img 
                  src={`/logo.png?v=${Date.now()}`}
                  alt=""
                  style={{ maxHeight: '48px', width: 'auto', objectFit: 'contain' }}
                  className="rounded-xl"
                  onError={(e) => {
                    e.currentTarget.style.display = 'none';
                  }}
                />
              </div>
              <div>
                <h1 className="text-xl font-black text-white leading-tight tracking-tight uppercase">Dashboard</h1>
                <p className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-[0.4em] mono">Quality Analytics Pro</p>
              </div>
            </div>
              <div className="flex items-center gap-4">
                {loading && (
                  <div className="flex items-center gap-2 px-3 py-1.5 bg-[#38bdf8]/10 border border-[#38bdf8]/20 rounded-full animate-pulse">
                    <RefreshCw className="w-3 h-3 text-[#38bdf8] animate-spin" />
                    <span className="text-[9px] font-black text-[#38bdf8] uppercase tracking-widest">Syncing</span>
                  </div>
                )}
                {data.length > 0 && (
                  <button 
                    onClick={handleCopyReport}
                    className="hidden sm:flex items-center gap-2 px-5 py-2.5 text-xs font-black text-[#e5e7eb] bg-[#22c55e]/10 hover:bg-[#22c55e]/20 rounded-xl transition-all border border-[#22c55e]/30 uppercase tracking-widest"
                  >
                    <Copy className="w-4 h-4 text-[#22c55e]" />
                    REPORT
                  </button>
                )}
                {config.url && (
                  <div className="flex flex-col items-end gap-1">
                    <button 
                      onClick={() => {
                        if (latestDataRef.current) syncLatestData();
                        else loadData(false);
                      }} 
                      className="flex items-center gap-2 px-5 py-2.5 text-xs font-black text-[#38bdf8] hover:bg-[#38bdf8]/10 rounded-xl transition-all border border-[#38bdf8]/20 disabled:opacity-50 uppercase tracking-widest" 
                      disabled={loading}
                    >
                      <RefreshCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
                      <span className="hidden md:inline">{loading ? 'LOADING...' : 'SYNC NOW'}</span>
                    </button>
                  </div>
                )}
                <button onClick={() => setIsSettingsOpen(true)} className="p-3 bg-[#1e232d] text-[#e5e7eb] hover:bg-[#2a313d] rounded-2xl border border-white/5 transition-all shadow-xl">
                  <Menu className="w-6 h-6" />
                </button>
              </div>
          </div>
        </div>
      </header>

      <main className="w-full px-4 sm:px-6 lg:px-8 mt-10">
        {error && (
          <div className="mb-10 p-6 bg-[#ef4444]/10 border border-[#ef4444]/30 rounded-2xl flex items-start gap-4">
            <AlertCircle className="w-6 h-6 text-[#ef4444] shrink-0 mt-1" />
            <div className="flex-1">
              <h4 className="text-sm font-bold text-[#ef4444] uppercase tracking-widest">System Error</h4>
              <p className="text-sm text-[#9ca3af] mt-2 font-medium">{error}</p>
              <button 
                onClick={() => loadData(false)}
                className="mt-4 px-4 py-2 bg-[#ef4444]/20 hover:bg-[#ef4444]/30 text-[#ef4444] text-xs font-bold rounded-lg transition-all uppercase tracking-widest flex items-center gap-2"
              >
                <RefreshCw className="w-3 h-3" />
                Retry Sync
              </button>
            </div>
          </div>
        )}

        {(!config.url && !loading) ? (
          <div className="py-32 flex flex-col items-center justify-center text-center">
            <div className="w-24 h-24 bg-[#161a23] rounded-3xl flex items-center justify-center mb-8 border border-white/5 shadow-2xl">
              <Database className="w-10 h-10 text-[#38bdf8]" />
            </div>
            <h2 className="text-3xl font-black text-white mb-4 uppercase tracking-tighter">No Stream Detected</h2>
            <p className="text-[#9ca3af] max-sm mx-auto mb-10 text-sm leading-relaxed">
              Connect to a valid Google Sheet to initialize analytical rendering.
            </p>
            <button onClick={() => setIsSettingsOpen(true)} className="px-10 py-4 bg-[#38bdf8] hover:bg-[#0ea5e9] text-white font-bold rounded-2xl shadow-xl transition-all">
              OPEN CONFIG
            </button>
          </div>
        ) : (
          <div className="animate-in fade-in duration-700 space-y-10 w-full">
            <MemoizedFilterSection 
              batches={allUniqueBatches} 
              selectedBatches={selectedBatches} 
              setSelectedBatches={handleSetSelectedBatches} 
              dateRange={dateRange} 
              setDateRange={handleSetDateRange}
              uidSearch={uidSearch}
              setUidSearch={handleSetUidSearch}
              loading={loading}
            />
            <MemoizedKPIGrid 
              stats={stats} 
              loading={loading} 
              onRejectedClick={handleOpenRejectionModal} 
              onAcceptedClick={handleOpenAcceptedModal}
              onWipClick={handleOpenWipModal}
              filteredData={filteredData}
              mapping={config.mapping || DEFAULT_MAPPING}
            />
            
            {data.length > 0 && (
              <>
                <div className="flex items-center justify-between px-6 py-4 bg-[#161a23] rounded-2xl border border-white/5">
                  <div className="flex items-center gap-6">
                    <p className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-[0.2em] mono">
                      Source: <span className="text-[#38bdf8]">{config.sheetName}</span>
                    </p>
                    <p className="hidden md:block text-[10px] font-bold text-[#9ca3af] uppercase tracking-[0.2em] mono border-l border-white/10 pl-6">
                      Last synced: <span className="text-white">{lastSyncTime.current.toLocaleTimeString()}</span>
                    </p>
                  </div>
                  <p className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-[0.2em] mono">
                    Total Records: <span className="text-white bg-white/5 px-2 py-0.5 rounded ml-2">{filteredData.length}</span>
                  </p>
                </div>
                
                <div style={{ contentVisibility: 'auto' }}>
                  <MemoizedSKUDetailsSection skuDetails={skuDetails} loading={loading} />
                </div>
                
                <div className="pb-10 min-h-[500px]" style={{ contentVisibility: 'auto' }}>
                  <MemoizedRejectionDetailsSection 
                    filteredData={filteredData} 
                    allData={data} 
                    mapping={config.mapping || DEFAULT_MAPPING} 
                    headers={headers} 
                    loading={loading}
                  />
                </div>
              </>
            )}
          </div>
        )}
      </main>

      {showCopyToast && (
        <div className="fixed bottom-8 left-1/2 -translate-x-1/2 z-50 animate-in slide-in-from-bottom-4 fade-in duration-300">
          <div className="px-8 py-4 bg-[#161a23] border border-[#22c55e]/30 rounded-2xl shadow-2xl flex items-center gap-3 backdrop-blur-xl">
            <Check className="w-5 h-5 text-[#22c55e]" />
            <span className="text-sm font-bold text-white uppercase tracking-widest">Report Copied to Clipboard</span>
          </div>
        </div>
      )}

      <MemoizedSettingsMenu config={{...config, mapping: config.mapping || DEFAULT_MAPPING}} headers={headers} onUpdate={handleConfigUpdate} isOpen={isSettingsOpen} setIsOpen={setIsSettingsOpen} isRefreshing={loading} />
      
      <MemoizedRejectionDrilldownModal 
        isOpen={isRejectionModalOpen} 
        onClose={handleCloseRejectionModal} 
        data={filteredData} 
        mapping={config.mapping || DEFAULT_MAPPING} 
      />

      <MemoizedAcceptedDrilldownModal 
        isOpen={isAcceptedModalOpen} 
        onClose={handleCloseAcceptedModal} 
        data={filteredData} 
        mapping={config.mapping || DEFAULT_MAPPING} 
      />

      <MemoizedWipDrilldownModal 
        isOpen={isWipModalOpen}
        onClose={handleCloseWipModal}
        data={filteredData}
        headers={headers}
      />
    </div>
    </ErrorBoundary>
  );
};

export default App;