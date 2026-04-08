import React, { useState, useEffect, useRef } from 'react';
import { createPortal } from 'react-dom';
import { Calendar, Filter, Check, ChevronDown, Clock, Target, RotateCcw, X, Search } from 'lucide-react';
import { format, subDays, startOfMonth, endOfMonth, parseISO, startOfDay } from 'date-fns';

interface FilterSectionProps {
  batches: string[];
  selectedBatches: string[];
  setSelectedBatches: (batches: string[]) => void;
  dateRange: { start: Date | null; end: Date | null };
  setDateRange: (range: { start: Date | null; end: Date | null }) => void;
  uidSearch: string;
  setUidSearch: (search: string) => void;
  loading?: boolean;
}

const FilterSection: React.FC<FilterSectionProps> = ({
  batches = [],
  selectedBatches = [],
  setSelectedBatches,
  dateRange,
  setDateRange,
  uidSearch,
  setUidSearch,
  loading = false
}) => {
  console.log('FilterSection: Rendered with batches:', batches?.length);
  const [isBatchOpen, setIsBatchOpen] = useState(false);
  const [showResetToast, setShowResetToast] = useState(false);
  const [localSearch, setLocalSearch] = useState(uidSearch);
  const batchButtonRef = useRef<HTMLButtonElement>(null);
  const [dropdownPos, setDropdownPos] = useState({ top: 0, left: 0, width: 0 });

  // Update dropdown position when opened and on scroll/resize
  useEffect(() => {
    const updatePos = () => {
      if (isBatchOpen && batchButtonRef.current) {
        const rect = batchButtonRef.current.getBoundingClientRect();
        setDropdownPos({
          top: rect.bottom,
          left: rect.left,
          width: rect.width
        });
      }
    };

    updatePos();
    if (isBatchOpen) {
      window.addEventListener('scroll', updatePos, true);
      window.addEventListener('resize', updatePos);
    }
    return () => {
      window.removeEventListener('scroll', updatePos, true);
      window.removeEventListener('resize', updatePos);
    };
  }, [isBatchOpen]);

  // Debounce search input
  useEffect(() => {
    const timer = setTimeout(() => {
      setUidSearch(localSearch);
    }, 300); // 300ms debounce as requested
    return () => clearTimeout(timer);
  }, [localSearch, setUidSearch]);

  // Sync local search with prop if prop changes externally (e.g. on clear all)
  useEffect(() => {
    setLocalSearch(uidSearch);
  }, [uidSearch]);

  const presets = [
    { label: 'ALL TIME', value: 'all' },
    { label: 'YESTERDAY', value: 'yesterday' },
    { label: 'TODAY', value: 'today' },
    { label: '7 DAYS', value: '7d' },
    { label: '30 DAYS', value: '30d' },
    { label: 'MONTH', value: 'month' },
  ];

  const parseInputDate = (val: string) => {
    if (!val) return null;
    
    // 1. Handle YYYY-MM-DD (Standard for HTML5 date input)
    const isoMatch = val.match(/^(\d{1,4})-(\d{1,2})-(\d{1,2})$/);
    if (isoMatch) {
      const y = parseInt(isoMatch[1], 10);
      const m = parseInt(isoMatch[2], 10);
      const d = parseInt(isoMatch[3], 10);
      
      if (!isNaN(y) && !isNaN(m) && !isNaN(d)) {
        // Use local time construction to match browser date input behavior
        const date = new Date(y, m - 1, d);
        date.setHours(0, 0, 0, 0);
        return isNaN(date.getTime()) ? null : date;
      }
    }
    
    // 2. Handle DD-MM-YYYY or MM-DD-YYYY or other common formats (Manual entry / Fallback)
    const parts = val.split(/[-/.]/);
    if (parts.length === 3) {
      const p1 = parseInt(parts[0], 10);
      const p2 = parseInt(parts[1], 10);
      const p3 = parseInt(parts[2], 10);
      
      if (parts[2].length === 4) {
        // Year is at the end (DD-MM-YYYY or MM-DD-YYYY)
        const y = p3;
        let m, d;
        if (p1 > 12) { // DD-MM-YYYY
          d = p1; m = p2;
        } else if (p2 > 12) { // MM-DD-YYYY
          m = p1; d = p2;
        } else {
          // Ambiguous, assume DD-MM-YYYY as it's common in many locales
          d = p1; m = p2;
        }
        
        const date = new Date(y, m - 1, d);
        date.setHours(0, 0, 0, 0);
        return isNaN(date.getTime()) ? null : date;
      }
      
      if (parts[0].length === 4) {
        // Year is at the start (YYYY-MM-DD)
        const y = p1;
        const m = p2;
        const d = p3;
        const date = new Date(y, m - 1, d);
        date.setHours(0, 0, 0, 0);
        return isNaN(date.getTime()) ? null : date;
      }
    }

    // 3. Fallback to native Date parsing
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }

    return null;
  };

  const handlePreset = (type: string) => {
    const today = startOfDay(new Date());
    switch (type) {
      case 'yesterday': {
        const yesterday = subDays(today, 1);
        setDateRange({ start: yesterday, end: yesterday });
        break;
      }
      case 'today': setDateRange({ start: today, end: today }); break;
      case '7d': setDateRange({ start: subDays(today, 7), end: today }); break;
      case '30d': setDateRange({ start: subDays(today, 30), end: today }); break;
      case 'month': setDateRange({ start: startOfMonth(today), end: endOfMonth(today) }); break;
      default: setDateRange({ start: null, end: null });
    }
  };

  const handleFilterOnly = (batch: string) => {
    setSelectedBatches([batch]);
    setIsBatchOpen(false);
  };

  const handleClearAll = () => {
    setDateRange({ start: null, end: null });
    setSelectedBatches(batches);
    setUidSearch('');
    setShowResetToast(true);
    setTimeout(() => setShowResetToast(false), 2000);
  };

  return (
    <div className="relative z-50 overflow-visible mb-8 bg-[#161a23] p-8 rounded-2xl border border-white/5 shadow-2xl animate-in fade-in duration-500">
      {/* Header with Search and Clear All Button */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-8">
        <div className="flex items-center gap-3">
          <Filter className="w-4 h-4 text-[#38bdf8]" />
          <h2 className="text-xs font-black text-white uppercase tracking-[0.2em] mono">Filter Panel</h2>
        </div>
        
        <div className="flex flex-wrap items-center gap-4 w-full md:w-auto">
          {/* UID Search Bar */}
          <div className="relative flex-1 md:w-64 min-w-[200px]">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[#38bdf8]" />
            <input
              type="text"
              placeholder="Search UID..."
              className="w-full pl-10 pr-10 py-2.5 bg-[#0f1117] border border-white/5 rounded-xl text-sm text-white focus:ring-2 focus:ring-[#38bdf8]/50 outline-none transition-all placeholder:text-[#9ca3af]/50 mono"
              value={localSearch}
              onChange={(e) => setLocalSearch(e.target.value)}
            />
            {localSearch && (
              <button
                onClick={() => setLocalSearch('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 p-1 text-[#9ca3af] hover:text-[#ef4444] transition-colors"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          <button
            onClick={handleClearAll}
            className="flex items-center gap-2 px-4 py-2.5 text-[10px] font-black text-[#9ca3af] hover:text-white bg-[#0f1117] hover:bg-[#ef4444]/10 border border-white/5 hover:border-[#ef4444]/30 rounded-xl transition-all uppercase tracking-widest"
          >
            <RotateCcw className="w-3 h-3" />
            CLEAR ALL
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-end">
        <div className="lg:col-span-8 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="space-y-3 lg:col-span-2">
            <label className="text-[11px] font-bold text-[#9ca3af] uppercase tracking-widest flex items-center gap-2">
               PRESETS
            </label>
            <div className="flex flex-wrap gap-2">
              {presets.map(p => (
                <button
                  key={p.value}
                  onClick={() => handlePreset(p.value)}
                  className="px-3 py-2 text-[10px] font-black rounded-xl border border-white/5 bg-[#0f1117] hover:bg-[#38bdf8]/10 hover:border-[#38bdf8]/30 transition-all text-[#e5e7eb] uppercase tracking-tighter"
                >
                  {p.label}
                </button>
              ))}
            </div>
          </div>
          
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-[11px] font-bold text-[#9ca3af] uppercase tracking-widest flex items-center gap-2">
                 START DATE
              </label>
              {dateRange.start && (
                <button 
                  onClick={() => setDateRange({ ...dateRange, start: null })}
                  className="text-[10px] text-[#ef4444] hover:underline uppercase font-bold"
                >
                  Clear
                </button>
              )}
            </div>
            <input
              type="date"
              className="w-full px-4 py-2.5 text-sm border border-white/5 rounded-xl focus:ring-2 focus:ring-[#38bdf8]/50 outline-none bg-[#0f1117] text-white mono"
              value={dateRange.start ? format(dateRange.start, 'yyyy-MM-dd') : ''}
              onChange={(e) => setDateRange({ ...dateRange, start: parseInputDate(e.target.value) })}
            />
          </div>

          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-[11px] font-bold text-[#9ca3af] uppercase tracking-widest flex items-center gap-2">
                 END DATE
              </label>
              {dateRange.end && (
                <button 
                  onClick={() => setDateRange({ ...dateRange, end: null })}
                  className="text-[10px] text-[#ef4444] hover:underline uppercase font-bold"
                >
                  Clear
                </button>
              )}
            </div>
            <input
              type="date"
              className="w-full px-4 py-2.5 text-sm border border-white/5 rounded-xl focus:ring-2 focus:ring-[#38bdf8]/50 outline-none bg-[#0f1117] text-white mono"
              value={dateRange.end ? format(dateRange.end, 'yyyy-MM-dd') : ''}
              onChange={(e) => setDateRange({ ...dateRange, end: parseInputDate(e.target.value) })}
            />
          </div>
        </div>

        <div className="lg:col-span-4 space-y-3">
          <label className="text-[11px] font-bold text-[#9ca3af] uppercase tracking-widest flex items-center gap-2">
             BATCH SELECTOR
          </label>
          <div className="relative">
            <button
              ref={batchButtonRef}
              type="button"
              disabled={loading}
              onMouseDown={(e) => {
                e.stopPropagation();
                if (!loading) setIsBatchOpen(!isBatchOpen);
              }}
              className={`relative z-30 w-full flex items-center justify-between px-5 py-2.5 text-sm border border-white/5 rounded-xl bg-[#0f1117] hover:border-[#38bdf8]/40 transition-colors text-white cursor-pointer ${loading ? 'opacity-50 cursor-not-allowed' : ''}`}
            >
              <span className="truncate font-bold">
                {loading ? 'SYNCING BATCHES...' : (batches.length === 0 ? 'NO BATCHES FOUND' : (selectedBatches.length === batches.length ? 'ALL ACTIVE' : `${selectedBatches.length} SELECTED`))}
              </span>
              <ChevronDown className={`w-4 h-4 text-[#38bdf8] transition-transform duration-200 ${isBatchOpen ? 'rotate-180' : ''} ${loading ? 'animate-pulse' : ''}`} />
            </button>
            
            {isBatchOpen && createPortal(
              <>
                <div 
                  className="fixed inset-0 z-[99998] bg-transparent pointer-events-auto" 
                  onMouseDown={(e) => {
                    e.stopPropagation();
                    setIsBatchOpen(false);
                  }} 
                />
                <div 
                  style={{ 
                    position: 'fixed',
                    top: dropdownPos.top + 8,
                    left: dropdownPos.left,
                    width: dropdownPos.width,
                    zIndex: 99999
                  }}
                  className="bg-[#161a23] border border-white/10 rounded-2xl shadow-2xl max-h-60 overflow-y-auto p-3 space-y-1 pointer-events-auto animate-in fade-in zoom-in-95 duration-200"
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  {batches.length === 0 ? (
                    <div className="py-4 px-2 text-center">
                      <p className="text-xs text-[#9ca3af] font-medium">No batch data found in the current sheet mapping.</p>
                    </div>
                  ) : (
                    <>
                      <div className="flex gap-2 p-1 mb-2">
                        <button 
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedBatches(batches);
                          }} 
                          className="flex-1 py-2 text-[10px] font-black uppercase tracking-widest text-[#38bdf8] bg-[#38bdf8]/10 rounded-xl hover:bg-[#38bdf8]/20 transition-all border border-[#38bdf8]/20"
                        >
                          All
                        </button>
                        <button 
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedBatches([]);
                          }} 
                          className="flex-1 py-2 text-[10px] font-black uppercase tracking-widest text-[#9ca3af] bg-[#0f1117] rounded-xl hover:bg-white/5 transition-all border border-white/5"
                        >
                          None
                        </button>
                      </div>
                      {batches.map(batch => (
                        <div key={batch} className="group flex items-center px-2 hover:bg-white/5 rounded-xl transition-all">
                          <label className="flex-1 flex items-center py-2.5 cursor-pointer text-sm text-[#e5e7eb]">
                            <input 
                              type="checkbox" 
                              className="sr-only" 
                              checked={selectedBatches.includes(batch)} 
                              onChange={(e) => {
                                e.stopPropagation();
                                if (selectedBatches.includes(batch)) setSelectedBatches(selectedBatches.filter(b => b !== batch));
                                else setSelectedBatches([...selectedBatches, batch]);
                              }} 
                            />
                            <div className={`w-5 h-5 border-2 rounded-lg mr-3 flex items-center justify-center transition-all ${selectedBatches.includes(batch) ? 'bg-[#38bdf8] border-[#38bdf8]' : 'bg-[#0f1117] border-white/10'}`}>
                              {selectedBatches.includes(batch) && <Check className="w-3 h-3 text-[#0f1117] stroke-[4]" />}
                            </div>
                            <span className="font-medium truncate">{batch || '(Blank)'}</span>
                          </label>
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              handleFilterOnly(batch);
                            }}
                            title="Show only this batch"
                            className="p-2 text-[#38bdf8] hover:bg-[#38bdf8]/10 rounded-lg opacity-0 group-hover:opacity-100 transition-all ml-2"
                          >
                            <Target className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      ))}
                    </>
                  )}
                </div>
              </>,
              document.body
            )}
          </div>
        </div>
      </div>

      {/* Internal Reset Notification */}
      {showResetToast && (
        <div className="absolute top-2 right-1/2 translate-x-1/2 z-30 animate-in slide-in-from-top-2 fade-in duration-300">
          <div className="px-4 py-2 bg-[#22c55e]/10 border border-[#22c55e]/30 rounded-full flex items-center gap-2 backdrop-blur-md">
            <Check className="w-3 h-3 text-[#22c55e]" />
            <span className="text-[10px] font-black text-white uppercase tracking-widest">Filters cleared</span>
          </div>
        </div>
      )}
    </div>
  );
};

export default FilterSection;