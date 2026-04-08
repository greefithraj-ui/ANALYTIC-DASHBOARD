import React, { useState, useMemo, memo } from 'react';
import { Search, AlertTriangle, ChevronUp, ChevronDown } from 'lucide-react';
import { DashboardRow, ColumnMapping } from '../types';
import TrendAnalysisChart from './TrendAnalysisChart';
import SKUTrendAnalysisChart from './SKUTrendAnalysisChart';
import { useDebounce } from '../hooks/useDebounce';
import { TableVirtuoso } from 'react-virtuoso';

interface RejectionDetailsSectionProps {
  filteredData: DashboardRow[];
  allData: DashboardRow[];
  mapping: ColumnMapping;
  headers: string[];
  loading?: boolean;
}

interface RejectionReasonGroup {
  reason: string;
  count: number;
}

const RejectionDetailsSection: React.FC<RejectionDetailsSectionProps> = ({ filteredData, allData, mapping, headers, loading }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const debouncedSearch = useDebounce(searchTerm, 300);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('desc');

  const rejectionGroups = useMemo(() => {
    if (!mapping) return null;
    const reasonKey = mapping.reason;
    const statusKey = mapping.ringStatus;

    if (!headers.includes(reasonKey) || !headers.includes(statusKey)) {
      return null;
    }

    const groups: Record<string, number> = {};
    let totalRejected = 0;

    filteredData.forEach(row => {
      const status = String(row[statusKey] || '').trim().toLowerCase();
      if (['rejected', 'nok', 'fail', '0', 'false', 'no'].includes(status)) {
        const reason = String(row[reasonKey] || '').trim();
        if (reason !== '') {
          groups[reason] = (groups[reason] || 0) + 1;
          totalRejected++;
        }
      }
    });

    const result: RejectionReasonGroup[] = Object.entries(groups).map(([reason, count]) => ({
      reason,
      count
    }));

    return { 
      data: result,
      totalCount: totalRejected
    };
  }, [filteredData, mapping, headers]);

  const filteredAndSorted = useMemo(() => {
    if (!rejectionGroups) return [];
    let list = [...rejectionGroups.data];
    if (debouncedSearch) {
      const lowerSearch = debouncedSearch.toLowerCase();
      list = list.filter(item => item.reason.toLowerCase().includes(lowerSearch));
    }
    list.sort((a, b) => sortDirection === 'desc' ? b.count - a.count : a.count - b.count);
    return list;
  }, [rejectionGroups, debouncedSearch, sortDirection]);

  const maxCount = useMemo(() => {
    if (!rejectionGroups || rejectionGroups.data.length === 0) return 0;
    return Math.max(...rejectionGroups.data.map(i => i.count));
  }, [rejectionGroups]);

  if (!rejectionGroups) {
    return (
      <div className="premium-card p-8 flex items-center justify-center h-full">
        <div className="flex items-center gap-4 text-[#ef4444]">
          <AlertTriangle className="w-5 h-5" />
          <h3 className="text-sm font-bold uppercase tracking-widest mono text-center">REASON column missing.</h3>
        </div>
      </div>
    );
  }

  return (
    <div className="premium-card p-8 flex flex-col h-full animate-in fade-in duration-700 hover:scale-[1.01] transition-transform">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-8">
        <div className="flex items-center gap-4">
          <div className="w-1.5 h-6 bg-[#ef4444] rounded-full shadow-[0_0_15px_rgba(239,68,68,0.5)]" />
          <h3 className="text-sm font-black text-[#e5e7eb] uppercase tracking-[0.2em] mono">
            REJECTION ANALYSIS
          </h3>
        </div>

        <div className="relative w-full md:w-64 group">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[#9ca3af] group-focus-within:text-[#ef4444] transition-colors" />
          <input
            type="text"
            placeholder="Search reasons..."
            className="w-full pl-10 pr-4 py-2.5 bg-[#0f1117] border border-white/5 rounded-xl text-sm text-white focus:ring-2 focus:ring-[#ef4444]/50 outline-none transition-all placeholder:text-[#9ca3af]/50 mono"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
      </div>

      <div className="flex-1 overflow-hidden">
        <div className="rounded-xl border border-white/5 bg-[#0f1117]/30 overflow-hidden">
          <div className="w-full overflow-x-auto">
            {loading && filteredData.length === 0 ? (
              <div className="p-8 space-y-4">
                {[...Array(3)].map((_, i) => (
                  <div key={i} className="flex gap-4">
                    <div className="h-10 flex-1 shimmer rounded-lg opacity-10" />
                    <div className="h-10 w-24 shimmer rounded-lg opacity-10" />
                  </div>
                ))}
              </div>
            ) : (
              <TableVirtuoso
                style={{ height: filteredAndSorted.length > 0 ? Math.min(filteredAndSorted.length * 56 + 56, 400) : 200 }}
              data={filteredAndSorted}
              fixedHeaderContent={() => (
                <tr className="bg-[#161a23] border-b border-white/5">
                  <th className="px-6 py-4 text-[10px] font-black text-[#9ca3af] uppercase tracking-widest">Rejection Pattern</th>
                  <th 
                    className="px-6 py-4 text-[10px] font-black text-[#9ca3af] uppercase tracking-widest text-right cursor-pointer hover:bg-white/[0.05] transition-colors"
                    onClick={() => setSortDirection(prev => prev === 'desc' ? 'asc' : 'desc')}
                  >
                    <div className="flex items-center justify-end gap-2">
                      Count
                      {sortDirection === 'desc' ? <ChevronDown className="w-3 h-3 text-[#ef4444]" /> : <ChevronUp className="w-3 h-3 text-[#ef4444]" />}
                    </div>
                  </th>
                </tr>
              )}
              itemContent={(idx, item) => (
                <>
                  <td className="px-6 py-4 text-sm font-black text-white mono group-hover:text-[#ef4444] transition-colors border-b border-white/5">
                    <div className="flex items-center gap-3">
                      {item.count === maxCount && item.count > 0 && (
                        <div className="w-2 h-2 bg-[#ef4444] rounded-full animate-pulse shadow-[0_0_10px_#ef4444]" />
                      )}
                      {item.reason}
                    </div>
                  </td>
                  <td className={`px-6 py-4 text-sm font-black text-right mono border-b border-white/5 ${item.count === maxCount ? 'text-[#ef4444]' : 'text-[#e5e7eb]'}`}>
                    {item.count.toLocaleString()}
                  </td>
                </>
              )}
              components={{
                Table: (props) => <table {...props} className="w-full text-left border-collapse" />,
                TableRow: (props) => <tr {...props} className="hover:bg-white/[0.03] transition-all group" />,
                EmptyPlaceholder: () => (
                  <div className="px-6 py-12 text-center text-[#9ca3af]/30 italic text-[10px] font-black uppercase tracking-[0.3em]">
                    {loading ? 'Analyzing Rejection Patterns...' : 'No rejection data recorded'}
                  </div>
                )
              }}
            />
            )}
          </div>
          {rejectionGroups.totalCount > 0 && (
            <div className="bg-white/[0.03] border-t border-white/5 flex justify-between items-center px-6 py-4">
              <span className="text-[9px] font-black text-[#9ca3af] uppercase tracking-widest mono">TOTAL NOK DETECTED</span>
              <span className="text-sm font-black text-[#ef4444] mono">{rejectionGroups.totalCount.toLocaleString()}</span>
            </div>
          )}
        </div>
      </div>
      
      <TrendAnalysisChart data={filteredData} mapping={mapping} />
      <SKUTrendAnalysisChart data={filteredData} mapping={mapping} />
    </div>
  );
};

export default memo(RejectionDetailsSection);
