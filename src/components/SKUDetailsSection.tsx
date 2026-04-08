import React, { useState, useMemo, memo } from 'react';
import { Search, ChevronUp, ChevronDown, Table as TableIcon, Minimize2, Maximize2 } from 'lucide-react';
import { SKUDetail } from '../types';
import { useDebounce } from '../hooks/useDebounce';
import { TableVirtuoso } from 'react-virtuoso';

interface SKUDetailsSectionProps {
  skuDetails: SKUDetail[];
  loading?: boolean;
}

type SortConfig = {
  key: keyof SKUDetail;
  direction: 'asc' | 'desc';
};

const SKUDetailsSection: React.FC<SKUDetailsSectionProps> = ({ skuDetails, loading }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const debouncedSearch = useDebounce(searchTerm, 300);
  const [sortConfig, setSortConfig] = useState<SortConfig>({ key: 'total', direction: 'desc' });
  const [isMinimized, setIsMinimized] = useState(false);

  const handleSort = (key: keyof SKUDetail) => {
    setSortConfig(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'desc' ? 'asc' : 'desc'
    }));
  };

  const filteredAndSortedData = useMemo(() => {
    let data = [...skuDetails];

    if (debouncedSearch) {
      const lowerSearch = debouncedSearch.toLowerCase();
      data = data.filter(item => 
        item.sku.toLowerCase().includes(lowerSearch)
      );
    }

    data.sort((a, b) => {
      const aVal = a[sortConfig.key];
      const bVal = b[sortConfig.key];
      
      if (typeof aVal === 'string' && typeof bVal === 'string') {
        return sortConfig.direction === 'asc' 
          ? aVal.localeCompare(bVal) 
          : bVal.localeCompare(aVal);
      }
      
      return sortConfig.direction === 'asc' 
        ? (aVal as number) - (bVal as number) 
        : (bVal as number) - (aVal as number);
    });

    return data;
  }, [skuDetails, debouncedSearch, sortConfig]);

  const SortIndicator = ({ column }: { column: keyof SKUDetail }) => {
    if (sortConfig.key !== column) return <div className="w-4 h-4 opacity-0" />;
    return sortConfig.direction === 'asc' 
      ? <ChevronUp className="w-4 h-4 text-[#38bdf8]" /> 
      : <ChevronDown className="w-4 h-4 text-[#38bdf8]" />;
  };

  return (
    <div className="premium-card p-8 animate-in fade-in duration-700 hover:scale-[1.01] transition-transform">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-8">
        <div className="flex items-center justify-between w-full md:w-auto gap-4">
          <div className="flex items-center gap-4">
            <div className="w-1.5 h-6 bg-[#38bdf8] rounded-full shadow-[0_0_15px_rgba(56,189,248,0.5)]" />
            <h3 className="text-sm font-black text-[#e5e7eb] uppercase tracking-[0.2em] mono">
              SKU PERFORMANCE
            </h3>
          </div>
          <button 
            onClick={() => setIsMinimized(!isMinimized)}
            className="p-2 hover:bg-white/5 rounded-lg transition-colors text-[#9ca3af] hover:text-[#38bdf8]"
            title={isMinimized ? "Maximize" : "Minimize"}
          >
            {isMinimized ? <Maximize2 className="w-4 h-4" /> : <Minimize2 className="w-4 h-4" />}
          </button>
        </div>

        {!isMinimized && (
          <div className="relative w-full md:w-64 group">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[#9ca3af] group-focus-within:text-[#38bdf8] transition-colors" />
            <input
              type="text"
              placeholder="Search SKU..."
              className="w-full pl-10 pr-4 py-2.5 bg-[#0f1117] border border-white/5 rounded-xl text-sm text-white focus:ring-2 focus:ring-[#38bdf8]/50 outline-none transition-all placeholder:text-[#9ca3af]/50 mono"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
        )}
      </div>

      {!isMinimized && (
        <div className="rounded-xl border border-white/5 bg-[#0f1117]/30 overflow-hidden">
          <div className="w-full overflow-x-auto">
            {loading && skuDetails.length === 0 ? (
              <div className="p-8 space-y-4">
                {[...Array(5)].map((_, i) => (
                  <div key={i} className="flex gap-4">
                    <div className="h-10 flex-1 shimmer rounded-lg opacity-10" />
                    <div className="h-10 w-24 shimmer rounded-lg opacity-10" />
                    <div className="h-10 w-24 shimmer rounded-lg opacity-10" />
                    <div className="h-10 w-24 shimmer rounded-lg opacity-10" />
                    <div className="h-10 w-32 shimmer rounded-lg opacity-10" />
                  </div>
                ))}
              </div>
            ) : (
              <TableVirtuoso
                style={{ height: filteredAndSortedData.length > 0 ? Math.min(filteredAndSortedData.length * 60 + 56, 600) : 200 }}
              data={filteredAndSortedData}
              fixedHeaderContent={() => (
                <tr className="bg-[#161a23] border-b border-white/5">
                  {(['sku', 'total', 'accepted', 'rejected', 'yield'] as const).map((key) => (
                    <th 
                      key={key}
                      onClick={() => handleSort(key)}
                      className="px-6 py-4 text-[10px] font-black text-[#9ca3af] uppercase tracking-widest cursor-pointer hover:bg-white/[0.05] transition-colors"
                    >
                      <div className="flex items-center gap-2">
                        {key === 'sku' ? 'SKU ID' : key.replace(/([A-Z])/g, ' $1')}
                        <SortIndicator column={key} />
                      </div>
                    </th>
                  ))}
                </tr>
              )}
              itemContent={(idx, item) => (
                <>
                  <td className="px-6 py-4 text-sm font-black text-white mono group-hover:text-[#38bdf8] transition-colors border-b border-white/5">{item.sku}</td>
                  <td className="px-6 py-4 text-sm font-medium text-[#9ca3af] mono border-b border-white/5">{item.total.toLocaleString()}</td>
                  <td className="px-6 py-4 text-sm font-bold text-[#22c55e] mono border-b border-white/5">{item.accepted.toLocaleString()}</td>
                  <td className="px-6 py-4 text-sm font-bold text-[#ef4444] mono border-b border-white/5">{item.rejected.toLocaleString()}</td>
                  <td className="px-6 py-4 border-b border-white/5">
                    <div className="flex items-center gap-3">
                      <div className="flex-1 h-1.5 bg-white/5 rounded-full overflow-hidden max-w-[120px] hidden sm:block">
                        <div 
                          className={`h-full transition-all duration-1000 ease-out ${item.yield > 95 ? 'bg-[#22c55e]' : item.yield > 80 ? 'bg-amber-500' : 'bg-[#ef4444]'}`}
                          style={{ 
                            width: `${item.yield}%`,
                            boxShadow: `0 0 10px ${item.yield > 95 ? '#22c55e' : item.yield > 80 ? '#f59e0b' : '#ef4444'}40`
                          }}
                        />
                      </div>
                      <span className={`text-sm font-black mono w-12 text-right ${item.yield > 95 ? 'text-[#22c55e]' : item.yield > 80 ? 'text-amber-400' : 'text-[#ef4444]'}`}>
                        {item.yield.toFixed(1)}%
                      </span>
                    </div>
                  </td>
                </>
              )}
              components={{
                Table: (props) => <table {...props} className="w-full text-left border-collapse" />,
                TableRow: (props) => <tr {...props} className="hover:bg-white/[0.03] transition-all group" />,
                EmptyPlaceholder: () => (
                  <div className="px-6 py-12 text-center text-[#9ca3af]/40 italic uppercase text-[10px] font-black tracking-[0.3em]">
                    {loading ? 'Fetching SKU Performance Data...' : 'No matching records found'}
                  </div>
                )
              }}
            />
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default memo(SKUDetailsSection);
