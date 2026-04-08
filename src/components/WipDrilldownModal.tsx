import React, { useState, useMemo, memo } from 'react';
import { X, Search, Database, List, Package, Hash } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { DashboardRow, RemainingQtyItem } from '../types';

interface WipDrilldownModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: DashboardRow[];
  headers: string[];
}

const WipDrilldownModal: React.FC<WipDrilldownModalProps> = ({ isOpen, onClose, data, headers }) => {
  const [searchTerm, setSearchTerm] = useState('');

  const remainingQtyData = useMemo<RemainingQtyItem[] | null>(() => {
    const hasSkus = headers.includes('SKUs');
    const hasRemainingQty = headers.includes('REMAINING QTY');

    if (!hasSkus || !hasRemainingQty) return null;

    // Group by SKU to avoid duplicates if filteredData has multiple entries for same SKU
    // and we want a unique list like the original table seemed to imply
    const skuMap: Record<string, number> = {};
    
    data.forEach(row => {
      const sku = String(row['SKUs'] || '').trim();
      const qty = Number(row['REMAINING QTY']) || 0;
      if (sku) {
        // If the sheet has multiple rows for the same SKU, 
        // we take the last one or sum them? 
        // The original logic just mapped them 1:1.
        // Let's stick to the original logic of mapping but maybe unique SKUs is better for a summary.
        // Actually, the original logic was: data.map(...).filter(...)
        // If 'data' had 1000 rows, it showed 1000 rows.
        skuMap[sku] = qty; 
      }
    });

    return Object.entries(skuMap).map(([sku, qty]) => ({ sku, qty }));
  }, [data, headers]);

  const filteredItems = useMemo(() => {
    if (!remainingQtyData) return [];
    if (!searchTerm) return remainingQtyData;
    return remainingQtyData.filter(item => 
      item.sku.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [remainingQtyData, searchTerm]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 sm:p-6">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-[#0f1117]/80 backdrop-blur-md"
      />
      
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 20 }}
        className="relative w-full max-w-3xl bg-[#161a23] border border-white/10 rounded-[32px] shadow-2xl overflow-hidden flex flex-col max-h-[85vh]"
      >
        {/* Header */}
        <div className="p-6 border-b border-white/5 flex items-center justify-between bg-[#1c212c]">
          <div className="flex items-center gap-4">
            <div className="p-3 bg-[#a855f7]/10 rounded-2xl border border-[#a855f7]/20">
              <Package className="w-6 h-6 text-[#a855f7]" />
            </div>
            <div>
              <h2 className="text-xl font-black text-white uppercase tracking-tight">WIP Inventory Details</h2>
              <p className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-[0.3em] mono">
                SKU Remaining Quantities
              </p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <div className="relative hidden sm:block group">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[#9ca3af] group-focus-within:text-[#a855f7] transition-colors" />
              <input
                type="text"
                placeholder="Search SKUs..."
                className="pl-10 pr-4 py-2 bg-[#0f1117] border border-white/5 rounded-xl text-xs text-white focus:ring-2 focus:ring-[#a855f7]/50 outline-none transition-all placeholder:text-[#9ca3af]/50 mono w-48"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            <button 
              onClick={onClose}
              className="p-2 hover:bg-white/5 rounded-xl transition-colors text-[#9ca3af] hover:text-white"
            >
              <X className="w-6 h-6" />
            </button>
          </div>
        </div>

        {/* Mobile Search */}
        <div className="p-4 border-b border-white/5 bg-[#161a23] sm:hidden">
          <div className="relative group">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[#9ca3af] group-focus-within:text-[#a855f7] transition-colors" />
            <input
              type="text"
              placeholder="Search SKUs..."
              className="w-full pl-10 pr-4 py-2.5 bg-[#0f1117] border border-white/5 rounded-xl text-sm text-white focus:ring-2 focus:ring-[#a855f7]/50 outline-none transition-all placeholder:text-[#9ca3af]/50 mono"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-6 custom-scrollbar">
          {!remainingQtyData ? (
            <div className="py-20 flex flex-col items-center justify-center text-[#ef4444] bg-[#ef4444]/5 rounded-3xl border border-[#ef4444]/20">
              <Database className="w-12 h-12 mb-4 opacity-50" />
              <p className="text-xs font-black uppercase tracking-[0.3em] text-center px-10">
                Required columns "SKUs" and "REMAINING QTY" not found in the data stream.
              </p>
            </div>
          ) : (
            <div className="overflow-hidden rounded-2xl border border-white/5 bg-[#0f1117]/30">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-white/[0.02] border-b border-white/5">
                    <th className="px-6 py-4 text-[10px] font-black text-[#9ca3af] uppercase tracking-widest">SKU Identifier</th>
                    <th className="px-6 py-4 text-[10px] font-black text-[#9ca3af] uppercase tracking-widest text-right">Remaining Qty</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-white/5">
                  {filteredItems.length > 0 ? (
                    filteredItems.map((item, idx) => (
                      <tr 
                        key={idx} 
                        className="hover:bg-white/[0.03] transition-all group"
                      >
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 flex items-center justify-center bg-white/5 rounded-lg text-[#9ca3af] group-hover:text-[#a855f7] transition-colors">
                              <Hash className="w-4 h-4" />
                            </div>
                            <span className="text-sm font-black text-white mono group-hover:text-[#a855f7] transition-colors">
                              {item.sku}
                            </span>
                          </div>
                        </td>
                        <td className={`px-6 py-4 text-sm font-black text-right mono ${item.qty < 10 ? 'text-[#ef4444]' : 'text-[#e5e7eb]'}`}>
                          {item.qty.toLocaleString()}
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={2} className="px-6 py-20 text-center text-[#9ca3af]/40 italic uppercase text-[10px] font-black tracking-[0.3em]">
                        No matching SKU records found
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="p-6 bg-[#1c212c] border-t border-white/5 flex justify-between items-center">
          <div className="flex items-center gap-4">
            <p className="text-[10px] font-bold text-[#4b5563] uppercase tracking-widest mono">
              Unique SKUs: <span className="text-[#a855f7]">{remainingQtyData?.length || 0}</span>
            </p>
            {filteredItems.length !== (remainingQtyData?.length || 0) && (
              <p className="text-[10px] font-bold text-[#4b5563] uppercase tracking-widest mono border-l border-white/10 pl-4">
                Filtered: <span className="text-white">{filteredItems.length}</span>
              </p>
            )}
          </div>
          <div className="flex items-center gap-2 text-[9px] font-black text-[#a855f7] uppercase tracking-widest">
            <List className="w-3 h-3" /> Inventory Sync Active
          </div>
        </div>
      </motion.div>
    </div>
  );
};

export default memo(WipDrilldownModal);
