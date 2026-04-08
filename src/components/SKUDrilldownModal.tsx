import React, { memo } from 'react';
import { X, Package, CheckCircle2, XCircle, Activity, ArrowUpRight } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface SKUDetail {
  accepted: number;
  rejected: number;
}

interface SKUDrilldownModalProps {
  isOpen: boolean;
  onClose: () => void;
  date: string;
  skuData: Record<string, SKUDetail>;
}

const SKUDrilldownModal: React.FC<SKUDrilldownModalProps> = ({ isOpen, onClose, date, skuData }) => {
  const skus = Object.entries(skuData).map(([name, stats]) => {
    const s = stats as SKUDetail;
    const total = s.accepted + s.rejected;
    const yieldRate = total > 0 ? (s.accepted / total) * 100 : 100;
    return { name, ...s, total, yieldRate };
  }).sort((a, b) => b.total - a.total);

  return (
    <AnimatePresence>
      {isOpen && (
        <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 md:p-8">
          {/* Backdrop */}
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={onClose}
            className="absolute inset-0 bg-black/60 backdrop-blur-md"
          />

          {/* Modal Container */}
          <motion.div
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.9, y: 20 }}
            className="relative w-full max-w-5xl max-h-[85vh] bg-[#0f1117] border border-white/10 rounded-3xl shadow-[0_0_50px_rgba(0,0,0,0.5)] flex flex-col overflow-hidden"
          >
            {/* Glow Accents */}
            <div className="absolute top-0 left-1/4 right-1/4 h-px bg-gradient-to-r from-transparent via-[#22c55e]/50 to-transparent" />
            <div className="absolute bottom-0 left-1/4 right-1/4 h-px bg-gradient-to-r from-transparent via-[#38bdf8]/50 to-transparent" />

            {/* Header */}
            <div className="p-6 md:p-8 border-b border-white/5 flex items-center justify-between bg-white/[0.02]">
              <div className="flex items-center gap-4">
                <div className="p-3 bg-[#22c55e]/10 rounded-2xl border border-[#22c55e]/20">
                  <Package className="w-6 h-6 text-[#22c55e]" />
                </div>
                <div>
                  <h2 className="text-xl font-black text-white uppercase tracking-widest mono">SKU Performance Deep-Dive</h2>
                  <p className="text-xs text-[#9ca3af] uppercase tracking-[0.2em] mt-1">Detailed Analytics for <span className="text-[#38bdf8]">{date}</span></p>
                </div>
              </div>
              <button
                onClick={onClose}
                className="p-2 hover:bg-white/5 rounded-xl text-[#9ca3af] hover:text-white transition-all border border-white/5 group"
              >
                <X className="w-6 h-6 group-hover:rotate-90 transition-transform duration-300" />
              </button>
            </div>

            {/* Content - Scrollable Table */}
            <div className="flex-1 overflow-y-auto p-6 md:p-8 custom-scrollbar">
              <div className="overflow-hidden rounded-2xl border border-white/5 bg-[#161a23]/50">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-white/[0.03] border-b border-white/5">
                      <th className="px-6 py-5 text-[10px] font-black text-[#9ca3af] uppercase tracking-[0.2em]">SKU Identifier</th>
                      <th className="px-6 py-5 text-[10px] font-black text-[#9ca3af] uppercase tracking-[0.2em] text-center">Total Output</th>
                      <th className="px-6 py-5 text-[10px] font-black text-[#9ca3af] uppercase tracking-[0.2em] text-center">Accepted (OK)</th>
                      <th className="px-6 py-5 text-[10px] font-black text-[#9ca3af] uppercase tracking-[0.2em] text-center">Rejected (NOK)</th>
                      <th className="px-6 py-5 text-[10px] font-black text-[#9ca3af] uppercase tracking-[0.2em] text-right">Yield Performance</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5">
                    {skus.map((sku, idx) => (
                      <tr key={idx} className="hover:bg-white/[0.02] transition-colors group">
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-3">
                            <div className="w-1.5 h-1.5 rounded-full bg-[#38bdf8] shadow-[0_0_8px_#38bdf8]" />
                            <span className="text-sm font-black text-white mono group-hover:text-[#38bdf8] transition-colors">{sku.name}</span>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <span className="text-sm font-black text-[#e5e7eb] mono">{sku.total.toLocaleString()}</span>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <div className="flex items-center justify-center gap-2">
                            <CheckCircle2 className="w-3 h-3 text-[#22c55e]" />
                            <span className="text-sm font-black text-[#22c55e] mono">{sku.accepted.toLocaleString()}</span>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <div className="flex items-center justify-center gap-2">
                            <XCircle className="w-3 h-3 text-[#ef4444]" />
                            <span className="text-sm font-black text-[#ef4444] mono">{sku.rejected.toLocaleString()}</span>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex flex-col items-end gap-1">
                            <div className="flex items-center gap-2">
                              <Activity className={`w-3 h-3 ${sku.yieldRate > 95 ? 'text-[#22c55e]' : sku.yieldRate > 85 ? 'text-yellow-400' : 'text-[#ef4444]'}`} />
                              <span className={`text-sm font-black mono ${sku.yieldRate > 95 ? 'text-[#22c55e]' : sku.yieldRate > 85 ? 'text-yellow-400' : 'text-[#ef4444]'}`}>
                                {sku.yieldRate.toFixed(2)}%
                              </span>
                            </div>
                            <div className="w-24 h-1 bg-white/5 rounded-full overflow-hidden flex">
                              <div 
                                className={`h-full ${sku.yieldRate > 95 ? 'bg-[#22c55e]' : sku.yieldRate > 85 ? 'bg-yellow-400' : 'bg-[#ef4444]'}`}
                                style={{ width: `${sku.yieldRate}%` }}
                              />
                            </div>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Footer */}
            <div className="p-6 border-t border-white/5 bg-white/[0.01] flex items-center justify-between">
              <div className="flex items-center gap-6">
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-[#22c55e]" />
                  <span className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-widest">High Performance (&gt;95%)</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-yellow-400" />
                  <span className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-widest">Warning (85-95%)</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-[#ef4444]" />
                  <span className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-widest">Critical (&lt;85%)</span>
                </div>
              </div>
              <div className="flex items-center gap-2 text-[#38bdf8]">
                <span className="text-[10px] font-black uppercase tracking-widest">Live Analysis Active</span>
                <ArrowUpRight className="w-3 h-3" />
              </div>
            </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );
};

export default memo(SKUDrilldownModal);
