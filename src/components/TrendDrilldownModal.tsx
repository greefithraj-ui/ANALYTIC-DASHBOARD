import React, { memo } from 'react';
import { X, AlertTriangle, Activity, ArrowUpRight, ListFilter, Info } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface TrendDrilldownModalProps {
  isOpen: boolean;
  onClose: () => void;
  date: string;
  total: number;
  nok: number;
  yieldRate: number;
  reasons: Record<string, number>;
}

const TrendDrilldownModal: React.FC<TrendDrilldownModalProps> = ({ isOpen, onClose, date, total, nok, yieldRate, reasons }) => {
  const sortedReasons = Object.entries(reasons)
    .sort((a, b) => (b[1] as number) - (a[1] as number));

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
            className="relative w-full max-w-4xl max-h-[80vh] bg-[#0f1117] border border-white/10 rounded-3xl shadow-[0_0_50px_rgba(0,0,0,0.5)] flex flex-col overflow-hidden"
          >
            {/* Glow Accents */}
            <div className="absolute top-0 left-1/4 right-1/4 h-px bg-gradient-to-r from-transparent via-[#ef4444]/50 to-transparent" />
            <div className="absolute bottom-0 left-1/4 right-1/4 h-px bg-gradient-to-r from-transparent via-[#38bdf8]/50 to-transparent" />

            {/* Header */}
            <div className="p-6 md:p-8 border-b border-white/5 flex items-center justify-between bg-white/[0.02]">
              <div className="flex items-center gap-4">
                <div className="p-3 bg-[#ef4444]/10 rounded-2xl border border-[#ef4444]/20">
                  <AlertTriangle className="w-6 h-6 text-[#ef4444]" />
                </div>
                <div>
                  <h2 className="text-xl font-black text-white uppercase tracking-widest mono">Trend Intelligence Deep-Dive</h2>
                  <p className="text-xs text-[#9ca3af] uppercase tracking-[0.2em] mt-1">Defect Analysis for <span className="text-[#38bdf8]">{date}</span></p>
                </div>
              </div>
              <button
                onClick={onClose}
                className="p-2 hover:bg-white/5 rounded-xl text-[#9ca3af] hover:text-white transition-all border border-white/5 group"
              >
                <X className="w-6 h-6 group-hover:rotate-90 transition-transform duration-300" />
              </button>
            </div>

            {/* Stats Overview */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 p-6 md:p-8 bg-white/[0.01] border-b border-white/5">
              <div className="bg-[#161a23] p-4 rounded-2xl border border-white/5">
                <span className="text-[10px] font-black text-[#9ca3af] uppercase tracking-widest block mb-1">Total Samples</span>
                <div className="text-2xl font-black text-white mono">{total.toLocaleString()}</div>
              </div>
              <div className="bg-[#161a23] p-4 rounded-2xl border border-white/5">
                <span className="text-[10px] font-black text-[#9ca3af] uppercase tracking-widest block mb-1">NOK Count</span>
                <div className="text-2xl font-black text-[#ef4444] mono">{nok.toLocaleString()}</div>
              </div>
              <div className="bg-[#161a23] p-4 rounded-2xl border border-white/5">
                <span className="text-[10px] font-black text-[#9ca3af] uppercase tracking-widest block mb-1">Yield Rate</span>
                <div className="flex items-center gap-3">
                  <div className="text-2xl font-black text-[#22c55e] mono">{yieldRate.toFixed(2)}%</div>
                  <div className="flex-1 h-1.5 bg-white/5 rounded-full overflow-hidden">
                    <div className="h-full bg-[#22c55e]" style={{ width: `${yieldRate}%` }} />
                  </div>
                </div>
              </div>
            </div>

            {/* Content - Defect Breakdown */}
            <div className="flex-1 overflow-y-auto p-6 md:p-8 custom-scrollbar">
              <div className="flex items-center gap-2 mb-6">
                <ListFilter className="w-4 h-4 text-[#38bdf8]" />
                <h3 className="text-xs font-black text-white uppercase tracking-widest mono">Defect Pattern Breakdown</h3>
              </div>

              {sortedReasons.length > 0 ? (
                <div className="grid grid-cols-1 gap-3">
                  {sortedReasons.map(([reason, count], idx) => {
                    const percentage = ((count as number) / (nok || 1) * 100).toFixed(1);
                    return (
                      <div key={idx} className="group bg-[#161a23]/50 border border-white/5 rounded-2xl p-4 hover:border-[#ef4444]/30 transition-all">
                        <div className="flex items-center justify-between mb-3">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 rounded-xl bg-[#ef4444]/10 flex items-center justify-center text-[10px] font-black text-[#ef4444] mono border border-[#ef4444]/20">
                              {idx + 1}
                            </div>
                            <span className="text-sm font-bold text-white group-hover:text-[#ef4444] transition-colors">{reason}</span>
                          </div>
                          <div className="text-right">
                            <div className="text-sm font-black text-white mono">{count.toLocaleString()} <span className="text-[10px] text-[#9ca3af] font-normal uppercase ml-1">Units</span></div>
                            <div className="text-[10px] font-bold text-[#ef4444] mono">{percentage}% of defects</div>
                          </div>
                        </div>
                        <div className="h-1.5 bg-white/5 rounded-full overflow-hidden">
                          <div 
                            className="h-full bg-gradient-to-r from-[#ef4444]/50 to-[#ef4444]" 
                            style={{ width: `${percentage}%` }} 
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="py-12 flex flex-col items-center justify-center text-[#9ca3af]/40 border border-dashed border-white/10 rounded-3xl">
                  <Info className="w-8 h-8 mb-3 opacity-20" />
                  <p className="text-[10px] font-black uppercase tracking-[0.3em]">No defect patterns recorded for this period</p>
                </div>
              )}
            </div>

            {/* Footer */}
            <div className="p-6 border-t border-white/5 bg-white/[0.01] flex items-center justify-between">
              <div className="text-[9px] font-black text-[#9ca3af] uppercase tracking-widest">
                Analysis based on {total} total samples
              </div>
              <div className="flex items-center gap-2 text-[#ef4444]">
                <span className="text-[10px] font-black uppercase tracking-widest">Defect Intelligence Active</span>
                <ArrowUpRight className="w-3 h-3" />
              </div>
            </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );
};

export default memo(TrendDrilldownModal);
