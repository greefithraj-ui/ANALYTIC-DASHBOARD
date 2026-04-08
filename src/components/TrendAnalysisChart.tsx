import React, { useState, useMemo, useEffect, memo } from 'react';
import { 
  XAxis, YAxis, CartesianGrid, Tooltip, 
  ResponsiveContainer, Area, Line, ComposedChart,
  Brush
} from 'recharts';
import { Calendar, Layers, Activity, TrendingUp, Filter, Maximize2, Minimize2, Plus, X, MousePointer2, Keyboard, AlertTriangle } from 'lucide-react';
import { format, startOfDay, startOfWeek, startOfMonth } from 'date-fns';
import { DashboardRow, ColumnMapping } from '../types';
import { parseDate } from '../services/sheetService';
import TrendDrilldownModal from './TrendDrilldownModal';

interface TrendAnalysisChartProps {
  data: DashboardRow[];
  mapping: ColumnMapping;
}

type Grain = 'daily' | 'weekly' | 'monthly';

const TrendAnalysisChart: React.FC<TrendAnalysisChartProps> = ({ data, mapping }) => {
  const [grain, setGrain] = useState<Grain>('daily');
  const [showNOK, setShowNOK] = useState(true);
  const [showYield, setShowYield] = useState(true);
  const [selectedReasons, setSelectedReasons] = useState<string[]>([]);
  const [isExpanded, setIsExpanded] = useState(false);

  // Drill-down state
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [drilldownData, setDrilldownData] = useState<any>(null);
  
  // Frozen tooltip state
  const [frozenTooltip, setFrozenTooltip] = useState<any>(null);
  
  // Hover tracking for hotkey
  const [currentHoverData, setCurrentHoverData] = useState<any>(null);

  // 1. Extract all unique rejection reasons for the filter
  const allReasons = useMemo(() => {
    const reasons = new Set<string>();
    const reasonKey = mapping.reason;
    const statusKey = mapping.ringStatus;

    data.forEach(row => {
      const status = String(row[statusKey] || '').trim().toLowerCase();
      if (['rejected', 'nok', 'fail', '0', 'false', 'no'].includes(status)) {
        const reason = String(row[reasonKey] || '').trim();
        if (reason) reasons.add(reason);
      }
    });
    return Array.from(reasons).sort();
  }, [data, mapping]);

  // 2. Process data based on grain and filters
  const chartData = useMemo(() => {
    const dateKey = mapping.date;
    const statusKey = mapping.ringStatus;
    const reasonKey = mapping.reason;

    const grouped: Record<string, { date: Date, total: number, nok: number, reasons: Record<string, number> }> = {};

    data.forEach(row => {
      const date = (row as any)._parsedDate || parseDate(String(row[dateKey] || ''));
      if (!date) return;

      let groupKey = '';
      let groupDate = date;

      if (grain === 'daily') {
        groupDate = startOfDay(date);
        groupKey = format(groupDate, 'yyyy-MM-dd');
      } else if (grain === 'weekly') {
        groupDate = startOfWeek(date);
        groupKey = format(groupDate, 'yyyy-ww');
      } else {
        groupDate = startOfMonth(date);
        groupKey = format(groupDate, 'yyyy-MM');
      }

      if (!grouped[groupKey]) {
        grouped[groupKey] = { date: groupDate, total: 0, nok: 0, reasons: {} };
      }

      grouped[groupKey].total++;
      
      const status = String(row[statusKey] || '').trim().toLowerCase();
      const isNOK = ['rejected', 'nok', 'fail', '0', 'false', 'no'].includes(status);
      
      if (isNOK) {
        const reason = String(row[reasonKey] || '').trim();
        
        // If reason filters are active, only count if it matches
        if (selectedReasons.length === 0 || selectedReasons.includes(reason)) {
          grouped[groupKey].nok++;
          if (reason) {
            grouped[groupKey].reasons[reason] = (grouped[groupKey].reasons[reason] || 0) + 1;
          }
        }
      }
    });

    return Object.entries(grouped)
      .map(([key, val]) => ({
        key,
        displayDate: grain === 'daily' ? format(val.date, 'MMM dd') : 
                     grain === 'weekly' ? `W${format(val.date, 'ww, MMM')}` : 
                     format(val.date, 'MMM yyyy'),
        fullDate: format(val.date, 'PPP'),
        nok: val.nok,
        total: val.total,
        yield: val.total > 0 ? Number(((val.total - val.nok) / val.total * 100).toFixed(2)) : 100,
        reasons: val.reasons,
        rawDate: val.date
      }))
      .sort((a, b) => a.rawDate.getTime() - b.rawDate.getTime());
  }, [data, mapping, grain, selectedReasons]);

  const toggleReason = (reason: string) => {
    setSelectedReasons(prev => 
      prev.includes(reason) ? prev.filter(r => r !== reason) : [...prev, reason]
    );
  };

  // Hotkey listener
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Support +, =, Enter
      const isTriggerKey = e.key === '+' || e.key === '=' || e.key === 'Enter';
      
      if (isTriggerKey) {
        // Priority 1: Use frozen data if available
        if (frozenTooltip && frozenTooltip.payload && frozenTooltip.payload.length > 0) {
          const d = frozenTooltip.payload[0].payload;
          setDrilldownData(d);
          setIsModalOpen(true);
          e.preventDefault();
          return;
        }
        
        // Priority 2: Use current hover data
        if (currentHoverData) {
          setDrilldownData(currentHoverData);
          setIsModalOpen(true);
          e.preventDefault();
        }
      }

      if (e.key === 'Escape') {
        setFrozenTooltip(null);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [currentHoverData, frozenTooltip]);

  const handleChartClick = (state: any) => {
    if (state && state.activePayload && state.activePayload.length > 0) {
      setFrozenTooltip({
        payload: state.activePayload,
        label: state.activeLabel,
        coordinate: state.coordinate
      });
    }
  };

  const handleMouseMove = (state: any) => {
    if (state && state.activePayload && state.activePayload.length > 0) {
      setCurrentHoverData(state.activePayload[0].payload);
    } else if (state && state.activeTooltipIndex !== undefined && chartData[state.activeTooltipIndex]) {
      setCurrentHoverData(chartData[state.activeTooltipIndex]);
    } else {
      setCurrentHoverData(null);
    }
  };

  const CustomTooltip = ({ active, payload, label }: any) => {
    const isFrozen = !!frozenTooltip;
    const effectiveActive = isFrozen ? true : active;
    const effectivePayload = isFrozen ? frozenTooltip.payload : payload;
    const effectiveLabel = isFrozen ? frozenTooltip.label : label;

    if (effectiveActive && effectivePayload && effectivePayload.length) {
      const d = effectivePayload[0].payload;
      return (
        <div 
          className={`bg-[#161a23]/95 border backdrop-blur-xl p-4 rounded-xl shadow-[0_0_30px_rgba(56,189,248,0.2)] pointer-events-auto transition-all duration-300 ${
            isFrozen ? 'border-[#38bdf8] ring-2 ring-[#38bdf8]/20' : 'border-[#38bdf8]/30'
          }`}
          onMouseEnter={(e) => e.stopPropagation()}
        >
          <div className="flex items-center justify-between gap-4 mb-3 pb-2 border-b border-white/10">
            <div className="flex items-center gap-2">
              <Calendar className="w-4 h-4 text-[#38bdf8]" />
              <span className="text-xs font-black text-white uppercase tracking-widest mono">{d.fullDate}</span>
              {isFrozen && (
                <div className="flex items-center gap-1 px-1.5 py-0.5 bg-[#38bdf8]/20 rounded border border-[#38bdf8]/30 ml-2">
                  <MousePointer2 className="w-2.5 h-2.5 text-[#38bdf8]" />
                  <span className="text-[8px] font-black text-[#38bdf8] uppercase tracking-tighter">Frozen</span>
                </div>
              )}
            </div>
            {isFrozen && (
              <button 
                onClick={(e) => {
                  e.stopPropagation();
                  setFrozenTooltip(null);
                }}
                className="p-1 hover:bg-white/10 rounded-md text-[#9ca3af] hover:text-white transition-colors"
                title="Unfreeze Tooltip (Esc)"
              >
                <X className="w-3 h-3" />
              </button>
            )}
          </div>
          
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-4">
              <div className="flex flex-col">
                <span className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-tighter">Yield Rate</span>
                <span className="text-lg font-black text-[#22c55e] mono">{d.yield}%</span>
              </div>
              <div className="flex flex-col">
                <span className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-tighter">NOK Count</span>
                <span className="text-lg font-black text-[#ef4444] mono">{d.nok}</span>
              </div>
            </div>

            {Object.keys(d.reasons).length > 0 && (
              <div className="pt-2 border-t border-white/5">
                <span className="text-[9px] font-black text-[#9ca3af] uppercase tracking-widest block mb-2">Defect Breakdown</span>
                <div className="space-y-1.5">
                  {Object.entries(d.reasons)
                    .sort((a: any, b: any) => b[1] - a[1])
                    .slice(0, 3)
                    .map(([reason, count]: any) => (
                      <div key={reason} className="flex items-center justify-between gap-4">
                        <span className="text-[10px] text-[#e5e7eb] truncate max-w-[120px] font-bold">{reason}</span>
                        <span className="text-[10px] font-black text-[#ef4444] mono">{count}</span>
                      </div>
                    ))}
                  {Object.keys(d.reasons).length > 3 && (
                    <span className="text-[9px] text-[#9ca3af] italic">+{Object.keys(d.reasons).length - 3} more patterns</span>
                  )}
                </div>
              </div>
            )}
          </div>

          <div className="mt-3 pt-3 border-t border-white/10 flex flex-col gap-3">
            <div className="flex justify-between items-center">
              <span className="text-[9px] text-[#9ca3af] uppercase">Total Samples</span>
              <span className="text-[10px] font-bold text-white mono">{d.total}</span>
            </div>

            <button
              onClick={(e) => {
                e.stopPropagation();
                setDrilldownData(d);
                setIsModalOpen(true);
              }}
              className="w-full py-2 bg-[#ef4444]/10 hover:bg-[#ef4444]/20 border border-[#ef4444]/30 rounded-lg flex items-center justify-center gap-2 transition-all group"
            >
              <Plus className="w-3 h-3 text-[#ef4444] group-hover:scale-110 transition-transform" />
              <span className="text-[9px] font-black text-[#ef4444] uppercase tracking-[0.2em]">[ + MORE DETAILS ]</span>
            </button>
          </div>

          {!isFrozen && (
            <div className="mt-2 flex items-center justify-center gap-1 opacity-40">
              <Keyboard className="w-2.5 h-2.5 text-[#9ca3af]" />
              <span className="text-[8px] font-bold text-[#9ca3af] uppercase tracking-tighter">Press + or Enter to drill-down</span>
            </div>
          )}
        </div>
      );
    }
    return null;
  };

  return (
    <div className={`flex flex-col gap-6 transition-all duration-500 ${isExpanded ? 'fixed inset-4 z-[100] bg-[#0f1117] p-8 rounded-3xl border border-white/10 shadow-[0_0_100px_rgba(0,0,0,0.8)]' : 'mt-8 pt-8 border-t border-white/5'}`}>
      {/* Header & Controls */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
        <div className="flex items-center gap-4">
          <div className="p-2 bg-[#38bdf8]/10 rounded-lg border border-[#38bdf8]/20">
            <TrendingUp className="w-5 h-5 text-[#38bdf8]" />
          </div>
          <div>
            <h3 className="text-sm font-black text-white uppercase tracking-[0.2em] mono">Trend Intelligence</h3>
            <p className="text-[10px] text-[#9ca3af] uppercase tracking-widest mt-0.5">Predictive Yield & Rejection Analysis</p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          {/* Grain Toggles */}
          <div className="flex bg-[#161a23] p-1 rounded-xl border border-white/5">
            {(['daily', 'weekly', 'monthly'] as Grain[]).map((g) => (
              <button
                key={g}
                onClick={() => setGrain(g)}
                className={`px-4 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all ${
                  grain === g 
                    ? 'bg-[#38bdf8] text-white shadow-[0_0_15px_rgba(56,189,248,0.4)]' 
                    : 'text-[#9ca3af] hover:text-white'
                }`}
              >
                {g}
              </button>
            ))}
          </div>

          {/* Metric Toggles */}
          <div className="flex gap-2">
            <button
              onClick={() => setShowNOK(!showNOK)}
              className={`flex items-center gap-2 px-3 py-1.5 rounded-xl border transition-all ${
                showNOK 
                  ? 'bg-[#ef4444]/10 border-[#ef4444]/30 text-[#ef4444]' 
                  : 'bg-transparent border-white/5 text-[#9ca3af]'
              }`}
            >
              <div className={`w-1.5 h-1.5 rounded-full ${showNOK ? 'bg-[#ef4444] shadow-[0_0_8px_#ef4444]' : 'bg-[#9ca3af]'}`} />
              <span className="text-[10px] font-black uppercase tracking-widest">NOK</span>
            </button>
            <button
              onClick={() => setShowYield(!showYield)}
              className={`flex items-center gap-2 px-3 py-1.5 rounded-xl border transition-all ${
                showYield 
                  ? 'bg-[#22c55e]/10 border-[#22c55e]/30 text-[#22c55e]' 
                  : 'bg-transparent border-white/5 text-[#9ca3af]'
              }`}
            >
              <div className={`w-1.5 h-1.5 rounded-full ${showYield ? 'bg-[#22c55e] shadow-[0_0_8px_#22c55e]' : 'bg-[#9ca3af]'}`} />
              <span className="text-[10px] font-black uppercase tracking-widest">Yield</span>
            </button>
          </div>

          <button 
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-2 bg-[#161a23] border border-white/5 rounded-xl text-[#9ca3af] hover:text-white transition-colors"
          >
            {isExpanded ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* Chart Area */}
      <div className={`relative bg-[#161a23]/30 rounded-2xl border border-white/5 p-4 ${isExpanded ? 'flex-1' : 'h-[400px]'}`}>
        {chartData.length === 0 ? (
          <div className="flex flex-col items-center justify-center h-full gap-4 animate-in fade-in duration-500">
            <div className="p-4 bg-white/5 rounded-full border border-white/5">
              <AlertTriangle className="w-8 h-8 text-[#9ca3af]/30" />
            </div>
            <div className="text-center">
              <p className="text-xs font-black text-[#e5e7eb] uppercase tracking-[0.2em] mb-1">No data available</p>
              <p className="text-[10px] font-bold text-[#9ca3af] uppercase tracking-widest">Adjust filters to visualize trends</p>
            </div>
          </div>
        ) : (
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart 
              data={chartData} 
              margin={{ top: 20, right: 20, bottom: 20, left: 0 }}
              onClick={handleChartClick}
              onMouseMove={handleMouseMove}
              onMouseLeave={() => setCurrentHoverData(null)}
              style={{ cursor: 'pointer' }}
            >
              <defs>
                <linearGradient id="colorNok" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#ef4444" stopOpacity={0.3}/>
                  <stop offset="95%" stopColor="#ef4444" stopOpacity={0}/>
                </linearGradient>
                <linearGradient id="colorYield" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#22c55e" stopOpacity={0.2}/>
                  <stop offset="95%" stopColor="#22c55e" stopOpacity={0}/>
                </linearGradient>
              </defs>
              
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.03)" vertical={false} />
              
              <XAxis 
                dataKey="displayDate" 
                axisLine={false}
                tickLine={false}
                tick={{ fill: '#9ca3af', fontSize: 10, fontWeight: 700 }}
                dy={10}
              />
              
              <YAxis 
                yAxisId="left"
                axisLine={false}
                tickLine={false}
                tick={{ fill: '#ef4444', fontSize: 10, fontWeight: 700 }}
                hide={!showNOK}
              />
              
              <YAxis 
                yAxisId="right"
                orientation="right"
                axisLine={false}
                tickLine={false}
                tick={{ fill: '#22c55e', fontSize: 10, fontWeight: 700 }}
                domain={[0, 100]}
                hide={!showYield}
              />
  
              <Tooltip 
                content={<CustomTooltip />} 
                cursor={{ stroke: 'rgba(56, 189, 248, 0.2)', strokeWidth: 2 }} 
                wrapperStyle={{ pointerEvents: 'auto', outline: 'none' }}
                isAnimationActive={false}
                position={frozenTooltip ? { x: frozenTooltip.coordinate.x, y: frozenTooltip.coordinate.y - 200 } : undefined}
              />
              
              {showNOK && (
                <Area
                  yAxisId="left"
                  type="monotone"
                  dataKey="nok"
                  stroke="#ef4444"
                  strokeWidth={3}
                  fillOpacity={1}
                  fill="url(#colorNok)"
                  className="pulse-glow"
                  animationDuration={2000}
                />
              )}
  
              {showYield && (
                <Line
                  yAxisId="right"
                  type="monotone"
                  dataKey="yield"
                  stroke="#22c55e"
                  strokeWidth={3}
                  dot={{ r: 4, fill: '#22c55e', strokeWidth: 2, stroke: '#0f1117', className: "pulse-glow" }}
                  activeDot={{ r: 6, strokeWidth: 0, fill: '#fff', className: "pulse-glow" }}
                  className="energy-flow"
                  animationDuration={2500}
                />
              )}
  
              <Brush 
                dataKey="displayDate" 
                height={30} 
                stroke="#38bdf8" 
                fill="#161a23"
                travellerWidth={10}
                gap={1}
              >
                <ComposedChart>
                  <Area dataKey="nok" fill="#ef4444" stroke="none" fillOpacity={0.2} />
                </ComposedChart>
              </Brush>
            </ComposedChart>
          </ResponsiveContainer>
        )}
      </div>

      {/* Interactive Legend / Reason Filter */}
      <div className="flex flex-col gap-3">
        <div className="flex items-center gap-2">
          <Filter className="w-3 h-3 text-[#9ca3af]" />
          <span className="text-[10px] font-black text-[#9ca3af] uppercase tracking-widest">Isolate Defect Patterns</span>
        </div>
        <div className="flex flex-wrap gap-2">
          {allReasons.map(reason => (
            <button
              key={reason}
              onClick={() => toggleReason(reason)}
              className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all border ${
                selectedReasons.includes(reason)
                  ? 'bg-[#ef4444] border-[#ef4444] text-white shadow-[0_0_10px_rgba(239,68,68,0.3)]'
                  : 'bg-[#161a23] border-white/5 text-[#9ca3af] hover:border-white/20'
              }`}
            >
              {reason}
            </button>
          ))}
          {allReasons.length === 0 && (
            <span className="text-[10px] text-[#9ca3af]/40 italic uppercase tracking-widest">No defect patterns detected in current range</span>
          )}
          {selectedReasons.length > 0 && (
            <button 
              onClick={() => setSelectedReasons([])}
              className="px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest text-[#ef4444] hover:bg-[#ef4444]/10 transition-colors"
            >
              Clear All
            </button>
          )}
        </div>
      </div>

      {/* Drilldown Modal */}
      <TrendDrilldownModal 
        isOpen={isModalOpen}
        onClose={() => setIsModalOpen(false)}
        date={drilldownData?.fullDate || ''}
        total={drilldownData?.total || 0}
        nok={drilldownData?.nok || 0}
        yieldRate={drilldownData?.yield || 0}
        reasons={drilldownData?.reasons || {}}
      />
    </div>
  );
};

export default memo(TrendAnalysisChart);
