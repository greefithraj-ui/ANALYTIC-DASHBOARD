import React, { useState, useEffect, useMemo, useRef, memo } from 'react';
import { Layers, CheckCircle, XCircle, Clock, Percent, Package } from 'lucide-react';
import { motion, useMotionValue, useTransform, useSpring, AnimatePresence } from 'motion/react';
import { LineChart, Line, ResponsiveContainer } from 'recharts';
import { KPIStats, DashboardRow, ColumnMapping } from '../types';

interface KPIGridProps {
  stats: KPIStats;
  loading?: boolean;
  onRejectedClick?: () => void;
  onAcceptedClick?: () => void;
  onWipClick?: () => void;
  filteredData: DashboardRow[];
  mapping: ColumnMapping;
}

const Odometer = memo(({ value, suffix = '' }: { value: number; suffix?: string }) => {
  const springValue = useSpring(0, { stiffness: 50, damping: 20 });
  const [displayValue, setDisplayValue] = useState(0);

  useEffect(() => {
    springValue.set(value);
  }, [value, springValue]);

  useEffect(() => {
    return springValue.on("change", (latest) => {
      setDisplayValue(latest);
    });
  }, [springValue]);

  const isPercent = suffix === '%';
  return (
    <span>
      {isPercent ? displayValue.toFixed(1) : Math.floor(displayValue).toLocaleString()}
      {suffix}
    </span>
  );
});

const TiltCard = memo(({ 
  children, 
  className, 
  onClick, 
  glowColor = 'rgba(56, 189, 248, 0.2)', 
  isPulsing = false 
}: { 
  children: React.ReactNode; 
  className?: string; 
  onClick?: () => void;
  glowColor?: string;
  isPulsing?: boolean;
}) => {
  const x = useMotionValue(0);
  const y = useMotionValue(0);

  const rotateX = useTransform(y, [-100, 100], [10, -10]);
  const rotateY = useTransform(x, [-100, 100], [-10, 10]);

  const springX = useSpring(rotateX, { stiffness: 150, damping: 20 });
  const springY = useSpring(rotateY, { stiffness: 150, damping: 20 });

  function handleMouseMove(event: React.MouseEvent<HTMLDivElement>) {
    const rect = event.currentTarget.getBoundingClientRect();
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    x.set(event.clientX - centerX);
    y.set(event.clientY - centerY);
  }

  function handleMouseLeave() {
    x.set(0);
    y.set(0);
  }

  return (
    <motion.div
      onMouseMove={handleMouseMove}
      onMouseLeave={handleMouseLeave}
      onClick={onClick}
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      whileHover={{ scale: 1.02 }}
      style={{
        rotateX: springX,
        rotateY: springY,
        transformStyle: "preserve-3d",
        boxShadow: isPulsing ? undefined : `0 0 20px ${glowColor}`,
      }}
      className={`${className} ${isPulsing ? 'animate-pulse-glow' : ''}`}
    >
      <div style={{ transform: "translateZ(30px)" }}>
        {children}
      </div>
    </motion.div>
  );
});

const KPIGrid: React.FC<KPIGridProps> = ({ stats, loading, onRejectedClick, onAcceptedClick, onWipClick, filteredData, mapping }) => {
  const trendData = useMemo(() => {
    if (!filteredData || filteredData.length === 0) return [];
    
    const groups: Record<string, { date: string; accepted: number; rejected: number; yield: number }> = {};
    
    filteredData.forEach(row => {
      const dateStr = row._parsedDate ? (row._parsedDate as Date).toISOString().split('T')[0] : 'Unknown';
      if (!groups[dateStr]) {
        groups[dateStr] = { date: dateStr, accepted: 0, rejected: 0, yield: 0 };
      }
      
      const status = String(row[mapping.ringStatus] || '').trim().toLowerCase();
      if (['accepted', 'ok', 'pass', '1', 'true', 'yes'].includes(status)) {
        groups[dateStr].accepted++;
      } else if (['rejected', 'nok', 'fail', '0', 'false', 'no'].includes(status)) {
        groups[dateStr].rejected++;
      }
    });

    const sorted = Object.values(groups).sort((a, b) => a.date.localeCompare(b.date));
    
    sorted.forEach(day => {
      const total = day.accepted + day.rejected;
      day.yield = total > 0 ? (day.accepted / total) * 100 : 0;
    });

    return sorted.slice(-7);
  }, [filteredData, mapping]);

  const cards = [
    { label: 'TOTAL ITEMS', value: stats.total, icon: Layers, color: 'blue', suffix: '', trendKey: 'accepted' },
    { label: 'ACCEPTED', value: stats.accepted, icon: CheckCircle, color: 'green', suffix: '', isClickable: true, onClick: onAcceptedClick, trendKey: 'accepted' },
    { label: 'REJECTED', value: stats.rejected, icon: XCircle, color: 'rose', suffix: '', isClickable: true, onClick: onRejectedClick, trendKey: 'rejected' },
    { label: 'WIP INVENTORY', value: stats.wip, icon: Clock, color: 'amber', suffix: '', isClickable: true, onClick: onWipClick, trendKey: 'accepted' },
    { label: 'MOVED TO INVENTORY', value: stats.movedToInventory, icon: Package, color: 'purple', suffix: '', trendKey: 'accepted' },
    { label: 'OVERALL YIELD', value: stats.yield, icon: Percent, color: 'cyan', suffix: '%', trendKey: 'yield' },
  ];

  const getColorClasses = (color: string) => {
    switch(color) {
      case 'blue': return { text: 'text-[#38bdf8]', bg: 'bg-[#38bdf8]/10', border: 'border-[#38bdf8]/20', hex: '#38bdf8' };
      case 'green': return { text: 'text-[#22c55e]', bg: 'bg-[#22c55e]/10', border: 'border-[#22c55e]/20', hex: '#22c55e' };
      case 'rose': return { text: 'text-[#ef4444]', bg: 'bg-[#ef4444]/10', border: 'border-[#ef4444]/20', hex: '#ef4444' };
      case 'cyan': return { text: 'text-[#38bdf8]', bg: 'bg-[#38bdf8]/10', border: 'border-[#38bdf8]/20', hex: '#38bdf8' };
      case 'amber': return { text: 'text-[#f59e0b]', bg: 'bg-[#f59e0b]/10', border: 'border-[#f59e0b]/20', hex: '#f59e0b' };
      case 'purple': return { text: 'text-[#a855f7]', bg: 'bg-[#a855f7]/10', border: 'border-[#a855f7]/20', hex: '#a855f7' };
      default: return { text: 'text-[#38bdf8]', bg: 'bg-[#38bdf8]/10', border: 'border-[#38bdf8]/20', hex: '#38bdf8' };
    }
  };

  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-5 mb-8 w-full perspective-1000">
      {cards.map((card, idx) => {
        const colors = getColorClasses(card.color);
        const isYieldLow = card.label === 'OVERALL YIELD' && stats.yield < 75;
        const isYieldHigh = card.label === 'OVERALL YIELD' && stats.yield > 85;
        
        const glowColor = isYieldLow 
          ? 'rgba(239, 68, 68, 0.4)' 
          : isYieldHigh 
            ? 'rgba(34, 197, 94, 0.4)' 
            : `${colors.hex}33`;

        return (
          <TiltCard 
            key={idx} 
            onClick={card.isClickable ? card.onClick : undefined}
            glowColor={glowColor}
            isPulsing={isYieldLow}
            className={`premium-card p-6 flex flex-col justify-between group ${card.isClickable ? 'cursor-pointer' : ''} transition-all duration-300`} 
          >
            {loading ? (
              <div className="space-y-4">
                <div className="flex justify-between items-center">
                  <div className="h-3 w-20 shimmer rounded-full opacity-20" />
                  <div className="h-8 w-8 shimmer rounded-lg opacity-20" />
                </div>
                <div className="h-10 w-32 shimmer rounded-lg opacity-20" />
              </div>
            ) : (
              <>
                <div className="flex items-center justify-between mb-4">
                  <span className="text-[11px] font-bold text-[#9ca3af] uppercase tracking-[0.15em]">{card.label}</span>
                  <div 
                    className={`p-2.5 rounded-xl border transition-all duration-500 group-hover:rotate-12 group-hover:scale-110 ${colors.text} ${colors.bg} ${colors.border}`}
                  >
                    <card.icon className="w-4 h-4" />
                  </div>
                </div>
                
                <div className="flex flex-col gap-1">
                  <h3 className="text-4xl font-black text-white tracking-tighter mono truncate drop-shadow-[0_0_15px_rgba(255,255,255,0.1)]">
                    <Odometer value={card.value} suffix={card.suffix} />
                  </h3>
                  
                  {/* Sparkline */}
                  {['ACCEPTED', 'REJECTED', 'OVERALL YIELD'].includes(card.label) && trendData.length > 1 && (
                    <div className="h-10 w-full mt-2">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={trendData}>
                          <defs>
                            <filter id={`glow-${idx}`} x="-20%" y="-20%" width="140%" height="140%">
                              <feGaussianBlur stdDeviation="2" result="blur" />
                              <feComposite in="SourceGraphic" in2="blur" operator="over" />
                            </filter>
                          </defs>
                          <Line 
                            type="monotone" 
                            dataKey={card.trendKey} 
                            stroke={colors.hex} 
                            strokeWidth={3} 
                            dot={false}
                            filter={`url(#glow-${idx})`}
                            isAnimationActive={true}
                          />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  )}
                </div>

                <div className="mt-4 h-1 w-full bg-white/5 rounded-full overflow-hidden">
                  <motion.div 
                    initial={{ width: 0 }}
                    animate={{ width: card.label === 'OVERALL YIELD' ? `${stats.yield}%` : '100%' }}
                    transition={{ duration: 1.5, ease: "easeOut" }}
                    className="h-full" 
                    style={{ 
                      backgroundColor: colors.hex,
                      opacity: 0.5,
                      boxShadow: `0 0 10px ${colors.hex}`
                    }} 
                  />
                </div>
              </>
            )}
          </TiltCard>
        );
      })}
    </div>
  );
};

export default memo(KPIGrid);
