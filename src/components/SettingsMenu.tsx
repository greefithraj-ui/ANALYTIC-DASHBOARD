
import React, { useState } from 'react';
import { Settings, X, RefreshCw, ChevronDown } from 'lucide-react';
import { SheetConfig, ColumnMapping } from '../types';

interface SettingsMenuProps {
  config: SheetConfig;
  headers: string[];
  onUpdate: (newConfig: SheetConfig) => void;
  isOpen: boolean;
  setIsOpen: (open: boolean) => void;
  isRefreshing: boolean;
}

const SettingsMenu: React.FC<SettingsMenuProps> = ({ 
  config, 
  headers, 
  onUpdate, 
  isOpen, 
  setIsOpen,
  isRefreshing
}) => {
  const [localConfig, setLocalConfig] = useState(config);

  const handleSave = () => {
    onUpdate(localConfig);
    setIsOpen(false);
  };

  const handleMappingChange = (key: keyof ColumnMapping, value: string) => {
    setLocalConfig(prev => ({
      ...prev,
      mapping: { ...prev.mapping, [key]: value }
    }));
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 overflow-hidden">
      <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" onClick={() => setIsOpen(false)} />
      <div className="absolute inset-y-0 right-0 max-w-md w-full bg-white shadow-2xl flex flex-col">
        <div className="p-6 border-b flex justify-between items-center bg-slate-50">
          <h2 className="text-xl font-bold flex items-center gap-2 text-slate-900">
            <Settings className="w-5 h-5 text-indigo-600" />
            Dashboard Settings
          </h2>
          <button onClick={() => setIsOpen(false)} className="p-2 hover:bg-slate-200 rounded-full transition-colors text-slate-500">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-6 space-y-8">
          {/* Source Settings */}
          <section className="space-y-4">
            <h3 className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Sheet Source</h3>
            <div className="space-y-3">
              <div>
                <label className="block text-sm font-medium text-slate-900 mb-1">Google Sheet URL</label>
                <input 
                  type="text" 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm text-slate-900"
                  placeholder="https://docs.google.com/spreadsheets/d/..."
                  value={localConfig.url}
                  onChange={e => setLocalConfig({...localConfig, url: e.target.value})}
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-slate-900 mb-1">Sheet Name</label>
                  <input 
                    type="text" 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm text-slate-900"
                    value={localConfig.sheetName}
                    onChange={e => setLocalConfig({...localConfig, sheetName: e.target.value})}
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-900 mb-1">Range</label>
                  <input 
                    type="text" 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm text-slate-900"
                    value={localConfig.range}
                    onChange={e => setLocalConfig({...localConfig, range: e.target.value})}
                  />
                </div>
              </div>
            </div>
          </section>

          {/* Column Mapping */}
          <section className="space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="text-xs font-semibold text-slate-500 uppercase tracking-wider">Column Mapping</h3>
              <span className="text-[10px] bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded-full">Auto-detected headers</span>
            </div>
            <div className="space-y-3">
              {localConfig.mapping && (Object.keys(localConfig.mapping) as Array<keyof ColumnMapping>).map((key) => (
                <div key={key} className="flex flex-col gap-1">
                  <label className="text-sm font-medium text-slate-900 capitalize">
                    {key.replace(/([A-Z])/g, ' $1').trim()}
                  </label>
                  <div className="relative">
                    <select 
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg appearance-none bg-white focus:ring-2 focus:ring-indigo-500 outline-none text-sm pr-10 text-slate-900"
                      value={localConfig.mapping[key]}
                      onChange={(e) => handleMappingChange(key, e.target.value)}
                    >
                      <option value="" className="text-slate-900">-- Select Header --</option>
                      {headers.map(h => (
                        <option key={h} value={h} className="text-slate-900">{h}</option>
                      ))}
                    </select>
                    <ChevronDown className="absolute right-3 top-2.5 w-4 h-4 text-slate-400 pointer-events-none" />
                  </div>
                </div>
              ))}
            </div>
          </section>
        </div>

        <div className="p-6 border-t bg-slate-50 space-y-3">
          <button 
            onClick={handleSave}
            className="w-full py-3 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold rounded-lg transition-colors flex items-center justify-center gap-2 disabled:opacity-50"
            disabled={isRefreshing}
          >
            {isRefreshing ? <RefreshCw className="w-4 h-4 animate-spin" /> : null}
            Apply Changes
          </button>
        </div>
      </div>
    </div>
  );
};

export default SettingsMenu;
