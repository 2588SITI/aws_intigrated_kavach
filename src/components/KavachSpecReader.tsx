import React, { useState } from 'react';
import { 
  BookOpen, 
  ShieldCheck, 
  Activity, 
  FileText, 
  Database, 
  Search, 
  Sliders, 
  Cpu, 
  Layers, 
  Radio, 
  Info,
  ChevronRight,
  ClipboardList
} from 'lucide-react';
import { motion } from 'motion/react';
import { KAVACH_OFFICIAL_SPECS, TDMA_SLOTS_AMDT10, TdmaSlot } from '../utils/kavachSpecs';
import { cn } from '../utils/cn';

export const KavachSpecReader: React.FC = () => {
  const [activeAnnex, setActiveAnnex] = useState<string>("Annexure-A2");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [selectedSlot, setSelectedSlot] = useState<TdmaSlot | null>(TDMA_SLOTS_AMDT10[1]); // Default to P2

  const annexList = Object.entries(KAVACH_OFFICIAL_SPECS.annexures);

  // Filter content based on search query
  const filteredAnnexList = annexList.filter(([id, data]) => {
    const s = searchQuery.toLowerCase();
    const matchesTitle = data.title.toLowerCase().includes(s) || id.toLowerCase().includes(s);
    const matchesDesc = data.description.toLowerCase().includes(s);
    const matchesParams = data.parameters?.some(p => 
      p.name.toLowerCase().includes(s) || p.desc.toLowerCase().includes(s)
    ) || false;
    const matchesFormats = data.dataFormats?.some(f => 
      f.field.toLowerCase().includes(s) || f.desc.toLowerCase().includes(s)
    ) || false;
    return matchesTitle || matchesDesc || matchesParams || matchesFormats;
  });

  return (
    <div className="space-y-6">
      {/* Title Header Block */}
      <div className="glass-card p-6 rounded-2xl border-t-2 border-emerald-500 relative overflow-hidden">
        <div className="absolute right-0 top-0 opacity-5 pointer-events-none transform translate-x-12 translate-y-2">
          <BookOpen className="w-96 h-96 text-emerald-400" />
        </div>
        <div className="space-y-2 relative">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-emerald-500/10 rounded-xl border border-emerald-500/20">
              <BookOpen className="w-6 h-6 text-emerald-400" />
            </div>
            <div>
              <span className="text-[10px] font-bold text-emerald-400 uppercase tracking-widest bg-emerald-500/10 px-2 py-0.5 rounded-full border border-emerald-500/10">
                Official Binder - Amendment 10 Certified
              </span>
              <h1 className="text-2xl font-black text-white tracking-tight mt-1 uppercase">
                RDSO Kavach Rules & Specifications Library
              </h1>
            </div>
          </div>
          <p className="text-slate-300 text-sm leading-relaxed max-w-4xl">
            This repository stores the official structural guidelines, packet formats, and timeout constraints extracted from the latest <span className="text-emerald-400 font-bold">{KAVACH_OFFICIAL_SPECS.version}</span> documents loaded onto this diagnostics suite. Use it to verify calculated outcomes or cross-reference alarm triggers with the safety standards of the Indian Railways.
          </p>
        </div>
      </div>

      {/* Interactive Timing Diagram Block */}
      <div className="glass-card p-6 rounded-2xl border border-white/5 space-y-6">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <Radio className="w-5 h-5 text-emerald-400" />
              <h2 className="text-lg font-bold text-white uppercase tracking-tight">FDMA / TDMA 2s Frame Slot Allocation (Amdt-10)</h2>
            </div>
            <p className="text-slate-400 text-xs font-medium">
              Amendment-10 optimized cycle to fit <span className="text-emerald-300">70 slots</span> of <span className="text-emerald-300">22.5ms (432 bits)</span>. Click a slot below to inspect its purpose.
            </p>
          </div>
          <div className="flex flex-wrap gap-2 text-[10px] font-bold uppercase tracking-wider">
            <span className="flex items-center gap-1.5 px-2 py-1 rounded bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
              <span className="w-2 h-2 rounded-full bg-emerald-400" /> Regular / Onboard (M)
            </span>
            <span className="flex items-center gap-1.5 px-2 py-1 rounded bg-blue-500/10 text-blue-400 border border-blue-500/20">
              <span className="w-2 h-2 rounded-full bg-blue-400" /> Mobile Broadcast (MBS)
            </span>
            <span className="flex items-center gap-1.5 px-2 py-1 rounded bg-rose-500/10 text-rose-400 border border-rose-500/20">
              <span className="w-2 h-2 rounded-full bg-rose-400" /> Mobile Emergency (ME)
            </span>
            <span className="flex items-center gap-1.5 px-2 py-1 rounded bg-orange-500/10 text-orange-400 border border-orange-500/20">
              <span className="w-2 h-2 rounded-full bg-orange-400" /> Station Emergency (SE)
            </span>
            <span className="flex items-center gap-1.5 px-2 py-1 rounded bg-purple-500/10 text-purple-400 border border-purple-500/20">
              <span className="w-2 h-2 rounded-full bg-purple-400" /> Access Authority (STS)
            </span>
          </div>
        </div>

        {/* The Grid / Timeline representing the 70 slots cycle */}
        <div className="bg-black/30 p-4 rounded-xl border border-white/5 space-y-4">
          <div className="overflow-x-auto">
            <div className="flex gap-1.5 min-w-[900px] mb-2 p-1 bg-white/5 rounded-lg">
              {TDMA_SLOTS_AMDT10.map(slot => {
                const isSelected = selectedSlot?.id === slot.id;
                let bgClass = "bg-emerald-500/10 border-emerald-500/20 text-emerald-400 hover:bg-emerald-500/20";
                if (slot.color === "blue") bgClass = "bg-blue-500/10 border-blue-500/20 text-blue-400 hover:bg-blue-500/20";
                if (slot.color === "rose") bgClass = "bg-rose-500/10 border-rose-500/20 text-rose-400 hover:bg-rose-500/20";
                if (slot.color === "orange") bgClass = "bg-orange-500/10 border-orange-500/20 text-orange-400 hover:bg-orange-500/20";
                if (slot.color === "purple") bgClass = "bg-purple-500/10 border-purple-500/20 text-purple-400 hover:bg-purple-500/20";
                if (slot.color === "amber") bgClass = "bg-amber-500/10 border-amber-500/20 text-amber-400 hover:bg-amber-500/20";

                return (
                  <button
                    key={slot.id}
                    onClick={() => setSelectedSlot(slot)}
                    className={cn(
                      "flex-1 py-3 px-2 rounded-md border text-center transition-all focus:outline-none cursor-pointer",
                      bgClass,
                      isSelected ? "ring-2 ring-emerald-400 border-transparent transform -translate-y-1 shadow-lg shadow-emerald-500/20" : ""
                    )}
                  >
                    <div className="font-black text-xs">{slot.id}</div>
                    <div className="text-[9px] font-bold opacity-80 mt-0.5 truncate">{slot.name.replace(/.*\((.*)\)/, "$1")}</div>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Details of the selected slot */}
          {selectedSlot && (
            <motion.div 
              key={selectedSlot.id}
              initial={{ opacity: 0, y: 5 }}
              animate={{ opacity: 1, y: 0 }}
              className="grid grid-cols-1 md:grid-cols-3 gap-4 bg-white/5 p-4 rounded-xl border border-white/10"
            >
              <div className="space-y-1">
                <span className="text-[10px] text-emerald-400 font-bold tracking-widest uppercase">Slot Identifier</span>
                <h4 className="text-xl font-black text-white">{selectedSlot.name}</h4>
                <div className="flex gap-2 text-xs font-mono font-medium text-slate-400 mt-1">
                  <span>Starts: {selectedSlot.pStart} ms</span>
                  <span>|</span>
                  <span>Width: {selectedSlot.widthMs} ms</span>
                </div>
              </div>
              <div className="space-y-1">
                <span className="text-[10px] text-slate-400 font-bold tracking-widest uppercase">Assigned Entity</span>
                <p className="text-sm font-semibold text-slate-200">{selectedSlot.assignedTo}</p>
              </div>
              <div className="space-y-1 md:border-l md:border-white/10 md:pl-4">
                <span className="text-[10px] text-slate-400 font-bold tracking-widest uppercase">Purpose / Amendment-10 Note</span>
                <p className="text-xs text-slate-300 font-medium leading-relaxed">{selectedSlot.desc}</p>
              </div>
            </motion.div>
          )}
        </div>
      </div>

      {/* Main Binder Explorer Tabs and Split Pane */}
      <div className="grid grid-cols-1 xl:grid-cols-4 gap-6">
        
        {/* Left Column: Selection lists & Searchbar */}
        <div className="xl:col-span-1 space-y-4">
          <div className="glass-card p-4 rounded-2xl border border-white/5 space-y-4">
            <div className="relative">
              <Search className="absolute left-3 top-3 w-4 h-4 text-slate-500" />
              <input
                type="text"
                placeholder="Search specs & rules..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full bg-white/5 border border-white/10 rounded-xl pl-9 pr-4 py-2.5 text-xs text-white focus:outline-none focus:border-emerald-500/50 transition-all placeholder:text-slate-600 font-medium"
              />
            </div>

            <div className="space-y-1.5">
              <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest px-2">Uploaded Annexures</span>
              <div className="space-y-1">
                {filteredAnnexList.map(([id, data]) => {
                  const isSelected = activeAnnex === id;
                  return (
                    <button
                      key={id}
                      onClick={() => setActiveAnnex(id)}
                      className={cn(
                        "w-full text-left p-3 rounded-xl transition-all border text-xs flex justify-between items-center group cursor-pointer",
                        isSelected 
                          ? "bg-emerald-500/10 border-emerald-500/20 text-emerald-400 font-bold" 
                          : "bg-white/5 border-transparent text-slate-300 hover:bg-white/10"
                      )}
                    >
                      <div className="space-y-0.5 truncate">
                        <div className="font-bold flex items-center gap-1.5">
                          <span className={cn(
                            "w-1.5 h-1.5 rounded-full",
                            isSelected ? "bg-emerald-400" : "bg-slate-500"
                          )} />
                          {id}
                        </div>
                        <div className="text-[10px] text-slate-400 opacity-90 truncate max-w-xs">{data.title}</div>
                      </div>
                      <ChevronRight className="w-3.5 h-3.5 text-slate-500 group-hover:text-slate-300 group-hover:transform group-hover:translate-x-0.5 transition-all" />
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Spec Document Details */}
        <div className="xl:col-span-3">
          {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex] ? (
            <motion.div
              key={activeAnnex}
              initial={{ opacity: 0, scale: 0.99 }}
              animate={{ opacity: 1, scale: 1 }}
              className="glass-card p-6 rounded-2xl border border-white/5 space-y-6"
            >
              {/* Doc Title Banner */}
              <div className="flex flex-col sm:flex-row justify-between items-start gap-4 pb-4 border-b border-white/5">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="text-xl font-extrabold text-white">{activeAnnex}: {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].title}</span>
                  </div>
                  <p className="text-xs text-slate-400 font-medium">
                    {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].description}
                  </p>
                </div>
                <span className="shrink-0 text-xs font-black text-emerald-400 bg-emerald-500/10 px-3 py-1 rounded-lg border border-emerald-500/20 uppercase tracking-widest">
                  {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].amendment}
                </span>
              </div>

              {/* If Annex has configurable parameters list */}
              {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].parameters && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Sliders className="w-4 h-4 text-emerald-400" />
                    <span className="text-xs font-bold text-slate-300 uppercase tracking-wider">Configurable Core Parameters Reference Table</span>
                  </div>
                  <div className="bg-black/20 rounded-xl border border-white/10 overflow-hidden text-xs">
                    <div className="grid grid-cols-12 bg-white/5 p-2 font-bold text-slate-400 border-b border-white/5 uppercase text-[10px] tracking-wider">
                      <div className="col-span-4">Parameter Name</div>
                      <div className="col-span-5">Specification Definition</div>
                      <div className="col-span-1.5 text-center">Default</div>
                      <div className="col-span-1.5 text-right">Range / Limit</div>
                    </div>
                    <div className="divide-y divide-white/5">
                      {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].parameters?.map(p => (
                        <div key={p.name} className="grid grid-cols-12 p-3 font-medium text-slate-300">
                          <div className="col-span-4 font-bold text-white text-[11px] font-mono leading-tight">{p.name}</div>
                          <div className="col-span-5 opacity-90 text-[11px] leading-relaxed">{p.desc}</div>
                          <div className="col-span-1.5 text-center font-mono text-emerald-400 font-bold bg-emerald-500/5 py-1 px-1.5 rounded">{p.defaultValue} {p.unit}</div>
                          <div className="col-span-1.5 text-right font-mono text-[10px] opacity-80">{p.limitRange}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* If Annex has Tag formats (Annexure-D) */}
              {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].dataFormats && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Database className="w-4 h-4 text-emerald-400" />
                    <span className="text-xs font-bold text-slate-300 uppercase tracking-wider">RFID 128-Bit Tag Payload Struct Layout</span>
                  </div>
                  <div className="bg-black/20 rounded-xl border border-white/10 overflow-hidden text-xs">
                    <div className="grid grid-cols-12 bg-white/5 p-2.5 font-bold text-slate-400 border-b border-white/5 uppercase text-[10px] tracking-wider">
                      <div className="col-span-3">Bit Fields</div>
                      <div className="col-span-2">Telemetry Bits</div>
                      <div className="col-span-2">Spec Default</div>
                      <div className="col-span-5 text-right">Functional Meaning & Directives</div>
                    </div>
                    <div className="divide-y divide-white/5">
                      {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].dataFormats?.map(f => (
                        <div key={f.field} className="grid grid-cols-12 p-3 text-slate-300 font-medium">
                          <div className="col-span-3 font-extrabold text-white text-[11px]">{f.field}</div>
                          <div className="col-span-2 font-mono text-xs text-slate-400">{f.bits}</div>
                          <div className="col-span-2 font-mono text-[11px] text-emerald-400">{f.defaultValue || "-"}</div>
                          <div className="col-span-5 text-right text-[11px] leading-relaxed opacity-90">{f.desc}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* If Annex has Protocols or Message maps (Annexure-G, Annexure-P, Annexure-A1) */}
              {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].protocols && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Layers className="w-4 h-4 text-emerald-400" />
                    <span className="text-xs font-bold text-slate-300 uppercase tracking-wider">Safety Protocols Messages & State Keys Map</span>
                  </div>
                  <div className="bg-black/20 rounded-xl border border-white/10 overflow-hidden text-xs">
                    <div className="grid grid-cols-12 bg-white/5 p-2.5 font-bold text-slate-400 border-b border-white/5 uppercase text-[10px] tracking-wider">
                      <div className="col-span-3">Hex Code / Key</div>
                      <div className="col-span-4">Protocol Event Title</div>
                      <div className="col-span-5 text-right">Operational Objective & Safety Behavior</div>
                    </div>
                    <div className="divide-y divide-white/5">
                      {KAVACH_OFFICIAL_SPECS.annexures[activeAnnex].protocols?.map(p => (
                        <div key={p.msgType} className="grid grid-cols-12 p-3 text-slate-300 font-medium">
                          <div className="col-span-3 font-mono font-black text-emerald-400 text-xs bg-white/5 py-1 px-2.5 rounded max-w-max border border-white/5">{p.msgType}</div>
                          <div className="col-span-4 font-bold text-white text-[11px] leading-tight">{p.value}</div>
                          <div className="col-span-5 text-right text-[11px] leading-relaxed opacity-90">{p.purpose}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {/* Safe-Side Formulas Callout specifically for Annexure-C/H/I */}
              {activeAnnex === "Annexure-D" && (
                <div className="bg-white/5 p-4 rounded-xl border border-white/10 space-y-3">
                  <div className="flex items-center gap-2 text-xs font-bold text-emerald-400 uppercase tracking-wider">
                    <Activity className="w-4 h-4" />
                    <span>L_DOUBTOVER & L_DOUBT_UNDER Safe-Side Bounds</span>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs font-medium">
                    <div className="bg-black/20 p-3 rounded border border-white/5 space-y-1">
                      <div className="font-bold text-white uppercase text-[11px]">L_DOUBTOVER Formulation:</div>
                      <p className="text-slate-300 leading-relaxed text-[11px]">
                        Over-reading + 5m location accuracy + 5% odometer error + Reader Offset in Rear End (ROR).
                      </p>
                    </div>
                    <div className="bg-black/20 p-3 rounded border border-white/5 space-y-1">
                      <div className="font-bold text-white uppercase text-[11px]">L_DOUBTUNDER Formulation:</div>
                      <p className="text-slate-300 leading-relaxed text-[11px]">
                        Under-reading + 5m location accuracy + 5% odometer error + Reader Offset from Front End (RORF).
                      </p>
                    </div>
                  </div>
                </div>
              )}
            </motion.div>
          ) : (
            <div className="glass-card p-12 rounded-2xl border border-white/5 text-center text-slate-500">
              <ClipboardList className="w-12 h-12 mx-auto text-slate-600 mb-3" />
              <p className="text-xs">No specifications match search filter.</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
