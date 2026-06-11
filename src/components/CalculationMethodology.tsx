import React from 'react';
import { Calculator, Info, CheckCircle2, AlertCircle, ArrowRight, Layers, Divide, BookOpen, ShieldCheck, Radio, Database, Activity, FileText } from 'lucide-react';
import { motion } from 'motion/react';
import { KAVACH_OFFICIAL_SPECS } from '../utils/kavachSpecs';

export const CalculationMethodology: React.FC = () => {
  return (
    <div className="space-y-8 animate-in fade-in duration-700">
      <div className="flex items-center gap-4 mb-2">
        <div className="w-12 h-12 bg-emerald-500/10 rounded-xl flex items-center justify-center border border-emerald-500/20">
          <Calculator className="w-6 h-6 text-emerald-400" />
        </div>
        <div>
          <h2 className="text-2xl font-black text-white tracking-tighter uppercase">Calculation Methodology</h2>
          <p className="text-slate-400 text-sm font-medium">How we ensure 100% accuracy in station performance metrics</p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Rule 1: Deduplication */}
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="glass-card p-6 rounded-2xl relative overflow-hidden group"
        >
          <div className="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
            <Layers className="w-24 h-24 text-emerald-400" />
          </div>
          
          <div className="relative z-10 space-y-4">
            <div className="flex items-center gap-3">
              <span className="flex items-center justify-center w-8 h-8 rounded-full bg-emerald-500 text-white font-black text-sm">1</span>
              <h3 className="text-xl font-bold text-white">Deduplicate First</h3>
            </div>
            
            <p className="text-slate-300 text-sm leading-relaxed">
              Every row in the raw XLS appears <span className="text-emerald-400 font-bold">exactly twice</span> due to the logging mechanism. 
              If we skip this step, both numerator and denominator double. While the percentage remains the same, 
              your packet counts become inflated and incorrect for deep analysis.
            </p>

            <div className="bg-white/5 rounded-xl p-4 border border-white/10">
              <div className="flex items-center justify-between text-xs font-mono mb-4">
                <span className="text-slate-500 uppercase">Raw Data (2x)</span>
                <ArrowRight className="w-4 h-4 text-slate-600" />
                <span className="text-emerald-400 uppercase">Processed Data (1x)</span>
              </div>
              
              <div className="space-y-2">
                <div className="flex gap-2 opacity-50">
                  <div className="h-4 w-full bg-white/10 rounded border border-white/5" />
                  <div className="h-4 w-12 bg-rose-500/20 rounded border border-rose-500/20" />
                </div>
                <div className="flex gap-2">
                  <div className="h-4 w-full bg-white/10 rounded border border-white/5" />
                  <div className="h-4 w-12 bg-emerald-500/40 rounded border border-emerald-500/40" />
                </div>
                <div className="flex gap-2 opacity-50">
                  <div className="h-4 w-full bg-white/10 rounded border border-white/5" />
                  <div className="h-4 w-12 bg-rose-500/20 rounded border border-rose-500/20" />
                </div>
              </div>
              <p className="text-[10px] text-slate-500 mt-3 italic text-center">
                We use a composite key (Loco + Station + Direction + Time + Date) to filter out duplicate logs.
              </p>
            </div>
          </div>
        </motion.div>

        {/* Rule 2: Sum then Divide */}
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          transition={{ delay: 0.1 }}
          className="glass-card p-6 rounded-2xl relative overflow-hidden group border-l-4 border-emerald-500"
        >
          <div className="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
            <Divide className="w-24 h-24 text-emerald-400" />
          </div>

          <div className="relative z-10 space-y-4">
            <div className="flex items-center gap-3">
              <span className="flex items-center justify-center w-8 h-8 rounded-full bg-emerald-500 text-white font-black text-sm">2</span>
              <h3 className="text-xl font-bold text-white">Sum Then Divide</h3>
            </div>
            
            <p className="text-slate-300 text-sm leading-relaxed">
              Never average the percentage column. A short 48-packet trip should not have the same "weight" as a 559-packet trip. 
              We sum all expected packets and all received packets separately, then perform a single division.
            </p>

            <div className="grid grid-cols-2 gap-4">
              <div className="p-3 bg-rose-500/5 rounded-xl border border-rose-500/10">
                <div className="flex items-center gap-2 mb-2">
                  <AlertCircle className="w-3 h-3 text-rose-400" />
                  <span className="text-[10px] font-black text-rose-400 uppercase">Wrong Way</span>
                </div>
                <div className="text-xs text-slate-400 font-mono">
                  (98.2% + 99.1% + ...) / N
                  <div className="mt-1 text-rose-400 font-bold">Result: 98.49%</div>
                </div>
              </div>
              <div className="p-3 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                <div className="flex items-center gap-2 mb-2">
                  <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                  <span className="text-[10px] font-black text-emerald-400 uppercase">Correct Way</span>
                </div>
                <div className="text-xs text-slate-400 font-mono">
                  Σ Received / Σ Expected
                  <div className="mt-1 text-emerald-400 font-bold">Result: 98.61%</div>
                </div>
              </div>
            </div>
          </div>
        </motion.div>
      </div>

      {/* Real-world Example */}
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        whileInView={{ opacity: 1, scale: 1 }}
        viewport={{ once: true }}
        className="bg-emerald-500/10 border border-emerald-500/20 rounded-2xl p-8"
      >
        <div className="flex flex-col md:flex-row items-center gap-8">
          <div className="shrink-0 text-center md:text-left">
            <h4 className="text-emerald-400 font-black uppercase tracking-widest text-xs mb-1">Case Study</h4>
            <div className="text-4xl font-black text-white tracking-tighter">BL Station</div>
            <p className="text-slate-400 text-sm mt-2 max-w-xs">
              Actual data comparison showing the impact of weighted averages on reporting accuracy.
            </p>
          </div>
          
          <div className="flex-1 w-full grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="bg-white/5 p-4 rounded-xl border border-white/10 flex flex-col justify-center">
              <span className="text-[10px] text-slate-500 uppercase font-bold mb-1">Total Expected</span>
              <span className="text-2xl font-mono font-bold text-white">2,735</span>
            </div>
            <div className="bg-white/5 p-4 rounded-xl border border-white/10 flex flex-col justify-center">
              <span className="text-[10px] text-slate-500 uppercase font-bold mb-1">Total Received</span>
              <span className="text-2xl font-mono font-bold text-white">2,697</span>
            </div>
            <div className="bg-emerald-500 p-4 rounded-xl flex flex-col justify-center shadow-lg shadow-emerald-500/20">
              <span className="text-[10px] text-white/70 uppercase font-bold mb-1">Final Weighted %</span>
              <span className="text-3xl font-mono font-black text-white">98.61%</span>
            </div>
          </div>
        </div>
      </motion.div>

      {/* RDSO KAVACH Specification Reference Library */}
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        whileInView={{ opacity: 1, y: 0 }}
        viewport={{ once: true }}
        className="glass-card p-6 rounded-2xl border-t-2 border-emerald-500 relative overflow-hidden group space-y-6"
      >
        <div className="flex items-center gap-3">
          <BookOpen className="w-6 h-6 text-emerald-400" />
          <h3 className="text-xl font-bold text-white uppercase tracking-tight">Kavach RDSO Specifications & Official Reference Library</h3>
        </div>

        <p className="text-slate-300 text-sm leading-relaxed max-w-4xl">
          Based on the official <span className="text-emerald-400 font-bold">{KAVACH_OFFICIAL_SPECS.version}</span> specification documents and <span className="text-emerald-400 font-bold">Amendment-{KAVACH_OFFICIAL_SPECS.amendmentNo}</span> uploaded to this system, the following operational and network communication rules are loaded into the diagnostic state engine to drive logic and scientific verdicts:
        </p>

        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
          {/* Box 1: Multiple Access Slotting (Amdt 10) */}
          <div className="bg-white/5 rounded-xl p-5 border border-white/10 space-y-3">
            <div className="flex items-center gap-2 text-emerald-400 font-bold">
              <Radio className="w-4 h-4" />
              <span className="text-sm font-bold uppercase text-white">Multiple Access scheme (Amdt-10)</span>
            </div>
            <ul className="text-xs text-slate-300 space-y-2 list-disc list-inside font-medium leading-relaxed">
              <li>Each Multiple Access frame cycle spans <span className="text-emerald-400 font-mono">2000 ms</span>.</li>
              <li>TDMA/FDMA/SDMA contains <span className="text-emerald-400 font-mono">70 time slots</span> (previously 68), width of <span className="text-emerald-400 font-mono">432 bits (22.5 ms)</span>.</li>
              <li>P2 starts exactly at <span className="text-emerald-300">45 ms</span> from cycle start. P47 starts at <span className="text-emerald-300">1320 ms</span>.</li>
              <li>Positions <span className="font-mono text-emerald-400">P1 & P46</span> are kept as reserve.</li>
              <li>Radio 1 transmission prefix: <span className="font-mono bg-black/40 px-1 py-0.5 rounded text-emerald-400">0xF1 0xA5 0xC3</span>.</li>
              <li>Radio 2 transmission prefix: <span className="font-mono bg-black/40 px-1 py-0.5 rounded text-emerald-400">0xF2 0xA5 0xC3</span>.</li>
            </ul>
          </div>

          {/* Box 2: Radio Failure Fallbacks (FRS 36.1) */}
          <div className="bg-white/5 rounded-xl p-5 border border-white/10 space-y-3">
            <div className="flex items-center gap-2 text-emerald-400 font-bold">
              <ShieldCheck className="w-4 h-4" />
              <span className="text-sm font-bold uppercase text-white">Radio Failure & Fallbacks (FRS 36.1)</span>
            </div>
            <ul className="text-xs text-slate-300 space-y-2 list-disc list-inside font-medium leading-relaxed">
              <li>If stationary packet is <span className="text-emerald-400 font-bold">&gt; 6 seconds</span> old, signal aspects turn blank.</li>
              <li>With active track profile (<span className="text-emerald-400 font-mono">&lt; 3000m</span>): FS/OS transits dynamically to <span className="text-emerald-400 font-bold">Limited Supervision (LS)</span>.</li>
              <li>Without active track profile: FS/OS transits dynamically to <span className="text-emerald-400 font-bold">Staff Responsible (SR)</span>.</li>
              <li>Stipulated Driver Ack Timeout: <span className="text-emerald-400 font-mono">15 seconds</span>.</li>
              <li>Failure to acknowledge transits: Automatic command to trigger <span className="text-rose-400 font-bold">Service Brakes (SB)</span>.</li>
            </ul>
          </div>

          {/* Box 3: Safe-Side Distance Calculations */}
          <div className="bg-white/5 rounded-xl p-5 border border-white/10 space-y-3 md:col-span-2 xl:col-span-1">
            <div className="flex items-center gap-2 text-emerald-400 font-bold">
              <Activity className="w-4 h-4" />
              <span className="text-sm font-bold uppercase text-white">Safe-Side Distance Formulas</span>
            </div>
            <div className="space-y-3 text-xs">
              <div className="space-y-1">
                <div className="text-emerald-400 font-mono font-bold uppercase">L_DOUBTOVER (Supervision of Targets):</div>
                <p className="text-slate-300 leading-relaxed bg-black/20 p-2 rounded border border-white/5">
                  Over-reading + 5m GPS/RFID accuracy + 5% odometer error + Reader Offset in Rear End (ROR). Evaluates PSR, TSR, Linking, Rear End Collisions.
                </p>
              </div>
              <div className="space-y-1">
                <div className="text-emerald-400 font-mono font-bold uppercase">L_DOUBTUNDER (Discarding Locations):</div>
                <p className="text-slate-300 leading-relaxed bg-black/20 p-2 rounded border border-white/5">
                  Under-reading + 5m GPS/RFID accuracy + 5% odometer error + Reader Offset from Front End (RORF). Evaluates Tag links, Head-on Collisions.
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Dynamic specifications mapping tables */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-4 border-t border-white/5">
          {/* Table A: Packet Types & Codes (Amdt L12) */}
          <div className="space-y-2">
            <div className="flex items-center gap-2 text-xs font-bold uppercase text-slate-400 tracking-wider">
              <Database className="w-3 h-3 text-slate-400" />
              <span>RF Protocol Packet Architecture (Amdt 10)</span>
            </div>
            <div className="bg-black/30 rounded-xl border border-white/10 overflow-hidden text-xs">
              <div className="grid grid-cols-4 bg-white/5 p-2 font-bold text-slate-400 border-b border-white/5">
                <div className="col-span-1">PKT_TYPE</div>
                <div className="col-span-3">Description / Operational Directive</div>
              </div>
              <div className="divide-y divide-white/5">
                {Object.values(KAVACH_OFFICIAL_SPECS.packetTypes).map(p => (
                  <div key={p.code} className="grid grid-cols-4 p-2.5 font-mono text-[11px]">
                    <div className="col-span-1 text-emerald-400 font-bold">{p.code}</div>
                    <div className="col-span-3 text-slate-300">{p.desc}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Table B: Tag Link Spacing Status codes */}
          <div className="space-y-2">
            <div className="flex items-center gap-2 text-xs font-bold uppercase text-slate-400 tracking-wider">
              <FileText className="w-3 h-3 text-slate-400" />
              <span>TAG_LINK_INFO Diagnostics Codes</span>
            </div>
            <div className="bg-black/30 rounded-xl border border-white/10 overflow-hidden text-xs">
              <div className="grid grid-cols-4 bg-white/5 p-2 font-bold text-slate-400 border-b border-white/5">
                <div className="col-span-1">TAG_LINK</div>
                <div className="col-span-3">Diagnostic Status Definition</div>
              </div>
              <div className="divide-y divide-white/5 max-h-56 overflow-y-auto">
                {Object.entries(KAVACH_OFFICIAL_SPECS.tagLinkInfo).map(([code, p]) => (
                  <div key={code} className="grid grid-cols-4 p-2 font-mono text-[11px]">
                    <div className="col-span-1 text-emerald-400 font-bold">{code}</div>
                    <div className="col-span-3 text-slate-300">{p.desc}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </motion.div>

      <div className="flex items-center gap-2 text-slate-500 text-xs italic justify-center">
        <Info className="w-3 h-3" />
        <span>This logic is applied globally across all station and locomotive performance reports in this dashboard.</span>
      </div>
    </div>
  );
};
