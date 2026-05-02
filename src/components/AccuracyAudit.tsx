import React from 'react';
import { CheckCircle2, AlertTriangle, XCircle, Info, Activity, Clock } from 'lucide-react';
import { cn } from '../lib/utils';

export function RootCauseAccuracyAudit() {
  const verifiableEvents = [
    { 
      id: "37352", time: "07:10:44", stn: "ST STATION", trans: "FS → StandBy", 
      verdict: "wrong", pdfExplanation: "LOCO HARDWARE FAULT — NMS code 8 caused internal processor sync loss",
      realCause: "MA Expiry. Speed=0, Signal=Red, InterTagDist on ALL 22 rows. Station not renewing MA.",
      finding: "PDF wrongly blames loco NMS=8 (background noise). Action should be ST station tags/MA logic."
    },
    { 
      id: "39018", time: "07:18:14", stn: "ST STATION", trans: "FS → StandBy", 
      verdict: "wrong", pdfExplanation: "MOMENTARY PACKET GAP — Active Radio switched (Ra... handover)",
      realCause: "No radio switch occurred (standard R1/R2 polling). Speed=0, InterTagDist on ALL rows.",
      finding: "Identical to 37352. Both should blame ST STATION MA expiry + InterTagDist."
    },
    { 
      id: "37943", time: "18:33:30", stn: "NVS STATION", trans: "FS → SR", 
      verdict: "wrong", pdfExplanation: "RADIO INTERNAL SYNC FAILURE — NMS 40/16 Vital Hardware Error",
      realCause: "NMS only 0/1 (no 16/40). Persistent InterTagDist on ALL rows.",
      finding: "NMS claim has no data support. Real cause is NVS RFID tags (Station side)."
    },
    { 
      id: "37086", time: "05:41:46", stn: "BL STATION", trans: "OS → SR", 
      verdict: "partial", pdfExplanation: "RADIO INTERNAL SYNC FAILURE — conflicting packets SR/OS",
      realCause: "Radio 1→2 switchover at exact moment + 1-packet link dropout.",
      finding: "This WAS a radio handover gap. PDF swapped this explanation with a wrong one for another event."
    },
    { 
      id: "37424", time: "05:18:48", stn: "BL STATION", trans: "FS → OS", 
      verdict: "partial", pdfExplanation: "RADIO INTERNAL SYNC FAILURE — NMS 40/16, OS/FS packets",
      realCause: "NMS 32 conflict is real, BUT NMS 40 was post-event and InterTagDist (16/16 rows) was ignored.",
      finding: "Evidence is real but timing is inverted and infrastructure pressure ignored."
    },
    { 
      id: "37424", time: "18:41:38", stn: "NVS STATION", trans: "FS → LS", 
      verdict: "correct", pdfExplanation: "KAVACH SAFETY TRIGGER — ReadEndCollision",
      realCause: "Emr Status = ReadEndCollision active on all rows.",
      finding: "Correct identification and action."
    }
  ];

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h2 className="text-xl font-black text-white uppercase tracking-tight flex items-center gap-2">
          <Activity className="w-6 h-6 text-emerald-400" />
          Root Cause Accuracy Audit
        </h2>
        <div className="flex items-center gap-4">
           <div className="flex items-center gap-1.5 px-3 py-1 bg-emerald-500/10 border border-emerald-500/20 rounded-full">
              <div className="w-2 h-2 rounded-full bg-emerald-500" />
              <span className="text-[10px] font-bold text-emerald-400 uppercase">Correct</span>
           </div>
           <div className="flex items-center gap-1.5 px-3 py-1 bg-amber-500/10 border border-amber-500/20 rounded-full">
              <div className="w-2 h-2 rounded-full bg-amber-500" />
              <span className="text-[10px] font-bold text-amber-400 uppercase">Debatable</span>
           </div>
           <div className="flex items-center gap-1.5 px-3 py-1 bg-rose-500/10 border border-rose-500/20 rounded-full">
              <div className="w-2 h-2 rounded-full bg-rose-500" />
              <span className="text-[10px] font-bold text-rose-400 uppercase">Incorrect</span>
           </div>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4">
        {verifiableEvents.map((ev, i) => (
          <div key={i} className={cn(
            "glass-card overflow-hidden border-l-4",
            ev.verdict === 'correct' ? "border-emerald-500" : ev.verdict === 'partial' ? "border-amber-500" : "border-rose-500"
          )}>
            <div className="bg-white/5 px-6 py-3 border-b border-white/5 flex items-center justify-between">
              <div className="flex items-center gap-3">
                {ev.verdict === 'correct' && <CheckCircle2 className="w-4 h-4 text-emerald-500" />}
                {ev.verdict === 'partial' && <AlertTriangle className="w-4 h-4 text-amber-500" />}
                {ev.verdict === 'wrong' && <XCircle className="w-4 h-4 text-rose-500" />}
                <span className="font-mono text-xs font-bold text-slate-400">{ev.id} | {ev.time} | {ev.stn}</span>
              </div>
              <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">{ev.trans}</span>
            </div>
            
            <div className="p-5 grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <p className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Report Explanation</p>
                <p className="text-sm text-slate-300 font-medium italic">"{ev.pdfExplanation}"</p>
              </div>
              <div className="space-y-2">
                <p className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Actual Log Ground-Truth</p>
                <div className="flex gap-2">
                   <Info className="w-4 h-4 text-blue-400 shrink-0 mt-0.5" />
                   <p className="text-sm text-white font-semibold">{ev.realCause}</p>
                </div>
              </div>
            </div>

            <div className={cn(
              "px-6 py-3 text-[11px] font-medium border-t border-white/5",
              ev.verdict === 'correct' ? "bg-emerald-500/5 text-emerald-400" : ev.verdict === 'partial' ? "bg-amber-500/5 text-amber-400" : "bg-rose-500/5 text-rose-400"
            )}>
              <span className="font-black uppercase mr-2 tracking-widest">Audit Finding:</span>
              {ev.finding}
            </div>
          </div>
        ))}
      </div>

      <div className="glass-card p-6 border-l-4 border-blue-500">
         <h4 className="text-sm font-black text-white uppercase tracking-wider mb-4">Core Algorithm Upgrade Notes</h4>
         <div className="grid grid-cols-1 md:grid-cols-2 gap-y-4 gap-x-8 text-xs text-slate-400">
           <div className="flex gap-3">
             <div className="w-1.5 h-1.5 rounded-full bg-blue-500 shrink-0 mt-1.5" />
             <p><span className="text-white font-bold">NMS Timing Guard:</span> Hardware errors are now split into Pre-Event and Post-Event. Causes are only attributed if errors exist BEFORE the transition.</p>
           </div>
           <div className="flex gap-3">
             <div className="w-1.5 h-1.5 rounded-full bg-blue-500 shrink-0 mt-1.5" />
             <p><span className="text-white font-bold">Infras. Weighting:</span> Tag spacing errors (InterTagDist) now override minor NMS alerts if present in {'>'}60% of data rows.</p>
           </div>
           <div className="flex gap-3">
             <div className="w-1.5 h-1.5 rounded-full bg-blue-500 shrink-0 mt-1.5" />
             <p><span className="text-white font-bold">Stationary Logic:</span> Specifically detects Speed 0 + StandBy fallback to identify Station MA expiry vs Loco hardware failure.</p>
           </div>
           <div className="flex gap-3">
             <div className="w-1.5 h-1.5 rounded-full bg-blue-500 shrink-0 mt-1.5" />
             <p><span className="text-white font-bold">Sync Precision:</span> Conflict detection now requires actual mode disagreement within the same second to trigger "Radio Sync Failure".</p>
           </div>
         </div>
      </div>
    </div>
  );
}
