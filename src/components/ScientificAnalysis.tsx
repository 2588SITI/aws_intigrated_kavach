import React from 'react';
import { motion } from 'motion/react';
import { 
  Calculator, 
  RefreshCw, 
  Zap, 
  ShieldAlert, 
  TrendingUp, 
  Activity,
  ArrowRight,
  ClipboardList,
  AlertTriangle,
  History
} from 'lucide-react';
import { DashboardStats } from '../types';

export function ScientificAnalysis({ stats }: { stats: DashboardStats }) {
  const insights = stats.scientificInsights;

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-8"
    >
      {/* Real-time Detections based on the Model */}
      {insights && insights.highRiskScenarios.length > 0 && (
        <div className="glass-card p-6 rounded-3xl border border-rose-500/20 bg-rose-500/5">
          <div className="flex items-center gap-3 mb-6">
            <div className="w-10 h-10 rounded-xl bg-rose-500/20 flex items-center justify-center border border-rose-500/20">
              <AlertTriangle className="w-5 h-5 text-rose-400" />
            </div>
            <div>
              <h3 className="text-lg font-black text-white uppercase tracking-tighter">Model Violations Detected</h3>
              <p className="text-rose-400/70 text-xs font-bold uppercase tracking-widest">Mathematical Proof Validation in Current Logs</p>
            </div>
          </div>

          <div className="space-y-3">
             {insights.highRiskScenarios.map((s, i) => (
               <div key={i} className="flex flex-col md:flex-row md:items-center justify-between p-4 bg-black/40 rounded-2xl border border-white/5 gap-4">
                 <div className="flex items-center gap-4">
                    <div className="text-center px-3 py-1 bg-white/5 rounded-lg border border-white/10">
                       <p className="text-[10px] font-bold text-slate-500 uppercase">Time</p>
                       <p className="text-xs font-black text-white tracking-widest leading-none">{s.time}</p>
                    </div>
                    <div>
                       <p className="text-sm font-bold text-white uppercase">{s.stationName}</p>
                       <p className="text-[10px] text-slate-400 font-medium">Loco {s.locoId} • {s.description}</p>
                    </div>
                 </div>
                 <div className="flex items-center gap-4">
                    <div className="text-right">
                       <p className="text-[10px] font-black text-rose-400 uppercase tracking-widest italic">Est. Uncertainty (σ²)</p>
                       <p className="text-lg font-black text-rose-400 tabular-nums leading-none">{s.estimatedUncertainty} m²</p>
                    </div>
                    <div className="w-12 h-12 rounded-full border-4 border-rose-500/20 border-t-rose-500 flex items-center justify-center">
                       <span className="text-[10px] font-black text-rose-400">KO</span>
                    </div>
                 </div>
               </div>
             ))}
          </div>
        </div>
      )}
      <div className="glass-card p-8 rounded-3xl border border-white/5 relative overflow-hidden">
        <div className="absolute top-0 right-0 p-12 opacity-5 pointer-events-none">
          <Calculator className="w-64 h-64" />
        </div>
        
        <div className="relative">
          <div className="flex items-center gap-4 mb-6">
            <div className="w-12 h-12 rounded-2xl bg-emerald-500/20 flex items-center justify-center border border-emerald-500/20">
              <Calculator className="w-6 h-6 text-emerald-400" />
            </div>
            <div>
              <h2 className="text-2xl font-black text-white tracking-tighter uppercase">Mathematical Proof</h2>
              <p className="text-slate-400 text-sm font-medium">RFID Tag Spacing Defect → TCAS Degradation Correlation</p>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            <div className="space-y-6">
              <div className="space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <Activity className="w-5 h-5 text-emerald-400" />
                  System Model Definition
                </h3>
                <p className="text-slate-400 text-sm leading-relaxed">
                  Let the TCAS positioning system be modeled as a Kalman Filter-based state estimator:
                </p>
                <div className="bg-black/40 p-6 rounded-2xl border border-white/5 font-mono text-emerald-400 text-center text-lg italic">
                  x̂ₖ = Fx̂ₖ₋₁ + Buₖ + wₖ
                </div>
                <div className="text-[11px] text-slate-500 italic space-y-1">
                  <p>x̂ₖ = estimated train position at time k</p>
                  <p>F = state transition matrix</p>
                  <p>wₖ ∼ N(0, Qₖ) = process noise (e.g., wheel slip/creep)</p>
                </div>
              </div>

              <div className="space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <ShieldAlert className="w-5 h-5 text-amber-400" />
                  Step 1: The Spacing Defect
                </h3>
                <div className="glass-card p-5 border-l-4 border-l-amber-500 bg-amber-500/5">
                  <p className="text-sm text-slate-300 font-medium italic">
                    "When INTER TAG DIST &gt; DUP TAG spacing, the tag read event arrives late."
                  </p>
                </div>
                <div className="bg-black/20 p-4 rounded-xl space-y-2 font-mono text-xs text-slate-400">
                  <p>Actual position: Tᵢᵈᵉᶠᵉᶜᵗ = Tᵢ + Δdᵢ (where Δdᵢ &gt; 0)</p>
                  <p>Timing error: Δtₑᵣᵣₒᵣ = Δdᵢ / v</p>
                  <p>Position fix error: εᵈᵉᶠᵉᶜᵗ = Δdᵢ ≫ δₜₕᵣₑₛₕₒₗᵈ</p>
                </div>
                <div className="flex items-center gap-2 text-emerald-400 font-black text-xs uppercase tracking-widest bg-emerald-500/10 w-fit px-3 py-1 rounded-full border border-emerald-500/10">
                   <TrendingUp className="w-3 h-3" /> Result: Persistent Positive Bias
                </div>
              </div>
            </div>

            <div className="space-y-6">
              <div className="space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <RefreshCw className="w-5 h-5 text-blue-400" />
                  Step 2: Kalman Gain Sluggishness
                </h3>
                <p className="text-slate-400 text-sm">
                  The Innovation residual (νₖ) inflates, causing the Kalman Gain (Kₖ) to decrease:
                </p>
                <div className="bg-black/40 p-6 rounded-2xl border border-white/5 font-mono text-blue-400 text-center text-lg italic">
                  E[νₖᵈᵉᶠᵉᶜᵗ] = Δdᵢ ≠ 0
                </div>
                <p className="text-[11px] text-slate-500 italic">
                  Result: The filter over-corrects or trusts measurements less (Kₖᵈᵉᶠᵉᶜᵗ &lt; Kₖⁿᵒʳᵐᵃˡ).
                </p>
              </div>

              <div className="space-y-4">
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <Zap className="w-5 h-5 text-rose-400" />
                  Step 3: Background Pressure Growth
                </h3>
                <p className="text-slate-400 text-sm">
                  The position error covariance (Pₖ) grows monotonically between fixes:
                </p>
                <div className="bg-black/40 p-6 rounded-2xl border border-white/5 font-mono text-rose-400 text-center text-lg italic">
                  Pₙᵈᵉᶠᵉᶜᵗ ∝ n · Δdᵢ²
                </div>
                <div className="bg-rose-500/10 border border-rose-500/20 p-4 rounded-xl">
                  <p className="text-xs text-rose-300 font-bold leading-relaxed">
                    ✅ Proved: Persistent defect causes monotonically increasing background error pressure, making mode degradation mathematically guaranteed during radio switches.
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="glass-card p-6 rounded-3xl border border-white/5">
          <h3 className="text-sm font-black text-white uppercase tracking-[0.2em] mb-4 flex items-center gap-2">
            <ClipboardList className="w-4 h-4 text-emerald-400" />
            Theoretical Foundation & Logic
          </h3>
          <div className="space-y-4">
            <div className="p-4 bg-emerald-500/5 rounded-2xl border border-emerald-500/10">
              <p className="text-xs text-emerald-300 font-bold mb-2">1. False Position Logic (The Bias)</p>
              <p className="text-[11px] text-slate-400 leading-relaxed italic">
                If a tag is placed forward of its designated coordinate, the system detects the "Fix" later than expected. This timing lag injects a persistent mathematical bias into the position estimator.
              </p>
            </div>
            <div className="p-4 bg-blue-500/5 rounded-2xl border border-blue-500/10">
              <p className="text-xs text-blue-300 font-bold mb-2">2. State Estimator Confusion (Kalman Filters)</p>
              <p className="text-[11px] text-slate-400 leading-relaxed italic">
                The Kalman Filter acts as the system's "intelligence." When tags arrive at inconsistent intervals, the Innovation Residual inflates, leading the system to distrust its own sensors—a state of reduced confidence.
              </p>
            </div>
            <div className="p-4 bg-rose-500/5 rounded-2xl border border-rose-500/10">
              <p className="text-xs text-rose-300 font-bold mb-2">3. Handover Breakdown Propagation</p>
              <p className="text-[11px] text-slate-400 leading-relaxed italic">
                The primary risk manifests during radio handovers. Because the system's "background pressure" or uncertainty was already elevated by previous RFID defects, any minor radio packet drop pushes the error beyond safety limits, triggering automatic degradation.
              </p>
            </div>
          </div>
        </div>

        <div className="glass-card p-6 rounded-3xl border border-white/5 flex flex-col">
          <h3 className="text-sm font-black text-white uppercase tracking-[0.2em] mb-4 flex items-center gap-2">
            <TrendingUp className="w-4 h-4 text-amber-400" />
            Numerical Impact Example
          </h3>
          <div className="flex-1 overflow-x-auto">
            <table className="w-full text-left text-[11px]">
               <thead>
                 <tr className="border-b border-white/10 text-slate-500">
                    <th className="py-2">Parameter</th>
                    <th className="py-2">Normal</th>
                    <th className="py-2 text-rose-400">With Defect</th>
                 </tr>
               </thead>
               <tbody className="text-slate-300 font-mono">
                 <tr className="border-b border-white/5">
                   <td className="py-2">Spacing Gap (Δdᵢ)</td>
                   <td className="py-2">0 m</td>
                   <td className="py-2 text-rose-400 font-bold">8 m (Defect)</td>
                 </tr>
                 <tr className="border-b border-white/5">
                   <td className="py-2">Entry pressure (Pₑₙₜᵣᵧ)</td>
                   <td className="py-2">2 m²</td>
                   <td className="py-2 text-rose-400 font-bold">322 m²</td>
                 </tr>
                 <tr className="border-b border-white/5">
                   <td className="py-2">Blackout (τ)</td>
                   <td className="py-2">5 s</td>
                   <td className="py-2">5 s</td>
                 </tr>
                 <tr>
                   <td className="py-2">Final Uncertainty</td>
                   <td className="py-2 text-emerald-400 font-bold">52 m² ✅ SAFE</td>
                   <td className="py-2 text-rose-400 font-bold italic underline">372 m² ❌ DEGRADE</td>
                 </tr>
               </tbody>
            </table>
          </div>
          <div className="mt-6 p-4 bg-black/40 rounded-2xl border border-white/5">
             <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mb-1 italic">Inventor Note:</p>
             <p className="text-[11px] text-slate-400 leading-relaxed font-medium">
                Rudolf Kálmán developed this algorithm for NASA to filter "Noise" and extract the true "State." In Kavach/TCAS, it maintains balance between wheel sensors and RFID tags. Any disruption in tag spacing destabilizes the entire state estimator.
             </p>
          </div>
        </div>
      </div>

      {/* Failure Mode Analysis: Why Degradation & Brakes Happen */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="glass-card p-6 rounded-3xl border border-rose-500/20 bg-rose-500/5">
          <div className="flex items-center gap-3 mb-4">
             <div className="w-8 h-8 rounded-lg bg-rose-500/20 flex items-center justify-center">
                <ShieldAlert className="w-4 h-4 text-rose-400" />
             </div>
             <h4 className="text-sm font-black text-white uppercase tracking-wider">Mode Degradation Logic</h4>
          </div>
          <div className="space-y-3 text-[11px] leading-relaxed text-slate-400">
             <p>
                <strong className="text-rose-300">Trigger:</strong> Uncertainty (σ²) &gt; Safety Threshold (Σ_max).
             </p>
             <p>
                <strong>Reasoning:</strong> When RFID spacing is defective, the Kalman Filter's "Self-Confidence" drops. During a radio handover, if the system cannot prove the train's position within a 5-meter margin (Confidence Interval), it can no longer support "Full Supervision."
             </p>
             <p className="italic bg-black/20 p-2 rounded-lg border border-white/5">
                Result: System drops to 'Limited Supervision' or 'Staff Responsible' where the driver must manually verify safety.
             </p>
             <div className="pt-2 border-t border-white/5 text-[9px] text-slate-500">
                <p>Standard Limit (Σ_max): ~5.0m (2500 cm²) for Full Supervision.</p>
                <p>Degradation Threshold: &gt; 10.0m (10000 cm²).</p>
             </div>
          </div>
        </div>

        <div className="glass-card p-6 rounded-3xl border border-amber-500/20 bg-amber-500/5">
          <div className="flex items-center gap-3 mb-4">
             <div className="w-8 h-8 rounded-lg bg-amber-500/20 flex items-center justify-center">
                <Zap className="w-4 h-4 text-amber-400" />
             </div>
             <h4 className="text-sm font-black text-white uppercase tracking-wider">Kavach Brake Trigger Logic</h4>
          </div>
          <div className="space-y-3 text-[11px] leading-relaxed text-slate-400">
             <p>
                <strong className="text-amber-300">Trigger:</strong> (Position + Offset Error) &gt; Safe Wall Distance.
             </p>
             <p>
                <strong>Reasoning:</strong> The RFID spacing bias (Δd) creates a "Ghost Position." If the speed is high, the system calculates that the "True Position" might have already entered a danger zone (Signal at Danger) due to the accumulated error.
             </p>
             <p className="italic bg-black/20 p-2 rounded-lg border border-white/5">
                Result: Automatic EB (Emergency Brake) application because the mathematical safety envelope has been breached.
             </p>
          </div>
        </div>
      </div>
      
      {/* Event-Specific Scientific Audit Trail */}
      {stats.scientificInsights?.eventAudits && stats.scientificInsights.eventAudits.length > 0 && (
        <div className="glass-card p-6 rounded-3xl border border-emerald-500/20 bg-emerald-500/5 mt-6">
             <div className="flex items-center gap-3 mb-6">
                <div className="w-10 h-10 rounded-xl bg-emerald-500/20 flex items-center justify-center border border-emerald-500/20">
                    <History className="w-5 h-5 text-emerald-400" />
                </div>
                <div>
                   <h3 className="text-lg font-black text-white uppercase tracking-tighter">Event-Specific Scientific Audit</h3>
                   <p className="text-emerald-400/70 text-xs font-bold uppercase tracking-widest">Mathematical Correlation of Mode Drops & Brakes</p>
                </div>
             </div>

             <div className="space-y-4">
                {stats.scientificInsights.eventAudits.map((event, idx) => (
                   <div key={idx} className="p-4 rounded-2xl bg-black/40 border border-white/5 hover:border-emerald-500/30 transition-all group">
                      <div className="flex flex-wrap items-center justify-between gap-4 mb-3">
                         <div className="flex items-center gap-3">
                            <span className={`px-2 py-0.5 rounded-md text-[9px] font-black uppercase tracking-tighter ${
                               event.type === 'Degradation' ? 'bg-rose-500/20 text-rose-400' : 'bg-amber-500/20 text-amber-400'
                            }`}>
                               {event.type}
                            </span>
                            <span className="text-xs font-mono font-bold text-white tracking-widest">{event.time}</span>
                            <span className="text-[10px] text-slate-500 font-bold uppercase">{event.station}</span>
                            <span className="text-[10px] text-slate-400 font-medium">LOCO {event.locoId}</span>
                         </div>
                         <div className="text-[10px] font-bold text-emerald-400/80 uppercase italic">
                            Trigger: {event.trigger}
                         </div>
                      </div>
                      <p className="text-[11px] text-slate-400 leading-relaxed font-medium bg-white/5 p-3 rounded-xl border border-white/5">
                         <span className="text-emerald-400/60 font-bold mr-2 uppercase text-[9px]">Scientific Verdict:</span>
                         {event.scientificVerdict}
                      </p>
                   </div>
                ))}
             </div>
        </div>
      )}

      {/* External System Fault Correlation */}
      {stats.faultLogs && stats.faultLogs.length > 0 && (
        <div className="glass-card p-6 rounded-3xl border border-rose-500/20 bg-rose-500/5">
          <div className="flex items-center gap-3 mb-6">
            <div className="w-10 h-10 rounded-xl bg-rose-500/20 flex items-center justify-center border border-rose-500/20">
              <ShieldAlert className="w-5 h-5 text-rose-400" />
            </div>
            <div>
              <h3 className="text-lg font-black text-white uppercase tracking-tighter">External System Faults (Correlated)</h3>
              <p className="text-rose-400/70 text-xs font-bold uppercase tracking-widest">Ground-Truth Audit from NMS V2 / Fault Monitoring</p>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-[11px]">
               <thead>
                 <tr className="border-b border-white/10 text-slate-500">
                    <th className="py-2 px-4">Time</th>
                    <th className="py-2 px-4">Station</th>
                    <th className="py-2 px-4">Loco</th>
                    <th className="py-2 px-4">Fault Description</th>
                    <th className="py-2 px-4">Status</th>
                 </tr>
               </thead>
               <tbody className="text-slate-300">
                 {stats.faultLogs.slice(0, 20).map((f, i) => (
                   <tr key={i} className="border-b border-white/5 hover:bg-white/5 transition-colors">
                     <td className="py-3 px-4 font-mono font-bold text-rose-400">{f.time}</td>
                     <td className="py-3 px-4 uppercase font-black text-white">{f.station}</td>
                     <td className="py-3 px-4">{f.locoId}</td>
                     <td className="py-3 px-4 font-medium text-slate-400">{f.faultMsg}</td>
                     <td className="py-3 px-4">
                       <span className={`px-2 py-0.5 rounded-full text-[9px] font-bold uppercase ${
                         f.status.toLowerCase().includes('fail') || f.status.toLowerCase().includes('fault')
                         ? 'bg-rose-500/20 text-rose-400 border border-rose-500/20'
                         : 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/20'
                       }`}>
                         {f.status}
                       </span>
                     </td>
                   </tr>
                 ))}
               </tbody>
            </table>
          </div>
        </div>
      )}
    </motion.div>
  );
}
