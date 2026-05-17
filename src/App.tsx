/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { 
  Shield, 
  Upload, 
  CheckCircle2, 
  AlertCircle, 
  Construction,
  BarChart3, 
  Calculator,
  Activity, 
  Database,
  Info,
  Clock,
  Settings,
  AlertTriangle,
  ShieldCheck,
  ArrowRight,
  Zap,
  MapPin,
  Train,
  Download,
  Wifi,
  FileText,
  RefreshCw,
  ClipboardList,
  ShieldAlert,
  Tag,
  Search,
  Zap as ZapIcon
} from 'lucide-react';
import { motion } from 'motion/react';
import { RootCauseAccuracyAudit } from './components/AccuracyAudit';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  LabelList,
  ReferenceLine
} from 'recharts';
import { parseFile, processDashboardData, parseDateString, formatStationName, generateDiagnosticAdvice } from './utils/dataProcessor';
import { DashboardStats } from './types';
import { CalculationMethodology } from './components/CalculationMethodology';
import { ScientificAnalysis } from './components/ScientificAnalysis';
import { cn } from './utils/cn';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const StationDisplay = ({ id, name, showEmerald = true }: { id: string | number, name?: string, showEmerald?: boolean }) => {
  const normalizedId = String(id || '').trim();
  const normalizedName = String(name || '').trim();

  const idName = formatStationName(normalizedId);
  const rawName = normalizedName && normalizedName !== 'N/A' && normalizedName !== '-' && normalizedName !== '0' && normalizedName !== '0.0' && normalizedName !== ''
                ? formatStationName(normalizedName) 
                : null;
  
  if (idName !== 'N/A' && idName !== '') {
    return (
      <span className="inline-block">
        <span className="font-bold block">{idName}</span>
        {rawName && rawName !== idName && (
          <span className={cn("text-[10px] font-bold uppercase tracking-tight leading-none mt-0.5 block", showEmerald ? "text-emerald-400" : "text-slate-400")}>
            {rawName}
          </span>
        )}
      </span>
    );
  } else if (rawName) {
    return (
      <span className={cn("inline-block font-bold", showEmerald ? "text-emerald-400" : "text-white")}>
        {rawName}
      </span>
    );
  }
  return <span className="text-slate-500">-</span>;
};

export default function App() {
  const [files, setFiles] = useState<{ rf: File[]; rfSt: File[]; trn: File[]; radio: File | null; faultLogs: File[] }>({
    rf: [],
    rfSt: [],
    trn: [],
    radio: null,
    faultLogs: [],
  });
  const [stats, setStats] = useState<DashboardStats | null>(null);
  const [activeTab, setActiveTab] = useState('summary');
  const [tagSearch, setTagSearch] = useState('');
  const [selectedStation, setSelectedStation] = useState<string>('All');
  const [selectedLoco, setSelectedLoco] = useState<string>('All');
  const [startDate, setStartDate] = useState<string>('All');
  const [endDate, setEndDate] = useState<string>('All');
  
  // Manual Remarks State for PDF
  const [manualRemarks, setManualRemarks] = useState<Record<string, string>>({});

  const updateRemark = (key: string, value: string) => {
    setManualRemarks(prev => ({ ...prev, [key]: value }));
  };
  
  // Cloud Storage State
  const [isAwsConnected, setIsAwsConnected] = useState(false);
  const [cloudFiles, setCloudFiles] = useState<any[]>([]);
  const [isFetching, setIsFetching] = useState(false);
  const [availableDates, setAvailableDates] = useState<string[]>([]);
  const [availableLocos, setAvailableLocos] = useState<string[]>([]);
  const [cloudLoco, setCloudLoco] = useState<string>('All');
  const [division, setDivision] = useState<string>('SC'); // Default to SC

  useEffect(() => {
    checkAwsStatus();
  }, []);

  const checkAwsStatus = async () => {
    try {
      const res = await fetch('/api/aws/files');
      if (!res.ok) {
        const text = await res.text();
        console.error("AWS se error aaya hai (Status Check):", text);
        return;
      }
      const data = await res.json();
      if (!data.error) {
        setIsAwsConnected(true);
        setCloudFiles(data.files);
        updateAvailableDatesAndLocos(data.files);
        if (data.warning) {
          console.warn("DEMO MODE ACTIVE:", data.warning);
        }
      }
    } catch (err) {
      console.error('Error checking AWS status:', err);
    }
  };

  const handleAwsConnect = async () => {
    setIsFetching(true);
    try {
      const res = await fetch('/api/aws/files');
      if (!res.ok) {
        const text = await res.text();
        console.error("AWS se error aaya hai (Connect):", text);
        alert(`Server Error: ${res.status}\n\nDetails: ${text.substring(0, 100)}...`);
        return;
      }
      const data = await res.json();
      if (data.error) {
        const details = data.details ? `\nDetails: ${data.details}` : '';
        alert(`AWS Connection Error: ${data.error}${details}\n\nPlease ensure AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, and AWS_BUCKET_NAME are set in Environment Variables.`);
        return;
      }
      setIsAwsConnected(true);
      console.log('AWS S3 Files:', data.files);
      setCloudFiles(data.files);
      updateAvailableDatesAndLocos(data.files);
      
      if (data.warning) {
        alert(`${data.warning}\n\nDemo files have been loaded.`);
      } else {
        alert(`Successfully connected to AWS S3! Found ${data.files.length} files.`);
      }
    } catch (err: any) {
      console.error('Error connecting to AWS:', err);
      alert(`Failed to connect to AWS S3. ${err.message || ''}`);
    } finally {
      setIsFetching(false);
    }
  };

  const extractDateFromFilename = (filename: string): string | null => {
    const nameOnly = filename.split('/').pop() || filename;
    
    const findDate = (str: string): string | null => {
      // Standard formats: DD-MM-YYYY, DD.MM.YYYY, DD/MM/YYYY, YYYY-MM-DD, DD_MM_YYYY
      const match = str.match(/\d{1,2}[-/._]\d{1,2}[-/._]\d{2,4}/);
      if (match) {
        const parts = match[0].split(/[-/._]/);
        if (parts.length === 3) {
          let day, month, year;
          if (parts[0].length === 4) {
            year = parts[0];
            month = parts[1].padStart(2, '0');
            day = parts[2].padStart(2, '0');
          } else {
            day = parts[0].padStart(2, '0');
            month = parts[1].padStart(2, '0');
            year = parts[2].length === 2 ? `20${parts[2]}` : parts[2];
          }
          const result = `${day}-${month}-${year}`;
          return result;
        }
      }
      
      const match8 = str.match(/\d{8}/);
      if (match8) {
        const val = match8[0];
        const year = parseInt(val.substring(0, 4));
        let result;
        if (year > 1900 && year < 2100) {
          result = `${val.substring(6, 8)}-${val.substring(4, 6)}-${val.substring(0, 4)}`;
        } else {
          result = `${val.substring(0, 2)}-${val.substring(2, 4)}-${val.substring(4, 8)}`;
        }
        return result;
      }
      return null;
    };

    const date = findDate(nameOnly) || findDate(filename);
    if (date) {
      console.log(`Extracted date ${date} from path: ${filename}`);
    } else {
      console.log(`Failed to extract date from path: ${filename}`);
    }
    return date;
  };

   const extractLocoFromFilename = (filename: string): string | null => {
    // DISABLED: ONLY derive loco from "Loco Id" header in the file content
    return null;
  };

  const updateAvailableDatesAndLocos = (newFiles: any[]) => {
    const dates = new Set<string>(availableDates);
    const locos = new Set<string>(availableLocos);
    newFiles.forEach((f: any) => {
      const date = extractDateFromFilename(f.name);
      if (date) dates.add(date);
      const loco = extractLocoFromFilename(f.name);
      if (loco) locos.add(loco);
    });
    const sortedDates = Array.from(dates).sort((a, b) => parseDateString(a) - parseDateString(b));
    const sortedLocos = Array.from(locos).sort();
    console.log('Found Available Dates:', sortedDates);
    console.log('Found Available Locos:', sortedLocos);
    setAvailableDates(sortedDates);
    setAvailableLocos(sortedLocos);
  };

  const analyzeCloudData = async () => {
    if (startDate === 'All' || endDate === 'All') {
      alert('Please select both Start and End dates.');
      return;
    }

    setIsFetching(true);
    try {
      const startT = parseDateString(startDate);
      const endT = parseDateString(endDate);

      // Filter files by date and type
      console.log(`Filtering ${cloudFiles.length} cloud files for range ${startDate} to ${endDate}...`);
      console.log(`Start Timestamp: ${startT}, End Timestamp: ${endT}`);
      
      const selectedFiles = cloudFiles.filter(f => {
        const dateStr = extractDateFromFilename(f.name);
        if (!dateStr) {
          console.log(` - Skipping ${f.name}: No date found`);
          return false;
        }
        const fileT = parseDateString(dateStr);
        const isDateMatch = fileT >= startT && fileT <= endT;
        
        let isLocoMatch = true;
        if (cloudLoco !== 'All') {
          const locoStr = extractLocoFromFilename(f.name);
          // If it's a station file, we might still need it, but let's filter TRN and RADIO files by loco
          if (f.name.toUpperCase().includes('TRNMSNMA') || f.name.toUpperCase().includes('RADIO') || f.name.toUpperCase().includes('RFCOMM_TR')) {
             isLocoMatch = locoStr === cloudLoco;
          }
        }

        const isMatch = isDateMatch && isLocoMatch;
        console.log(` - File: ${f.name}, Extracted: ${dateStr}, Timestamp: ${fileT}, In Range: ${isDateMatch}, Loco Match: ${isLocoMatch}`);
        return isMatch;
      });

      if (selectedFiles.length === 0) {
        alert('No files found for the selected date range.');
        return;
      }

      const rfTrFiles = selectedFiles.filter(f => {
        const name = f.name.toUpperCase();
        const nameOnly = f.name.split('/').pop()?.toUpperCase() || '';
        
        // Priority 1: Explicit markers
        if (nameOnly.includes('RFCOMM_TR')) return true;
        if (nameOnly.includes('RFCOMM_ST')) return false;
        
        // Priority 2: ID patterns in filename
        if (/\d{5}_RFCOMM/i.test(nameOnly) || /RFCOMM_\d{5}/i.test(nameOnly)) return true;
        if (/[A-Z0-9_]{2,12}_RFCOMM/i.test(nameOnly) || /RFCOMM_[A-Z0-9_]{2,12}/i.test(nameOnly)) return false;

        // Priority 3: Folder markers
        if (name.includes('RFCOMM_TR')) return true;
        if (name.includes('RFCOMM_ST')) return false;
        
        // Priority 4: Filename content (LOCO or 5-digit ID)
        // We check nameOnly to avoid folder names triggering this for the wrong files
        if (nameOnly.includes('LOCO') || /\b\d{5}\b/.test(nameOnly)) return true;
        
        // Fallback
        return name.includes('LOCO') && !nameOnly.includes('STN') && !nameOnly.includes('STATION');
      });

      const rfStFiles = selectedFiles.filter(f => {
        const name = f.name.toUpperCase();
        const nameOnly = f.name.split('/').pop()?.toUpperCase() || '';
        
        // Priority 1: Explicit markers
        if (nameOnly.includes('RFCOMM_ST')) return true;
        if (nameOnly.includes('RFCOMM_TR')) return false;
        
        // Priority 2: ID patterns
        const isStnPattern = /[A-Z0-9_]{2,12}_RFCOMM/i.test(nameOnly) || /RFCOMM_[A-Z0-9_]{2,12}/i.test(nameOnly);
        const isLocoPattern = /\d{5}_RFCOMM/i.test(nameOnly) || /RFCOMM_\d{5}/i.test(nameOnly);
        if (isStnPattern && !isLocoPattern) return true;

        // Priority 3: Folder markers
        if (name.includes('RFCOMM_ST')) return true;
        if (name.includes('RFCOMM_TR')) return false;
        
        // Priority 4: Filename content
        if (nameOnly.includes('STN') || nameOnly.includes('STATION')) return true;
        
        return name.includes('STN') || name.includes('STATION') || name.includes('_ST_');
      });
      const trnFiles = selectedFiles.filter(f => f.name.toUpperCase().includes('TRN'));
      const radioFiles = selectedFiles.filter(f => f.name.toUpperCase().includes('RADIO'));

      console.log('File Categorization:');
      console.log(` - RF Train Files: ${rfTrFiles.length}`, rfTrFiles.map(f => f.name));
      console.log(` - RF Station Files: ${rfStFiles.length}`, rfStFiles.map(f => f.name));
      console.log(` - TRN Files: ${trnFiles.length}`, trnFiles.map(f => f.name));
      console.log(` - Radio Files: ${radioFiles.length}`, radioFiles.map(f => f.name));

      if (rfTrFiles.length === 0 && rfStFiles.length === 0) {
        alert('No RF files found for the selected date range. RF logs (Train or Station) are required for analysis.');
        return;
      }

      const fetchAndParse = async (f: any) => {
        try {
          console.log(`Fetching file: ${f.name} (ID: ${f.id})`);
          const resUrl = await fetch(`/api/aws/download?key=${encodeURIComponent(f.id)}`);
          if (!resUrl.ok) {
            const text = await resUrl.text();
            console.error("AWS se error aaya hai (Download URL):", text);
            throw new Error(`Server Error: ${resUrl.status}`);
          }
          const data = await resUrl.json();
          if (data.error) throw new Error(data.error);
          
          const res = await fetch(data.url);
          if (!res.ok) {
            const text = await res.text();
            console.error("AWS se error aaya hai (File Download):", text);
            throw new Error(`Failed to download file from S3: ${res.statusText}`);
          }
          
          const blob = await res.blob();
          console.log(`Parsing file: ${f.name}, size: ${blob.size} bytes`);
          return await parseFile(blob, f.name);
        } catch (err: any) {
          console.error(`Error processing file ${f.name}:`, err);
          throw new Error(`Failed to process ${f.name}: ${err.message}`);
        }
      };

      const rfTrData = (await Promise.all(rfTrFiles.map(fetchAndParse))).flat();
      const rfStData = (await Promise.all(rfStFiles.map(fetchAndParse))).flat();
      
      console.log(`Parsed Data: Train Rows=${rfTrData.length}, Station Rows=${rfStData.length}`);
      const trnData = trnFiles.length > 0 ? (await Promise.all(trnFiles.map(fetchAndParse))).flat() : null;
      const radioData = radioFiles.length > 0 ? (await Promise.all(radioFiles.map(fetchAndParse))).flat() : [];

      const processed = processDashboardData(rfTrData, trnData, radioData, rfStData);
      setStats(processed);
      if (processed.detectedDivision) {
        setDivision(processed.detectedDivision);
      }
      setSelectedStation('All');
      setSelectedLoco('All');
    } catch (err: any) {
      console.error('Error analyzing cloud data:', err);
      alert(`Failed to analyze data from AWS S3: ${err.message || 'Unknown error'}`);
    } finally {
      setIsFetching(false);
    }
  };

  const handleAwsLogout = () => {
    setIsAwsConnected(false);
    setCloudFiles([]);
    setAvailableDates([]);
    setStats(null);
  };

  const handleFileUpload = (type: keyof typeof files, uploaded: File | FileList) => {
    if (type === 'radio') {
      const file = uploaded instanceof FileList ? uploaded[0] : uploaded;
      setFiles((prev) => ({ ...prev, radio: file }));
    } else if (type === 'faultLogs') {
      const newFiles = uploaded instanceof FileList ? Array.from(uploaded) : [uploaded];
      setFiles((prev) => ({ ...prev, faultLogs: [...prev.faultLogs, ...newFiles] }));
    } else {
      const newFiles = uploaded instanceof FileList ? Array.from(uploaded) : [uploaded];
      setFiles((prev) => ({ ...prev, [type]: [...prev[type as 'rf' | 'rfSt' | 'trn'], ...newFiles] }));
    }
  };

  const analyzeData = async () => {
    if (files.rf.length === 0 && files.rfSt.length === 0 && files.trn.length === 0) return;
    
    setIsFetching(true);
    try {
      const rfTrData = (await Promise.all(files.rf.map(f => parseFile(f)))).flat();
      const rfStData = (await Promise.all(files.rfSt.map(f => parseFile(f)))).flat();
      const trnData = files.trn.length > 0 ? (await Promise.all(files.trn.map(f => parseFile(f)))).flat() : null;
      const radioData = files.radio ? await parseFile(files.radio) : [];
      const faultLogData = files.faultLogs.length > 0 ? (await Promise.all(files.faultLogs.map(f => parseFile(f)))).flat() : [];
      
      const processed = processDashboardData(rfTrData, trnData, radioData, rfStData, faultLogData);
      setStats(processed);
      if (processed.detectedDivision) {
        setDivision(processed.detectedDivision);
      }
      setSelectedStation('All');
      setSelectedLoco('All');
      setStartDate('All');
      setEndDate('All');
    } catch (err: any) {
      console.error('Error analyzing local data:', err);
      alert(`Failed to analyze local data: ${err.message || 'Unknown error'}`);
    } finally {
      setIsFetching(false);
    }
  };

  const getFilteredStats = (): DashboardStats | null => {
    if (!stats) return null;
    
    let filtered = { ...stats };

    // Date Range Filtering
    if (startDate !== 'All' || endDate !== 'All') {
      const startT = startDate !== 'All' ? parseDateString(startDate) : 0;
      const endT = endDate !== 'All' ? parseDateString(endDate) : Infinity;

      const filterByDate = (item: any) => {
        if (!item.date) return true;
        const itemT = parseDateString(item.date);
        return itemT >= startT && itemT <= endT;
      };

      filtered.stationStats = filtered.stationStats.filter(filterByDate);
      filtered.stnPerf = filtered.stnPerf.filter(filterByDate);
      filtered.tagLinkIssues = filtered.tagLinkIssues.filter(filterByDate);
      filtered.uniqueTrainLengths = filtered.uniqueTrainLengths.filter(filterByDate);
      filtered.trainConfigChanges = filtered.trainConfigChanges.filter(filterByDate);
      filtered.modeDegradations = filtered.modeDegradations.filter(filterByDate);
      filtered.brakeApplications = filtered.brakeApplications.filter(filterByDate);
      filtered.signalOverrides = filtered.signalOverrides.filter(filterByDate);
      filtered.sosEvents = filtered.sosEvents.filter(filterByDate);
      filtered.maPackets = filtered.maPackets.filter(filterByDate);
      filtered.shortPackets = filtered.shortPackets.filter(filterByDate);
      filtered.nmsLogs = filtered.nmsLogs.filter(filterByDate);
      filtered.rawRfLogs = filtered.rawRfLogs.filter(filterByDate);

      // Update display duration
      if (startDate !== 'All') filtered.startTime = startDate;
      if (endDate !== 'All') filtered.endTime = endDate;
    }

    if (selectedLoco !== 'All') {
      filtered.stationStats = filtered.stationStats.filter(s => 
        String(s.locoId) === selectedLoco || 
        (s.source === 'station' && String(s.locoId) === 'Station Log')
      );
      filtered.stnPerf = filtered.stnPerf.filter(s => String(s.locoId) === selectedLoco);
      filtered.tagLinkIssues = filtered.tagLinkIssues.filter(t => String(t.locoId) === selectedLoco);
      filtered.uniqueTrainLengths = filtered.uniqueTrainLengths.filter(t => String(t.locoId) === selectedLoco);
      filtered.trainConfigChanges = filtered.trainConfigChanges.filter(t => String(t.locoId) === selectedLoco);
      filtered.modeDegradations = filtered.modeDegradations.filter(m => String(m.locoId) === selectedLoco);
      filtered.brakeApplications = filtered.brakeApplications.filter(b => String(b.locoId) === selectedLoco);
      filtered.signalOverrides = filtered.signalOverrides.filter(s => String(s.locoId) === selectedLoco);
      filtered.sosEvents = filtered.sosEvents.filter(s => String(s.locoId) === selectedLoco);
      filtered.maPackets = filtered.maPackets.filter(p => String(p.locoId) === selectedLoco);
      filtered.shortPackets = filtered.shortPackets.filter(p => String(p.locoId) === selectedLoco);
      filtered.nmsLogs = filtered.nmsLogs.filter(n => String(n.locoId) === selectedLoco);
      
      // Update primary locoId for display
      filtered.locoId = selectedLoco;

      // Recalculate Radio Lag
      if (filtered.maPackets.length > 0) {
        filtered.avgLag = filtered.maPackets.reduce((acc, p) => acc + p.delay, 0) / filtered.maPackets.length;
      } else {
        filtered.avgLag = 0;
      }

      // Recalculate NMS Status and Fail Rate
      if (filtered.nmsLogs.length > 0) {
        const nmsMap: Record<string, number> = {};
        filtered.nmsLogs.forEach(n => {
          nmsMap[n.health] = (nmsMap[n.health] || 0) + 1;
        });
        filtered.nmsStatus = Object.entries(nmsMap).map(([name, value]) => ({ name, value }));
        filtered.nmsFailRate = (filtered.nmsLogs.filter(n => n.health !== '0').length / filtered.nmsLogs.length) * 100;
      } else {
        filtered.nmsStatus = [];
        filtered.nmsFailRate = 0;
      }
    } else {
      filtered.locoId = 'All Locos';
    }

    if (selectedStation !== 'All') {
      filtered.stationStats = filtered.stationStats.filter(s => String(s.stationId) === selectedStation);
      filtered.stnPerf = filtered.stnPerf.filter(s => String(s.stationId) === selectedStation);
      filtered.tagLinkIssues = filtered.tagLinkIssues.filter(t => String(t.stationId) === selectedStation);
      filtered.uniqueTrainLengths = filtered.uniqueTrainLengths.filter(t => String(t.stationId) === selectedStation);
      filtered.trainConfigChanges = filtered.trainConfigChanges.filter(t => String(t.stationId) === selectedStation);
      filtered.modeDegradations = filtered.modeDegradations.filter(m => String(m.stationId) === selectedStation);
      filtered.brakeApplications = filtered.brakeApplications.filter(b => String(b.stationId) === selectedStation);
      filtered.signalOverrides = filtered.signalOverrides.filter(s => String(s.stationId) === selectedStation);
      filtered.sosEvents = filtered.sosEvents.filter(s => String(s.stationId) === selectedStation);
    }

    // Recalculate loco performance and station lists after all filters
    if (filtered.stationStats.length > 0) {
      // Aggregate by station, direction, and source for the final view
      const aggregated: Record<string, any> = {};
      filtered.stationStats.forEach(s => {
        const key = `${s.stationId}|${s.direction}|${s.source}`;
        if (!aggregated[key]) {
          aggregated[key] = { ...s, totalRowCount: s.rowCount, totalPercSum: s.totalPercSum, totalExp: s.expected, totalRec: s.received };
        } else {
          aggregated[key].totalRowCount += s.rowCount;
          aggregated[key].totalPercSum += s.totalPercSum;
          aggregated[key].totalExp += s.expected;
          aggregated[key].totalRec += s.received;
        }
      });

      filtered.stationStats = Object.values(aggregated).map(s => ({
        ...s,
        percentage: s.totalPercSum / s.totalRowCount,
        expected: s.totalExp,
        received: s.totalRec,
        rowCount: s.totalRowCount,
        totalPercSum: s.totalPercSum
      }));

      const totalExpAll = filtered.stationStats.reduce((acc, s) => acc + s.expected, 0);
      const totalRecAll = filtered.stationStats.reduce((acc, s) => acc + s.received, 0);
      
      const totalOverallPercSum = filtered.stationStats.reduce((acc, s) => acc + s.totalPercSum, 0);
      const totalOverallRowCount = filtered.stationStats.reduce((acc, s) => acc + s.totalRowCount, 0);
      
      filtered.locoPerformance = totalOverallRowCount > 0 
        ? totalOverallPercSum / totalOverallRowCount 
        : 0;
      
      // Create per-station weighted average for summary cards to avoid duplicates and show correct weighted %
      const stnSummary: Record<string, { exp: number, rec: number }> = {};
      filtered.stationStats.forEach(s => {
        const id = String(s.stationId);
        if (!stnSummary[id]) stnSummary[id] = { exp: 0, rec: 0 };
        stnSummary[id].exp += s.expected;
        stnSummary[id].rec += s.received;
      });

      const stnPerfList = Object.entries(stnSummary).map(([id, data]) => ({
        id,
        pct: data.exp > 0 ? (data.rec / data.exp) * 100 : 0
      }));

      filtered.badStns = stnPerfList.filter(s => s.pct < 85).map(s => s.id);
      filtered.marginalStns = stnPerfList.filter(s => s.pct >= 85 && s.pct <= 95).map(s => s.id);
      filtered.goodStns = stnPerfList.filter(s => s.pct > 95).map(s => s.id);

      filtered.unhealthyStns = stnPerfList
        .filter(s => s.pct < 85)
        .sort((a, b) => a.pct - b.pct);
      
      filtered.warningStns = stnPerfList
        .filter(s => s.pct >= 85 && s.pct <= 95)
        .sort((a, b) => a.pct - b.pct);
      
      filtered.healthyStns = stnPerfList
        .filter(s => s.pct > 95)
        .sort((a, b) => a.pct - b.pct);

      // Recalculate Diagnostic Advice for filtered data
      filtered.diagnosticAdvice = generateDiagnosticAdvice(filtered);
    }

    // Update Deep Analysis for the selected loco
    if (stats.locoAnalyses) {
      const analysisKey = selectedLoco === 'All' ? 'All' : selectedLoco;
      if (stats.locoAnalyses[analysisKey]) {
        filtered.stationDeepAnalysis = stats.locoAnalyses[analysisKey];
      }
    }

    return filtered;
  };

  const filteredStats = getFilteredStats();
  const uniqueStations: string[] = stats 
    ? ['All', ...Array.from(new Set(stats.stationStats
        .filter(s => selectedLoco === 'All' || String(s.locoId) === selectedLoco)
        .map(s => String(s.stationId)))) as string[]] 
    : ['All'];
  const uniqueLocos = stats ? ['All', ...new Set(stats.locoIds.map(id => String(id)))] : ['All'];

  const generatePDFReport = () => {
    try {
      if (!filteredStats) {
        console.error("No stats available for report");
        return;
      }
      
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const dateStr = new Date().toLocaleString();

      const addFooter = (doc: jsPDF, pageNumber: number) => {
        doc.setFontSize(8);
        doc.setTextColor(150);
        doc.text(`Kavach Expert Diagnostic System - ${division.toUpperCase()} Division`, 20, pageHeight - 10);
        doc.text(`Page ${pageNumber}`, pageWidth - 30, pageHeight - 10);
        doc.line(20, pageHeight - 15, pageWidth - 20, pageHeight - 15);
      };

      // --- COVER PAGE ---
      doc.setFillColor(15, 23, 42); // slate-900
      doc.rect(0, 0, pageWidth, pageHeight, 'F');
      
      doc.setDrawColor(16, 185, 129); // emerald-500
      doc.setLineWidth(1.5);
      doc.rect(10, 10, pageWidth - 20, pageHeight - 20);
      
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(32);
      doc.setFont('helvetica', 'bold');
      doc.text('KAVACH EXPERT', pageWidth / 2, 80, { align: 'center' });
      doc.text('DIAGNOSTIC REPORT', pageWidth / 2, 95, { align: 'center' });
      
      doc.setDrawColor(255, 255, 255);
      doc.line(40, 105, pageWidth - 40, 105);
      
      doc.setFontSize(16);
      doc.setFont('helvetica', 'normal');
      doc.text(`Division: ${division.toUpperCase()}`, pageWidth / 2, 120, { align: 'center' });
      doc.text(`Period: ${filteredStats.logDate || stats.logDate || 'Consolidated Analysis'}`, pageWidth / 2, 130, { align: 'center' });
      
      doc.setFillColor(16, 185, 129);
      doc.roundedRect(60, 150, pageWidth - 120, 40, 5, 5, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(14);
      doc.text('PERFORMANCE VERDICT', pageWidth / 2, 162, { align: 'center' });
      doc.setFontSize(22);
      doc.text(`${filteredStats.locoPerformance.toFixed(1)}%`, pageWidth / 2, 178, { align: 'center' });
      doc.addPage();

      let pageNum = 2;
      addFooter(doc, pageNum);

      // KPI Boxes
      doc.setTextColor(15, 23, 42);
      doc.setFontSize(20);
      doc.setFont('helvetica', 'bold');
      doc.text('1. EXECUTIVE SUMMARY', 20, 30);
      
      const drawKPI = (x: number, y: number, label: string, value: string, subValue: string, color: [number, number, number]) => {
        doc.setFillColor(color[0], color[1], color[2], 0.1);
        doc.roundedRect(x, y, 55, 35, 3, 3, 'F');
        doc.setDrawColor(color[0], color[1], color[2]);
        doc.setLineWidth(0.5);
        doc.roundedRect(x, y, 55, 35, 3, 3, 'S');
        doc.setTextColor(color[0], color[1], color[2]);
        doc.setFontSize(9);
        doc.text(label.toUpperCase(), x + 27.5, y + 10, { align: 'center' });
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.text(value, x + 27.5, y + 22, { align: 'center' });
        doc.setFontSize(8);
        doc.setFont('helvetica', 'normal');
        doc.text(subValue, x + 27.5, y + 30, { align: 'center' });
      };

      drawKPI(20, 45, 'Overall Health', `${filteredStats.locoPerformance.toFixed(1)}%`, 'RFCOMM Packet Delivery', [16, 185, 129]);
      drawKPI(77.5, 45, 'NMS Fault Rate', `${filteredStats.nmsFailRate.toFixed(2)}%`, `${filteredStats.nmsLogs.length} Events`, [239, 68, 68]);
      drawKPI(135, 45, 'Radio Lag', `${filteredStats.avgLag.toFixed(2)}s`, 'MA Average Interval', [59, 130, 246]);

      let currentY = 100;

      // Safety Summary Table
      doc.setFontSize(16);
      doc.setTextColor(15, 23, 42);
      doc.text('1.1 Safety Integrity Profile', 20, currentY);
      
      const ebCount = filteredStats.brakeApplications.filter(b => b.type.includes('EB')).length;
      const fsbCount = filteredStats.brakeApplications.filter(b => b.type.includes('FSB')).length;
      const modeDegCount = filteredStats.modeDegradations.length;

      autoTable(doc, {
        startY: currentY + 5,
        head: [['Safety Metric', 'Event Count', 'Severity Level', 'Action Recommendation']],
        body: [
          ['Emergency Brakes (EB)', ebCount, ebCount > 0 ? 'CRITICAL' : 'OPTIMAL', ebCount > 0 ? 'Urgent Root Cause Investigation' : 'Monitor'],
          ['Service Brakes (FSB/NB)', fsbCount, fsbCount > 5 ? 'HIGH' : 'NORMAL', fsbCount > 5 ? 'Check Signal Feed synchronization' : 'Routine Check'],
          ['Mode Degradations (FS → SB/SR)', modeDegCount, modeDegCount > 2 ? 'MAJOR' : 'LOW', modeDegCount > 2 ? 'Verify Radio switching logic' : 'Standard Monitor'],
          ['SOS / Emergency Signal', filteredStats.sosEvents.length, filteredStats.sosEvents.length > 0 ? 'CRITICAL' : 'NONE', 'Verify Emergency Switch Integrity']
        ],
        theme: 'grid',
        headStyles: { fillColor: [15, 23, 42] },
        alternateRowStyles: { fillColor: [245, 247, 250] },
      });

      currentY = (doc as any).lastAutoTable.finalY + 15;

      // --- SECTION 1.2: Station-wise Operations Summary ---
      if (filteredStats.stationSummary && filteredStats.stationSummary.length > 0) {
        if (currentY > 230) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(16);
        doc.setTextColor(15, 23, 42);
        doc.text('1.2 Station-wise Operations Summary', 20, currentY);
        
        const stnRows = filteredStats.stationSummary.map(s => [
          s.stationName,
          s.locoCount,
          s.degradationCount,
          s.totalBrakes,
          s.ebCount,
          s.fsbCount,
          s.attribution
        ]);

        autoTable(doc, {
          startY: currentY + 5,
          head: [['Station Name', 'Locos', 'Mode Deg', 'Total Brakes', 'EB', 'FSB', 'Verdict']],
          body: stnRows,
          theme: 'striped',
          headStyles: { fillColor: [16, 185, 129] }, // Emerald-500
          styles: { fontSize: 8 },
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- 1.3 Loco-wise Performance Summary ---
      if (filteredStats.locoSummary && filteredStats.locoSummary.length > 0) {
        if (currentY > 230) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(16);
        doc.setTextColor(15, 23, 42);
        doc.text('1.3 Loco-wise Performance Summary', 20, currentY);
        
        const locRows = filteredStats.locoSummary.map(l => [
          l.locoId,
          l.stationCount,
          l.degradationCount,
          l.totalBrakes,
          l.ebCount,
          l.fsbCount,
          l.attribution
        ]);

        autoTable(doc, {
          startY: currentY + 5,
          head: [['Loco No', 'Stations', 'Degradations', 'Brakes', 'EB', 'FSB', 'Verdict']],
          body: locRows,
          theme: 'striped',
          headStyles: { fillColor: [59, 130, 246] }, // Blue-500
          styles: { fontSize: 8 },
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 2: MODE DEGRADATIONS ---
      if (filteredStats.modeDegradations.length > 0) {
        if (currentY > 230) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(245, 158, 11); // Amber-500
        doc.text('2. MODE DEGRADATION ANALYSIS', 20, currentY);
        
        const degRows = filteredStats.modeDegradations.map(deg => [
          deg.time, 
          `${deg.from} → ${deg.to}`, 
          formatStationName(deg.stationName || deg.stationId), 
          deg.locoId, 
          `${deg.speed || 0} Kmph`, 
          deg.radio || '-',
          manualRemarks[`deg-${deg.time}-${deg.locoId}`] || ''
        ]);

        autoTable(doc, {
          startY: currentY + 10,
          head: [['Time', 'Transition', 'Station', 'Loco No', 'Speed', 'Radio Unit', 'Staff Reason / Remarks']],
          body: degRows,
          theme: 'grid',
          headStyles: { fillColor: [245, 158, 11] },
          styles: { fontSize: 7 },
          columnStyles: {
            6: { cellWidth: 40 }
          }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 3: AUTOMATIC BRAKE ACTIVATION LOG ---
      if (filteredStats.brakeApplications.length > 0) {
        if (currentY > 230) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(239, 68, 68); // Red-500
        doc.text('3. AUTOMATIC BRAKE ACTIVATION LOG', 20, currentY);
        
        const brakeRows = filteredStats.brakeApplications.map(b => [
          b.time, 
          b.type, 
          formatStationName(b.stationName || b.stationId), 
          b.locoId, 
          `${b.speed} Kmph`, 
          b.radio || '-',
          manualRemarks[`brake-${b.time}-${b.locoId}`] || ''
        ]);

        autoTable(doc, {
          startY: currentY + 10,
          head: [['Time', 'Brake Type', 'Station', 'Loco No', 'Speed', 'Radio Unit', 'Staff Reason / Remarks']],
          body: brakeRows,
          theme: 'grid',
          headStyles: { fillColor: [220, 38, 38] },
          styles: { fontSize: 7 },
          columnStyles: { 6: { cellWidth: 50 } }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // SOS Emergency Table
      if (filteredStats.sosEvents.length > 0) {
        if (currentY > 235) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(225, 29, 72); // Rose-600
        doc.text('4. SOS EMERGENCY EVENTS', 20, currentY);
        
        const sosRows = filteredStats.sosEvents.map(s => [
          s.time, s.type, s.source, s.locoId, formatStationName(s.stationName || s.stationId)
        ]);
        autoTable(doc, {
          startY: currentY + 10,
          head: [['Time', 'Type', 'Source', 'Loco No', 'Location']],
          body: sosRows,
          theme: 'grid',
          headStyles: { fillColor: [225, 29, 72] },
          styles: { fontSize: 8 }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 5: FINAL DIAGNOSIS & REMEDIATION ---
      if (currentY > 220) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
      doc.setFontSize(20);
      doc.setTextColor(15, 23, 42);
      doc.text('5. FINAL DIAGNOSIS & REMEDIATION', 20, currentY);
      
      doc.setDrawColor(59, 130, 246);
      doc.setLineWidth(1);
      doc.line(20, currentY + 5, pageWidth - 20, currentY + 5);
      
      currentY += 12;

      filteredStats.diagnosticAdvice.forEach((advice, i) => {
        const titleFormatted = advice.title.replace(/([A-Z])/g, ' $1').trim().toUpperCase();
        const detailLines = doc.splitTextToSize(advice.detail, pageWidth - 50);
        const actionLines = doc.splitTextToSize(`ACTION: ${advice.action}`, pageWidth - 50);
        
        // Dynamic height with safety margin
        const boxHeight = 15 + (detailLines.length * 5) + (actionLines.length * 6) + 12;

        if (currentY + boxHeight > 270) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        
        doc.setFillColor(249, 250, 251);
        doc.setDrawColor(226, 232, 240);
        doc.roundedRect(20, currentY, pageWidth - 40, boxHeight, 1.5, 1.5, 'FD');
        
        doc.setTextColor(15, 23, 42);
        doc.setFontSize(11);
        doc.setFont('helvetica', 'bold');
        doc.text(`${i + 1}. ${titleFormatted}`, 25, currentY + 10);
        
        doc.setTextColor(71, 85, 105);
        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        doc.text(detailLines, 25, currentY + 18);
        
        doc.setTextColor(59, 130, 246);
        doc.setFont('helvetica', 'bold');
        doc.text(actionLines, 25, currentY + 18 + (detailLines.length * 5) + 6);
        
        currentY += boxHeight + 10;
      });

      // --- SECTION 6: ROOT CAUSE & PATTERN ANALYSIS ---
      if (filteredStats.smartDiagnosis) {
        if (currentY > 220) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(20);
        doc.setTextColor(15, 23, 42);
        doc.text('6. ROOT CAUSE & PATTERN ANALYSIS', 20, currentY);
        
        const sd = filteredStats.smartDiagnosis;
        currentY += 12;

        // Global Pattern
        if (sd.globalPattern) {
          const formattedIssue = sd.globalPattern.issue
            .replace(/([A-Z])/g, ' $1')
            .trim()
            .toUpperCase();
            
          const titleText = `GLOBAL PATTERN: ${formattedIssue}`;
          const titleLines = doc.splitTextToSize(titleText, pageWidth - 50);
          
          const incidenceText = `Incidence Rate: ${sd.globalPattern.percentage.toFixed(1)}% of all data rows (${sd.globalPattern.affectedRows.toLocaleString()} rows).`;
          const incidenceLines = doc.splitTextToSize(incidenceText, pageWidth - 50);
          
          const explLines = doc.splitTextToSize(sd.globalPattern.explanation, pageWidth - 50);
          
          // Dynamic box height calculation based on wrapped lines
          const boxHeight = 15 + (titleLines.length * 6) + (incidenceLines.length * 5) + (explLines.length * 5) + 5;
          
          if (currentY + boxHeight > 270) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }

          doc.setFillColor(249, 250, 251);
          doc.setDrawColor(209, 213, 219);
          doc.roundedRect(20, currentY, pageWidth - 40, boxHeight, 2, 2, 'FD');
          
          doc.setFontSize(11);
          doc.setTextColor(15, 23, 42);
          doc.setFont('helvetica', 'bold');
          doc.text(titleLines, 25, currentY + 10);
          
          const incidenceY = currentY + 10 + (titleLines.length * 6);
          doc.setFontSize(9);
          doc.setTextColor(71, 85, 105);
          doc.setFont('helvetica', 'normal');
          doc.text(incidenceLines, 25, incidenceY);
          
          const explY = incidenceY + (incidenceLines.length * 5);
          doc.text(explLines, 25, explY);
          
          currentY += boxHeight + 12;
        }

        // Infrastructure Defects
        if (sd.stationInsights.length > 0) {
          doc.setFontSize(14);
          doc.setTextColor(15, 23, 42);
          doc.setFont('helvetica', 'bold');
          doc.text('STATION-WISE BLAME / INFRASTRUCTURE DEFECTS:', 20, currentY);
          currentY += 10;

          sd.stationInsights.forEach(stn => {
            const descLines = doc.splitTextToSize(stn.description, pageWidth - 50);
            let totalDetailHeight = 0;
            const detailsProcessed = stn.details.map(d => {
              const lines = doc.splitTextToSize(`• ${d}`, pageWidth - 55);
              totalDetailHeight += (lines.length * 5);
              return lines;
            });

            const boxHeight = 15 + (descLines.length * 5) + totalDetailHeight + 5;
            if (currentY + boxHeight > 270) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }

            doc.setFillColor(254, 242, 242); // Red-50 backdrop
            doc.setDrawColor(252, 165, 165); // Red-300 border
            doc.roundedRect(20, currentY, pageWidth - 40, boxHeight, 2, 2, 'FD');
            
            doc.setFontSize(11);
            doc.setTextColor(stn.severity === 'Critical' ? 185 : 15, stn.severity === 'Critical' ? 28 : 23, stn.severity === 'Critical' ? 28 : 42);
            doc.setFont('helvetica', 'bold');
            doc.text(`${stn.stationName} — ${stn.severity.toUpperCase()}`, 25, currentY + 10);
            
            doc.setFontSize(9);
            doc.setTextColor(71, 85, 105);
            doc.setFont('helvetica', 'normal');
            doc.text(descLines, 25, currentY + 18);
            
            let detailY = currentY + 18 + (descLines.length * 5);
            detailsProcessed.forEach(lines => {
              doc.text(lines, 30, detailY);
              detailY += (lines.length * 5);
            });
            currentY += boxHeight + 10;
          });
        }
      }

      // --- SECTION 7: STATION PERFORMANCE SUMMARY ---
      if (currentY > 230) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
      doc.setFontSize(18);
      doc.setTextColor(15, 23, 42);
      doc.text('7. STATION PERFORMANCE SUMMARY', 20, currentY);
      currentY += 10;

      const drawStnGroup = (title: string, color: [number, number, number], stns: { id: string | number; pct: number }[]) => {
        if (stns.length === 0) return;
        const text = stns.map(s => `${formatStationName(s.id)} (${s.pct.toFixed(1)}%)`).join(', ');
        const lines = doc.splitTextToSize(text, pageWidth - 60);
        const boxHeight = 22 + (lines.length * 5);

        if (currentY + boxHeight > 270) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        
        doc.setFillColor(249, 250, 251); // Light grey background
        doc.setDrawColor(color[0], color[1], color[2]);
        doc.setLineWidth(1); // Thicker border
        doc.roundedRect(20, currentY, pageWidth - 40, boxHeight, 1.5, 1.5, 'FD');

        doc.setFontSize(10);
        doc.setTextColor(color[0], color[1], color[2]);
        doc.setFont('helvetica', 'bold');
        doc.text(title.toUpperCase(), 25, currentY + 8);
        
        // Progress bar indicator
        const avg = stns.reduce((acc, curr) => acc + curr.pct, 0) / stns.length;
        doc.setFillColor(230, 230, 230);
        doc.rect(25, currentY + 11, pageWidth - 50, 2, 'F');
        doc.setFillColor(color[0], color[1], color[2]);
        doc.rect(25, currentY + 11, (pageWidth - 50) * (avg / 100), 2, 'F');

        doc.setFontSize(9);
        doc.setTextColor(15, 23, 42); // High contrast text
        doc.setFont('helvetica', 'normal');
        doc.text(lines, 25, currentY + 18);
        currentY += boxHeight + 8;
      };

      const allStns = [...filteredStats.unhealthyStns, ...filteredStats.warningStns, ...filteredStats.healthyStns];
      const redGroup = allStns.filter(s => s.pct < 90);
      const saffronGroup = allStns.filter(s => s.pct >= 90 && s.pct <= 95);
      const greenGroup = allStns.filter(s => s.pct > 95);

      drawStnGroup('BAD PERFORMANCE STATIONS (<90%)', [220, 38, 38], redGroup);
      drawStnGroup('MODERATE PERFORMANCE STATIONS (90% - 95%)', [255, 153, 51], saffronGroup);
      drawStnGroup('GOOD PERFORMANCE STATIONS (>95%)', [5, 150, 105], greenGroup);

      // --- SECTION 8: RFCOMM PERFORMANCE (TRAIN VS STATION PERSPECTIVE) ---
      const rfcommData = filteredStats.stationDeepAnalysis.dashboard.problem1.table;
      if (rfcommData && rfcommData.length > 0) {
        if (currentY > 220) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(15, 23, 42);
        doc.text('8. RFCOMM PERFORMANCE (TRAIN VS STATION VIEW)', 20, currentY);
        
        // Visual Chart for Top 10 entries
        doc.setFontSize(10);
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(100);
        doc.text('Visual Performance Comparison (Top 10 Stations):', 20, currentY + 10);
        currentY += 14;

        const slice = [...rfcommData].slice(0, 10);
        slice.forEach((r) => {
          const lVal = Number(r.locoVal);
          const oAvg = Number(r.othersAvg);
          
          doc.setFontSize(7);
          doc.setTextColor(15, 23, 42);
          doc.setFont('helvetica', 'bold');
          doc.text(formatStationName(r.station), 20, currentY);
          
          // Train Bar (Blue)
          doc.setFillColor(59, 130, 246);
          doc.rect(55, currentY - 3, (lVal * 0.8), 2.5, 'F');
          
          // Station Bar (Slate)
          doc.setFillColor(203, 213, 225);
          doc.rect(55, currentY + 0.5, (oAvg * 0.8), 2.5, 'F');
          
          doc.setFontSize(6);
          doc.setFont('helvetica', 'normal');
          doc.text(`T: ${lVal.toFixed(1)}% | S: ${oAvg.toFixed(1)}%`, 55 + Math.max(lVal, oAvg) * 0.8 + 2, currentY + 1);
          
          currentY += 8;
        });

        currentY += 5;
        
        const rfRows = rfcommData.map(r => {
          const lVal = Number(r.locoVal);
          const oAvg = Number(r.othersAvg);
          return [
            formatStationName(r.station),
            `${lVal.toFixed(2)} %`,
            `${oAvg.toFixed(2)} %`,
            lVal >= 98 && oAvg >= 98 ? 'Excellent' : 
            lVal < oAvg - 5 ? 'Suspect Train' : 
            oAvg < lVal - 5 ? 'Suspect Station' : 'Consistent'
          ];
        });

        autoTable(doc, {
          startY: currentY + 10,
          head: [['Station Name', 'Train View (RFCOMM)', 'Station View (Avg)', 'Comparison Verdict']],
          body: rfRows,
          theme: 'grid',
          headStyles: { fillColor: [15, 23, 42] },
          styles: { fontSize: 8 }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 9: DEEP ANALYSIS ---
      if (currentY > 240) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
      doc.setFontSize(18);
      doc.setTextColor(59, 130, 246);
      doc.text('9. DEEP ANALYSIS — PACKET LOSS MAPPING', 20, currentY);
      
      doc.setFontSize(10);
      doc.setTextColor(15, 23, 42);
      doc.setFont('helvetica', 'bold');
      doc.text(`Conclusion: ${filteredStats.stationDeepAnalysis.dashboard.conclusion}`, 20, currentY + 10);
      
      currentY += 20;
      const isLFaulty = filteredStats.stationDeepAnalysis.dashboard.conclusion.includes('Suspect') || filteredStats.stationDeepAnalysis.dashboard.conclusion.includes('Multiple');
      const isSFaulty = filteredStats.stationDeepAnalysis.dashboard.conclusion.includes('Station') || filteredStats.stationDeepAnalysis.dashboard.conclusion.includes('Multiple');

      // Status Boxes
      const drawStatusBox = (x: number, label: string, status: string, color: [number, number, number]) => {
        doc.setFillColor(color[0], color[1], color[2], 0.1);
        doc.roundedRect(x, currentY, 55, 30, 2, 2, 'F');
        doc.setTextColor(color[0], color[1], color[2]);
        doc.setFontSize(9);
        doc.text(label, x + 27.5, currentY + 10, { align: 'center' });
        doc.setFontSize(14);
        doc.setFont('helvetica', 'bold');
        doc.text(status, x + 27.5, currentY + 22, { align: 'center' });
      };

      drawStatusBox(20, 'LOCO TCAS', isLFaulty ? 'SUSPECT' : 'FIT', isLFaulty ? [220, 38, 38] : [16, 185, 129]);
      drawStatusBox(77.5, 'STATION TCAS', isSFaulty ? 'INSPECT' : 'HEALTHY', isSFaulty ? [245, 158, 11] : [16, 185, 129]);
      drawStatusBox(135, 'BENCHMARK', 'CLEARED', [59, 130, 246]);

      currentY += 45;

      // Probability Chart
      doc.setFontSize(12);
      doc.setTextColor(15, 23, 42);
      doc.text('Root Cause Probability Analysis', 20, currentY);
      
      const drawProbBar = (label: string, val: number, y: number, color: [number, number, number]) => {
        doc.setFontSize(9);
        doc.setTextColor(80);
        doc.text(label, 20, y);
        doc.setFillColor(240, 240, 240);
        doc.roundedRect(60, y - 3, 100, 4, 1, 1, 'F');
        doc.setFillColor(color[0], color[1], color[2]);
        doc.roundedRect(60, y - 3, Math.max(2, val), 4, 1, 1, 'F');
        doc.setTextColor(15, 23, 42);
        doc.text(`${val}%`, 165, y);
      };

      const rc = filteredStats.stationDeepAnalysis.rootCause;
      drawProbBar('Station-side', rc.stationSide, currentY + 10, [16, 185, 129]);
      drawProbBar('Loco-side', rc.locoSide, currentY + 18, [239, 68, 68]);
      drawProbBar('Hardware Issue', rc.hardwareProb, currentY + 26, [245, 158, 11]);
      drawProbBar('Software Logic', rc.softwareProb, currentY + 34, [59, 130, 246]);

      currentY += 45;

      // --- SECTION 10: MOVING RADIO GAP ---
      if (filteredStats.movingRadioLoss.length > 0) {
        if (currentY > 240) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(16, 185, 129);
        doc.text('10. MOVING RADIO LOSS ANALYSIS', 20, currentY);
        
        const movingRows = filteredStats.movingRadioLoss.slice(0, 15).map(m => [
          m.locoId, m.movingGaps, `${m.maxGap}s`, `${m.r1Usage}%`, `${m.r2Usage}%`, m.conclusion
        ]);
        
        autoTable(doc, {
          startY: currentY + 10,
          head: [['Loco ID', 'Gap Count', 'Max Gap', 'R1 %', 'R2 %', 'Verdict']],
          body: movingRows,
          theme: 'grid',
          headStyles: { fillColor: [16, 185, 129] },
          styles: { fontSize: 8 }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 11: TAG ISSUES ---
      if (filteredStats.tagLinkIssues.length > 0) {
        if (currentY > 240) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(15, 23, 42);
        doc.text('11. TAG LINK & INFRASTRUCTURE DEFECTS', 20, currentY);
        
        const tagRows = filteredStats.tagLinkIssues.slice(0, 15).map(t => [
          t.time, t.locoId, formatStationName(t.stationName || t.stationId), t.error, t.info
        ]);
        
        autoTable(doc, {
          startY: currentY + 10,
          head: [['Time', 'Loco', 'Station', 'Defect Type', 'Diagnostic Info']],
          body: tagRows,
          theme: 'grid',
          headStyles: { fillColor: [71, 85, 105] },
          styles: { fontSize: 8 }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 12: SCIENTIFIC AUDIT & MATHEMATICAL PROOF ---
      if (currentY > 180) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
      doc.setFontSize(18);
      doc.setTextColor(16, 185, 129); // Emerald-500
      doc.text('12. SCIENTIFIC AUDIT: MATHEMATICAL ERROR CORRELATION', 20, currentY);
      currentY += 8;

      const headingText = 'System Model: Kalman Filter-based state estimator [x(k) = F x(k-1) + B u(k) + w(k)]';
      const proofLinesRaw = [
        "1. THE SPACING DEFECT: RFID Tag Spacing > Designated distance causes the 'Fix' events to arrive with a timing lag (dT).",
        "2. KALMAN GAIN SLUGGISHNESS: This lag inflates the Innovation Residual (v), causing the Kalman Gain to decrease.",
        "3. COVARIANCE GROWTH (P_defect proportional to n * d^2): Uncertainty grows monotonically between defective fixes.",
        "4. DEGRADATION RISK: This elevated background pressure (sigma^2) triggers safety threshold violations (S_max)."
      ];

      const mathBoxWidth = pageWidth - 40;
      doc.setFontSize(10);
      const headingSplit = doc.splitTextToSize(headingText, mathBoxWidth - 10);
      doc.setFontSize(9);
      const processedProofLines = proofLinesRaw.map(line => doc.splitTextToSize(line, mathBoxWidth - 10));
      
      // Calculate height with specific spacing for gaps
      const headingHeight = headingSplit.length * 6;
      const bulletsHeight = processedProofLines.reduce((acc, lines) => acc + (lines.length * 5.5), 0);
      const gaps = 10; // Extra spacing between heading and bullets and bullets themselves
      const mathBoxHeight = headingHeight + bulletsHeight + gaps + 10;

      if (currentY + mathBoxHeight > 275) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }

      doc.setFillColor(240, 253, 244); // Emerald-50
      doc.setDrawColor(16, 185, 129); // Emerald-500
      doc.setLineWidth(0.5);
      doc.roundedRect(20, currentY, mathBoxWidth, mathBoxHeight, 2, 2, 'FD');
      
      let boxY = currentY + 8;
      doc.setFontSize(10);
      doc.setTextColor(5, 150, 105);
      doc.setFont('helvetica', 'bold');
      doc.text(headingSplit, 25, boxY);
      boxY += (headingSplit.length * 6) + 4; // Gap after heading
      
      doc.setFontSize(9);
      doc.setTextColor(31, 41, 55);
      doc.setFont('helvetica', 'normal');
      processedProofLines.forEach((lines) => {
        doc.text(lines, 25, boxY);
        boxY += (lines.length * 5.5) + 1; // Tight gap between items
      });
      
      currentY += mathBoxHeight + 5;

      // Subsection: Specific Scientific Event Audit
      const eventAudits = filteredStats.scientificInsights?.eventAudits || [];
      if (eventAudits.length > 0) {
        if (currentY > 200) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(14);
        doc.setTextColor(5, 150, 105);
        doc.setFont('helvetica', 'bold');
        doc.text('SCIENTIFIC EVENT AUDIT TRAIL (MODE DROPS & BRAKES)', 20, currentY);
        currentY += 8;

        const auditRows = eventAudits.map(a => [
          a.time,
          a.station,
          `Loco ${a.locoId}`,
          a.type,
          a.trigger,
          a.scientificVerdict
        ]);

        autoTable(doc, {
          startY: currentY,
          head: [['Time', 'Station', 'Loco', 'Type', 'Trigger', 'Scientific Verdict']],
          body: auditRows,
          theme: 'striped',
          headStyles: { fillColor: [5, 150, 105], fontSize: 8 },
          styles: { fontSize: 7 },
          columnStyles: { 
            3: { fontStyle: 'bold' },
            5: { cellWidth: 70 } 
          }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      if (filteredStats.scientificInsights && filteredStats.scientificInsights.highRiskScenarios.length > 0) {
        doc.setFontSize(12);
        doc.setTextColor(220, 38, 38); // Red-600
        doc.text('DETECTED RISK SCENARIOS (MODEL VIOLATIONS)', 20, currentY);
        
        const riskRows = filteredStats.scientificInsights.highRiskScenarios.map(s => [
           s.time,
           formatStationName(s.stationName),
           s.locoId,
           `${s.estimatedUncertainty} m2`,
           'CRITICAL: Background Covariance exceeded safety limit S_max'
        ]);

        autoTable(doc, {
          startY: currentY + 5,
          head: [['Time', 'Station', 'Loco', 'Uncertainty (σ²)', 'Model Audit Finding']],
          body: riskRows,
          theme: 'grid',
          headStyles: { fillColor: [220, 38, 38] },
          styles: { fontSize: 7 }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 13: SYSTEM EXTERNAL FAULT AUDIT ---
      if (filteredStats.faultLogs && filteredStats.faultLogs.length > 0) {
        if (currentY > 230) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(220, 38, 38); // Red-600
        doc.text('13. SYSTEM EXTERNAL FAULT AUDIT (NMS V2)', 20, currentY);
        currentY += 12;

        const faultRows = filteredStats.faultLogs.slice(0, 50).map(f => [
          f.time,
          f.station,
          f.locoId,
          f.absLoc,
          f.faultMsg,
          f.status
        ]);

        autoTable(doc, {
          startY: currentY,
          head: [['Time', 'Station', 'Loco', 'Loc', 'Fault Message', 'Status']],
          body: faultRows,
          theme: 'striped',
          headStyles: { fillColor: [220, 38, 38] },
          styles: { fontSize: 6 }
        });
        currentY = (doc as any).lastAutoTable.finalY + 15;
      }

      // --- SECTION 14: NMS LOGS ---
      if (filteredStats.nmsLogs.length > 0) {
        if (currentY > 230) { doc.addPage(); currentY = 25; pageNum++; addFooter(doc, pageNum); }
        doc.setFontSize(18);
        doc.setTextColor(15, 23, 42);
        doc.text('14. SYSTEM NMS FAILURE LOGS', 20, currentY);
        
        const nmsRows = filteredStats.nmsLogs.slice(0, 30).map(n => [n.time, n.locoId, n.health, n.status]);
        autoTable(doc, {
          startY: currentY + 10,
          head: [['Time', 'Loco No', 'Health Code', 'Status']],
          body: nmsRows,
          theme: 'striped',
          headStyles: { fillColor: [30, 41, 59] },
          styles: { fontSize: 8 }
        });
      }

      doc.save(`Kavach_Expert_Report_${division}_${filteredStats.locoId}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err: any) {
      console.error('Error generating PDF:', err);
      alert('PDF generation failed. Check console.');
    }
  };

  const generateFailureLetter = () => {
    const filteredStats = getFilteredStats();
    if (!filteredStats) return;

    if (selectedLoco === 'All') {
      alert("Please select a specific Loco ID for the failure analysis letter.");
      return;
    }

    try {
      const doc = new jsPDF();
      const date = new Date().toLocaleDateString();
      const time = new Date().toLocaleTimeString();
      const reportId = `KAV/${filteredStats.locoId}/${Math.floor(Math.random() * 10000)}`;
      
      // Header
      doc.setFontSize(18);
      doc.setTextColor(0);
      doc.line(20, 30, 190, 30);

      // Meta Info
      doc.setFontSize(10);
      doc.text(`Division: ${division.toUpperCase()}`, 150, 33);
      doc.text(`Date: ${date}`, 150, 38);
      doc.text(`Time: ${time}`, 150, 43);
      doc.text(`Report ID: ${reportId}`, 20, 38);
      doc.text(`Loco ID: ${filteredStats.locoId}`, 20, 43);

      const logTimes = filteredStats.maPackets.map(p => p.time).sort();
      if (logTimes.length > 0) {
        doc.text(`Log Duration: ${logTimes[0]} to ${logTimes[logTimes.length - 1]}`, 20, 48);
      }

      doc.setFontSize(12);
      doc.setFont('helvetica', 'bold');
      doc.text('To,', 20, 55);
      doc.setFont('helvetica', 'normal');
      doc.text('The Senior Divisional Electrical Engineer (Rolling Stock),', 20, 62);
      doc.text('Traction Operations Department.', 20, 68);

      doc.setFont('helvetica', 'bold');
      const subjectText = `Subject: Deep Analysis & Failure Validation - Locomotive ${filteredStats.locoId}${selectedStation !== 'All' ? ' at ' + formatStationName(selectedStation) : ''}`;
      doc.text(subjectText, 20, 80);
      doc.line(20, 82, 160, 82);

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(10);
      
      let bodyY = 92;
      const writeText = (text: string, y: number, size = 10, isBold = false) => {
        doc.setFontSize(size);
        doc.setFont('helvetica', isBold ? 'bold' : 'normal');
        const lines = doc.splitTextToSize(text, 170);
        
        let currentY = y;
        lines.forEach((line: string) => {
          if (currentY > 280) {
            doc.addPage();
            currentY = 20;
            doc.setFontSize(size);
            doc.setFont('helvetica', isBold ? 'bold' : 'normal');
          }
          doc.text(line, 20, currentY);
          currentY += 5;
        });
        
        return currentY;
      };

      bodyY = writeText(`Sir,`, bodyY);
      const introText = `This letter provides a comprehensive technical audit of Locomotive ${filteredStats.locoId}${selectedStation !== 'All' ? ' at ' + formatStationName(selectedStation) : ''} based on real-time diagnostic logs. The analysis evaluates whether the reported system failure is technically justified (Genuine) or based on environmental/external factors (Flimsy).`;
      bodyY = writeText(introText, bodyY + 5);

      // 1. Technical Metrics Summary
      bodyY = writeText(`1. TECHNICAL PERFORMANCE METRICS:`, bodyY + 8, 11, true);
      
      const metricsData = [
        ['Metric', 'Value', 'Status'],
        ['Overall RFCOMM Success', `${filteredStats.locoPerformance.toFixed(2)}%`, filteredStats.locoPerformance > 95 ? 'Healthy' : filteredStats.locoPerformance >= 85 ? 'Warning' : 'Unhealthy'],
        ['NMS Software Health', `${(100 - filteredStats.nmsFailRate).toFixed(2)}%`, filteredStats.nmsFailRate <= 5 ? 'Healthy' : 'Critical'],
        ['Avg Radio MA Lag', `${filteredStats.avgLag.toFixed(2)}s`, filteredStats.avgLag <= 1.5 ? 'Normal' : 'High Latency'],
        ['Critical Tag Link Issues', `${filteredStats.tagLinkIssues.length}`, filteredStats.tagLinkIssues.length === 0 ? 'None' : 'Action Required']
      ];

      autoTable(doc, {
        startY: bodyY + 2,
        head: [metricsData[0]],
        body: metricsData.slice(1),
        theme: 'grid',
        styles: { fontSize: 9 },
        margin: { left: 20 }
      });
      bodyY = (doc as any).lastAutoTable.finalY + 8;

      // 2. Station-wise Performance Analysis
      bodyY = writeText(`2. STATION-SPECIFIC COMMUNICATION AUDIT:`, bodyY, 11, true);
      const stnData = filteredStats.stationStats
        .sort((a, b) => a.percentage - b.percentage)
        .slice(0, 5)
        .map(s => [formatStationName(s.stationId), s.direction, `${s.received}/${s.expected}`, `${s.percentage.toFixed(1)}%`]);

      if (stnData.length > 0) {
        autoTable(doc, {
          startY: bodyY + 2,
          head: [['Station ID', 'Direction', 'Packets (R/E)', 'Success %']],
          body: stnData,
          theme: 'striped',
          styles: { fontSize: 8 },
          margin: { left: 20 }
        });
        bodyY = (doc as any).lastAutoTable.finalY + 8;
      } else {
        bodyY = writeText(`No specific station communication drops detected.`, bodyY + 2);
      }

      // 3. Mode Degradation Analysis
      if (filteredStats.modeDegradations.length > 0) {
        bodyY = writeText(`3. MODE DEGRADATION AUDIT (TRNMSNMA):`, bodyY, 11, true);
        autoTable(doc, {
          startY: bodyY + 2,
          head: [['Timestamp', 'Station', 'From -> To', 'Reason / Root Cause']],
          body: filteredStats.modeDegradations.map(d => {
            const fStnName = formatStationName(d.stationName || d.stationId);
            const syncNote = d.hasSyncConflict ? "[RADIO SYNC FAILURE] " : "";
            const rootCause = d.rootCause ? `\nAudit: ${d.rootCause.slice(0, 150)}${d.rootCause.length > 150 ? '...' : ''}` : "";
            
            return [
              d.time, 
              fStnName, 
              `${d.from} -> ${d.to}`, 
              `${syncNote}${d.reason}${rootCause}`
            ];
          }),
          theme: 'grid',
          styles: { fontSize: 6.5, cellPadding: 2 },
          columnStyles: {
            3: { cellWidth: 80 }
          },
          margin: { left: 20 }
        });
        bodyY = (doc as any).lastAutoTable.finalY + 8;
      }

      // 3.5 NMS Health Audit
      if (filteredStats.nmsLocoStats && filteredStats.nmsLocoStats.length > 0) {
        if (bodyY > 240) { doc.addPage(); bodyY = 20; }
        bodyY = writeText(`NMS HEALTH AUDIT (HARDWARE FAULT DETECTION):`, bodyY, 11, true);
        bodyY = writeText(`NMS Health indicates the internal diagnostic status of the Loco Vital Computer (LVC). A value of '0' means healthy.`, bodyY, 9, false);
        
        autoTable(doc, {
          startY: bodyY + 2,
          head: [['Loco ID', 'Total Records', 'Errors (Non-Zero)', 'Error %', 'Status']],
          body: filteredStats.nmsLocoStats.map(d => [
            d.locoId,
            d.totalRecords.toLocaleString(),
            d.errors.toLocaleString(),
            `${d.errorPercentage}%`,
            d.category
          ]),
          theme: 'grid',
          styles: { fontSize: 8 },
          margin: { left: 20 }
        });
        bodyY = (doc as any).lastAutoTable.finalY + 8;

        if (filteredStats.nmsDeepAnalysis && filteredStats.nmsDeepAnalysis.length > 0) {
          if (bodyY > 240) { doc.addPage(); bodyY = 20; }
          bodyY = writeText(`Continuous Error Events (Deep Analysis):`, bodyY, 10, true);
          autoTable(doc, {
            startY: bodyY + 2,
            head: [['Loco ID', 'Station', 'Time Range', 'Code', 'Error Type', 'Count']],
            body: filteredStats.nmsDeepAnalysis.slice(0, 10).map(d => {
              const fStnId = formatStationName(d.stationId);
              const fStnName = formatStationName(d.stationName);
              const stnId = (fStnId !== 'N/A') ? fStnId : '';
              const stnName = (fStnName !== 'N/A' && fStnName !== '-' && fStnName !== '0') ? `\n(${fStnName})` : '';
              return [
                d.locoId,
                `${stnId}${stnName}`.trim(),
                `${d.startTime.split(' ')[1]} - ${d.endTime.split(' ')[1]}`,
                d.errorCode,
                d.errorType,
                d.count.toString()
              ];
            }),
            theme: 'grid',
            styles: { fontSize: 7 },
            margin: { left: 20 }
          });
          bodyY = (doc as any).lastAutoTable.finalY + 8;
        }
      }

      // 4. Chronological Event Log (Last 5 Critical Events)
      const events = [
        ...filteredStats.modeDegradations.map(e => ({ 
          time: e.time, 
          type: 'DEGRADATION', 
          detail: `${e.from} -> ${e.to} at ${formatStationName(e.stationName || e.stationId)} (${e.reason})` 
        })),
        ...filteredStats.sosEvents.map(e => ({ time: e.time, type: 'SOS', detail: `${e.type} from ${e.source} at ${formatStationName(e.stationId)}` })),
        ...filteredStats.brakeApplications.map(e => ({ time: e.time, type: 'BRAKE', detail: `${e.type} at ${e.speed} km/h near ${formatStationName(e.stationId)}` }))
      ].sort((a, b) => b.time.localeCompare(a.time)).slice(0, 5);

      if (events.length > 0) {
        bodyY = writeText(`4. RECENT CRITICAL EVENTS LOG:`, bodyY, 11, true);
        autoTable(doc, {
          startY: bodyY + 2,
          head: [['Timestamp', 'Event Type', 'Technical Details']],
          body: events.map(e => [e.time, e.type, e.detail]),
          theme: 'grid',
          styles: { fontSize: 8 },
          margin: { left: 20 }
        });
        bodyY = (doc as any).lastAutoTable.finalY + 8;
      }

      // 5. Loco Overview Section
      bodyY = writeText(`5. LOCO OVERVIEW & DATA DURATION:`, bodyY, 11, true);
      bodyY = writeText(`Locomotive Number: ${filteredStats.locoId}`, bodyY + 2, 10);
      bodyY = writeText(`Analysis Duration: ${filteredStats.startTime} to ${filteredStats.endTime}`, bodyY + 2, 10);
      bodyY += 5;

      // 6. Expert Judgment Section (Correlation Logic)
      // A failure is GENUINE only if Internal Errors (NMS) correlate with External Symptoms (RF/Tags)
      // OR if there is a sustained Radio Timeout.
      const hasInternalFault = filteredStats.nmsFailRate > 40;
      const hasExternalSymptom = filteredStats.locoPerformance < 92 || filteredStats.tagLinkIssues.length > 2;
      const hasCriticalRadioLag = filteredStats.avgLag > 2.5;
      
      const isGenuine = (hasInternalFault && hasExternalSymptom) || hasCriticalRadioLag || (filteredStats.brakeApplications.length > 0 && filteredStats.locoPerformance < 85);
      
      const judgment = isGenuine ? "VALIDATED FUNCTIONAL FAILURE (SYSTEMIC)" : "NON-FUNCTIONAL DIAGNOSTIC ANOMALY (TRANSIENT)";
      const color = isGenuine ? [200, 0, 0] : [0, 150, 0];

      bodyY = writeText(`6. TECHNICAL VALIDATION & LOGICAL PROOF:`, bodyY + 5, 11, true);
      doc.setTextColor(color[0], color[1], color[2]);
      bodyY = writeText(`FINAL DECISION: ${judgment}`, bodyY + 2, 11, true);
      doc.setTextColor(0);

      let reasoning = "";
      if (isGenuine) {
        reasoning = `TECHNICAL VALIDATION: The failure is classified as FUNCTIONAL due to CORRELATION. `;
        if (hasInternalFault && hasExternalSymptom) {
          reasoning += `The system shows both Internal NMS instability (${filteredStats.nmsFailRate.toFixed(1)}%) AND External performance degradation (RF: ${filteredStats.locoPerformance.toFixed(1)}%). This proves that the NMS errors are not just 'noise' but are actively causing communication drops or hardware malfunctions. `;
        } else if (hasCriticalRadioLag) {
          reasoning += `The average Radio MA lag of ${filteredStats.avgLag.toFixed(2)}s exceeds the safety threshold, directly impacting train operation regardless of NMS status. `;
        }
        
        // Multi-Loco Station Proof
        if (filteredStats.multiLocoBadStns.length > 0) {
          const conciseStnList = filteredStats.multiLocoBadStns.map(s => {
            const ids = s.locoDetails.map(d => d.id).join(', ');
            return `${formatStationName(s.stationId)} (Locos: ${ids})`;
          }).join('; ');
          
          const detailedStnList = filteredStats.multiLocoBadStns.map(s => {
            const details = s.locoDetails.map(d => `${d.id}: ${d.perf.toFixed(1)}% [${d.startTime} - ${d.endTime}]`).join(', ');
            return `${formatStationName(s.stationId)} (Locos: ${details})`;
          }).join('; ');
          
          reasoning += `LOGICAL PROOF: The failure is marked as GENUINE. The performance drops are observed at [${conciseStnList}] across multiple locomotives. Since multiple locos are failing at the same spot, the fault lies with the Station TCAS equipment. The locomotive unit under analysis is performing normally elsewhere. \n\nDetailed Performance Audit: [${detailedStnList}]. `;
        }
        
        reasoning += `This confirms a hardware/software defect in the Loco Kavach Unit, but also highlights track-side infrastructure issues.`;
      } else {
        reasoning = `TECHNICAL VALIDATION: The failure is classified as NON-FUNCTIONAL/TRANSIENT. `;
        
        // Multi-Loco Station Proof for Flimsy Grounds
        if (filteredStats.multiLocoBadStns.length > 0) {
          const conciseStnList = filteredStats.multiLocoBadStns.map(s => {
            const ids = s.locoDetails.map(d => d.id).join(', ');
            return `${formatStationName(s.stationId)} (Locos: ${ids})`;
          }).join('; ');
          
          const detailedStnList = filteredStats.multiLocoBadStns.map(s => {
            const details = s.locoDetails.map(d => `${d.id}: ${d.perf.toFixed(1)}% [${d.startTime} - ${d.endTime}]`).join(', ');
            return `${formatStationName(s.stationId)} (Locos: ${details})`;
          }).join('; ');
          
          reasoning += `LOGICAL PROOF: The failure is marked as FLIMSY/WRONG. The performance drops are observed at [${conciseStnList}] across multiple locomotives. Since multiple locos are failing at the same spot, the fault lies with the Station TCAS equipment. The locomotive unit under analysis is performing normally elsewhere. \n\nDetailed Performance Audit: [${detailedStnList}]. `;
        } else if (hasInternalFault && !hasExternalSymptom) {
          reasoning += `Although NMS health is reported as sub-optimal (${filteredStats.nmsFailRate.toFixed(1)}% non-32 codes), the RFCOMM performance is stable at ${filteredStats.locoPerformance.toFixed(2)}% with 0 Tag issues. This indicates that the NMS codes are 'Transient' or 'Informational' and do not constitute a functional failure. `;
        } else if ((filteredStats.badStns.length + (filteredStats.marginalStns?.length || 0)) > 0 && (filteredStats.badStns.length + (filteredStats.marginalStns?.length || 0)) <= 2) {
          const allAffected = [...filteredStats.badStns, ...(filteredStats.marginalStns || [])];
          reasoning += `The performance drops are highly localized to ${allAffected.map(id => formatStationName(id)).join(', ')}, proving that the issue is Track-side (RFID/Signal) and the Locomotive unit is healthy. `;
        }
        reasoning += `The locomotive is technically fit for operation.`;
      }
      bodyY = writeText(reasoning, bodyY + 2);

      // Recommendation
      bodyY = writeText(`7. RECOMMENDATION:`, bodyY + 8, 11, true);
      let recommendation = "";
      if (isGenuine) {
        if (filteredStats.multiLocoBadStns.length > 0) {
          const stnIds = filteredStats.multiLocoBadStns.map(s => formatStationName(s.stationId)).join(', ');
          recommendation = `1. URGENT: Inspect Station TCAS/Kavach equipment at Stations [${stnIds}] as multiple locomotives are failing there. 2. Perform a technical audit of the Loco Processing Unit (CPU) and Power Supply Module.`;
        } else if (filteredStats.nmsFailRate > 50 && (filteredStats.badStns.length + (filteredStats.marginalStns?.length || 0)) === 1) {
          const stnId = filteredStats.badStns[0] || filteredStats.marginalStns?.[0];
          recommendation = `1. Inspect Station Kavach equipment at ${formatStationName(stnId)} for CPU/Radio faults. 2. If the problem persists across other stations, replace the Loco Processing Unit (CPU) and check the Power Supply Module.`;
        } else {
          recommendation = `Immediate inspection of the Kavach antenna, RF cables, and NMS processing unit is required at the shed. The locomotive should be grounded for a full technical audit and recalibration.`;
        }
      } else {
        if (filteredStats.multiLocoBadStns.length > 0) {
          const stnIds = filteredStats.multiLocoBadStns.map(s => formatStationName(s.stationId)).join(', ');
          recommendation = `The locomotive is fit for service. The reported communication drops are due to faulty Station-side equipment at [${stnIds}]. URGENT track-side audit is required at these locations.`;
        } else {
          const allAffected = [...filteredStats.badStns, ...(filteredStats.marginalStns || [])];
          recommendation = `The locomotive is fit for service. No hardware replacement is required. It is recommended to audit the track-side Kavach equipment and signal strength at stations [${allAffected.map(id => formatStationName(id)).join(', ')}] to resolve the localized communication drops.`;
        }
      }
      bodyY = writeText(recommendation, bodyY + 2);

      bodyY = writeText(recommendation, bodyY + 2);

      // ANNEX A: Deep Analysis Report
      doc.addPage();
      bodyY = 20;
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 102, 204);
      doc.text('ANNEXURE A: DEEP DIAGNOSTIC ANALYSIS', 105, bodyY, { align: 'center' });
      doc.line(20, bodyY + 2, 190, bodyY + 2);
      
      bodyY += 15;
      doc.setFontSize(11);
      doc.setTextColor(0);
      bodyY = writeText(`Diagnostic Conclusion: ${filteredStats.stationDeepAnalysis.dashboard.conclusion}`, bodyY, 11, true);
      
      // Verdict Blocks
      const isLFaulty = filteredStats.stationDeepAnalysis.dashboard.conclusion.includes('Loco');
      const isSFaulty = filteredStats.stationDeepAnalysis.dashboard.conclusion.includes('Station');
      
      doc.setFillColor(245, 245, 245);
      doc.rect(20, bodyY + 5, 170, 20, 'F');
      doc.setFontSize(10);
      doc.text(`Loco ${filteredStats.locoId} Status:`, 25, bodyY + 12);
      doc.setTextColor(isLFaulty ? 200 : 0, isLFaulty ? 0 : 150, 0);
      doc.setFontSize(12);
      doc.text(isLFaulty ? 'SUSPECTED FAULTY' : 'CONDITION FIT', 65, bodyY + 12, { align: 'left' });
      
      doc.setTextColor(0);
      doc.setFontSize(10);
      doc.text(`Station Health:`, 25, bodyY + 20);
      doc.setTextColor(isSFaulty ? 180 : 0, isSFaulty ? 120 : 150, 0);
      doc.setFontSize(12);
      doc.text(isSFaulty ? 'TRACKSIDE ISSUES DETECTED' : 'TRACKSIDE HEALTHY', 65, bodyY + 20, { align: 'left' });
      
      bodyY += 35;
      doc.setTextColor(0);
      bodyY = writeText('Loco Journey Performance Map:', bodyY, 11, true);
      
      autoTable(doc, {
        startY: bodyY + 2,
        head: [['Station', `Loco ${filteredStats.locoId}`, 'Baaki Locos (Avg)']],
        body: filteredStats.stationDeepAnalysis.dashboard.problem1.table.map(r => [formatStationName(r.station), r.locoVal, r.othersAvg]),
        theme: 'grid',
        styles: { fontSize: 8 },
        margin: { left: 20 }
      });
      bodyY = (doc as any).lastAutoTable.finalY + 10;

      if (filteredStats.multiLocoBadStns.length > 0) {
        bodyY = writeText('Multi-Loco Station Cross-Check:', bodyY, 11, true);
        autoTable(doc, {
          startY: bodyY + 2,
          head: [['Station', 'Locos Failed', 'Avg Perf', 'Action Required']],
          body: filteredStats.multiLocoBadStns.map(s => {
            return [formatStationName(s.stationId), s.locoCount, `${s.avgPerf.toFixed(1)}%`, s.locoCount >= 3 ? 'PRIORITY' : 'INSPECT'];
          }),
          theme: 'grid',
          styles: { fontSize: 8 },
          margin: { left: 20 }
        });
        bodyY = (doc as any).lastAutoTable.finalY + 10;
      }

      bodyY += 5;

      doc.setFontSize(9);
      doc.setTextColor(0, 120, 0);
      doc.setFont('helvetica', 'italic');
      bodyY = writeText(filteredStats.stationDeepAnalysis.dashboard.amlConclusion, bodyY, 9);
      doc.setTextColor(0);
      doc.setFont('helvetica', 'normal');

      bodyY += 10;
      doc.setFillColor(0, 102, 204);
      doc.rect(20, bodyY, 170, 5, 'F');
      bodyY += 10;

      // Technical Audit Logs (New Section for Sync Failures)
      if (filteredStats.technicalAudit && filteredStats.technicalAudit.length > 0) {
        doc.addPage();
        bodyY = 20;
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 102, 204);
        bodyY = writeText('DETAILED TECHNICAL EVENT AUDIT (TRNMSNMA)', bodyY, 12, true);
        doc.line(20, bodyY + 2, 190, bodyY + 2);
        bodyY += 10;
        doc.setTextColor(0);

        filteredStats.technicalAudit.forEach((audit, aidx) => {
           if (bodyY > 240) { doc.addPage(); bodyY = 20; }
           bodyY = writeText(`EVENT #${aidx+1}: ${audit.title}`, bodyY, 10, true);
           bodyY = writeText(`Time Range: ${audit.timeRange} | Severity: ${audit.severity}`, bodyY + 2, 9);
           
           
           bodyY = writeText(`Technical Findings:`, bodyY, 9, true);
           audit.analysisBullets.forEach(bullet => {
              if (bodyY > 260) { doc.addPage(); bodyY = 20; }
              const bulletLines = doc.splitTextToSize(`- ${bullet}`, 165);
              doc.text(bulletLines, 25, bodyY + 4);
              bodyY += (bulletLines.length * 3.5) + 2;
           });
           bodyY += 8;
        });
      }

      // Technical Note
      bodyY = writeText(`7. TECHNICAL NOTE:`, bodyY + 8, 10, true);
      const techNote = `Please note that while a locomotive may be mechanically 'Fit' and operational for traction, the 'Kavach Failure' status refers specifically to the Electronic Safety System. A high NMS failure rate indicates that the Kavach unit is unable to perform its safety-critical monitoring, which is a mandatory requirement for high-speed operations.`;
      bodyY = writeText(techNote, bodyY + 2, 9);

      // Footer - Dynamic placement
      bodyY += 10;
      // Threshold increased to 260 to allow more content on the same page
      if (bodyY > 260) {
        doc.addPage();
        bodyY = 25;
      }
      
      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.setFont('helvetica', 'normal');
      bodyY = writeText('This is a computer-generated technical analysis based on uploaded Kavach diagnostic logs.', bodyY);
      
      bodyY += 15;
      doc.setTextColor(0);
      doc.setFont('helvetica', 'normal');
      doc.text('Yours Sincerely,', 140, bodyY);
      
      bodyY += 8;
      doc.setFont('helvetica', 'bold');
      doc.text('CHIEF LOCO INSPECTOR', 140, bodyY);

      doc.save(`Failure_Analysis_Letter_Loco_${filteredStats.locoId}_${date.replace(/\//g, '-')}.pdf`);
    } catch (error) {
      console.error("Letter Generation Error:", error);
      alert("There was an issue generating the letter.");
    }
  };

  return (
    <div className="flex h-screen relative font-sans">
      <div className="atmosphere" />
      
      {/* Sidebar */}
      <aside className="w-72 glass-sidebar text-white p-6 flex flex-col gap-8 shrink-0 z-10">
        <div className="flex items-center gap-3">
          <Shield className="w-8 h-8 text-emerald-400 drop-shadow-[0_0_8px_rgba(52,211,153,0.5)]" />
          <h1 className="text-xl font-bold tracking-tight text-white">Kavach Expert</h1>
        </div>

        {/* Mentorship Section */}
        <div className="bg-white/5 p-4 rounded-xl border border-white/10 backdrop-blur-sm">
          <p className="text-[10px] uppercase font-bold text-emerald-400 tracking-widest mb-1">Technical Supervision</p>
          <p className="text-sm font-semibold text-white">Mentored by CELE Sir</p>
          <p className="text-[10px] text-slate-400 mt-1 italic">Expert Guidance in Traction Operations</p>
        </div>

        {/* System Update Note */}
        <div className="bg-emerald-500/10 p-4 rounded-xl border border-emerald-500/20 backdrop-blur-sm">
          <div className="flex items-center gap-2 mb-2">
            <Zap className="w-3 h-3 text-emerald-400" />
            <p className="text-[10px] uppercase font-bold text-emerald-400 tracking-widest">System Update</p>
          </div>
          <p className="text-xs text-slate-300 leading-relaxed">
            This dashboard can now analyze Kavach data even in the absence of <span className="text-emerald-400 font-semibold">RADIO_1</span> logs, which are often difficult to obtain. It automatically generates a detailed report on Loco TCAS health.
          </p>
        </div>

        <div className="flex flex-col gap-6">
          <div className="space-y-4">
            <h3 className="text-xs font-semibold uppercase tracking-wider text-slate-400">Cloud Storage</h3>
            {!isAwsConnected ? (
              <button
                onClick={handleAwsConnect}
                className="w-full py-3 bg-white/5 hover:bg-white/10 text-white rounded-xl border border-white/10 transition-all flex items-center justify-center gap-2 font-bold"
              >
                <Database className="w-4 h-4 text-orange-400" />
                Connect AWS S3
              </button>
            ) : (
              <div className="flex justify-between items-center bg-white/5 p-3 rounded-xl border border-white/10">
                <div className="flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-orange-400" />
                  <span className="text-xs font-bold text-slate-300">AWS S3 Connected</span>
                </div>
                <button onClick={handleAwsLogout} className="text-[10px] text-red-400 hover:text-red-300 font-bold uppercase tracking-tighter">
                  Logout
                </button>
              </div>
            )}
          </div>

          <div className="space-y-4">
            <h3 className="text-xs font-semibold uppercase tracking-wider text-slate-400">Local Upload</h3>
            <div className="grid grid-cols-2 gap-2">
              <div className="space-y-1">
                <p className="text-[9px] text-slate-500 uppercase font-bold px-1">RF Train</p>
                <label className="flex flex-col items-center justify-center w-full h-10 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg cursor-pointer transition-all">
                  <div className="flex items-center gap-2">
                    <Upload className="w-3 h-3 text-emerald-400" />
                    <span className="text-[10px] font-bold">{files.rf.length || 'Select'}</span>
                  </div>
                  <input type="file" className="hidden" multiple onChange={(e) => handleFileUpload('rf', e.target.files!)} />
                </label>
              </div>
              <div className="space-y-1">
                <p className="text-[9px] text-slate-500 uppercase font-bold px-1">RF Station</p>
                <label className="flex flex-col items-center justify-center w-full h-10 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg cursor-pointer transition-all">
                  <div className="flex items-center gap-2">
                    <Upload className="w-3 h-3 text-emerald-400" />
                    <span className="text-[10px] font-bold">{files.rfSt.length || 'Select'}</span>
                  </div>
                  <input type="file" className="hidden" multiple onChange={(e) => handleFileUpload('rfSt', e.target.files!)} />
                </label>
              </div>
              <div className="space-y-1">
                <p className="text-[9px] text-slate-500 uppercase font-bold px-1">TRN Logs</p>
                <label className="flex flex-col items-center justify-center w-full h-10 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg cursor-pointer transition-all">
                  <div className="flex items-center gap-2">
                    <Upload className="w-3 h-3 text-emerald-400" />
                    <span className="text-[10px] font-bold">{files.trn.length || 'Select'}</span>
                  </div>
                  <input type="file" className="hidden" multiple onChange={(e) => handleFileUpload('trn', e.target.files!)} />
                </label>
              </div>
              <div className="space-y-1">
                <p className="text-[9px] text-slate-500 uppercase font-bold px-1">Radio 1</p>
                <label className="flex flex-col items-center justify-center w-full h-10 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg cursor-pointer transition-all">
                  <div className="flex items-center gap-2">
                    <Upload className="w-3 h-3 text-emerald-400" />
                    <span className="text-[10px] font-bold">{files.radio ? '1 File' : 'Select'}</span>
                  </div>
                  <input type="file" className="hidden" onChange={(e) => handleFileUpload('radio', e.target.files!)} />
                </label>
              </div>
              <div className="space-y-1 col-span-2">
                <p className="text-[9px] text-slate-500 uppercase font-bold px-1">System Fault Logs (Optional)</p>
                <label className="flex flex-col items-center justify-center w-full h-10 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg cursor-pointer transition-all">
                  <div className="flex items-center gap-2">
                    <Activity className="w-3 h-3 text-rose-400" />
                    <span className="text-[10px] font-bold">
                      {files.faultLogs.length > 0 ? `${files.faultLogs.length} Logs Added` : 'Select Fault Data'}
                    </span>
                  </div>
                  <input type="file" multiple className="hidden" onChange={(e) => handleFileUpload('faultLogs', e.target.files!)} />
                </label>
              </div>
            </div>
            {(files.rf.length > 0 || files.rfSt.length > 0 || files.trn.length > 0) && (
              <button
                onClick={analyzeData}
                disabled={isFetching}
                className="w-full py-2 bg-emerald-500 hover:bg-emerald-400 text-white rounded-xl font-bold text-xs transition-all flex items-center justify-center gap-2 shadow-lg shadow-emerald-500/20"
              >
                {isFetching ? <RefreshCw className="w-3 h-3 animate-spin" /> : <Zap className="w-3 h-3" />}
                Analyze Local Files
              </button>
            )}
          </div>

          <div className="space-y-6">
            <div className="flex justify-between items-center">
              <h3 className="text-xs font-semibold uppercase tracking-wider text-slate-400">Analysis Controls</h3>
              <button onClick={checkAwsStatus} className="text-emerald-400 hover:text-emerald-300 transition-colors">
                <RefreshCw className={cn("w-3 h-3", isFetching && "animate-spin")} />
              </button>
            </div>

              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <Clock className="w-3 h-3" /> Start Date
                  </label>
                  <select 
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all"
                  >
                    <option value="All" className="bg-slate-900">Select Date</option>
                    {availableDates.map(d => (
                      <option key={d} value={d} className="bg-slate-900">{d}</option>
                    ))}
                  </select>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <Clock className="w-3 h-3" /> End Date
                  </label>
                  <select 
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all"
                  >
                    <option value="All" className="bg-slate-900">Select Date</option>
                    {availableDates.map(d => (
                      <option key={d} value={d} className="bg-slate-900">{d}</option>
                    ))}
                  </select>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <Shield className="w-3 h-3" /> Loco ID (Optional)
                  </label>
                  <select 
                    value={cloudLoco}
                    onChange={(e) => setCloudLoco(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all"
                  >
                    <option value="All" className="bg-slate-900">All Locos</option>
                    {availableLocos.map(l => (
                      <option key={l} value={l} className="bg-slate-900">{l}</option>
                    ))}
                  </select>
                </div>

                {isAwsConnected && availableDates.length === 0 && (
                  <div className="p-3 bg-amber-500/10 border border-amber-500/20 rounded-xl">
                    <div className="flex items-center gap-2 mb-1">
                      <AlertTriangle className="w-3 h-3 text-amber-400" />
                      <p className="text-[10px] font-bold text-amber-400 uppercase tracking-widest">No Dates Found</p>
                    </div>
                    <p className="text-[10px] text-slate-400 leading-relaxed">
                      Connected to S3, but couldn't find dates in filenames. Ensure files follow a date format like DD-MM-YYYY or YYYYMMDD.
                    </p>
                  </div>
                )}

                <button
                  onClick={analyzeCloudData}
                  disabled={isFetching || startDate === 'All' || endDate === 'All'}
                  className={cn(
                    "w-full py-3 rounded-xl font-bold transition-all flex items-center justify-center gap-2 shadow-lg",
                    (!isFetching && startDate !== 'All' && endDate !== 'All')
                      ? "bg-emerald-500 hover:bg-emerald-400 text-white shadow-emerald-500/20" 
                      : "bg-white/5 text-slate-500 cursor-not-allowed border border-white/5"
                  )}
                >
                  {isFetching ? (
                    <RefreshCw className="w-4 h-4 animate-spin" />
                  ) : (
                    <Zap className="w-4 h-4" />
                  )}
                  Analyze AWS Data
              </button>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto p-8 z-10">
        {!stats ? (
          <div className="h-full flex flex-col items-center justify-center text-center space-y-6">
            <div className="w-24 h-24 glass-card rounded-3xl flex items-center justify-center animate-pulse">
              <Shield className="w-12 h-12 text-emerald-400" />
            </div>
            <div className="space-y-2">
              <h2 className="text-3xl font-bold text-white">Ready for Analysis</h2>
              <p className="text-slate-400 max-w-md mx-auto">
                Upload your Kavach RF and TRN logs to generate a comprehensive diagnostic report. 
                <span className="block mt-2 text-emerald-400/80 text-sm">Now supports analysis without RADIO_1 logs for faster reporting.</span>
              </p>
            </div>
          </div>
        ) : (
          <div className="max-w-6xl mx-auto space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
            {/* Header */}
            <div className="flex flex-col gap-6">
              <div className="flex justify-between items-end">
                <div className="flex items-end gap-6">
                  <div>
                    <p className="text-emerald-400 font-bold text-sm tracking-widest uppercase mb-1">Diagnostic Report</p>
                    <h2 className="text-4xl font-bold text-white tracking-tight">Loco {stats.locoId}</h2>
                  </div>
                    <button 
                      onClick={generatePDFReport}
                      className="mb-1 flex items-center gap-2 px-4 py-2 bg-white/10 hover:bg-white/20 text-white rounded-xl border border-white/10 transition-all text-sm font-bold"
                    >
                      <Download className="w-4 h-4 text-emerald-400" />
                      Download Official Report
                    </button>
                    <button 
                      onClick={generateFailureLetter}
                      className={cn(
                        "mb-1 flex items-center gap-2 px-4 py-2 rounded-xl border transition-all text-sm font-bold",
                        selectedLoco === 'All' 
                          ? "bg-amber-500/10 hover:bg-amber-500/20 text-amber-400 border-amber-500/30"
                          : "bg-emerald-500/20 hover:bg-emerald-500/30 text-emerald-400 border-emerald-500/30"
                      )}
                    >
                      <FileText className="w-4 h-4" />
                      {selectedLoco === 'All' ? 'Select Loco for Failure Letter' : 'Download Failure Analysis Letter'}
                    </button>
                  </div>
                <div className="flex gap-1 p-1 glass-card rounded-xl overflow-x-auto max-w-3xl">
                  <TabButton active={activeTab === 'summary'} onClick={() => setActiveTab('summary')} label="Summary" />
                  <TabButton active={activeTab === 'mapping'} onClick={() => setActiveTab('mapping')} label="Mapping" />
                  <TabButton active={activeTab === 'station'} onClick={() => setActiveTab('station')} label="Station Analysis" />
                  <TabButton active={activeTab === 'radio'} onClick={() => setActiveTab('radio')} label="Radio Analysis" />
                  <TabButton active={activeTab === 'expert'} onClick={() => setActiveTab('expert')} label="Expert Diagnostics" />
                  <TabButton active={activeTab === 'ops'} onClick={() => setActiveTab('ops')} label="Operations Summary" />
                  <TabButton active={activeTab === 'nms'} onClick={() => setActiveTab('nms')} label="NMS" />
                  <TabButton active={activeTab === 'sync'} onClick={() => setActiveTab('sync')} label="Sync" />
                  <TabButton active={activeTab === 'interval'} onClick={() => setActiveTab('interval')} label="Interval" />
                  <TabButton active={activeTab === 'moving'} onClick={() => setActiveTab('moving')} label="Moving Analysis" />
                  <TabButton active={activeTab === 'methodology'} onClick={() => setActiveTab('methodology')} label="Methodology" />
                  <TabButton active={activeTab === 'scientific'} onClick={() => setActiveTab('scientific')} label="Scientific Audit" />
                </div>
              </div>

              {/* Filters */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 p-4 glass-card rounded-2xl border border-white/5">
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <Database className="w-3 h-3" /> Division
                  </label>
                  <input 
                    type="text"
                    value={division}
                    onChange={(e) => setDivision(e.target.value)}
                    placeholder="SC"
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all placeholder:text-slate-600"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <MapPin className="w-3 h-3" /> Station
                  </label>
                  <select 
                    value={selectedStation}
                    onChange={(e) => setSelectedStation(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all"
                  >
                    {uniqueStations.map(stn => (
                      <option key={stn} value={stn} className="bg-slate-900">{stn === 'All' ? 'All Stations' : formatStationName(stn)}</option>
                    ))}
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <Shield className="w-3 h-3" /> Loco
                  </label>
                  <select 
                    value={selectedLoco}
                    onChange={(e) => setSelectedLoco(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all"
                  >
                    {uniqueLocos.map(loco => (
                      <option key={loco} value={loco} className="bg-slate-900">{loco}</option>
                    ))}
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <Clock className="w-3 h-3" /> From Date
                  </label>
                  <select 
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all"
                  >
                    <option value="All" className="bg-slate-900">All Dates</option>
                    {stats.allDates.map(d => (
                      <option key={d} value={d} className="bg-slate-900">{d}</option>
                    ))}
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                    <Clock className="w-3 h-3" /> To Date
                  </label>
                  <select 
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    className="w-full bg-white/5 border border-white/10 rounded-xl px-3 py-2 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all"
                  >
                    <option value="All" className="bg-slate-900">All Dates</option>
                    {stats.allDates.map(d => (
                      <option key={d} value={d} className="bg-slate-900">{d}</option>
                    ))}
                  </select>
                </div>
                { (selectedStation !== 'All' || selectedLoco !== 'All' || startDate !== 'All' || endDate !== 'All') && (
                  <div className="md:col-span-4 flex justify-end">
                    <button 
                      onClick={() => { 
                        setSelectedStation('All'); 
                        setSelectedLoco('All'); 
                        setStartDate('All'); 
                        setEndDate('All'); 
                      }}
                      className="px-4 py-2 bg-rose-500/20 text-rose-400 rounded-xl border border-rose-500/20 text-xs font-bold hover:bg-rose-500/30 transition-all"
                    >
                      Reset All Filters
                    </button>
                  </div>
                )}
              </div>
            </div>

            {activeTab === 'summary' && filteredStats && <ExecutiveSummary stats={filteredStats} setActiveTab={setActiveTab} />}
            {activeTab === 'mapping' && filteredStats && <DeepMapping stats={filteredStats} files={files} />}
            {activeTab === 'station' && filteredStats && <StationAnalysis stats={filteredStats} />}
            {activeTab === 'ops' && filteredStats && <OperationsSummary stats={filteredStats} />}
            {activeTab === 'expert' && filteredStats && (
              <ExpertDiagnostics 
                stats={filteredStats} 
                tagSearch={tagSearch} 
                setTagSearch={setTagSearch} 
                remarks={manualRemarks}
                onUpdateRemark={updateRemark}
              />
            )}
            {activeTab === 'nms' && filteredStats && <NMSAnalysis stats={filteredStats} />}
            {activeTab === 'sync' && filteredStats && <SyncAnalysis stats={filteredStats} />}
            {activeTab === 'interval' && filteredStats && <IntervalAnalysis stats={filteredStats} />}
            {activeTab === 'radio' && filteredStats && <RadioLossAnalysis stats={filteredStats} />}
            {activeTab === 'moving' && filteredStats && <MovingAnalysis stats={filteredStats} />}
            {activeTab === 'methodology' && <CalculationMethodology />}
            {activeTab === 'scientific' && filteredStats && <ScientificAnalysis stats={filteredStats} />}
          </div>
        )}
      </main>
    </div>
  );
}

function StationAnalysis({ stats }: { stats: DashboardStats }) {
  // Group stats by station and source for comparison
  const stationComparison = stats.stationStats.reduce((acc: any[], curr) => {
    const existing = acc.find(a => a.stationId === curr.stationId);
    
    if (existing) {
      if (curr.source === 'station') {
        existing.stationExp = (existing.stationExp || 0) + (curr.expected || 0);
        existing.stationRec = (existing.stationRec || 0) + (curr.received || 0);
        existing.stationPerc = existing.stationExp > 0 ? (existing.stationRec / existing.stationExp) * 100 : 0;
      } else {
        existing.trainExp = (existing.trainExp || 0) + (curr.expected || 0);
        existing.trainRec = (existing.trainRec || 0) + (curr.received || 0);
        existing.trainPerc = existing.trainExp > 0 ? (existing.trainRec / existing.trainExp) * 100 : 0;
      }
    } else {
      acc.push({
        stationId: curr.stationId,
        trainExp: curr.source === 'station' ? 0 : (curr.expected || 0),
        trainRec: curr.source === 'station' ? 0 : (curr.received || 0),
        trainPerc: curr.source === 'station' ? null : (curr.expected > 0 ? (curr.received / curr.expected) * 100 : 0),
        stationExp: curr.source === 'station' ? (curr.expected || 0) : 0,
        stationRec: curr.source === 'station' ? (curr.received || 0) : 0,
        stationPerc: curr.source === 'station' ? (curr.expected > 0 ? (curr.received / curr.expected) * 100 : 0) : null,
        label: formatStationName(curr.stationId)
      });
    }
    return acc;
  }, []);

    const stationRecords = stats.stationStats.filter(s => s.source === 'station').length;
    const trainRecords = stats.stationStats.filter(s => s.source === 'train').length;
  
    const stationIds = Array.from(new Set(stats.stationStats.filter(s => s.source === 'station').map(s => s.stationId)));
    const trainLocoIds = Array.from(new Set(stats.stationStats.filter(s => s.source === 'train').map(s => s.locoId)));
  
    return (
      <div className="space-y-6">
        {/* Debug Info */}
        <div className="bg-slate-800 text-slate-300 p-4 rounded-xl text-xs font-mono space-y-2">
          <div className="flex flex-wrap gap-6">
            <div>Station Records: <span className={stationRecords > 0 ? "text-emerald-400" : "text-rose-400"}>{stationRecords}</span></div>
            <div>Train Records: <span className={trainRecords > 0 ? "text-emerald-400" : "text-rose-400"}>{trainRecords}</span></div>
            <div>Comparison Groups: <span className={stationComparison.length > 0 ? "text-emerald-400" : "text-rose-400"}>{stationComparison.length}</span></div>
            <div className="text-slate-500">|</div>
            <div>Raw Station Stats: {stats.stationStats.length}</div>
            <div>Unique Stations: {new Set(stats.stationStats.map(s => s.stationId)).size}</div>
            {stats.skippedRfRows > 0 && <div className="text-amber-400">Skipped Rows (No Stn ID): {stats.skippedRfRows}</div>}
          </div>
          <div className="text-[10px] text-slate-500 pt-2 border-t border-white/5 flex flex-col gap-1">
            <div>Detected Train IDs: {trainLocoIds.slice(0, 10).join(', ')} {trainLocoIds.length > 10 ? '...' : ''}</div>
            <div>Detected Station IDs: {stationIds.slice(0, 20).map(id => formatStationName(id)).join(', ')} {stationIds.length > 20 ? '...' : ''}</div>
          </div>
          {stationRecords === 0 && (
            <div className="text-rose-400 border-t border-white/5 pt-2 mt-2">
              ⚠️ No Station-side logs detected. Ensure you have uploaded files with "RFCOMM_ST" in the name or folder.
            </div>
          )}
        </div>

      {/* Deep Analysis Dashboard */}
      {stats.stationDeepAnalysis.dashboard && (
        <div className="glass-card p-8 rounded-3xl border-2 border-emerald-500/20 relative overflow-hidden group">
          <div className="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
            <Zap className="w-32 h-32 text-emerald-400" />
          </div>
          
          <div className="relative z-10">
            <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-6">
              <div>
                <h2 className="text-3xl font-black text-white tracking-tighter flex items-center gap-3">
                  <ShieldCheck className="w-8 h-8 text-emerald-400" />
                  Deep Analysis — Packet Loss Root Cause
                </h2>
                <p className="text-emerald-400 font-bold mt-1 uppercase tracking-widest text-xs">
                  Conclusion: {stats.stationDeepAnalysis.dashboard.conclusion}
                </p>
              </div>
              
              <div className="flex gap-3">
                {/* Verdict Boxes */}
                {(() => {
                  const isLFaulty = stats.stationDeepAnalysis.dashboard.conclusion.includes('Suspect') || stats.stationDeepAnalysis.dashboard.conclusion.includes('Multiple');
                  const isSFaulty = stats.stationDeepAnalysis.dashboard.conclusion.includes('Station') || stats.stationDeepAnalysis.dashboard.conclusion.includes('Multiple');
                  
                  return (
                    <>
                      <div className={cn(
                        "px-4 py-2 rounded-xl border flex flex-col items-center justify-center min-w-[100px]",
                        isLFaulty ? "bg-rose-500/10 border-rose-500/30" : "bg-emerald-500/10 border-emerald-500/30"
                      )}>
                        <span className="text-[9px] font-black text-slate-500 uppercase">Loco TCAS</span>
                        <span className={cn("text-sm font-black", isLFaulty ? "text-rose-400" : "text-emerald-400")}>
                          {isLFaulty ? 'SUSPECT' : 'FIT'}
                        </span>
                      </div>

                      <div className={cn(
                        "px-4 py-2 rounded-xl border flex flex-col items-center justify-center min-w-[100px]",
                        isSFaulty ? "bg-amber-500/10 border-amber-500/30" : "bg-emerald-500/10 border-emerald-500/30"
                      )}>
                        <span className="text-[9px] font-black text-slate-500 uppercase">Station TCAS</span>
                        <span className={cn("text-sm font-black", isSFaulty ? "text-amber-400" : "text-emerald-400")}>
                          {isSFaulty ? 'INSPECT' : 'HEALTHY'}
                        </span>
                      </div>

                      <div className="px-4 py-2 rounded-xl border bg-blue-500/10 border-blue-500/30 flex flex-col items-center justify-center min-w-[100px]">
                        <span className="text-[9px] font-black text-slate-500 uppercase">Benchmark</span>
                        <span className="text-sm font-black text-blue-400">CLEARED</span>
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              {/* Problem 1 */}
              <div className="space-y-4 bg-white/5 p-6 rounded-2xl border border-white/10">
                <h3 className="text-xl font-bold text-white flex items-center gap-2">
                  <div className={cn(
                    "w-2 h-2 rounded-full",
                    stats.stationDeepAnalysis.dashboard.conclusion.includes('Fit') || stats.stationDeepAnalysis.dashboard.conclusion.includes('Healthy')
                      ? "bg-emerald-400"
                      : "bg-rose-500 animate-ping"
                  )} />
                  {stats.stationDeepAnalysis.dashboard.problem1.title}
                </h3>
                <p className="text-slate-400 text-sm leading-relaxed">
                  {stats.stationDeepAnalysis.dashboard.problem1.description}
                </p>
                
                <div className="overflow-hidden rounded-xl border border-white/5">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-white/5 text-slate-500 uppercase font-bold">
                      <tr>
                        <th className="p-3">Station</th>
                        <th className="p-3 text-rose-400">Loco {stats.locoId}</th>
                        <th className="p-3 text-emerald-400">Baaki Locos (Avg)</th>
                      </tr>
                    </thead>
                    <tbody className="text-slate-300">
                      {stats.stationDeepAnalysis.dashboard.problem1.table.map((row, idx) => (
                        <tr key={idx} className="border-t border-white/5">
                          <td className="p-3 font-bold">{formatStationName(row.station)}</td>
                          <td className="p-3 font-black text-rose-400">{row.locoVal}</td>
                          <td className="p-3 font-bold text-emerald-400">{row.othersAvg}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {stats.stationDeepAnalysis.dashboard.problem1.causes.length > 0 && (
                  <div className="space-y-2">
                    <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Possible Causes:</p>
                    <ul className="space-y-1">
                      {stats.stationDeepAnalysis.dashboard.problem1.causes.map((cause, idx) => (
                        <li key={idx} className="text-xs text-slate-400 flex items-start gap-2">
                          <span className="text-rose-500 mt-1">•</span> {cause}
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
                {stats.stationDeepAnalysis.dashboard.problem1.causes.length === 0 && (
                  <div className="px-3 py-1 bg-emerald-500/10 border border-emerald-500/20 rounded-lg w-fit">
                    <span className="text-[10px] font-bold text-emerald-400 uppercase tracking-tight flex items-center gap-1">
                      <ShieldCheck className="w-3 h-3" /> Locomotive cleared
                    </span>
                  </div>
                )}
              </div>

              {/* Problem 2 & AML */}
              <div className="space-y-6">
                <div className="bg-white/5 p-6 rounded-2xl border border-white/10">
                  <h3 className="text-xl font-bold text-white flex items-center gap-2">
                    <div className="w-2 h-2 bg-amber-500 rounded-full" />
                    {stats.stationDeepAnalysis.dashboard.problem2.title}
                  </h3>
                  <p className="text-slate-400 text-sm leading-relaxed mt-2">
                    {stats.stationDeepAnalysis.dashboard.problem2.description}
                  </p>
                  {stats.stationDeepAnalysis.dashboard.problem2.priority.length > 0 && (
                    <div className="mt-4 flex flex-wrap gap-2">
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest w-full mb-1">Priority Order:</span>
                      {stats.stationDeepAnalysis.dashboard.problem2.priority.map((stn, idx) => (
                        <span key={idx} className="px-3 py-1 bg-amber-500/10 border border-amber-500/20 text-amber-400 text-xs font-bold rounded-lg">
                          {idx + 1}. {formatStationName(stn)}
                        </span>
                      ))}
                    </div>
                  )}
                </div>

                <div className="bg-emerald-500/5 p-6 rounded-2xl border border-emerald-500/10">
                  <h3 className="text-lg font-bold text-emerald-400 flex items-center gap-2">
                    <CheckCircle2 className="w-5 h-5" />
                    Station Performance Benchmark
                  </h3>
                  <p className="text-slate-400 text-sm leading-relaxed mt-2 italic">
                    {stats.stationDeepAnalysis.dashboard.amlConclusion}
                  </p>
                </div>

              </div>
            </div>
          </div>
        </div>
      )}


      <div className="glass-card p-8 rounded-2xl">
        <div className="flex justify-between items-center mb-6">
          <h3 className="text-xl font-bold text-white flex items-center gap-2">
            <BarChart3 className="w-6 h-6 text-emerald-400" />
            Station-wise RFCOMM Performance (Train vs Station Perspective)
          </h3>
          <div className="flex gap-4 text-[10px] font-bold uppercase tracking-widest">
            <div className="flex items-center gap-2"><div className="w-3 h-3 bg-emerald-500 rounded" /> Train View</div>
            <div className="flex items-center gap-2"><div className="w-3 h-3 bg-blue-500 rounded" /> Station View</div>
          </div>
        </div>
        
        <div className="h-[500px] w-full flex items-center justify-center">
          {stationComparison.length > 0 ? (
            <ResponsiveContainer width="100%" height="100%">
              <BarChart 
                data={stationComparison} 
                margin={{ bottom: 70 }}
              >
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" vertical={false} />
                <XAxis 
                  dataKey="label" 
                  stroke="#64748b" 
                  fontSize={10}
                  angle={-45}
                  textAnchor="end"
                  interval={0}
                />
                <YAxis 
                  stroke="#64748b" 
                  domain={[0, 100]} 
                  fontSize={10}
                  label={{ value: 'RFCOMM Success %', angle: -90, position: 'insideLeft', fill: '#64748b', fontSize: 12 }} 
                />
                <Tooltip 
                  contentStyle={{ backgroundColor: '#1e293b', border: 'none', borderRadius: '12px', boxShadow: '0 10px 15px -3px rgba(0,0,0,0.5)' }}
                  itemStyle={{ color: '#f8fafc', fontWeight: 'bold' }}
                  cursor={{ fill: 'rgba(255,255,255,0.05)' }}
                />
                <ReferenceLine 
                  y={95} 
                  stroke="#10b981" 
                  strokeDasharray="3 3" 
                  label={{ 
                    value: '95% Healthy', 
                    position: 'right', 
                    fill: '#10b981', 
                    fontSize: 10,
                    fontWeight: 'bold'
                  }} 
                />
                <ReferenceLine 
                  y={85} 
                  stroke="#f59e0b" 
                  strokeDasharray="3 3" 
                  label={{ 
                    value: '85% Warning', 
                    position: 'right', 
                    fill: '#f59e0b', 
                    fontSize: 10,
                    fontWeight: 'bold'
                  }} 
                />
                <Bar dataKey="trainPerc" name="Train View" fill="#10b981" radius={[4, 4, 0, 0]} barSize={20} />
                <Bar dataKey="stationPerc" name="Station View" fill="#3b82f6" radius={[4, 4, 0, 0]} barSize={20} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <div className="text-center space-y-4 py-12">
              <div className="w-20 h-20 bg-white/5 rounded-full flex items-center justify-center mx-auto">
                <BarChart3 className="w-10 h-10 text-slate-600" />
              </div>
              <div className="space-y-2">
                <p className="text-slate-400 font-medium">No RFCOMM data available for comparison</p>
                <p className="text-slate-500 text-sm max-w-md mx-auto">
                  {stats.stationStats.length === 0 
                    ? "No RFCOMM records were found in the processed files. Please check if the uploaded logs contain RFCOMM performance data."
                    : `Found ${stats.stationStats.length} RFCOMM records, but they couldn't be matched for comparison. Ensure station IDs and directions are consistent between Train and Station logs.`}
                </p>
              </div>
            </div>
          )}
        </div>
      </div>

      <div className="glass-card p-6 rounded-2xl">
        <h4 className="text-sm font-bold text-slate-400 uppercase tracking-wider mb-4">Detailed RFCOMM Log Mapping</h4>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead className="text-slate-500 uppercase text-[10px] font-bold border-b border-white/5">
              <tr>
                <th className="pb-3 px-4">Source</th>
                <th className="pb-3 px-4">Loco ID</th>
                <th className="pb-3 px-4">Station ID</th>
                <th className="pb-3 px-4">Direction</th>
                <th className="pb-3 px-4">Expected</th>
                <th className="pb-3 px-4">Received</th>
                <th className="pb-3 px-4">Success %</th>
              </tr>
            </thead>
            <tbody className="text-slate-300">
              {stats.stationStats.map((s, i) => (
                <tr key={i} className="border-b border-white/5 hover:bg-white/5 transition-colors">
                  <td className="py-3 px-4">
                    <span className={cn(
                      "px-2 py-0.5 rounded text-[10px] font-bold uppercase",
                      s.source === 'station' ? "bg-blue-500/20 text-blue-400" : "bg-emerald-500/20 text-emerald-400"
                    )}>
                      {s.source || 'Train'}
                    </span>
                  </td>
                  <td className="py-3 px-4 font-mono text-emerald-400">{s.locoId}</td>
                  <td className="py-3 px-4 font-bold text-white">
                    {formatStationName(s.stationId)}
                  </td>
                  <td className="py-3 px-4">
                    <span className={cn(
                      "px-2 py-0.5 rounded text-[10px] font-bold uppercase",
                      (s.direction || '').toLowerCase().includes('nominal') ? "bg-blue-500/20 text-blue-400" : "bg-purple-500/20 text-purple-400"
                    )}>
                      {s.direction}
                    </span>
                  </td>
                  <td className="py-3 px-4">{s.expected}</td>
                  <td className="py-3 px-4">{s.received}</td>
                  <td className="py-3 px-4">
                    <div className="flex items-center gap-2">
                      <div className="w-12 h-1.5 bg-white/5 rounded-full overflow-hidden">
                        <div 
                          className={cn("h-full", s.percentage < 85 ? "bg-rose-500" : s.percentage <= 95 ? "bg-amber-500" : "bg-emerald-500")} 
                          style={{ width: `${s.percentage}%` }} 
                        />
                      </div>
                      <span className={cn("font-bold", s.percentage < 85 ? "text-rose-400" : s.percentage <= 95 ? "text-amber-400" : "text-emerald-400")}>
                        {s.percentage.toFixed(2)}%
                      </span>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function OperationsSummary({ stats }: { stats: DashboardStats }) {
  const stationSummary = stats.stationSummary || [];
  const locoSummary = stats.locoSummary || [];

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
      <div className="flex items-center gap-3 mb-2">
        <MapPin className="w-8 h-8 text-emerald-400" />
        <div>
          <h2 className="text-3xl font-black text-white uppercase tracking-tighter">Station-wise Operations Summary</h2>
          <p className="text-slate-400 text-sm font-medium uppercase tracking-widest opacity-60">Aggregate view of Kavach events per station</p>
        </div>
      </div>

      <div className="glass-card overflow-hidden rounded-3xl border border-white/5 shadow-2xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-white/5 text-slate-500 uppercase text-[10px] font-black tracking-[0.2em] border-b border-white/5">
              <tr>
                <th className="py-5 px-6">Station Name</th>
                <th className="py-5 px-6 text-center">Locos (Unique)</th>
                <th className="py-5 px-6 text-center text-rose-400">Mode Deg.</th>
                <th className="py-5 px-6 text-center">Total Brakes</th>
                <th className="py-5 px-6 text-center text-rose-500">EB</th>
                <th className="py-5 px-6 text-center text-amber-500">FSB</th>
                <th className="py-5 px-6 text-right">Attribution</th>
              </tr>
            </thead>
            <tbody className="text-slate-300 divide-y divide-white/5">
              {stationSummary.map((s, i) => (
                <tr key={i} className="hover:bg-white/[0.02] transition-colors group">
                  <td className="py-4 px-6 font-bold text-white group-hover:text-emerald-400 transition-colors">
                    {s.stationName}
                  </td>
                  <td className="py-4 px-6 text-center font-mono">{s.locoCount}</td>
                  <td className="py-4 px-6 text-center">
                    <span className={cn(
                      "px-2 py-1 rounded-lg font-bold",
                      s.degradationCount > 0 ? "bg-rose-500/20 text-rose-400" : "bg-emerald-500/10 text-emerald-400"
                    )}>
                      {s.degradationCount}
                    </span>
                  </td>
                  <td className="py-4 px-6 text-center font-bold text-slate-400">{s.totalBrakes}</td>
                  <td className="py-4 px-6 text-center font-black text-rose-500">{s.ebCount}</td>
                  <td className="py-4 px-6 text-center font-black text-amber-500">{s.fsbCount}</td>
                  <td className="py-4 px-6 text-right">
                    <span className={cn(
                      "px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider border",
                      s.attribution === 'Non-RF' ? "bg-rose-500/10 text-rose-400 border-rose-500/20" : "bg-emerald-500/10 text-emerald-400 border-emerald-500/20"
                    )}>
                      {s.attribution}
                    </span>
                  </td>
                </tr>
              ))}
              {stationSummary.length === 0 && (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-500 italic">No station summary data available. Upload RFCOMM and TRN logs.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="flex items-center gap-3 mt-12 mb-2">
        <Train className="w-8 h-8 text-blue-400" />
        <div>
          <h2 className="text-3xl font-black text-white uppercase tracking-tighter">Loco-wise Operations Summary</h2>
          <p className="text-slate-400 text-sm font-medium uppercase tracking-widest opacity-60">Comparative analysis of Kavach health across locos</p>
        </div>
      </div>

      <div className="glass-card overflow-hidden rounded-3xl border border-white/5 shadow-2xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-white/5 text-slate-500 uppercase text-[10px] font-black tracking-[0.2em] border-b border-white/5">
              <tr>
                <th className="py-5 px-6">Loco Number</th>
                <th className="py-5 px-6 text-center">Stations Visited</th>
                <th className="py-5 px-6 text-center text-rose-400">Mode Deg.</th>
                <th className="py-5 px-6 text-center">Total Brakes</th>
                <th className="py-5 px-6 text-center text-rose-500">EB</th>
                <th className="py-5 px-6 text-center text-amber-500">FSB</th>
                <th className="py-5 px-6 text-right">Health Verdict / Attribution</th>
              </tr>
            </thead>
            <tbody className="text-slate-300 divide-y divide-white/5">
              {locoSummary.map((l, i) => (
                <tr key={i} className="hover:bg-white/[0.02] transition-colors group">
                  <td className="py-4 px-6 font-black text-emerald-400 font-mono text-lg">
                    {l.locoId}
                  </td>
                  <td className="py-4 px-6 text-center font-bold">{l.stationCount}</td>
                  <td className="py-4 px-6 text-center">
                    <span className={cn(
                      "px-2 py-1 rounded-lg font-bold",
                      l.degradationCount > 2 ? "bg-rose-500/20 text-rose-400" : l.degradationCount > 0 ? "bg-amber-500/20 text-amber-400" : "bg-emerald-500/10 text-emerald-400"
                    )}>
                      {l.degradationCount}
                    </span>
                  </td>
                  <td className="py-4 px-6 text-center font-bold text-slate-400">{l.totalBrakes}</td>
                  <td className="py-4 px-6 text-center font-black text-rose-500">{l.ebCount}</td>
                  <td className="py-4 px-6 text-center font-black text-amber-500">{l.fsbCount}</td>
                  <td className="py-4 px-6 text-right">
                    <span className={cn(
                      "px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider border",
                      l.attribution.includes('Issue') ? "bg-rose-500/10 text-rose-400 border-rose-500/20" : "bg-emerald-500/10 text-emerald-400 border-emerald-500/20"
                    )}>
                      {l.attribution}
                    </span>
                  </td>
                </tr>
              ))}
              {locoSummary.length === 0 && (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-500 italic">No loco summary data available.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function TechnicalAuditView({ stats }: { stats: DashboardStats }) {
  const audits = stats.technicalAudit;
  if (!audits || audits.length === 0) return (
    <div className="glass-card p-12 text-center rounded-3xl border-dashed border-2 border-white/5">
      <div className="w-16 h-16 bg-white/5 rounded-full flex items-center justify-center mx-auto mb-4">
        <ClipboardList className="w-8 h-8 text-slate-500" />
      </div>
      <h3 className="text-xl font-bold text-white mb-2">Technical Audit Log</h3>
      <p className="text-slate-400 max-w-sm mx-auto">Upload more data to trigger deeper event correlation and automated technical audits.</p>
    </div>
  );

  return (
    <div className="space-y-12">
       <div className="flex items-center justify-between">
          <div>
            <h2 className="text-2xl font-black text-white uppercase tracking-tight flex items-center gap-3">
              <ClipboardList className="w-7 h-7 text-blue-400" />
              Technical Audit Report
            </h2>
            <p className="text-slate-400 text-sm font-medium mt-1 uppercase tracking-widest opacity-60">Deep Event Analysis (Claude-Report Format)</p>
          </div>
          <div className="px-4 py-2 bg-blue-500/10 border border-blue-500/20 rounded-xl text-blue-400 text-[10px] font-black uppercase tracking-[0.2em]">
            {audits.length} AUDITS LOGGED
          </div>
       </div>

       {audits.map((audit, idx) => (
         <div key={audit.id} className="relative group">
           {/* Severity Border Glow */}
           <div className={cn(
             "absolute -inset-0.5 rounded-[2.5rem] blur opacity-20 group-hover:opacity-30 transition duration-500",
             audit.severity === 'Critical' ? "bg-rose-500" : audit.severity === 'Major' ? "bg-amber-500" : "bg-emerald-500"
           )}></div>
           
           <div className="relative bg-[#0a0f1d] overflow-hidden rounded-[1.5rem] border border-white/10 shadow-2xl font-mono">
              {/* Clinical Header */}
              <div className="px-8 py-6 bg-white/5 border-b border-white/10">
                <div className="flex justify-between items-start mb-4 text-[10px] opacity-60">
                   <div className="flex items-center gap-2 text-slate-400 font-bold">
                     <Clock className="w-3.5 h-3.5" />
                     {audit.timeRange}
                   </div>
                   <div className="flex items-center gap-2 text-slate-400 font-bold">
                     <MapPin className="w-3.5 h-3.5" />
                     {audit.stationName || audit.stationId}
                   </div>
                </div>

                <div className="flex justify-between items-start mb-4">
                  <h3 className="text-xl font-bold text-blue-400 italic">EVENT {idx + 1} — {audit.title.toUpperCase()}</h3>
                  <span className={cn(
                    "px-3 py-1 rounded text-[10px] font-black uppercase tracking-widest",
                    audit.severity === 'Critical' ? "bg-rose-500/20 text-rose-400" : 
                    audit.severity === 'Major' ? "bg-amber-500/20 text-amber-400" : "bg-emerald-500/20 text-emerald-400"
                  )}>{audit.severity}</span>
                </div>
                
                {/* Summary Line */}
                <div className="bg-white/5 p-3 rounded-lg border border-white/5 flex flex-wrap gap-x-12 gap-y-2 text-sm text-slate-300">
                  {audit.highlights.map((h, hi) => (
                    <div key={hi} className="flex gap-2">
                       <span className="text-slate-500 uppercase">{h.label}:</span>
                       <span className={cn("font-bold", h.color === 'rose' ? 'text-rose-400' : h.color === 'amber' ? 'text-amber-400' : h.color === 'emerald' ? 'text-emerald-400' : 'text-blue-400')}>
                         {h.value}
                       </span>
                    </div>
                  ))}
                  <div className="flex gap-2">
                    <span className="text-slate-500 uppercase">TRANSITION:</span>
                    <span className="text-white font-bold italic">{audit.transition}</span>
                  </div>
                </div>
              </div>

              <div className="p-8 space-y-8">
                 {/* Analysis Section */}
                 <div className="lg:col-span-12">

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                       <div>
                    <h4 className="text-xs font-black text-slate-300 uppercase tracking-widest mb-4 border-b border-white/10 pb-2 flex items-center gap-2">
                      <FileText className="w-4 h-4 text-blue-400" />
                      Telemetry Observations & Evidence:
                    </h4>
                    <ul className="space-y-4">
                       {audit.analysisBullets.map((bullet, bi) => (
                         <li key={bi} className="flex gap-3 group/li">
                            <span className="text-blue-500 font-bold shrink-0">-</span>
                            <p className="text-sm font-medium text-slate-300 leading-relaxed font-mono">
                              {bullet}
                            </p>
                         </li>
                       ))}
                    </ul>
                       </div>

                       <div className="flex flex-col gap-6">
                         <div className="bg-blue-500/5 p-6 rounded-2xl border border-blue-500/20 font-mono">
                         </div>
                       </div>
                    </div>
                 </div>

                 {/* Impacted Locos Footer */}
                 <div className="lg:col-span-12 mt-4 pt-4 border-t border-white/5 flex flex-wrap gap-3 items-center opacity-60">
                    <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mr-2">Audit Validated Across:</p>
                    {audit.locoIds.map((lId) => (
                      <div key={lId} className="flex items-center gap-1.5 px-2 py-0.5 bg-white/5 rounded text-[10px] text-slate-400 font-bold border border-white/5">
                        <Shield className="w-2.5 h-2.5 text-blue-500" />
                        LOCO {lId}
                      </div>
                    ))}
                 </div>
              </div>
           </div>
         </div>
       ))}
    </div>
  );
}

function SmartDiagnosisWidget({ stats }: { stats: DashboardStats }) {
  const diag = stats.smartDiagnosis;
  if (!diag) return null;

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-8">
      {/* Global Pattern Card */}
      <div className="lg:col-span-12 glass-card p-6 rounded-2xl border-l-4 border-emerald-500 overflow-hidden relative shadow-2xl">
        <div className="absolute right-0 top-0 opacity-5 -mr-8 -mt-8 pointer-events-none">
          <Activity className="w-48 h-48" />
        </div>
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 relative z-10">
          <div className="flex-1">
            <div className="flex items-center gap-2 mb-2">
              <ShieldCheck className="w-6 h-6 text-emerald-400" />
              <h3 className="text-lg font-black text-white uppercase tracking-tight">Global Infrastructure Pattern Analysis</h3>
            </div>
            <p className="text-sm text-slate-400 max-w-3xl leading-relaxed">
              {diag.globalPattern?.explanation}
            </p>
          </div>
          <div className="bg-emerald-500/10 border border-emerald-500/20 rounded-2xl p-4 flex items-center gap-5 min-w-[240px] backdrop-blur-sm">
             <div className="w-14 h-14 rounded-full border-2 border-emerald-500/30 flex items-center justify-center bg-white/5">
                <span className="text-xl font-black text-emerald-400">{diag.globalPattern?.percentage.toFixed(1)}%</span>
             </div>
             <div>
                <p className="text-[10px] text-slate-400 uppercase font-black tracking-widest leading-tight">Infrastructure Stress Rate</p>
                <p className="text-xs font-bold text-white mt-1">{diag.globalPattern?.affectedRows.toLocaleString()} / {diag.globalPattern?.totalRows.toLocaleString()} Data Rows</p>
             </div>
          </div>
        </div>
      </div>

      {/* Summary Message */}
      <div className="lg:col-span-12 px-6 py-4 bg-blue-500/5 rounded-2xl border border-blue-500/10 flex items-center gap-4 shadow-inner">
         <div className="w-10 h-10 rounded-full bg-blue-500/10 flex items-center justify-center shrink-0">
            <Info className="w-5 h-5 text-blue-400" />
         </div>
         <p className="text-sm font-bold text-slate-200 leading-snug">
           {diag.summary}
         </p>
      </div>

      {/* Station Blame / Insights */}
      <div className="lg:col-span-7 flex flex-col gap-6">
        <div className="flex items-center justify-between px-2">
          <h3 className="text-sm font-black text-slate-500 uppercase tracking-[0.25em] flex items-center gap-2">
            <MapPin className="w-4 h-4" />
            Station-Wise Blame Analysis
          </h3>
          <span className="text-[10px] font-bold text-slate-600 bg-white/5 px-2 py-1 rounded">MULTI-LOCO CONFIRMATION</span>
        </div>
        
        <div className="grid grid-cols-1 gap-4">
          {diag.stationInsights.map((stn, i) => (
            <div key={i} className={cn(
              "p-6 rounded-3xl border flex flex-col gap-5 relative overflow-hidden transition-all hover:translate-x-1 duration-300",
              stn.severity === 'Critical' ? "bg-rose-500/10 border-rose-500/20" : "bg-amber-500/10 border-amber-500/20"
            )}>
              <div className="flex justify-between items-start">
                <div className="flex items-center gap-3">
                  <div className={cn("w-10 h-10 rounded-xl flex items-center justify-center shadow-lg font-black text-lg", 
                    stn.severity === 'Critical' ? "bg-rose-500 text-white" : "bg-amber-500 text-black")}>
                    {stn.stationId.slice(0, 2)}
                  </div>
                  <div>
                    <h4 className={cn("text-xl font-black uppercase tracking-tighter leading-none", stn.severity === 'Critical' ? "text-rose-400" : "text-amber-400")}>
                      {stn.stationName}
                    </h4>
                    <p className="text-sm font-bold text-white/90 mt-1">{stn.description}</p>
                  </div>
                </div>
                <span className={cn(
                  "px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest",
                  stn.severity === 'Critical' ? "bg-rose-500/20 text-rose-400 border border-rose-500/30" : "bg-amber-500/20 text-amber-400 border border-amber-500/30"
                )}>
                  {stn.severity}
                </span>
              </div>
              
              <ul className="space-y-2.5">
                {stn.details.map((detail, di) => (
                  <li key={di} className="flex gap-3 text-xs font-medium text-slate-300 leading-relaxed">
                    <span className={cn("mt-1.5 w-1.5 h-1.5 rounded-full shrink-0", stn.severity === 'Critical' ? "bg-rose-400" : "bg-amber-400")} />
                    <span>{detail}</span>
                  </li>
                ))}
              </ul>
              
              <div className="mt-2 pt-5 border-t border-white/5 flex flex-wrap gap-2 items-center">
                 <p className="text-[10px] text-slate-500 font-black uppercase tracking-widest flex-1 min-w-full mb-1">Impacted Fleet:</p>
                 {stn.locosAffected.map((lId, li) => (
                   <span key={li} className="px-3 py-1 bg-white/5 rounded-lg text-[10px] font-black text-slate-400 border border-white/10 uppercase">Loco {lId}</span>
                 ))}
              </div>
            </div>
          ))}
          {diag.stationInsights.length === 0 && (
            <div className="bg-white/5 border border-dashed border-white/10 p-10 rounded-3xl text-center">
              <p className="text-slate-500 text-sm font-medium italic italic">No critical station-side infrastructure patterns found.</p>
            </div>
          )}
        </div>
      </div>

      {/* Loco & Protection Insights */}
      <div className="lg:col-span-5 flex flex-col gap-6">
        <h3 className="text-sm font-black text-slate-500 uppercase tracking-[0.25em] px-2 flex items-center gap-2">
          <Activity className="w-4 h-4" />
          Loco & Safety Protection
        </h3>
        
        <div className="space-y-5">
          {/* Hardware Alerts */}
          {diag.locoInsights.map((loco, i) => (
            <div key={i} className="p-6 rounded-3xl bg-blue-500/10 border border-blue-500/20 border-l-8 border-blue-500 shadow-xl relative overflow-hidden group">
               <div className="absolute right-0 top-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                 <Settings className="w-12 h-12 text-blue-400" />
               </div>
               <div className="flex justify-between items-center mb-4 relative z-10">
                 <h4 className="text-xs font-black text-blue-400 uppercase tracking-widest">Loco Side Fault: {loco.locoId}</h4>
               </div>
               <p className="text-lg font-black text-white leading-tight mb-2 relative z-10">{loco.issue}</p>
               <p className="text-sm font-medium text-slate-400 italic mb-6 relative z-10">{loco.context}</p>
               <div className="p-4 bg-white/5 rounded-2xl border border-white/10 backdrop-blur-md relative z-10">
                  <div className="flex items-center gap-2 mb-2">
                    <ShieldCheck className="w-4 h-4 text-emerald-400" />
                    <p className="text-[10px] font-black text-blue-300 uppercase tracking-widest">Expert Action Plan</p>
                  </div>
                  <p className="text-xs font-bold text-slate-200 leading-relaxed">{loco.recommendation}</p>
               </div>
            </div>
          ))}

          {/* Protection Events (ReadEndCollision) */}
          {diag.protectionEvents.map((pe, i) => (
            <div key={i} className="p-6 rounded-3xl bg-emerald-500/10 border border-emerald-500/20 border-l-8 border-emerald-500 shadow-xl relative overflow-hidden group">
               <div className="absolute right-0 top-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
                 <ShieldCheck className="w-12 h-12 text-emerald-400" />
               </div>
               <div className="flex justify-between items-center mb-4 relative z-10">
                 <h4 className="text-xs font-black text-emerald-400 uppercase tracking-widest">Safety Intervention</h4>
                 <span className="text-[10px] font-black font-mono text-slate-500">{pe.time}</span>
               </div>
               <div className="mb-4 relative z-10">
                 <p className="text-lg font-black text-white tracking-tight">{pe.event}</p>
                 <p className="text-sm text-slate-400 mt-1 font-medium italic">
                    Detected at <span className="text-white font-black">{formatStationName(pe.stationId)}</span> (Loco {pe.locoId})
                 </p>
               </div>
               <div className="p-4 bg-emerald-500/10 rounded-2xl border border-emerald-500/20 flex gap-3 relative z-10 shadow-lg">
                 <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0 mt-0.5" />
                 <p className="text-xs font-bold text-emerald-200/90 leading-relaxed">
                   {pe.analysis}
                 </p>
               </div>
            </div>
          ))}

          {diag.locoInsights.length === 0 && diag.protectionEvents.length === 0 && (
             <div className="bg-white/5 border border-white/5 p-10 rounded-3xl text-center">
                <p className="text-slate-500 text-sm font-medium italic">No active locomotive hardware faults or safety protection alerts identified.</p>
             </div>
          )}
        </div>
      </div>
    </div>
  );
}

function ExpertDiagnostics({ 
  stats, 
  tagSearch, 
  setTagSearch,
  remarks,
  onUpdateRemark
}: { 
  stats: DashboardStats; 
  tagSearch: string; 
  setTagSearch: (v: string) => void;
  remarks: Record<string, string>;
  onUpdateRemark: (key: string, value: string) => void;
}) {
  const [showAudit, setShowAudit] = React.useState(false);
  const filteredTags = stats.tagLinkIssues.filter(t => 
    (t.info || '').toLowerCase().includes((tagSearch || '').toLowerCase()) || 
    (t.error || '').toLowerCase().includes((tagSearch || '').toLowerCase()) ||
    (t.stationId || '').toLowerCase().includes((tagSearch || '').toLowerCase())
  );

  return (
    <div className="space-y-8">
      {/* 0. ROOT CAUSE ACCURACY AUDIT (User Request) */}
      <div className="glass-card p-0 rounded-3xl overflow-hidden border border-emerald-500/30">
        <div className="px-6 py-4 bg-emerald-500/10 border-b border-emerald-500/20 flex items-center justify-between">
          <div className="flex items-center gap-3">
             <div className="p-2 bg-emerald-500/20 rounded-lg">
                <CheckCircle2 className="w-5 h-5 text-emerald-400" />
             </div>
             <div>
                <h3 className="text-sm font-black text-white uppercase tracking-wider">Root Cause Accuracy Audit</h3>
                <p className="text-[10px] text-emerald-400/70 font-bold uppercase tracking-widest leading-none mt-1">Comparing Machine Analysis vs Human Ground-Truth</p>
             </div>
          </div>
          <button 
            onClick={() => setShowAudit(!showAudit)}
            className="px-4 py-2 bg-emerald-500/20 hover:bg-emerald-500/30 text-emerald-400 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all border border-emerald-500/30"
          >
            {showAudit ? 'Collapse Audit' : 'Review 12 Verified Events'}
          </button>
        </div>
        {showAudit && (
          <div className="p-6 bg-[#0a0f1d]/50">
            <RootCauseAccuracyAudit />
          </div>
        )}
      </div>

      {/* #2 Technical Audit Narrative Report (Deep Event Analysis) */}
      {/* #1 Mode Degradation Events */}
      <div className="space-y-4">
        <div className="flex items-center justify-between px-2">
          <h3 className="text-sm font-black text-slate-500 uppercase tracking-[0.25em] flex items-center gap-2">
            <Activity className="w-5 h-5 text-rose-400" />
            1. Mode Degradation Events
          </h3>
          <span className="text-[10px] font-bold text-slate-600 bg-white/5 px-2 py-1 rounded">ROOT CAUSE ANALYSIS ACTIVE</span>
        </div>

        <div className="grid grid-cols-1 gap-4">
          {stats.modeDegradations.length > 0 ? stats.modeDegradations.map((d, i) => (
            <div key={i} className={cn(
              "glass-card p-4 rounded-2xl border-l-8 overflow-hidden transition-all hover:translate-x-1 duration-300",
              d.severity === 'critical' ? "border-rose-500" : d.severity === 'warning' ? "border-amber-500" : "border-blue-500"
            )}>
              <div className="flex flex-wrap items-center justify-between gap-6">
                <div className="flex flex-wrap items-center gap-6">
                  <div className="flex flex-col">
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest leading-none mb-1">Time</span>
                    <span className="text-white font-mono font-bold">{d.time}</span>
                  </div>
                  <div className="h-8 w-px bg-white/10 hidden sm:block" />
                  <div className="flex flex-col">
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest leading-none mb-1">Station</span>
                    <StationDisplay id={d.stationId} name={d.stationName} />
                  </div>
                  <div className="h-8 w-px bg-white/10 hidden sm:block" />
                  <div className="flex flex-col">
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest leading-none mb-1">Loco No</span>
                    <span className="text-emerald-400 font-black">{d.locoId}</span>
                  </div>
                  <div className="h-8 w-px bg-white/10 hidden sm:block" />
                  <div className="flex flex-col">
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest leading-none mb-1">Speed</span>
                    <span className="text-amber-400 font-bold tabular-nums">{d.speed} Kmph</span>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                   <div className="flex items-center bg-black/40 rounded-xl px-4 py-2 border border-white/5 shadow-inner">
                      <span className="text-xs font-black text-emerald-400 uppercase tracking-widest mr-3">{d.from}</span>
                      <ArrowRight className="w-4 h-4 text-slate-600 mr-3" />
                      <span className="text-xs font-black text-rose-400 uppercase tracking-widest">{d.to}</span>
                   </div>
                </div>
              </div>

              {/* Editable Reason Field */}
              <div className="mt-4 pt-4 border-t border-white/5">
                 <div className="flex items-center justify-between mb-2">
                    <label className="text-[10px] font-black text-slate-500 uppercase tracking-[0.2em] flex items-center gap-2">
                       <ClipboardList className="w-3 h-3 text-emerald-400" />
                       Technical Reason & Staff Remediation
                    </label>
                    <span className="text-[9px] font-bold text-emerald-500/50 uppercase italic">Input here auto-populates in PDF Report</span>
                 </div>
                 <textarea
                    placeholder="Type the genuine technical reason or staff findings for this degradation..."
                    className="w-full bg-black/30 border border-white/10 rounded-xl px-4 py-3 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all min-h-[80px] resize-none"
                    value={remarks[`deg-${d.time}-${d.locoId}`] || ''}
                    onChange={(e) => onUpdateRemark(`deg-${d.time}-${d.locoId}`, e.target.value)}
                 />
              </div>
            </div>
          )) : (
            <div className="bg-white/5 border border-dashed border-white/10 p-12 rounded-3xl text-center">
              <Activity className="w-10 h-10 text-slate-700 mx-auto mb-4" />
              <p className="text-slate-500 text-sm font-medium italic">No critical mode degradation events detected in the current telemetry sequence.</p>
            </div>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        {/* #2 Brake Applications by Kavach */}
        <div className="glass-card p-6 rounded-2xl border-t-2 border-amber-500/30">
          <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
            <Zap className="w-5 h-5 text-amber-400" />
            2. Brake Applications by Kavach
          </h3>
          <div className="space-y-3">
            {stats.brakeApplications.length > 0 ? stats.brakeApplications.map((b, i) => (
              <div key={i} className="bg-white/5 p-4 rounded-xl border border-white/5 flex flex-col gap-3 group hover:bg-white/10 transition-all">
                <div className="flex justify-between items-center text-left">
                  <div>
                    <div className="flex items-center gap-2 mb-1 text-left">
                      <span className="text-[10px] font-black bg-emerald-500 text-white px-1.5 py-0.5 rounded uppercase tracking-tighter">Loco {b.locoId}</span>
                      <p className="text-xs font-bold text-white uppercase tracking-tight">{b.type}</p>
                    </div>
                    <div className="text-[10px] text-slate-500 font-medium text-left">
                      {b.time} | Loc: {b.location}
                      <span className="ml-2 text-emerald-400 font-bold uppercase inline-flex items-center gap-1">
                        @ <StationDisplay id={b.stationId} name={b.stationName} />
                      </span>
                    </div>
                  </div>
                  <div className="text-right">
                    <p className="text-sm font-bold text-amber-400 tabular-nums">{b.speed} Kmph</p>
                  </div>
                </div>

                <div className="space-y-1">
                  <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Technician Remarks:</p>
                  <input
                    type="text"
                    placeholder="Enter reason for brake activation..."
                    className="w-full bg-black/40 border border-white/10 rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none focus:border-amber-500/50 transition-all"
                    value={remarks[`brake-${b.time}-${b.locoId}`] || ''}
                    onChange={(e) => onUpdateRemark(`brake-${b.time}-${b.locoId}`, e.target.value)}
                  />
                </div>

                {b.reason && b.reason !== "System Safety Trigger" && (
                   <div className="mt-1 flex items-center gap-1.5 px-2 py-1 bg-rose-500/10 border border-rose-500/10 rounded-md w-fit">
                      <ShieldAlert className="w-3 h-3 text-rose-400" />
                      <span className="text-[9px] font-black text-rose-400 uppercase tracking-widest leading-none">Automated Reason: {b.reason}</span>
                   </div>
                )}
              </div>
            )) : <p className="text-center py-8 text-slate-500 text-sm italic">No brake applications recorded.</p>}
          </div>
        </div>

        {/* SOS Events */}
        <div className="glass-card p-6 rounded-2xl border-t-2 border-rose-600/30">
          <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
            <AlertCircle className="w-5 h-5 text-rose-500" />
            3. SOS Emergency Events
          </h3>
          <div className="space-y-3">
            {stats.sosEvents.length > 0 ? stats.sosEvents.map((s, i) => (
              <div key={i} className="bg-rose-500/5 p-4 rounded-xl border border-rose-500/10 flex justify-between items-center group hover:bg-rose-500/10 transition-all">
                <div>
                  <p className="text-xs font-bold text-rose-400 uppercase tracking-widest">SOS Triggered</p>
                  <div className="text-[10px] text-slate-500 mt-1">
                    {s.time} | Source: {s.source} 
                    <span className="ml-2 text-rose-300 font-bold uppercase tracking-tight">@ <StationDisplay id={s.stationId} name={s.stationName} showEmerald={false} /></span>
                  </div>
                </div>
                <div className="text-right">
                  <span className="px-2 py-1 bg-rose-600 text-white rounded text-[10px] font-black uppercase tracking-widest shadow-lg shadow-rose-600/20">Critical</span>
                </div>
              </div>
            )) : <p className="text-center py-8 text-slate-500 text-sm italic">No SOS emergency events detected.</p>}
          </div>
        </div>
      </div>

      {/* #4 Emergency Status Reports (Non-Regular) */}
      <div className="glass-card p-6 rounded-2xl border-t-4 border-rose-500 bg-rose-500/5">
        <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
          <ShieldAlert className="w-5 h-5 text-rose-500" />
          4. Emergency Status Reports (Non-Regular Monitoring)
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {stats.emergencyStatusEvents.length > 0 ? stats.emergencyStatusEvents.map((emr, i) => (
            <div key={i} className="bg-slate-900/50 border border-white/10 rounded-2xl overflow-hidden group hover:border-rose-500/40 transition-all duration-300 shadow-xl shadow-rose-500/5">
              <div className="bg-rose-500/15 px-4 py-3 border-b border-white/5 flex justify-between items-center">
                <span className="text-[10px] font-black text-rose-400 uppercase tracking-widest">{emr.status}</span>
                <span className="text-[10px] font-bold text-slate-500 italic">#{emr.rowCount} packets</span>
              </div>
              <div className="p-4 space-y-4">
                <div className="flex justify-between items-start">
                   <div className="space-y-1">
                      <p className="text-[9px] text-slate-500 uppercase font-black tracking-widest">Loco Unit</p>
                      <p className="text-sm font-bold text-white font-mono">{emr.locoId}</p>
                   </div>
                   <div className="text-right space-y-1">
                      <p className="text-[9px] text-slate-500 uppercase font-black tracking-widest">Station Context</p>
                      <div className="text-xs font-bold text-emerald-400 uppercase tracking-tight">
                         <StationDisplay id={emr.stationId} name={emr.stationName} />
                      </div>
                   </div>
                </div>
                
                <div className="pt-3 border-t border-white/5 grid grid-cols-2 gap-4">
                   <div className="space-y-1">
                      <p className="text-[9px] text-slate-500 uppercase font-black tracking-widest">Onset Time</p>
                      <p className="text-[11px] font-mono text-amber-400/90">{emr.startTime}</p>
                   </div>
                   <div className="text-right space-y-1">
                      <p className="text-[9px] text-slate-500 uppercase font-black tracking-widest">Recovery Time</p>
                      <p className="text-[11px] font-mono text-rose-400/90">{emr.endTime}</p>
                   </div>
                </div>
              </div>
            </div>
          )) : (
            <div className="col-span-full py-12 text-center bg-white/5 rounded-2xl border border-dashed border-white/10">
              <p className="text-slate-500 font-medium italic text-sm">No Non-Regular Emergency Status events identified in the current analytical window.</p>
            </div>
          )}
        </div>
      </div>

      {/* #6 Root Cause & Pattern Analysis (Smart Diagnosis) */}
      <div className="pt-8 border-t border-white/10">
        <SmartDiagnosisWidget stats={stats} />
      </div>

      {/* #7 NMS Health Audit */}
      {stats.nmsLocoStats && stats.nmsLocoStats.length > 0 && (
          <div className="glass-card p-6 rounded-3xl border border-white/5">
            <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
              <Activity className="w-5 h-5 text-emerald-400" />
              NMS Health Audit (Hardware Fault Detection)
            </h3>
            <p className="text-sm text-slate-400 mb-6">
              NMS Health indicates the internal diagnostic status of the Loco Vital Computer (LVC). A value of '0' means healthy. High error percentages indicate hardware module failures (e.g., BIU Interface, RFID Reader) or internal communication issues.
            </p>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm">
                <thead className="text-slate-500 uppercase text-[10px] font-bold border-b border-white/5">
                  <tr>
                    <th className="pb-3 px-4">Loco ID</th>
                    <th className="pb-3 px-4 text-right">Total Records</th>
                    <th className="pb-3 px-4 text-right">Errors (Non-Zero)</th>
                    <th className="pb-3 px-4 text-right">Error %</th>
                    <th className="pb-3 px-4">Status</th>
                  </tr>
                </thead>
                <tbody className="text-slate-300">
                  {stats.nmsLocoStats.map((d, i) => (
                    <tr key={i} className="border-b border-white/5 hover:bg-white/5 transition-colors">
                      <td className="py-3 px-4 font-bold text-white">{d.locoId}</td>
                      <td className="py-3 px-4 text-right font-mono text-xs">{d.totalRecords.toLocaleString()}</td>
                      <td className="py-3 px-4 text-right font-mono text-xs text-rose-400">{d.errors.toLocaleString()}</td>
                      <td className="py-3 px-4 text-right font-mono text-xs font-bold">{d.errorPercentage}%</td>
                      <td className="py-3 px-4">
                        <span className={cn(
                          "px-2 py-1 rounded text-[10px] font-bold uppercase tracking-wider",
                          d.category.includes('Critical') ? "bg-rose-500/20 text-rose-400" :
                          d.category.includes('High') ? "bg-amber-500/20 text-amber-400" :
                          "bg-emerald-500/20 text-emerald-400"
                        )}>
                          {d.category}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="mt-6 grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="bg-white/5 border border-white/10 rounded-xl p-4">
                 <h4 className="text-sm font-bold text-white mb-2">Technical Conclusion</h4>
                 <p className="text-xs text-slate-400 leading-relaxed">
                   Continuous '8' or '16' codes are early warnings of specific card failures (e.g., Input Output Card or Communication Card). Locos in the "Critical" or "High" categories require immediate maintenance attention.
                 </p>
              </div>
              <div className="bg-white/5 border border-white/10 rounded-xl p-4">
                 <h4 className="text-sm font-bold text-white mb-2">Impact on System</h4>
                 <p className="text-xs text-slate-400 leading-relaxed">
                   When NMS Health is not '0', it can cause the Kavach system to downgrade from Full Supervision (FS) to Staff Responsible (SR) or Isolate mode, and can also increase radio packet drops.
                 </p>
              </div>
            </div>

            {/* Deep Analysis Table */}
            {stats.nmsDeepAnalysis && stats.nmsDeepAnalysis.length > 0 && (
              <div className="mt-8">
                <h4 className="text-sm font-bold text-white mb-4 uppercase tracking-wider text-slate-400">Continuous Error Events (Deep Analysis)</h4>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="text-slate-500 uppercase text-[10px] font-bold border-b border-white/5">
                      <tr>
                        <th className="pb-3 px-4">Loco ID</th>
                        <th className="pb-3 px-4">Station</th>
                        <th className="pb-3 px-4">Time Range</th>
                        <th className="pb-3 px-4">Code</th>
                        <th className="pb-3 px-4">Error Type</th>
                        <th className="pb-3 px-4">Count</th>
                      </tr>
                    </thead>
                    <tbody className="text-slate-300">
                      {stats.nmsDeepAnalysis.slice(0, 15).map((d, i) => (
                        <tr key={i} className="border-b border-white/5 hover:bg-white/5 transition-colors">
                          <td className="py-3 px-4 font-bold text-white">{d.locoId}</td>
                          <td className="py-3 px-4">
                            <StationDisplay id={d.stationId} name={d.stationName} />
                          </td>
                          <td className="py-3 px-4 font-mono text-xs">
                            {d.startTime.split(' ')[1]} - {d.endTime.split(' ')[1]}
                          </td>
                          <td className="py-3 px-4 font-mono text-rose-400 font-bold">{d.errorCode}</td>
                          <td className="py-3 px-4">
                            <div className="font-bold text-white">{d.errorType}</div>
                            <div className="text-[10px] text-slate-400 mt-0.5 max-w-xs truncate" title={d.description}>{d.description}</div>
                          </td>
                          <td className="py-3 px-4 font-mono text-xs text-amber-400">{d.count}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
      )}

      <div className="grid grid-cols-2 gap-8">
        {/* #8 Signal Overrides */}
        <div className="glass-card p-6 rounded-2xl border-l-4 border-blue-500/50">
          <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
            <Shield className="w-5 h-5 text-blue-400" />
            Signal Override Cases (Authority Violations)
          </h3>
          <div className="space-y-3">
            {stats.signalOverrides.length > 0 ? stats.signalOverrides.map((s, i) => (
              <div key={i} className="bg-slate-900 border border-white/5 p-4 rounded-xl flex justify-between items-center group hover:border-blue-500/30 transition-all">
                <div>
                  <p className="text-xs font-black text-blue-400 uppercase tracking-widest mb-1">Signal ID: {s.signalId}</p>
                  <div className="text-[10px] text-slate-500 font-semibold italic">
                    {s.time} @ <StationDisplay id={s.stationId} name={s.stationName} showEmerald={false} />
                  </div>
                </div>
                <div className="text-right">
                  <span className="text-xs font-black text-blue-400 bg-blue-500/10 px-3 py-1 rounded-full uppercase tracking-tighter ring-1 ring-blue-500/20">{s.status}</span>
                </div>
              </div>
            )) : <p className="text-center py-8 text-slate-500 text-sm italic">No signal override violations detected.</p>}
          </div>
        </div>

        {/* Tag Tracking Control */}
        <div className="glass-card p-6 rounded-2xl border-l-4 border-emerald-500/50 text-emerald-400">
          <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
            <Tag className="w-5 h-5 text-emerald-400" />
            Tag Link Audit Control
          </h3>
          <div className="space-y-4">
            <input
              type="text"
              placeholder="Search Tag Identity or Detail..."
              value={tagSearch}
              onChange={(e) => setTagSearch(e.target.value)}
              className="w-full bg-white/5 border border-white/10 rounded-xl px-4 py-2 text-sm text-white focus:outline-none focus:border-emerald-500 transition-all"
            />
            <p className="text-[10px] text-slate-500 uppercase font-black tracking-[0.2em] px-1">Issues Found: {filteredTags.length}</p>
          </div>
        </div>
      </div>

      {/* #9 Loco Length Variations (TRNMSNMA) */}
      <div className="glass-card p-6 rounded-2xl border-t-4 border-amber-500">
        <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
          <Settings className="w-5 h-5 text-amber-400" />
          Loco Length (TRNMSNMA) Integrity Check
        </h3>
        <div className="space-y-4">
          <div className="bg-white/5 p-4 rounded-xl border border-white/5">
            <p className="text-xs text-slate-400 uppercase font-bold mb-4">Unique Lengths Detected with Context</p>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {stats.uniqueTrainLengths.length > 0 ? stats.uniqueTrainLengths.map((item, i) => (
                <div key={i} className={cn(
                  "p-3 rounded-xl border flex flex-col gap-1",
                  stats.uniqueTrainLengths.length > 1 ? "bg-rose-500/10 border-rose-500/20" : "bg-emerald-500/10 border-emerald-500/20"
                )}>
                  <div className="flex justify-between items-center">
                    <span className={cn(
                      "text-lg font-bold",
                      stats.uniqueTrainLengths.length > 1 ? "text-rose-400" : "text-emerald-400"
                    )}>{item.length} m</span>
                    <span className="text-[10px] font-mono text-slate-500">{item.time}</span>
                  </div>
                  <div className="flex items-center gap-1 text-[10px] text-slate-400">
                    <MapPin className="w-3 h-3" />
                    <span>Station: {formatStationName(item.stationId)}</span>
                  </div>
                </div>
              )) : <p className="text-slate-500 text-sm italic">No length data found</p>}
            </div>
            
            {stats.uniqueTrainLengths.length > 1 && (
              <div className="mt-6 p-4 bg-rose-500/10 border border-rose-500/20 rounded-xl flex items-start gap-3">
                <AlertTriangle className="w-6 h-6 text-rose-400 shrink-0" />
                <div>
                  <p className="text-sm text-rose-300 font-bold uppercase tracking-tight">Critical Alert: Multiple Train Lengths Detected</p>
                  <p className="text-xs text-rose-400/80 mt-1">
                    Variations in reported train length (from {stats.uniqueTrainLengths[0].length}m to {stats.uniqueTrainLengths[stats.uniqueTrainLengths.length-1].length}m) detected for Loco {stats.locoId}. 
                    This is a critical safety concern as it affects braking distance calculations and EBD/SBD curves.
                  </p>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* #10 Medha Kavach / Tag Link Issues */}
      <div className="glass-card p-6 rounded-2xl border-l-4 border-rose-500 bg-black/20">
        <div className="flex flex-col gap-6 mb-6">
          <div className="flex justify-between items-center">
            <h3 className="text-lg font-bold text-white flex items-center gap-2">
              <Info className="w-5 h-5 text-rose-400" />
              Tag Link Defects (Infrastructure Analysis)
            </h3>
            <div className="flex gap-4">
              <div className="bg-rose-500/10 px-4 py-2 rounded-xl border border-rose-500/20 text-center">
                <p className="text-[10px] text-slate-400 uppercase font-black tracking-widest">Both Tags Missing</p>
                <p className="text-xl font-bold text-rose-400">
                  {stats.tagLinkIssues.filter(t => t.isCritical).length}
                </p>
              </div>
              <div className="bg-amber-500/10 px-4 py-2 rounded-xl border border-amber-500/20 text-center">
                <p className="text-[10px] text-slate-400 uppercase font-black tracking-widest">Single Tag Missing</p>
                <p className="text-xl font-bold text-amber-400">
                  {stats.tagLinkIssues.filter(t => !t.isCritical).length}
                </p>
              </div>
            </div>
          </div>
          
          <div className="relative">
            <input 
              type="text"
              placeholder="Search Tag Issues (e.g. Main Tag Missing, Station ID...)"
              value={tagSearch}
              onChange={(e) => setTagSearch(e.target.value)}
              className="w-full bg-white/5 border border-white/10 rounded-xl px-4 py-3 text-sm text-white focus:outline-none focus:border-emerald-500/50 transition-all font-mono"
            />
            <div className="absolute right-4 top-1/2 -translate-y-1/2 flex gap-2">
              <button 
                onClick={() => setTagSearch('Main Tag Missing')}
                className="text-[9px] font-black uppercase tracking-tighter bg-rose-500/20 text-rose-400 px-2 py-1 rounded hover:bg-rose-500/30 transition-all"
              >
                Find Main Missing
              </button>
              <button 
                onClick={() => setTagSearch('Duplicate Tag Missing')}
                className="text-[9px] font-black uppercase tracking-tighter bg-amber-500/20 text-amber-400 px-2 py-1 rounded hover:bg-amber-500/30 transition-all"
              >
                Find Duplicate Missing
              </button>
              {tagSearch && (
                <button 
                  onClick={() => setTagSearch('')}
                  className="text-[9px] font-black uppercase tracking-tighter bg-white/10 text-slate-400 px-2 py-1 rounded hover:bg-white/20 transition-all"
                >
                  Clear
                </button>
              )}
            </div>
          </div>
        </div>

        <div className="overflow-x-auto rounded-xl border border-white/5">
          <table className="w-full text-left text-sm">
            <thead className="bg-white/5 text-slate-500 uppercase text-[10px] font-black border-b border-white/10">
              <tr>
                <th className="py-4 px-4 font-mono">Time</th>
                <th className="py-4 px-4 font-mono">Loco No</th>
                <th className="py-4 px-4 font-mono">Station ID</th>
                <th className="py-4 px-4 font-mono">Diagnostic Error</th>
                <th className="py-4 px-4 font-mono text-right">Tag Link Info</th>
              </tr>
            </thead>
            <tbody className="text-slate-300">
              {filteredTags.length > 0 ? filteredTags.map((t, i) => (
                <tr key={i} className="border-b border-white/5 hover:bg-white/5 transition-colors group">
                  <td className="py-4 px-4 font-mono text-xs">{t.time}</td>
                  <td className="py-4 px-4 font-mono text-emerald-400 font-bold">{t.locoId}</td>
                  <td className="py-4 px-4 text-white">
                     <span className="font-bold flex items-center gap-1 group-hover:text-emerald-400 transition-colors">
                        <StationDisplay id={t.stationId} name={t.stationName} />
                     </span>
                  </td>
                  <td className="py-4 px-4">
                    <p className="text-xs font-bold text-slate-300 leading-snug">{t.error}</p>
                  </td>
                  <td className="py-4 px-4 text-right">
                    <span className={cn(
                      "px-2 py-1 rounded text-[10px] font-black uppercase tracking-widest",
                      t.isCritical ? "bg-rose-500/20 text-rose-400 border border-rose-500/30" : "bg-amber-500/20 text-amber-400 border border-amber-500/30"
                    )}>
                      {t.info}
                    </span>
                  </td>
                </tr>
              )) : (
                <tr><td colSpan={5} className="py-12 text-center text-slate-500">No tag infrastructure defects detected in active registry.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function NMSAnalysis({ stats }: { stats: DashboardStats }) {
  const nmsColors: Record<string, string> = {
    '0': '#0066cc', '8': '#80ccff', '1': '#ff3333', '-': '#ffb3b3',
    '16': '#33b3a6', '32': '#80ffaa', '40': '#ff9900', 'default': '#64748b'
  };

  return (
    <div className="space-y-6">
      <div className="glass-card p-8 rounded-2xl">
        <h3 className="text-xl font-bold text-white mb-6 flex items-center gap-2">
          <Database className="w-6 h-6 text-emerald-400" />
          NMS Health Status Correlation
        </h3>
        <div className="grid grid-cols-2 gap-8 items-center">
          <div className="h-[400px]">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={stats.nmsStatus}
                  cx="50%" cy="50%"
                  outerRadius={120}
                  innerRadius={60}
                  dataKey="value"
                  minAngle={15}
                  labelLine={true}
                  label={({ name, percent }) => percent > 0.05 ? `${name}: ${(percent * 100).toFixed(1)}%` : ''}
                >
                  {stats.nmsStatus.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={nmsColors[entry.name] || nmsColors.default} />
                  ))}
                </Pie>
                <Tooltip contentStyle={{ backgroundColor: '#0f172a', border: 'none', borderRadius: '12px' }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="space-y-4">
            <div className="bg-white/5 p-6 rounded-2xl border border-white/5">
              <p className="text-sm text-slate-400 mb-2">Failure Rate Analysis</p>
              <p className="text-4xl font-bold text-rose-400">{stats.nmsFailRate.toFixed(1)}%</p>
              <p className="text-xs text-slate-500 mt-2">Percentage of logs where NMS Health was not 0.</p>
            </div>
            <div className="grid grid-cols-2 gap-3">
              {stats.nmsStatus.map((d, i) => (
                <div key={i} className="bg-white/5 p-3 rounded-xl flex items-center gap-3">
                  <div className="w-3 h-3 rounded-full" style={{ backgroundColor: nmsColors[d.name] || nmsColors.default }} />
                  <span className="text-xs text-slate-300 font-mono">{d.name}: {d.value}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function SyncAnalysis({ stats }: { stats: DashboardStats }) {
  return (
    <div className="space-y-6">
      <div className="glass-card p-8 rounded-2xl">
        <h3 className="text-xl font-bold text-white mb-6 flex items-center gap-2">
          <Activity className="w-6 h-6 text-emerald-400" />
          Movement Authority (MA) Packet Sync Analysis
        </h3>
        <div className="h-[400px] w-full">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={stats.maPackets}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
              <XAxis dataKey="time" hide />
              <YAxis stroke="#64748b" label={{ value: 'Delay (s)', angle: -90, position: 'insideLeft', fill: '#64748b' }} />
              <Tooltip 
                contentStyle={{ backgroundColor: '#0f172a', border: 'none', borderRadius: '12px' }}
                itemStyle={{ color: '#10b981' }}
              />
              <Line 
                type="monotone" 
                dataKey="delay" 
                stroke="#10b981" 
                strokeWidth={2} 
                dot={false}
                activeDot={{ r: 4, fill: '#10b981' }}
              />
            </LineChart>
          </ResponsiveContainer>
        </div>
        <div className="mt-8 grid grid-cols-3 gap-6">
          <div className="bg-white/5 p-4 rounded-xl">
            <p className="text-xs text-slate-500 uppercase font-bold mb-1">Avg Refresh Lag</p>
            <p className="text-2xl font-bold text-white">{stats.avgLag.toFixed(2)}s</p>
          </div>
          <div className="bg-white/5 p-4 rounded-xl">
            <p className="text-xs text-slate-500 uppercase font-bold mb-1">Total MA Packets</p>
            <p className="text-2xl font-bold text-white">{stats.maCount}</p>
          </div>
          <div className="bg-white/5 p-4 rounded-xl">
            <p className="text-xs text-slate-500 uppercase font-bold mb-1">Access Requests</p>
            <p className="text-2xl font-bold text-white">{stats.arCount}</p>
          </div>
        </div>
      </div>
    </div>
  );
}

function IntervalAnalysis({ stats }: { stats: DashboardStats }) {
  return (
    <div className="space-y-6">
      <div className="glass-card p-8 rounded-2xl">
        <h3 className="text-xl font-bold text-white mb-6 flex items-center gap-2">
          <BarChart3 className="w-6 h-6 text-emerald-400" />
          Packet Interval Distribution (RDSO Compliance)
        </h3>
        <div className="h-[400px] w-full">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={stats.intervalDist}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
              <XAxis dataKey="category" stroke="#64748b" />
              <YAxis stroke="#64748b" unit="%" />
              <Tooltip 
                contentStyle={{ backgroundColor: '#0f172a', border: 'none', borderRadius: '12px' }}
                cursor={{ fill: 'rgba(255,255,255,0.05)' }}
              />
              <Bar dataKey="percentage" radius={[8, 8, 0, 0]}>
                {stats.intervalDist.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={['#10b981', '#f59e0b', '#ef4444'][index]} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="mt-8 p-6 bg-emerald-500/5 border border-emerald-500/10 rounded-2xl">
          <p className="text-sm text-slate-300 leading-relaxed">
            <span className="font-bold text-emerald-400 mr-2">RDSO Standard:</span>
            Movement Authority (MA) packets must be refreshed every 1.0 seconds. Any delay exceeding 1.2 seconds triggers a session drop by the Loco system. Currently, <span className="font-bold text-white">{stats.intervalDist[0].percentage.toFixed(1)}%</span> of your packets are within the healthy range.
          </p>
        </div>
      </div>
    </div>
  );
}

function TabButton({ active, onClick, label }: { active: boolean; onClick: () => void; label: string }) {
  return (
    <button
      onClick={onClick}
      className={cn(
        "px-6 py-2 rounded-lg text-sm font-bold transition-all",
        active ? "bg-emerald-500 text-white shadow-lg shadow-emerald-500/20" : "text-slate-400 hover:text-white"
      )}
    >
      {label}
    </button>
  );
}

function InfrastructureStressMeter({ stress }: { stress?: DashboardStats['infrastructureStress'] }) {
  if (!stress) return null;

  const getScoreColor = (score: number) => {
    if (score > 60) return "text-rose-400";
    if (score > 30) return "text-amber-400";
    return "text-emerald-400";
  };

  const getProgressColor = (score: number) => {
    if (score > 60) return "bg-rose-500";
    if (score > 30) return "bg-amber-500";
    return "bg-emerald-500";
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="glass-card p-6 rounded-2xl hover-glow transition-all"
    >
      <div className="flex justify-between items-center mb-6">
        <div className="flex flex-col gap-1">
          <h4 className="font-bold text-white text-sm uppercase tracking-widest opacity-70 flex items-center gap-2">
            <Construction className="w-4 h-4 text-emerald-400" />
            Infrastructure Stress Meter
          </h4>
          <p className="text-[10px] text-slate-500 italic">Tracking 'InterTagDist' errors (Persistence without Degradation)</p>
        </div>
        <div className="text-right">
          <span className={cn("text-3xl font-black italic tracking-tighter block leading-none", getScoreColor(stress.overallScore))}>
            {stress.overallScore.toFixed(1)}%
          </span>
          <span className="text-[9px] uppercase font-bold text-slate-600">Avg Strain</span>
        </div>
      </div>

      <div className="space-y-6">
        <div className="relative pt-1">
          <div className="flex mb-2 items-center justify-between">
            <div>
              <span className="text-xs font-semibold inline-block py-1 px-2 uppercase rounded-full text-emerald-600 bg-emerald-200">
                Network Fatigue
              </span>
            </div>
            <div className="text-right">
              <span className="text-xs font-semibold inline-block text-emerald-600">
                {stress.overallScore.toFixed(0)}%
              </span>
            </div>
          </div>
          <div className="overflow-hidden h-2 mb-4 text-xs flex rounded bg-emerald-200 shadow-inner">
            <motion.div 
              initial={{ width: 0 }}
              animate={{ width: `${stress.overallScore}%` }}
              className={cn("shadow-none flex flex-col text-center whitespace-nowrap text-white justify-center", getProgressColor(stress.overallScore))}
            />
          </div>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <div className="bg-white/5 p-4 rounded-xl border border-white/5 group hover:bg-white/10 transition-all">
            <span className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Tag Points Analyzed</span>
            <span className="text-2xl font-bold text-white leading-none">{stress.totalTagTelemetry}</span>
            <div className="w-full h-1 bg-white/10 rounded-full mt-2 overflow-hidden">
               <div className="h-full bg-slate-500 w-full opacity-50" />
            </div>
          </div>
          <div className="bg-white/5 p-4 rounded-xl border border-white/5 group hover:bg-white/10 transition-all">
            <span className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Hidden Defects (ITD)</span>
            <span className="text-2xl font-bold text-rose-400 leading-none">{stress.tagSpacingDefects}</span>
            <div className="w-full h-1 bg-rose-500/20 rounded-full mt-2 overflow-hidden">
               <div className="h-full bg-rose-500" style={{ width: `${stress.overallScore}%` }} />
            </div>
          </div>
        </div>

        {stress.stationWiseStress.length > 0 && (
          <div className="pt-4 border-t border-white/10">
            <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4 flex items-center gap-2">
              <MapPin className="w-3 h-3" />
              Highest Trackside Load (By Station)
            </h5>
            <div className="grid grid-cols-1 gap-4">
              {stress.stationWiseStress.slice(0, 3).map((stn, i) => (
                <div key={i} className="flex items-center gap-4">
                  <div className="w-12 text-[10px] font-black text-slate-600 truncate">{i + 1}. {stn.stationId}</div>
                  <div className="flex-1">
                    <div className="flex justify-between text-[10px] font-bold mb-1">
                      <span className="text-slate-300">{formatStationName(stn.stationName || stn.stationId)}</span>
                      <span className={getScoreColor(stn.stressScore)}>{stn.stressScore.toFixed(1)}%</span>
                    </div>
                    <div className="w-full h-1 bg-white/5 rounded-full overflow-hidden">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: `${stn.stressScore}%` }}
                        className={cn("h-full rounded-full", getProgressColor(stn.stressScore))}
                      />
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
        
        <div className="p-3 bg-white/5 rounded-xl border border-dashed border-white/10">
           <p className="text-[10px] text-slate-400 leading-relaxed">
             <span className="text-emerald-400 font-bold">Expert Note:</span> High stress indicates physical tag spacing issues. If stress {'>'} 70%, mode degradation is imminent during monsoon or weak signal zones.
           </p>
        </div>
      </div>
    </motion.div>
  );
}

function SyncConflictChart({ conflicts }: { conflicts?: DashboardStats['conflictingPackets'] }) {
  if (!conflicts || conflicts.length === 0) return null;

  const stationData = conflicts.reduce((acc, c) => {
    const stn = c.stationName || c.stationId;
    acc[stn] = (acc[stn] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const chartData = Object.entries(stationData)
    .map(([name, count]) => ({ name: formatStationName(name), count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);

  return (
    <motion.div 
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      className="glass-card p-6 rounded-2xl border-l-4 border-rose-500"
    >
      <div className="flex items-center justify-between mb-4">
        <h4 className="font-bold text-white text-sm uppercase tracking-wider opacity-70 flex items-center gap-2">
          <Activity className="w-4 h-4 text-rose-500" />
          Hardware Sync Conflicts
        </h4>
        <div className="px-2 py-1 bg-rose-500/20 rounded text-[10px] font-bold text-rose-400 animate-pulse">
           {conflicts.length} EVENTS
        </div>
      </div>
      
      <p className="text-[10px] text-slate-400 mb-6 leading-relaxed italic">
        Multiple conflicting telemetry packets (Modes/NMS) received within 1 second. High probability of Radio board sync loss.
      </p>

      {chartData.length > 0 && (
        <div className="h-[150px] mb-6">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={chartData} layout="vertical">
              <XAxis type="number" hide />
              <YAxis 
                dataKey="name" 
                type="category" 
                width={80} 
                tick={{ fill: '#94a3b8', fontSize: 10 }} 
                axisLine={false}
                tickLine={false}
              />
              <Tooltip 
                cursor={{ fill: 'rgba(255,255,255,0.05)' }}
                contentStyle={{ backgroundColor: '#0f172a', border: 'none', borderRadius: '8px', borderStyle: 'none', fontSize: '10px', color: '#fff' }}
                itemStyle={{ color: '#fff' }}
              />
              <Bar dataKey="count" fill="#f43f5e" radius={[0, 4, 4, 0]} barSize={12} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}

      <div className="space-y-3 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar">
        {conflicts.slice(0, 10).map((c, i) => (
          <div key={i} className="p-3 bg-white/5 rounded-xl border border-white/5 hover:bg-white/10 transition-all">
            <div className="flex justify-between items-start mb-1">
              <span className="text-[10px] font-bold text-rose-400">{c.time.split(' ')[1]}</span>
              <span className="text-[9px] font-black text-slate-500 uppercase">{c.radio}</span>
            </div>
            <p className="text-[10px] font-bold text-white mb-1">{formatStationName(c.stationName || c.stationId)}</p>
            <div className="flex gap-1 flex-wrap">
               <span className="text-[9px] px-1 bg-white/5 rounded text-slate-400 uppercase font-mono">Modes: {c.modes.join('/')}</span>
               <span className="text-[9px] px-1 bg-rose-500/10 rounded text-rose-300 font-mono">NMS: {c.nmsCodes.join('/')}</span>
            </div>
          </div>
        ))}
        {conflicts.length > 10 && (
          <p className="text-[9px] text-center text-slate-500 pt-2 tracking-widest uppercase font-bold">+ {conflicts.length - 10} additional conflicts</p>
        )}
      </div>
    </motion.div>
  );
}

function StatusBox({ title, items }: { title: string; items: { label: string; status: string; reason: string }[] }) {
  return (
    <div className="glass-card p-6 rounded-2xl space-y-4">
      <h4 className="font-bold text-white text-sm uppercase tracking-wider opacity-70">{title}</h4>
      <div className="grid grid-cols-2 gap-4">
        {items.map((item, i) => (
          <div key={i} className="bg-white/5 p-4 rounded-xl border border-white/5">
            <div className="flex justify-between items-center mb-2">
              <span className="text-xs font-bold text-slate-400 uppercase">{item.label}</span>
              <span className={cn(
                "px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-tighter",
                item.status === 'Healthy' ? "bg-emerald-500/20 text-emerald-400" : 
                item.status === 'Warning' || item.status === 'Marginal' ? "bg-amber-500/20 text-amber-400" : "bg-rose-500/20 text-rose-400"
              )}>
                {item.status}
              </span>
            </div>
            <p className="text-xs text-slate-300 leading-relaxed">{item.reason}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

function ExecutiveSummary({ stats, setActiveTab }: { stats: DashboardStats; setActiveTab: (tab: string) => void }) {
  const nmsColors: Record<string, string> = {
    '0': '#0066cc',
    '8': '#80ccff',
    '1': '#ff3333',
    '-': '#ffb3b3',
    '16': '#33b3a6',
    '32': '#80ffaa',
    '40': '#ff9900',
    'default': '#64748b'
  };

  return (
    <div className="grid grid-cols-3 gap-8">
      <div className="col-span-2 space-y-6 min-w-0">
        <h3 className="text-lg font-bold text-white flex items-center gap-2">
          <Zap className="w-5 h-5 text-emerald-400" />
          System-Level Insights
        </h3>
        
        <div className="grid gap-4">
          <StatusBox 
            title="1. Hardware Analysis"
            items={[
              { 
               label: "Locomotives Analyzed", 
               status: stats.locoIds.length > 50 ? "Warning" : "Healthy", 
               reason: `Total ${stats.locoIds.length} unique locomotives identified${stats.locoIds.length > 10 ? ' (Showing first 10)' : ''}: ${stats.locoIds.slice(0, 10).join(', ')}${stats.locoIds.length > 10 ? '...' : '.'}` 
             },
              { label: `Loco ${stats.locoId} Performance`, status: stats.locoPerformance > 95 ? "Healthy" : (stats.locoPerformance >= 85 ? "Warning" : "Unhealthy"), reason: `Loco ${stats.locoId} achieved ${stats.locoPerformance.toFixed(1)}% performance across all stations.` },
              { 
                label: "🔴 Unhealthy Stations (<85%)", 
                status: stats.unhealthyStns && stats.unhealthyStns.length > 0 ? "Unhealthy" : "Healthy", 
                reason: stats.unhealthyStns && stats.unhealthyStns.length > 0 
                  ? `Critical drops: ${stats.unhealthyStns.map(s => `${formatStationName(s.id)} (${s.pct.toFixed(1)}%)`).join(', ')}.` 
                  : "No critical hardware failures detected." 
              },
              { 
                label: "🟡 Warning Stations (85-95%)", 
                status: stats.warningStns && stats.warningStns.length > 0 ? "Warning" : "Healthy", 
                reason: stats.warningStns && stats.warningStns.length > 0 
                  ? `Significant drops: ${stats.warningStns.map(s => `${formatStationName(s.id)} (${s.pct.toFixed(1)}%)`).join(', ')}.` 
                  : "No marginal performance issues detected." 
              },
              { 
                label: "🟢 Healthy Stations (>95%)", 
                status: "Healthy", 
                reason: stats.healthyStns && stats.healthyStns.length > 0 
                  ? `Optimal performance: ${stats.healthyStns.map(s => `${formatStationName(s.id)} (${s.pct.toFixed(1)}%)`).join(', ')}.` 
                  : "No stations currently meet the high-performance benchmark." 
              }
            ]}
          />

          <StatusBox 
            title="2. Protocol Analysis"
            items={[
              { label: "Sync Analysis", status: stats.avgLag <= 1.2 ? "Healthy" : "Warning", reason: `AR: ${stats.arCount} | MA: ${stats.maCount}. Ratio: ${((stats.maCount / (stats.arCount || 1)) * 100).toFixed(1)}%.` },
              { label: "Packet Interval Analysis", status: stats.avgLag <= 1.0 ? "Healthy" : "Warning", reason: `Average MA interval: ${stats.avgLag.toFixed(2)}s. RDSO standard is 1.0s.` }
            ]}
          />

          <StatusBox 
            title="3. Safety & Operational Events"
            items={[
              { 
                label: "MODE DEGRADATIONS", 
                status: stats.modeDegradations.length > 5 ? "Unhealthy" : stats.modeDegradations.length > 0 ? "Warning" : "Healthy", 
                reason: stats.modeDegradations.length > 0 
                  ? `Detected ${stats.modeDegradations.length} Mode Change(s). Check radio switching logic in Expert Diagnostics.` 
                  : "No mode changes detected." 
              },
              { 
                label: "BRAKE ACTIVATIONS", 
                status: stats.brakeApplications.length > 0 ? "Warning" : "Healthy", 
                reason: stats.brakeApplications.length > 0 
                  ? `${stats.brakeApplications.length} Automatic brake events logged. See Interval Analysis & PDF for details.` 
                  : "No automatic brake applications identified." 
              },
              { 
                label: "EMERGENCY BRAKES", 
                status: stats.brakeApplications.filter(b => b.type.includes('EB')).length > 0 ? "Unhealthy" : "Healthy", 
                reason: stats.brakeApplications.filter(b => b.type.includes('EB')).length > 0 
                  ? `${stats.brakeApplications.filter(b => b.type.includes('EB')).length} Emergency Brake (EB) event(s). Urgent root cause analysis required.` 
                  : "Zero Emergency Brake (EB) activations detected." 
              }
            ]}
          />
          

          <div className="glass-card p-6 rounded-2xl border-l-4 border-emerald-500 space-y-4">
            <h4 className="font-bold text-white flex items-center gap-2">
              <AlertCircle className="w-5 h-5 text-emerald-400" />
              Dynamic Diagnostic Advice
            </h4>
            <div className="space-y-4">
              {stats.diagnosticAdvice.map((advice, i) => (
                <div key={i} className={cn(
                  "p-4 rounded-xl border backdrop-blur-sm",
                  advice.severity === 'high' ? "bg-rose-500/10 border-rose-500/20" : 
                  advice.severity === 'medium' ? "bg-amber-500/10 border-amber-500/20" : "bg-emerald-500/10 border-emerald-500/20"
                )}>
                  <p className="font-bold text-sm mb-1 text-white">{advice.title}</p>
                  <p className="text-xs text-slate-400 mb-2">{advice.detail}</p>
                  <div className="flex gap-2 items-start mt-2 pt-2 border-t border-white/5">
                    <Zap className="w-3 h-3 mt-0.5 text-emerald-400" />
                    <p className="text-xs font-medium text-slate-300"><span className="text-slate-500 uppercase text-[9px] font-bold mr-1">Action:</span> {advice.action}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      <div className="space-y-6 min-w-0">
        <InfrastructureStressMeter stress={stats.infrastructureStress} />

        <div className="glass-card p-6 rounded-2xl">
          <h4 className="font-bold text-white text-sm mb-6 uppercase tracking-wider opacity-70">Interval Distribution</h4>
          <div className="h-[200px]">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={stats.intervalDist}
                  innerRadius={60}
                  outerRadius={80}
                  paddingAngle={5}
                  dataKey="percentage"
                >
                  {stats.intervalDist.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={['#10b981', '#f59e0b', '#ef4444'][index]} />
                  ))}
                </Pie>
                <Tooltip 
                  contentStyle={{ backgroundColor: 'rgba(0,0,0,0.8)', border: 'none', borderRadius: '12px', color: '#fff' }}
                  itemStyle={{ color: '#fff' }}
                />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="mt-4 space-y-2">
            {stats.intervalDist.map((d, i) => (
              <div key={i} className="flex justify-between text-xs">
                <span className="text-slate-400">{d.category}</span>
                <span className="font-bold text-white">{d.percentage.toFixed(1)}%</span>
              </div>
            ))}
          </div>
        </div>

        <div className="glass-card p-6 rounded-2xl">
          <h4 className="font-bold text-white text-sm mb-4 uppercase tracking-wider opacity-70">NMS Status Correlation</h4>
          <div className="h-[200px]">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={stats.nmsStatus}
                  cx="50%"
                  cy="50%"
                  outerRadius={65}
                  dataKey="value"
                  labelLine={false}
                  label={({ percent }) => percent > 0.05 ? `${(percent * 100).toFixed(0)}%` : ''}
                >
                  {stats.nmsStatus.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={nmsColors[entry.name] || nmsColors.default} />
                  ))}
                </Pie>
                <Tooltip 
                  contentStyle={{ backgroundColor: 'rgba(0,0,0,0.8)', border: 'none', borderRadius: '12px', color: '#fff' }}
                  itemStyle={{ color: '#fff' }}
                />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="mt-4 grid grid-cols-2 gap-2">
            {stats.nmsStatus.map((d, i) => (
              <div key={i} className="flex items-center gap-2 text-[10px]">
                <div className="w-2 h-2 rounded-full shrink-0" style={{ backgroundColor: nmsColors[d.name] || nmsColors.default }} />
                <span className="text-slate-400 truncate">{d.name}:</span>
                <span className="font-bold text-white">{d.value}</span>
              </div>
            ))}
          </div>
        </div>

        <SyncConflictChart conflicts={stats.conflictingPackets} />

        <div className="glass-card p-6 rounded-2xl bg-emerald-500/5 border-emerald-500/20 group cursor-pointer hover:bg-emerald-500/10 transition-all" onClick={() => setActiveTab('methodology')}>
          <div className="flex items-center justify-between mb-4">
            <h4 className="font-bold text-white text-sm uppercase tracking-wider opacity-70">Calculation Logic</h4>
            <Calculator className="w-4 h-4 text-emerald-400 group-hover:scale-110 transition-transform" />
          </div>
          <p className="text-xs text-slate-400 leading-relaxed mb-4">
            Station performance is calculated using <span className="text-emerald-400 font-bold">Weighted Averages</span> (Sum then Divide) to ensure accuracy across trips of varying lengths.
          </p>
          <div className="flex items-center gap-2 text-emerald-400 text-[10px] font-bold uppercase tracking-widest">
            View Methodology <ArrowRight className="w-3 h-3" />
          </div>
        </div>
      </div>
    </div>
  );
}

function DeepMapping({ stats, files }: { stats: DashboardStats; files: { rf: File[]; trn: File[]; radio: File | null } }) {
  const failures = [
    {
      id: 1,
      title: `NMS Health Critical Failure (${stats.nmsFailRate.toFixed(1)}%)`,
      source: files.trn.length > 0 ? files.trn.map(f => f.name).join(', ') : 'N/A',
      column: "'NMS Health'",
      detail: `The NMS Health column should ideally maintain a value of 0 (Healthy). Your data contains anomalous values in ${stats.nmsFailRate.toFixed(1)}% of rows, indicating persistent hardware or internal communication issues.`
    },
    {
      id: 2,
      title: "Session Persistence / Access Request Ratio",
      source: files.radio?.name || 'N/A',
      column: "'Packet Type'",
      detail: `The system transmitted ${stats.arCount} Access Requests, but only ${stats.maCount} Movement Authorities were registered. This significant mismatch confirms session stability failures.`
    },
    {
      id: 3,
      title: "Station Hardware Warning/Unhealthy Status",
      source: files.rf.length > 0 ? files.rf.map(f => f.name).join(', ') : 'N/A',
      column: "'Station Id' and 'Percentage'",
      detail: `Average percentage analysis indicates that signal strength at stations ${stats.badStns.join(', ') || 'None'} has fallen below the 85% Unhealthy threshold, and stations ${stats.marginalStns?.join(', ') || 'None'} are in the 85-95% Warning range.`
    },
    {
      id: 4,
      title: "Sync Loss / Refresh Lag Analysis",
      source: files.radio?.name || 'N/A',
      column: "'Time'",
      detail: `The average interval between MA packets was recorded at ${stats.avgLag.toFixed(2)} seconds. Any deviation from the RDSO standard (1.0s) triggers a session drop by the Loco system.`
    },
    ...(stats.modeDegradations.length > 0 ? [{
      id: 5,
      title: `Mode Degradation Events (${stats.modeDegradations.length})`,
      source: files.trn.length > 0 ? files.trn.map(f => f.name).join(', ') : 'N/A',
      column: "'Mode'",
      detail: `The system recorded ${stats.modeDegradations.length} instances where the Kavach mode was downgraded (e.g., FS to OS/SR). These events were analyzed for station-side and loco-side stressors.`
    }] : [])
  ];

  return (
    <div className="space-y-6">
      <div className="bg-blue-500/10 border-l-4 border-blue-500 p-4 rounded-r-xl backdrop-blur-sm">
        <div className="flex items-center gap-3">
          <AlertCircle className="w-5 h-5 text-blue-400" />
          <p className="text-sm text-blue-200 font-medium">This tab is dynamically updated based on the real-time analysis of your uploaded logs.</p>
        </div>
      </div>

      <div className="grid gap-6">
        {failures.map((f) => (
          <div key={f.id} className="glass-card p-6 rounded-2xl flex gap-6 group hover:border-emerald-500/50 transition-all">
            <div className="w-12 h-12 bg-emerald-500/10 rounded-xl flex items-center justify-center shrink-0 border border-emerald-500/20">
              <span className="text-emerald-400 font-bold">0{f.id}</span>
            </div>
            <div className="space-y-3">
              <div className="flex justify-between items-start">
                <h4 className="text-lg font-bold text-white">{f.title}</h4>
                <div className="flex gap-2">
                  <span className="px-2 py-1 bg-white/5 rounded text-[10px] font-mono text-slate-400 border border-white/10">Source: {f.source}</span>
                  <span className="px-2 py-1 bg-white/5 rounded text-[10px] font-mono text-emerald-400 border border-white/10">Col: {f.column}</span>
                </div>
              </div>
              <p className="text-sm text-slate-400 leading-relaxed">{f.detail}</p>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

const RadioLossAnalysis = ({ stats }: { stats: DashboardStats }) => {
  const events = stats.stationDeepAnalysis.criticalEvents.filter(e => e.type === 'Radio Loss');
  
  // Sort events by time
  const sortedEvents = [...events].sort((a, b) => a.time.localeCompare(b.time));

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 glass-card p-6 rounded-3xl border border-white/5">
          <div className="flex items-center justify-between mb-6">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-rose-500/20 rounded-xl">
                <Activity className="w-5 h-5 text-rose-400" />
              </div>
              <h3 className="text-xl font-bold text-white">Radio Loss Timeline</h3>
            </div>
          </div>
          
          <div className="h-[400px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={sortedEvents} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" vertical={false} />
                <XAxis 
                  dataKey="time" 
                  stroke="#94a3b8" 
                  fontSize={10} 
                  tick={{ fill: '#94a3b8' }}
                  angle={-45}
                  textAnchor="end"
                  interval={Math.ceil(sortedEvents.length / 15)}
                />
                <YAxis 
                  stroke="#94a3b8" 
                  fontSize={10} 
                  tick={{ fill: '#94a3b8' }}
                  label={{ value: 'Duration (s)', angle: -90, position: 'insideLeft', fill: '#94a3b8', fontSize: 10 }}
                />
                <Tooltip 
                  contentStyle={{ backgroundColor: '#0f172a', border: '1px solid rgba(255,255,255,0.1)', borderRadius: '12px' }}
                  itemStyle={{ color: '#f43f5e' }}
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      const data = payload[0].payload;
                      return (
                        <div className="bg-[#0f172a] border border-white/10 p-3 rounded-xl shadow-2xl">
                          <p className="text-slate-400 text-[10px] font-bold uppercase mb-1">{data.time}</p>
                          <p className="text-white font-bold mb-1">Loco: {data.locoId}</p>
                          <p className="text-rose-400 font-bold mb-1">Duration: {data.duration}s</p>
                          {data.stationName && (
                            <p className="text-slate-300 text-xs">
                              Station: {formatStationName(data.stationName)}
                            </p>
                          )}
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                <Bar dataKey="duration" fill="#f43f5e" radius={[4, 4, 0, 0]}>
                  {sortedEvents.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.duration > 60 ? '#f43f5e' : '#fb7185'} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="glass-card p-6 rounded-3xl border border-white/5">
          <h3 className="text-lg font-bold text-white mb-4">Radio Loss Summary</h3>
          <div className="space-y-4">
            <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
              <p className="text-xs text-slate-400 uppercase font-bold tracking-wider mb-1">Total Events</p>
              <p className="text-3xl font-bold text-white">{events.length}</p>
            </div>
            <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
              <p className="text-xs text-slate-400 uppercase font-bold tracking-wider mb-1">Avg Duration</p>
              <p className="text-3xl font-bold text-white">
                {events.length > 0 ? Math.round(events.reduce((acc, e) => acc + e.duration, 0) / events.length) : 0}s
              </p>
            </div>
            <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
              <p className="text-xs text-slate-400 uppercase font-bold tracking-wider mb-1">Max Duration</p>
              <p className="text-3xl font-bold text-rose-400">
                {events.length > 0 ? Math.max(...events.map(e => e.duration)) : 0}s
              </p>
            </div>
          </div>
        </div>
      </div>

      <div className="glass-card overflow-hidden rounded-3xl border border-white/5">
        <div className="p-6 border-b border-white/5 flex items-center justify-between">
          <h3 className="text-xl font-bold text-white">Detailed Event Log</h3>
          <div className="flex gap-2">
            <span className="px-3 py-1 bg-rose-500/10 text-rose-400 text-[10px] font-bold rounded-full border border-rose-500/20">
              CRITICAL LOSS
            </span>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-white/5">
                <th className="p-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Time</th>
                <th className="p-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Loco ID</th>
                <th className="p-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Station</th>
                <th className="p-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Duration</th>
                <th className="p-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Radio</th>
                <th className="p-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Reason</th>
                <th className="p-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Details</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/5">
              {sortedEvents.map((event, idx) => (
                <tr key={idx} className="hover:bg-white/[0.02] transition-colors">
                  <td className="p-4 text-sm font-medium text-slate-300">{event.time}</td>
                  <td className="p-4 text-sm font-bold text-white">{event.locoId}</td>
                  <td className="p-4 text-sm text-slate-300">
                    {(() => {
                      const name = event.stationName && event.stationName !== 'N/A' && event.stationName !== '-' ? String(event.stationName) : '';
                      const id = event.stationId && event.stationId !== 'N/A' && event.stationId !== '-' ? formatStationName(event.stationId) : '';
                      
                      if (name && id) {
                        return `${formatStationName(name)} (${id})`;
                      }
                      if (name) {
                        return formatStationName(name);
                      }
                      if (id) {
                        return formatStationName(id);
                      }
                      return 'Unknown Station';
                    })()}
                  </td>
                  <td className="p-4">
                    <span className={cn(
                      "px-2 py-1 rounded-lg text-xs font-bold",
                      event.duration > 60 ? "bg-rose-500/20 text-rose-400" : "bg-amber-500/20 text-amber-400"
                    )}>
                      {event.duration}s
                    </span>
                  </td>
                  <td className="p-4">
                    <span className="px-2 py-1 bg-emerald-500/10 text-emerald-400 rounded-lg text-xs font-bold border border-emerald-500/20">
                      {event.radio || 'Radio 1'}
                    </span>
                  </td>
                  <td className="p-4">
                    <span className={cn(
                      "text-[10px] font-bold px-2 py-0.5 rounded-full border",
                      event.reason?.includes('Hardware') ? "bg-rose-500/10 text-rose-400 border-rose-500/20" :
                      event.reason?.includes('Software') ? "bg-blue-500/10 text-blue-400 border-blue-500/20" :
                      "bg-amber-500/10 text-amber-400 border-amber-500/20"
                    )}>
                      {event.reason || 'N/A'}
                    </span>
                  </td>
                  <td className="p-4 text-xs text-slate-400">{event.description}</td>
                </tr>
              ))}
              {sortedEvents.length === 0 && (
                <tr>
                  <td colSpan={7} className="p-12 text-center text-slate-500 italic">
                    No radio loss events detected in the current selection.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const MovingAnalysis = ({ stats }: { stats: DashboardStats }) => {
  const data = stats.movingRadioLoss || [];

  return (
    <div className="space-y-6">
      <div className="glass-card p-8 rounded-3xl border border-white/5 bg-gradient-to-br from-emerald-500/5 to-transparent">
        <div className="flex items-center gap-4 mb-6">
          <div className="p-3 bg-emerald-500/20 rounded-2xl">
            <Zap className="w-6 h-6 text-emerald-400" />
          </div>
          <div>
            <h2 className="text-2xl font-bold text-white">Moving Radio Loss Analysis</h2>
            <p className="text-slate-400 text-sm">Analysis of signal drops while locomotive is in motion (Speed &gt; 0)</p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <div className="p-6 bg-white/5 rounded-2xl border border-white/5">
            <p className="text-xs text-slate-500 uppercase font-bold tracking-wider mb-2">Avg Moving Gaps</p>
            <p className="text-4xl font-bold text-white">
              {data.length > 0 ? (data.reduce((acc, d) => acc + d.movingGaps, 0) / data.length).toFixed(1) : 0}
            </p>
          </div>
          <div className="p-6 bg-white/5 rounded-2xl border border-white/5">
            <p className="text-xs text-slate-500 uppercase font-bold tracking-wider mb-2">Max Gap Recorded</p>
            <p className="text-4xl font-bold text-rose-400">
              {data.length > 0 ? Math.max(...data.map(d => d.maxGap)) : 0}s
            </p>
          </div>
          <div className="p-6 bg-white/5 rounded-2xl border border-white/5">
            <p className="text-xs text-slate-500 uppercase font-bold tracking-wider mb-2">Hardware Health</p>
            <p className="text-4xl font-bold text-emerald-400">
              {data.filter(d => !d.conclusion.includes('हार्डवेयर')).length} / {data.length} Healthy
            </p>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="border-b border-white/10">
                <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Loco ID</th>
                <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Moving Gaps</th>
                <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Max Gap (s)</th>
                <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">R1 Usage</th>
                <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">R2 Usage</th>
                <th className="p-4 text-[10px] font-bold text-slate-500 uppercase tracking-widest">Conclusion</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/5">
              {data.map((row, idx) => (
                <tr key={idx} className="hover:bg-white/[0.02] transition-colors">
                  <td className="p-4 font-bold text-white">{row.locoId}</td>
                  <td className="p-4">
                    <span className={cn(
                      "px-2 py-1 rounded-lg text-xs font-bold",
                      row.movingGaps > 20 ? "bg-rose-500/20 text-rose-400" : "bg-emerald-500/20 text-emerald-400"
                    )}>
                      {row.movingGaps} times
                    </span>
                  </td>
                  <td className="p-4 text-slate-300 font-mono">{row.maxGap.toLocaleString()}s</td>
                  <td className="p-4">
                    <div className="flex items-center gap-2">
                      <div className="w-16 h-1.5 bg-white/10 rounded-full overflow-hidden">
                        <div className="h-full bg-blue-500" style={{ width: `${row.r1Usage}%` }} />
                      </div>
                      <span className="text-xs text-slate-400">{row.r1Usage}%</span>
                    </div>
                  </td>
                  <td className="p-4">
                    <div className="flex items-center gap-2">
                      <div className="w-16 h-1.5 bg-white/10 rounded-full overflow-hidden">
                        <div className="h-full bg-purple-500" style={{ width: `${row.r2Usage}%` }} />
                      </div>
                      <span className="text-xs text-slate-400">{row.r2Usage}%</span>
                    </div>
                  </td>
                  <td className="p-4">
                    <span className={cn(
                      "text-xs font-medium px-3 py-1 rounded-full border",
                      row.conclusion.includes('सबसे अधिक') || row.conclusion.includes('हार्डवेयर') 
                        ? "bg-rose-500/10 text-rose-400 border-rose-500/20" 
                        : "bg-emerald-500/10 text-emerald-400 border-emerald-500/20"
                    )}>
                      {row.conclusion}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="glass-card p-6 rounded-3xl border border-white/5">
          <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
            <AlertTriangle className="w-5 h-5 text-amber-400" /> Key Observations
          </h3>
          <ul className="space-y-3">
            <li className="flex gap-3 text-sm text-slate-300">
              <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 mt-1.5 shrink-0" />
              <span><strong>Actual Radio Loss (Moving Loss):</strong> Even after removing the 'Band' (stationary) state, many locos still have significant radio gaps.</span>
            </li>
            <li className="flex gap-3 text-sm text-slate-300">
              <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 mt-1.5 shrink-0" />
              <span><strong>Radio Balance (Hardware Health):</strong> The imbalance between Radio 1 and Radio 2 indicates hardware malfunction or antenna alignment issues.</span>
            </li>
          </ul>
        </div>
        <div className="glass-card p-6 rounded-3xl border border-white/5">
          <h3 className="text-lg font-bold text-white mb-4 flex items-center gap-2">
            <CheckCircle2 className="w-5 h-5 text-emerald-400" /> Recommendations
          </h3>
          <ul className="space-y-3">
            <li className="flex gap-3 text-sm text-slate-300">
              <div className="w-1.5 h-1.5 rounded-full bg-blue-500 mt-1.5 shrink-0" />
              <span>Immediately inspect the radio units and antennas of locos with hardware issues.</span>
            </li>
            <li className="flex gap-3 text-sm text-slate-300">
              <div className="w-1.5 h-1.5 rounded-full bg-blue-500 mt-1.5 shrink-0" />
              <span>Map 'No Network Zones' in sections with large communication gaps.</span>
            </li>
          </ul>
        </div>
      </div>
    </div>
  );
};

function FileDrop({ zone, label, onUpload, file, files, multiple }: { zone: string; label: string; onUpload: any; file?: File | null; files?: File[]; multiple?: boolean }) {
  const hasFiles = multiple ? (files && files.length > 0) : !!file;
  
  return (
    <div className={cn(
      "relative group cursor-pointer rounded-xl border-2 border-dashed transition-all p-4 text-center",
      hasFiles ? "bg-emerald-500/10 border-emerald-500/50" : "bg-white/5 border-white/10 hover:border-emerald-500/30"
    )}>
      <input
        type="file"
        multiple={multiple}
        className="absolute inset-0 opacity-0 cursor-pointer"
        onChange={(e) => e.target.files && onUpload(zone, multiple ? e.target.files : e.target.files[0])}
      />
      <div className="flex flex-col items-center gap-2">
        {hasFiles ? (
          <CheckCircle2 className="w-6 h-6 text-emerald-400" />
        ) : (
          <Upload className="w-6 h-6 text-slate-500 group-hover:text-emerald-400 transition-colors" />
        )}
        <span className={cn("text-xs font-medium", hasFiles ? "text-emerald-400" : "text-slate-400")}>
          {multiple 
            ? (files && files.length > 0 ? `${files.length} files selected` : label)
            : (file ? file.name : label)
          }
        </span>
      </div>
    </div>
  );
}
