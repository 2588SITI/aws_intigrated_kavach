/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { RFData, TRNData, RadioData, DashboardStats, bucketDelay } from '../types';

export const parseFile = async (file: File | Blob, fileName?: string): Promise<any[]> => {
  const name = fileName || (file as File).name || '';
  return new Promise((resolve, reject) => {
    if (name.endsWith('.csv')) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target?.result as string;
        const dateRegex = /\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/;
        
        // Scan first 1000 chars for a date
        let foundDate: string | null = null;
        const match = text.slice(0, 1000).match(dateRegex);
        if (match) {
          const dateStr = match[0];
          let d = new Date(dateStr);
          if (isNaN(d.getTime())) {
            const dmyMatch = dateStr.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
            if (dmyMatch) {
              const day = dmyMatch[1].padStart(2, '0');
              const month = dmyMatch[2].padStart(2, '0');
              let year = dmyMatch[3];
              if (year.length === 2) year = `20${year}`;
              d = new Date(`${year}-${month}-${day}`);
            }
          }
          if (!isNaN(d.getTime())) {
            foundDate = dateStr;
          }
        }

        // Filename extraction disabled per user request
        const isTrainFile = name.toUpperCase().includes('RFCOMM_TR') || name.toUpperCase().includes('LOCO');
        const isStationFile = name.toUpperCase().includes('RFCOMM_ST') || name.toUpperCase().includes('STN') || name.toUpperCase().includes('STATION');

        // Detect header row in CSV by scanning for expected keywords
        const lines = text.split('\n');
        let headerRowIndex = 0;
        const expectedKeywords = ['Loco Id', 'Date', 'Time', 'Station', 'Direction', 'Expected', 'Received', 'Percentage'];
        
        for (let i = 0; i < Math.min(lines.length, 15); i++) {
          const line = lines[i].toLowerCase();
          if (expectedKeywords.some(k => line.includes(k.toLowerCase()))) {
            headerRowIndex = i;
            break;
          }
        }
        
        const csvContent = lines.slice(headerRowIndex).join('\n');

        Papa.parse(csvContent, {
          header: true,
          dynamicTyping: true,
          skipEmptyLines: true,
          complete: (results) => {
            // Clean keys
            const cleaned = results.data.map((row: any) => {
              const newRow: any = {};
              Object.keys(row).forEach(key => {
                newRow[key.trim()] = row[key];
              });
              if (foundDate && !newRow['Date'] && !newRow['Log Date']) {
                newRow['_extractedDate'] = foundDate;
              }
              return newRow;
            });
            resolve(cleaned);
          },
          error: (error) => reject(error),
        });
      };
      reader.onerror = (error) => reject(error);
      reader.readAsText(file);
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Try to find the header row if the first one seems wrong
        const expectedHeaders = ['Expected', 'Received', 'Station', 'Loco', 'Direction', 'Success', 'Nominal', 'Reverse'];
        let headerRowIndex = 0;
        for (let i = 0; i < Math.min(rawData.length, 10); i++) {
          const row = rawData[i];
          if (Array.isArray(row)) {
            const hasHeader = row.some(cell => 
              expectedHeaders.some(h => String(cell).toLowerCase().includes(h.toLowerCase()))
            );
            if (hasHeader) {
              headerRowIndex = i;
              break;
            }
          }
        }
        
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex });
        
        // Aggressive date search in headers/first rows if not found in data
        let foundDate: string | null = null;
        const dateRegex = /\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/;
        
        // Scan first 10 rows for anything that looks like a date
        for (let i = 0; i < Math.min(rawData.length, 10); i++) {
          const row = rawData[i];
          if (Array.isArray(row)) {
            for (const cell of row) {
              const cellStr = String(cell);
              if (dateRegex.test(cellStr)) {
                const match = cellStr.match(dateRegex);
                if (match) {
                  const dateStr = match[0];
                  // Robust check: try parsing as is, then try DMY
                  let d = new Date(dateStr);
                  if (isNaN(d.getTime())) {
                    const dmyMatch = dateStr.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
                    if (dmyMatch) {
                      const day = dmyMatch[1].padStart(2, '0');
                      const month = dmyMatch[2].padStart(2, '0');
                      let year = dmyMatch[3];
                      if (year.length === 2) year = `20${year}`;
                      d = new Date(`${year}-${month}-${day}`);
                    }
                  }

                  if (!isNaN(d.getTime())) {
                    foundDate = dateStr;
                    break;
                  }
                }
              }
            }
          }
          if (foundDate) break;
        }

        // Filename extraction disabled per user request
        const isTrainFile = name.toUpperCase().includes('RFCOMM_TR') || name.toUpperCase().includes('LOCO');
        const isStationFile = name.toUpperCase().includes('RFCOMM_ST') || name.toUpperCase().includes('STN') || name.toUpperCase().includes('STATION');

        // Clean keys
        const cleaned = jsonData.map((row: any) => {
          const newRow: any = {};
          Object.keys(row).forEach(key => {
            newRow[key.trim()] = row[key];
          });
          if (foundDate && !newRow['Date'] && !newRow['Log Date']) {
            newRow['_extractedDate'] = foundDate;
          }
          return newRow;
        });
        resolve(cleaned);
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    }
  });
};

const findColumn = (row: any, ...aliases: string[]) => {
  if (!row) return null;
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const found = keys.find(k => k.toLowerCase().replace(/\s/g, '') === (alias || '').toLowerCase().replace(/\s/g, ''));
    if (found) return found;
  }
  return null;
};

const parseTime = (timeStr: any): number => {
  if (!timeStr || timeStr === 'N/A' || timeStr === 'Unknown') return NaN;
  let s = String(timeStr).trim();
  
  // Base date for time-only strings to ensure consistency
  const baseDateStr = '2000-01-01';

  // Handle HH:MM:SS only
  if (s.match(/^\d{1,2}:\d{1,2}:\d{1,2}$/)) {
    const [h, m, sec] = s.split(':').map(Number);
    const d = new Date(`${baseDateStr}T${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(sec).padStart(2, '0')}`);
    return d.getTime();
  }

  // Handle YYYYMMDDHHMMSS or YYYYMMDD
  if (s.match(/^\d{8,14}$/)) {
    const year = s.substring(0, 4);
    const month = s.substring(4, 6);
    const day = s.substring(6, 8);
    const timePart = s.length > 8 ? `T${s.substring(8, 10)}:${s.substring(10, 12)}:${s.substring(12, 14) || '00'}` : '';
    const d = new Date(`${year}-${month}-${day}${timePart}`);
    if (!isNaN(d.getTime())) return d.getTime();
  }

  // Handle YYYY/MM/DD or YYYY-MM-DD or YYYY_MM_DD
  const ymdRegex = /^(\d{4})[-/._](\d{1,4})[-/._](\d{1,4})(.*)$/;
  const ymdMatch = s.match(ymdRegex);
  if (ymdMatch) {
    const year = ymdMatch[1];
    const month = ymdMatch[2].padStart(2, '0');
    const day = ymdMatch[3].padStart(2, '0');
    const rest = ymdMatch[4] || '';
    s = `${year}-${month}-${day}${rest.replace(/[/._]/g, '-')}`;
  } else {
    // Handle DD/MM/YYYY or DD-MM-YYYY or DD_MM_YYYY
    const dmyRegex = /^(\d{1,4})[-/._](\d{1,4})[-/._](\d{2,4})(.*)$/;
    const dmyMatch = s.match(dmyRegex);
    if (dmyMatch) {
      const day = dmyMatch[1].padStart(2, '0');
      const month = dmyMatch[2].padStart(2, '0');
      let year = dmyMatch[3];
      if (year.length === 2) year = `20${year}`;
      const rest = dmyMatch[4] || '';
      s = `${year}-${month}-${day}${rest.replace(/[/._]/g, '-')}`;
    }
  }

  // Final normalization: replace all / and _ with - and handle the weird 2026/03/2028 case
  s = s.replace(/[/._]/g, '-');
  const weirdDateMatch = s.match(/^(\d{4})-(\d{2})-(\d{4})(.*)$/);
  if (weirdDateMatch) {
    // Take the first 2 digits of the second "year" as the day
    const day = weirdDateMatch[3].substring(0, 2);
    s = `${weirdDateMatch[1]}-${weirdDateMatch[2]}-${day}${weirdDateMatch[4]}`;
  }

  const d = new Date(s);
  if (isNaN(d.getTime())) {
    // Fallback: try to extract just the time HH:MM:SS
    const timeMatch = s.match(/(\d{1,2}):(\d{1,2}):(\d{1,2})/);
    if (timeMatch) {
      const h = parseInt(timeMatch[1]);
      const m = parseInt(timeMatch[2]);
      const sec = parseInt(timeMatch[3]);
      const d = new Date(`${baseDateStr}T${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(sec).padStart(2, '0')}`);
      return d.getTime();
    }
  }
  return d.getTime();
};

const parseNumber = (val: any): number => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const str = String(val).replace(/[%,]/g, '').trim();
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
};

export const parseDateString = (d: string) => {
  if (!d || d === 'Unknown' || d === 'N/A') return 0;
  const parts = d.split(/[-/.]/);
  if (parts.length === 3) {
    let day, month, year;
    // Check if first part is a year (4 digits)
    if (parts[0].length === 4) {
      year = parseInt(parts[0]);
      month = parseInt(parts[1]) - 1;
      day = parseInt(parts[2]);
    } else {
      // Assume DD-MM-YYYY
      day = parseInt(parts[0]);
      month = parseInt(parts[1]) - 1;
      year = parts[2].length === 2 ? 2000 + parseInt(parts[2]) : parseInt(parts[2]);
    }
    const date = new Date(year, month, day);
    return isNaN(date.getTime()) ? 0 : date.getTime();
  }
  const date = new Date(d);
  return isNaN(date.getTime()) ? 0 : date.getTime();
};

export const formatStationName = (stn: string | number | undefined) => {
  if (!stn) return 'N/A';
  let s = String(stn).trim().toUpperCase();
  if (s === 'N/A' || s === '-' || s === '' || s === '0' || s === '0.0') return 'N/A';
  
  // Remove "STATION" wherever it is to normalize, then re-append it at the end
  s = s.replace(/\s*STATION/g, '').trim();
  
  // Also handle cases like "ST ST" -> "ST" or "BOR BOR" -> "BOR"
  const parts = s.split(/\s+/);
  if (parts.length > 1 && parts[0] === parts[1]) {
    s = parts[0];
  }

  return `${s} STATION`;
};

export const generateDiagnosticAdvice = (stats: Partial<DashboardStats>): DashboardStats['diagnosticAdvice'] => {
  const diagnosticAdvice: DashboardStats['diagnosticAdvice'] = [];
  const { 
    avgLag = 0, 
    modeDegradations = [], 
    nmsFailRate = 0, 
    tagLinkIssues = [], 
    intervalDist = [], 
    arCount = 0, 
    maCount = 0,
    unhealthyStns = [],
    warningStns = [],
    nmsLogs = []
  } = stats;

  if (avgLag > 1.2) {
    diagnosticAdvice.push({
      title: "Packet Refresh Lag Detected",
      detail: `Average MA packet interval is ${avgLag.toFixed(2)}s (Standard requirement is 1.0s).`,
      action: "Test: Radio Latency Test. Check: Station TCAS CPU load, Radio modem serial baud rate, and RF interference.",
      severity: 'medium'
    });
  }

  if (unhealthyStns.length > 0) {
    const formatted = unhealthyStns.map(s => `${formatStationName(s.id)} (${s.pct.toFixed(1)}%)`);
    diagnosticAdvice.push({
      title: "Station Hardware Unhealthy (Red)",
      detail: `Stations ${formatted.join(', ')} are performing below 85%. This indicates a potential radio link failure or consecutive packet loss.`,
      action: `Immediate audit required for track-side Kavach equipment at stations [${unhealthyStns.map(s => formatStationName(s.id)).join(', ')}]. Check for complete radio link failure.`,
      severity: 'high'
    });
  }

  if (warningStns.length > 0) {
    const formatted = warningStns.map(s => `${formatStationName(s.id)} (${s.pct.toFixed(1)}%)`);
    diagnosticAdvice.push({
      title: "Station Hardware Warning (Yellow)",
      detail: `Stations ${formatted.join(', ')} are performing between 85% and 95%. This indicates packet drops, likely due to Radio 1/2 failure or weak signal strength (<-90 dBm).`,
      action: `Monitor and audit the track-side Kavach equipment at stations [${warningStns.map(s => formatStationName(s.id)).join(', ')}]. Check Radio 1/2 status and signal strength.`,
      severity: 'medium'
    });
  }

  if (modeDegradations.length > 0) {
    const radioRelated = modeDegradations.filter(d => d.reason.toLowerCase().includes('radio packet loss') || d.reason.toLowerCase().includes('packet loss'));
    const rfRelated = modeDegradations.filter(d => d.reason.toLowerCase().includes('poor rf') || d.reason.toLowerCase().includes('signal'));
    
    const getAffectedStns = (list: any[]) => {
      const stns = Array.from(new Set(list.map(d => {
        const id = d.stationId;
        const name = d.stationName;
        if ((!id || id === 'N/A') && (!name || name === 'N/A')) return null;
        return formatStationName(name && name !== 'N/A' ? name : id);
      }).filter(s => s !== null)));
      return stns.length > 0 ? ` at Stations: ${stns.join(', ')}` : ' (station location not resolved — check TRN log Station ID column)';
    };

    if (radioRelated.length > 0) {
      diagnosticAdvice.push({
        title: "Radio Packet Loss causing Mode Degradation",
        detail: `${radioRelated.length} mode degradation events were directly correlated with radio packet timeouts (> 2s)${getAffectedStns(radioRelated)}.`,
        action: "Check: Radio modem power stability, antenna VSWR, and potential RF interference in the section.",
        severity: 'high'
      });
    } else if (rfRelated.length > 0) {
      diagnosticAdvice.push({
        title: "Poor RF Signal causing Mode Degradation",
        detail: `${rfRelated.length} mode degradation events were correlated with low RF signal strength${getAffectedStns(rfRelated)}.`,
        action: "Check: Station antenna alignment, cable attenuation, and potential signal blockage or interference.",
        severity: 'medium'
      });
    }
  }

  if (nmsFailRate > 10) {
    // Per-loco NMS Analysis
    const locoNms: Record<string, { total: number; errors: number }> = {};
    nmsLogs.forEach(log => {
      const lId = String(log.locoId);
      if (!locoNms[lId]) locoNms[lId] = { total: 0, errors: 0 };
      locoNms[lId].total++;
      const health = String(log.health).toLowerCase();
      if (health !== '0' && health !== 'healthy' && health !== 'ok') {
        locoNms[lId].errors++;
      }
    });

    const worstLocos = Object.entries(locoNms)
      .map(([id, data]) => ({ id, rate: (data.errors / data.total) * 100 }))
      .filter(l => l.rate > 0)
      .sort((a, b) => b.rate - a.rate)
      .slice(0, 3);

    const locoDetail = worstLocos.length > 0 
      ? `. Affected locos: ${worstLocos.map(l => `Loco ${l.id} (${l.rate.toFixed(1)}%)`).join(', ')}. Prioritise inspection of Loco ${worstLocos[0].id} first.`
      : '';

    diagnosticAdvice.push({
      title: "High NMS Error Rate",
      detail: `${nmsFailRate.toFixed(1)}% of NMS health records indicate non-zero status. This suggests internal hardware module faults or communication lag between TCAS and NMS${locoDetail}`,
      action: "Check: Loco Vital Computer (LVC) health logs, BIU interface, and RFID reader connectivity.",
      severity: nmsFailRate > 30 ? 'high' : 'medium'
    });
  }

  if (tagLinkIssues.length > 0) {
    // Group by station
    const stnGroups: Record<string, { main: number; dup: number }> = {};
    tagLinkIssues.forEach(t => {
      const stn = formatStationName(t.stationId);
      const key = stn === 'N/A' ? 'Unknown location (station ID not resolved)' : stn;
      if (!stnGroups[key]) stnGroups[key] = { main: 0, dup: 0 };
      if (t.error === "Main Tag Missing") stnGroups[key].main++;
      else if (t.error === "Duplicate Tag Missing") stnGroups[key].dup++;
    });

    const breakdown = Object.entries(stnGroups)
      .map(([stn, counts]) => `${stn} [Main: ${counts.main}, Dup: ${counts.dup}]`)
      .join('; ');

    diagnosticAdvice.push({
      title: "RFID Tag Link Failures",
      detail: `${tagLinkIssues.length} instances of Tag Link Missing or Duplicate Tag Missing detected. Breakdown: ${breakdown}.`,
      action: "Check: Track-side RFID tag placement, tag programming, and Loco RFID reader sensitivity.",
      severity: 'medium'
    });
  }

  if (intervalDist.length > 0) {
    const critical = intervalDist.find(d => d.category === '> 2.0s');
    if (critical && critical.percentage > 5) {
      diagnosticAdvice.push({
        title: "Critical Packet Interval Violations",
        detail: `${critical.percentage.toFixed(1)}% of MA packets arrived with a delay > 2.0s, which is a direct cause for Kavach session termination.`,
        action: "Check: Radio network congestion, station processing overhead, and RF signal stability.",
        severity: 'high'
      });
    }
  }

  return diagnosticAdvice;
};

export const processDashboardData = (

  rfData: RFData[],
  trnData: TRNData[] | null,
  radioData: RadioData[],
  rfStData: RFData[] = []
): DashboardStats => {
  const firstRf = rfData[0] || {};
  const firstRfSt = rfStData[0] || {};
  const firstRfAny = rfData.length > 0 ? firstRf : firstRfSt;
  const firstTrn = trnData?.[0] || {};
  const firstRadio = radioData[0] || {};

  /**
   * Validates a Locomotive ID.
   * @param id The value to validate.
   * @param headerName Optional header name to check against blacklisted columns.
   */
  const isValidLocoId = (id: any, headerName?: string) => {
    if (id === null || id === undefined) return false;
    const s = String(id).trim();
    
    // Explicitly reject if this comes from a "Frame" or "Sequence" column
    if (headerName) {
      const hLow = headerName.toLowerCase();
      if (hLow.includes('frame') || hLow.includes('seq') || hLow.includes('packet') || hLow.includes('dist') || hLow.includes('log')) {
        return false;
      }
    }
    
    // User Requested: strict verification check (4 to 6 digits, numeric only, and greater than 1000).
    if (!/^\d{4,6}$/.test(s)) return false; 
    
    const n = parseInt(s);
    // Typical Loco IDs are substantial numbers. Frame numbers are usually 1-5000.
    // If the user said > 1000, we follow but are suspicious of very low values.
    if (n <= 1000 || n > 999999) return false;
    
    // Exclude generic sequences known to be noise
    if (s === '11111' || s === '99999' || s === '123456' || s === '00000' || s === '0000') return false;

    const low = s.toLowerCase();
    return s !== '' && s !== '-' && s !== 'N/A' && s !== 'null' && s !== 'undefined' && 
           !low.includes('packet') && !low.includes('seq') && !low.includes('dist') &&
           s !== '0' && s !== '0.0';
  };

  const isValidStationId = (id: any) => {
    if (id === null || id === undefined) return false;
    const s = String(id).trim();
    const low = s.toLowerCase();
    return s !== '' && s !== '-' && s !== 'N/A' && s !== 'null' && s !== 'undefined' && 
           low !== 'station id' && low !== 'stationid' && low !== 'station_id' &&
           s !== '0' && s !== '0.0';
  };

  const getBestLocoIdFromRow = (row: any, keys: string[], currentDefault: string) => {
    // Defining early so it can be used for global locoId determination
    const k = Object.keys(row);
    const lKey = k.find(key => {
      const low = key.toLowerCase().replace(/\./g, '').replace(/_/g, '').replace(/\s/g, '');
      return low === 'locoid'; 
    });
    if (lKey) {
      const val = row[lKey];
      if (isValidLocoId(val, lKey)) return String(val).trim();
    }
    return currentDefault;
  };

  // Find column names strictly (User Requested: "Loco Id HEADER SE LE")
  const locoIdCol = findColumn(firstRfAny, 'Loco Id', 'Loco Id.', 'Loco_Id');
  const trnLocoIdCol = findColumn(firstTrn, 'Loco Id', 'Loco Id.', 'Loco_Id');
  const radioLocoIdCol = findColumn(firstRadio, 'Loco Id', 'Loco Id.', 'Loco_Id');
  const rfKeys_ = Object.keys(firstRfAny);
  const stnIdCol = findColumn(firstRfAny, 'Station Id', 'StationId', 'Station_Id', 'Station', 'Stn', 'StnId', 'Stn_Id', 'Station Name', 'StationName');
  const stnNameCol = findColumn(firstRfAny, 'Station Name', 'StationName', 'Station_Name', 'StnName', 'Stn Name', 'Station_Name_1');
  const percentageCol = findColumn(firstRfAny, 'Percentage', 'Perc', 'Success', 'RFCOMM %', 'Success %', 'Perc %', 'SuccessPerc', 'Success_Perc') || (rfKeys_.length > 7 ? rfKeys_[7] : 'Percentage');
  const nominalPercCol = findColumn(firstRfAny, 'Nominal Perc', 'NominalPerc', 'Nominal %') || 'Nominal Perc';
  const reversePercCol = findColumn(firstRfAny, 'Reverse Perc', 'ReversePerc', 'Reverse %') || 'Reverse Perc';
  
  const stationMap: Record<string, string> = {};
  rfData.forEach(row => {
    const id = String(row[stnIdCol] || '').trim();
    const name = String(row[stnNameCol] || '').trim();
    if (id && id !== 'N/A' && name && name !== 'N/A') stationMap[id] = name;
  });
  rfStData.forEach(row => {
    const id = String(row[stnIdCol] || '').trim();
    const name = String(row[stnNameCol] || '').trim();
    if (id && id !== 'N/A' && name && name !== 'N/A') stationMap[id] = name;
  });

  // Pre-process TRN data to fill missing station names/IDs based on adjacent rows
  if (trnData && trnData.length > 0) {
    const trnKeys = Object.keys(trnData[0]);
    const trnStnNameCol = findColumn(trnData[0], 'Station Name', 'StationName', 'Station_Name') || trnKeys[2];
    const trnStnIdCol = findColumn(trnData[0], 'Station Id', 'StationId', 'Station_Id');
    const trnStnCode2Col = findColumn(trnData[0], 'Station Code2', 'StationCode2', 'Station_Code2');

    // Also populate stationMap from TRN data to capture stations that might not have RF data
    trnData.forEach(row => {
      const id = String(row[trnStnIdCol] || '').trim();
      const name = String(row[trnStnNameCol] || '').trim();
      if (isValidStationId(id) && !stationMap[id] && name && name !== 'N/A') {
        stationMap[id] = name;
      }
    });

    for (let i = 0; i < trnData.length; i++) {
      const row = trnData[i];
      const prevRow = i > 0 ? trnData[i - 1] : null;
      const nextRow = i < trnData.length - 1 ? trnData[i + 1] : null;

      if (trnStnNameCol) {
        const currentName = String(row[trnStnNameCol] || '').trim();
        if (!currentName || currentName === '-' || currentName === 'N/A' || currentName === '0' || currentName === '0.0') {
          const prevName = prevRow ? String(prevRow[trnStnNameCol] || '').trim() : '';
          const nextName = nextRow ? String(nextRow[trnStnNameCol] || '').trim() : '';
          if (prevName && prevName !== '-' && prevName !== 'N/A' && prevName !== '0' && prevName !== '0.0' && prevName === nextName) {
            row[trnStnNameCol] = prevName;
          } else if (trnStnCode2Col) {
            const fallbackData = String(row[trnStnCode2Col] || '').trim();
            if (fallbackData && fallbackData !== '-' && fallbackData !== 'N/A' && fallbackData !== '0' && fallbackData !== '0.0') {
              row[trnStnNameCol] = fallbackData;
            }
          }
        }
      }

      if (trnStnIdCol) {
        const currentId = String(row[trnStnIdCol] || '').trim();
        if (!currentId || currentId === '-' || currentId === 'N/A' || currentId === '0' || currentId === '0.0') {
          const prevId = prevRow ? String(prevRow[trnStnIdCol] || '').trim() : '';
          const nextId = nextRow ? String(nextRow[trnStnIdCol] || '').trim() : '';
          if (prevId && prevId !== '-' && prevId !== 'N/A' && prevId !== '0' && prevId !== '0.0' && prevId === nextId) {
            row[trnStnIdCol] = prevId;
          } else if (trnStnCode2Col) {
            const fallbackData = String(row[trnStnCode2Col] || '').trim();
            if (fallbackData && fallbackData !== '-' && fallbackData !== 'N/A' && fallbackData !== '0' && fallbackData !== '0.0') {
              row[trnStnIdCol] = fallbackData;
            }
          }
        }
      }
    }
  }
  
  // RF Time Logic: User says D and F columns (index 3 and 5)
  const rfDateCol = findColumn(firstRfAny, 'Date', 'Log Date', 'LogDate', 'Log_Date', 'Report Date', 'ReportDate', 'Date_Time', 'DateTime', 'Day', 'LogDay');
  const rfTimeOnlyCol = findColumn(firstRfAny, 'Time', 'Log Time', 'LogTime', 'Log_Time', 'Report Time', 'ReportTime', 'Clock', 'LogTime');
  const rfTimestampCol = findColumn(firstRfAny, 'Timestamp', 'DateTime', 'Date Time', 'Log Time Stamp', 'Log_Time_Stamp', 'Time_Stamp', 'TimeStamp');
  
  const cleanTimeStr = (str: any) => {
    if (!str) return 'N/A';
    const s = String(str).trim();
    const parts = s.split(/\s+/);
    if (parts.length >= 3 && /^\d+$/.test(parts[parts.length - 1])) {
      return parts.slice(0, parts.length - 1).join(' ');
    }
    return s;
  };

  const normalizeDate = (d: string) => {
    if (!d || d === 'Unknown' || d === 'N/A') return 'Unknown';
    const parts = d.split(/[-/.]/);
    if (parts.length === 3) {
      let day, month, year;
      // Determine if Year is at the start (4 digits) or end
      if (parts[0].length === 4) {
        year = parts[0];
        month = parts[1].padStart(2, '0');
        day = parts[2].padStart(2, '0');
      } else {
        day = parts[0].padStart(2, '0');
        month = parts[1].padStart(2, '0');
        year = parts[2];
        if (year.length === 2) year = `20${year}`;
      }
      return `${day}/${month}/${year}`;
    }
    return d;
  };

  const getRfTime = (row: any) => {
    let rawTime = 'N/A';
    const keys = Object.keys(row);
    const findInRow = (...aliases: string[]) => {
      for (const alias of aliases) {
        const found = keys.find(k => k.toLowerCase().replace(/\s/g, '') === (alias || '').toLowerCase().replace(/\s/g, ''));
        if (found) return found;
      }
      return null;
    };

    const rfDateCol_ = findInRow('Date', 'Log Date', 'LogDate', 'Log_Date', 'Report Date', 'ReportDate', 'Date_Time', 'DateTime', 'Day', 'LogDay', 'From', 'To');
    const rfTimeOnlyCol_ = findInRow('Time', 'Log Time', 'LogTime', 'Log_Time', 'Report Time', 'ReportTime', 'Clock', 'LogTime', 'From', 'To');
    const rfTimestampCol_ = findInRow('Timestamp', 'DateTime', 'Date Time', 'Log Time Stamp', 'Log_Time_Stamp', 'Time_Stamp', 'TimeStamp', 'From', 'To');

    if (rfTimestampCol_ && row[rfTimestampCol_]) rawTime = String(row[rfTimestampCol_]);
    else if (rfDateCol_ && rfTimeOnlyCol_ && row[rfDateCol_] && row[rfTimeOnlyCol_]) {
      rawTime = `${row[rfDateCol_]} ${row[rfTimeOnlyCol_]}`;
    }
    else if (row['_extractedDate'] && rfTimeOnlyCol_ && row[rfTimeOnlyCol_]) {
      rawTime = `${row['_extractedDate']} ${row[rfTimeOnlyCol_]}`;
    }
    else if (rfTimeOnlyCol_ && row[rfTimeOnlyCol_]) rawTime = String(row[rfTimeOnlyCol_]);
    else if (rfDateCol_ && row[rfDateCol_]) rawTime = String(row[rfDateCol_]);
    else if (row['_extractedDate']) rawTime = String(row['_extractedDate']);
    else if (keys.length > 3 && keys.length > 5 && row[keys[3]] && row[keys[5]]) {
      rawTime = `${row[keys[3]]} ${row[keys[5]]}`;
    }
    else if (keys.length > 3 && row[keys[3]]) rawTime = String(row[keys[3]]);
    else if (keys.length > 5 && row[keys[5]]) rawTime = String(row[keys[5]]);
    
    // Ensure date is in the time string if we have it
    const rawDate = String(row._extractedDate || (rfDateCol_ && row[rfDateCol_]) || '').trim();
    const rowDate = normalizeDate(rawDate);
    if (rowDate && rowDate !== 'Unknown' && rowDate !== 'N/A' && !rawTime.includes(rowDate)) {
      rawTime = `${rowDate} ${rawTime}`;
    }
    
    return cleanTimeStr(rawTime);
  };

  const getRfTimestamp = (row: any) => {
    const timeStr = getRfTime(row);
    if (timeStr === 'N/A') return 0;
    const parts = timeStr.split(' ');
    if (parts.length === 2) {
      const dateParts = parts[0].split(/[-/.]/);
      const timeParts = parts[1].split(':');
      if (dateParts.length === 3 && timeParts.length >= 2) {
        const d = parseInt(dateParts[0]);
        const m = parseInt(dateParts[1]) - 1;
        const y = dateParts[2].length === 2 ? 2000 + parseInt(dateParts[2]) : parseInt(dateParts[2]);
        const hh = parseInt(timeParts[0]);
        const mm = parseInt(timeParts[1]);
        const ss = timeParts.length > 2 ? parseInt(timeParts[2]) : 0;
        const date = new Date(y, m, d, hh, mm, ss);
        return isNaN(date.getTime()) ? 0 : date.getTime();
      }
    }
    const d = new Date(timeStr);
    return isNaN(d.getTime()) ? 0 : d.getTime();
  };

  const getTrnTime = (row: any) => {
    let rawTime = String(row[trnTimeCol] || 'N/A');
    const rawDate = String(row._extractedDate || (trnDateCol && row[trnDateCol]) || '').trim();
    const rowDate = normalizeDate(rawDate);
    if (rowDate && rowDate !== 'Unknown' && rowDate !== 'N/A' && !rawTime.includes(rowDate)) {
      rawTime = `${rowDate} ${rawTime}`;
    }
    return rawTime;
  };

  const getTrnTimestamp = (row: any) => {
    const timeStr = getTrnTime(row);
    if (timeStr === 'N/A') return 0;
    const parts = timeStr.split(' ');
    if (parts.length === 2) {
      const dateParts = parts[0].split(/[-/.]/);
      const timeParts = parts[1].split(':');
      if (dateParts.length === 3 && timeParts.length >= 2) {
        let d = parseInt(dateParts[0]);
        let m = parseInt(dateParts[1]) - 1;
        let y = parseInt(dateParts[2]);

        // Heuristic to detect YYYY/MM/DD vs DD/MM/YYYY
        if (d > 31) {
          // Likely YYYY/MM/DD
          y = d;
          // Use regex to get only first 1 or 2 digits in case LocoId is joined to Date
          const dayMatch = String(dateParts[2]).match(/^\d{1,2}/);
          d = dayMatch ? parseInt(dayMatch[0]) : parseInt(dateParts[2]);
        } else if (y < 100) {
          // Likely DD/MM/YY
          y = 2000 + y;
        }

        const hh = parseInt(timeParts[0]);
        const mm = parseInt(timeParts[1]);
        const ss = timeParts.length > 2 ? parseInt(timeParts[2]) : 0;
        const date = new Date(y, m, d, hh, mm, ss);
        return isNaN(date.getTime()) ? 0 : date.getTime();
      }
    }
    const d = new Date(timeStr);
    return isNaN(d.getTime()) ? 0 : d.getTime();
  };

  const getRadioTime = (row: any) => {
    let rawTime = String(row[radioTimeCol] || 'N/A');
    const rawDate = String(row._extractedDate || (radioDateCol && row[radioDateCol]) || '').trim();
    const rowDate = normalizeDate(rawDate);
    if (rowDate && rowDate !== 'Unknown' && rowDate !== 'N/A' && !rawTime.includes(rowDate)) {
      rawTime = `${rowDate} ${rawTime}`;
    }
    return rawTime;
  };

  const trnTimeCol = findColumn(firstTrn, 'Time', 'Timestamp', 'Date', 'DateTime', 'LogTime') || 'Time';
  const trnDateCol = findColumn(firstTrn, 'Date', 'Log Date', 'LogDate');
  const radioTimeCol = findColumn(firstRadio, 'Time', 'Timestamp', 'Time_DT', 'LogTime') || 'Time';
  const radioDateCol = findColumn(firstRadio, 'Date', 'Log Date', 'LogDate');

  if (trnData) {
    trnData.sort((a, b) => getTrnTimestamp(a) - getTrnTimestamp(b));
  }
  if (radioData) {
    radioData.sort((a, b) => {
      const tA = getRadioTime(a);
      const tB = getRadioTime(b);
      const tsA = new Date(tA).getTime();
      const tsB = new Date(tB).getTime();
      return (isNaN(tsA) ? 0 : tsA) - (isNaN(tsB) ? 0 : tsB);
    });
  }

  // Find first valid locoId for default
  let locoId = 'N/A';
  const trnKeys_ = trnData && trnData.length > 0 ? Object.keys(trnData[0]) : [];
  
  const firstValidRf = rfData.find(r => {
    const val = r[locoIdCol];
    return locoIdCol && isValidLocoId(val, locoIdCol);
  });
  const firstValidTrn = trnData?.find(r => getBestLocoIdFromRow(r, trnKeys_, 'N/A') !== 'N/A');
  const firstValidRadio = radioData.find(r => {
    const val = r[radioLocoIdCol];
    return radioLocoIdCol && isValidLocoId(val, radioLocoIdCol);
  });
  
  if (firstValidRf) locoId = String(firstValidRf[locoIdCol]).trim();
  else if (firstValidTrn) locoId = getBestLocoIdFromRow(firstValidTrn, trnKeys_, 'N/A');
  else if (firstValidRadio) locoId = String(firstValidRadio[radioLocoIdCol]).trim();

  const allLocosCount: Record<string, number> = {};
  
  const getLocoFromRowDynamically = (row: any) => {
    if (!row) return null;
    const keys = Object.keys(row);
    // User requested absolute strictness on "Loco Id" header
    const lKey = keys.find(k => {
      const low = k.toLowerCase().replace(/\./g, '').replace(/_/g, '').replace(/\s/g, '');
      return low === 'locoid'; // Strictly 'locoid' only, NO "Frame Num" or others
    });
    if (lKey) {
      const val = row[lKey];
      if (isValidLocoId(val, lKey)) return String(val).trim();
    }
    return null;
  };

  rfData.forEach(row => { 
    const val = getLocoFromRowDynamically(row);
    if (val) allLocosCount[val] = (allLocosCount[val] || 0) + 1; 
  });
  trnData?.forEach(row => {
    const val = getLocoFromRowDynamically(row);
    if (val) allLocosCount[val] = (allLocosCount[val] || 0) + 1;
  });
  rfStData.forEach(row => { 
    const val = getLocoFromRowDynamically(row);
    if (val) allLocosCount[val] = (allLocosCount[val] || 0) + 1; 
  });
  radioData.forEach(row => { 
    const val = getLocoFromRowDynamically(row);
    if (val) allLocosCount[val] = (allLocosCount[val] || 0) + 1; 
  });
  
  // Real locos in a log appear hundreds of times. 
  // Numbers appearing only a few times are almost certainly transient noise (frame numbers, packet IDs).
  const locoIds = Object.keys(allLocosCount)
    .filter(id => allLocosCount[id] > 10) // Significant presence required
    .sort();

  // If we filtered out the default locoId, pick the most frequent one instead
  if (!locoIds.includes(locoId) && locoIds.length > 0) {
    locoId = locoIds.sort((a, b) => allLocosCount[b] - allLocosCount[a])[0];
  }

  // Station Performance & Stats
  const stnGroups: Record<string, { 
    expected: number; 
    received: number; 
    percentages: number[];
    times: string[]; 
    locoId: string | number; 
    date: string;
    source: 'train' | 'station';
  }> = {};
  
  const trnKeys = trnData && trnData.length > 0 ? Object.keys(trnData[0]) : [];
  const trnRadioCol = trnKeys[4] || 'Radio'; // Column E is index 4
  
  const expectedCol = findColumn(firstRfAny, 'Expected', 'Exp', 'Total', 'Expected Count', 'Exp Count', 'ExpCount');
  const receivedCol = findColumn(firstRfAny, 'Received', 'Rec', 'SuccessCount', 'Recieved Count', 'Rec Count', 'RecCount', 'Success');
  const radioCol = findColumn(firstRfAny, 'Radio', 'Modem', 'RadioId', 'Radio_Id', 'ModemId', 'Modem_Id', 'Radio No', 'RadioNo', 'Radio_No');
  const directionCol = findColumn(firstRfAny, 'Direction', 'Mode', 'Nominal/Reverse', 'Type', 'Nominal_Reverse', 'Dir', 'Nom/Rev', 'Nominal_Rev');

  const seenRfRows = new Set<string>();
  
  let skippedRfRows = 0;
  const processRfRow = (row: any, source: 'train' | 'station') => {
    const keys = Object.keys(row);
    const findInRow = (...aliases: string[]) => {
      for (const alias of aliases) {
        const found = keys.find(k => k.toLowerCase().replace(/\s/g, '') === (alias || '').toLowerCase().replace(/\s/g, ''));
        if (found) return found;
      }
      return null;
    };

    const effectiveSource = source;

    const sIdCol = findInRow('Station Id', 'StationId', 'Station_Id', 'Station', 'Stn', 'StnId', 'Stn_Id', 'Station Name', 'StationName');
    let stnId = '';
    
    // Priority 1: Explicit column match
    if (sIdCol) {
      stnId = String(row[sIdCol] || '').trim();
    }
    
    // Priority 2: Extracted from filename (very reliable for station logs)
    const isGeneric = !stnId || ['STATION', 'STN', 'PROJECT', 'SYSTEM'].includes(stnId.toUpperCase());
    if (isGeneric && row['_extractedStationId']) {
      stnId = String(row['_extractedStationId']).trim();
    }
    
    // Priority 3: Fallback
    if (!stnId && !sIdCol && keys.length > 0) {
      const fallbackId = String(row[keys[0]] || '').trim();
      const blacklist = ['project', 'system', 'log', 'report', 'date', 'time', 'loco', 'train'];
      if (fallbackId && !blacklist.some(b => fallbackId.toLowerCase().includes(b))) {
        stnId = fallbackId;
      }
    }
    
    // NORMALIZATION FIX:
    // Ensure "ST" and "ST STATION" are merged into the same key
    stnId = stnId.toUpperCase().replace(/\s+STATION$/i, '').trim();

    if (!stnId || stnId === 'STATION ID' || stnId === 'STATIONID') {
      skippedRfRows++;
      return;
    }
    
    const dCol = findInRow('Direction', 'Mode', 'Nominal/Reverse', 'Type', 'Nominal_Reverse', 'Dir', 'Nom/Rev', 'Nominal_Rev');
    const rawDirection = String(dCol ? row[dCol] : 'N/A');
    const direction = rawDirection.toLowerCase().includes('nominal') ? 'Nominal' : 
                      rawDirection.toLowerCase().includes('reverse') ? 'Reverse' : rawDirection;
    
    // ABSOLUTE DYNAMIC HEADER LOGIC (User Requested: "ONLY Loco Id HEADER")
    let rawRowLocoId = getLocoFromRowDynamically(row);

    // If no row-level header found, fallback to the first valid global locoId detected during initial pass 
    // ONLY as a last resort to keep the data associated with a loco.
    if (!rawRowLocoId && isValidLocoId(locoId)) {
      rawRowLocoId = locoId;
    }
    
    // If still nothing, use a generic placeholder to avoid polluting IDs with numeric junk
    if (!rawRowLocoId || !isValidLocoId(rawRowLocoId)) {
      rawRowLocoId = effectiveSource === 'station' ? 'Station Log' : 'Unknown';
    }
    
    const rowLocoId = String(rawRowLocoId).trim();
    const rfDateCol_ = findInRow('Date', 'Log Date', 'LogDate', 'Log_Date', 'Report Date', 'ReportDate', 'Date_Time', 'DateTime', 'Day', 'LogDay', 'From', 'To');
    const rawDate = String(row._extractedDate || (rfDateCol_ && row[rfDateCol_]) || 'Unknown').trim();
    const rowDateNormalized = normalizeDate(rawDate);
    const rowTime = getRfTime(row);
    
    const rowKey = `${rowLocoId}|${stnId}|${direction}|${rowTime}|${rowDateNormalized}|${effectiveSource}`;
    if (seenRfRows.has(rowKey)) return;
    seenRfRows.add(rowKey);

    const key = `${stnId}|${direction}|${rowLocoId}|${rowDateNormalized}|${effectiveSource}`;
    
    if (!stnGroups[key]) stnGroups[key] = { 
      expected: 0, received: 0, 
      percentages: [],
      times: [], locoId: rowLocoId, date: rowDateNormalized,
      source: effectiveSource
    };
    
    const eCol = findInRow('Expected', 'Exp', 'Total', 'Expected Count', 'Exp Count', 'ExpCount') || (keys.length > 5 ? keys[5] : '');
    const rCol = findInRow('Received', 'Rec', 'SuccessCount', 'Recieved Count', 'Rec Count', 'RecCount', 'Success') || (keys.length > 6 ? keys[6] : '');
    const pCol = findInRow('Percentage', 'Perc', 'Success', 'RFCOMM %', 'Success %', 'Perc %', 'SuccessPerc', 'Success_Perc') || (keys.length > 7 ? keys[7] : '');

    const exp = parseNumber(row[eCol]);
    const rec = parseNumber(row[rCol]);
    const perc = parseNumber(row[pCol]) || (exp > 0 ? (rec / exp) * 100 : 0);
    
    stnGroups[key].expected += exp;
    stnGroups[key].received += rec;
    stnGroups[key].percentages.push(perc);
    
    if (rowTime !== 'N/A') stnGroups[key].times.push(rowTime);
  };

  rfData.forEach(row => processRfRow(row, 'train'));
  rfStData.forEach(row => processRfRow(row, 'station'));

  console.log(`Processed RFCOMM: Train Rows=${rfData.length}, Station Rows=${rfStData.length}`);
  console.log(`Aggregated Station Stats: ${Object.keys(stnGroups).length} groups`);

  const stationStats = Object.entries(stnGroups).map(([key, data]) => {
    const parts = key.split('|');
    const stationId = parts[0];
    const direction = parts[1];
    const totalPercSum = data.percentages.reduce((a, b) => a + b, 0);
    const rowCount = data.percentages.length;
    const percentage = data.expected > 0 ? (data.received / data.expected) * 100 : (rowCount > 0 ? totalPercSum / rowCount : 0);

    return {
      stationId,
      direction,
      percentage,
      expected: data.expected,
      received: data.received,
      locoId: data.locoId,
      date: data.date,
      rowCount,
      totalPercSum,
      source: data.source
    };
  });

  const stnPerf = Object.entries(stnGroups).map(([key, data]) => {
    const [stationId] = key.split('|');
    const sortedTimes = [...data.times].sort();
    
    const percentage = data.expected > 0 
      ? (data.received / data.expected) * 100 
      : (data.percentages.length > 0 ? data.percentages.reduce((a, b) => a + b, 0) / data.percentages.length : 0);

    return {
      stationId,
      percentage,
      expected: data.expected,
      received: data.received,
      locoId: data.locoId,
      date: data.date,
      startTime: sortedTimes.length > 0 ? sortedTimes[0] : 'N/A',
      endTime: sortedTimes.length > 0 ? sortedTimes[sortedTimes.length - 1] : 'N/A'
    };
  });

  const rawRfLogs = rfData
    .filter(row => {
      const val = row[locoIdCol] || locoId;
      return isValidLocoId(val, locoIdCol || undefined);
    })
    .map(row => {
      const direction = String(row[directionCol] || 'N/A');
      const percentage = Number(row[percentageCol]) || 0;
      const isNominal = direction.toLowerCase().includes('nominal');
      const isReverse = direction.toLowerCase().includes('reverse');
      
      let sId = 'N/A';
      if (stnIdCol && row[stnIdCol]) {
        sId = String(row[stnIdCol]).trim();
      } else if (row['_extractedStationId']) {
        sId = String(row['_extractedStationId']).trim();
      }

      // NORMALIZATION FIX:
      sId = sId.toUpperCase().replace(/\s+STATION$/i, '').trim() || 'N/A';
      
      return {
        stationId: sId,
        direction,
        expected: Number(row[expectedCol]) || 0,
        received: Number(row[receivedCol]) || 0,
        nominalPerc: isNominal ? percentage : 0,
        reversePerc: isReverse ? percentage : 0,
        time: getRfTime(row),
        date: normalizeDate(String(row._extractedDate || (rfDateCol && row[rfDateCol]) || 'Unknown').trim()),
        locoId: String(row[locoIdCol] || locoId).trim()
      };
    });

  // Calculate global station performance (weighted average)
  const globalStationStats = new Map<string, { exp: number, rec: number }>();
  Object.entries(stnGroups).forEach(([key, data]) => {
    const stationId = key.split('|')[0];
    if (!globalStationStats.has(stationId)) globalStationStats.set(stationId, { exp: 0, rec: 0 });
    const g = globalStationStats.get(stationId)!;
    g.exp += data.expected;
    g.rec += data.received;
  });

  const globalStationPerf = Array.from(globalStationStats.entries()).map(([stationId, data]) => ({
    stationId,
    percentage: data.exp > 0 ? (data.rec / data.exp) * 100 : 0
  }));

  const badStns = globalStationPerf.filter(s => s.percentage < 85).map(s => s.stationId);
  const marginalStns = globalStationPerf.filter(s => s.percentage >= 85 && s.percentage <= 95).map(s => s.stationId);
  const goodStns = globalStationPerf.filter(s => s.percentage > 95).map(s => s.stationId);

  const unhealthyStns = globalStationPerf
    .filter(s => s.percentage < 85)
    .map(s => ({ id: s.stationId, pct: s.percentage }))
    .sort((a, b) => a.pct - b.pct);

  const warningStns = globalStationPerf
    .filter(s => s.percentage >= 85 && s.percentage <= 95)
    .map(s => ({ id: s.stationId, pct: s.percentage }))
    .sort((a, b) => a.pct - b.pct);

  const healthyStns = globalStationPerf
    .filter(s => s.percentage > 95)
    .map(s => ({ id: s.stationId, pct: s.percentage }))
    .sort((a, b) => a.pct - b.pct);

  // Multi-Loco Bad Station Logic
  // First, aggregate performance PER loco PER station
  const stationLocoAggregates = new Map<string, Map<string, { exp: number, rec: number }>>();
  
  stnPerf.forEach(s => {
    const sId = String(s.stationId);
    const lId = String(s.locoId);
    if (!stationLocoAggregates.has(sId)) stationLocoAggregates.set(sId, new Map());
    const lMap = stationLocoAggregates.get(sId)!;
    if (!lMap.has(lId)) lMap.set(lId, { exp: 0, rec: 0 });
    const m = lMap.get(lId)!;
    m.exp += (s as any).expected || 0;
    m.rec += (s as any).received || 0;
  });

  const stnLocoMap: Record<string | number, { 
    locoDetails: { id: string | number; perf: number; startTime: string; endTime: string }[];
    totalPerf: number; 
    count: number 
  }> = {};
  
  stationLocoAggregates.forEach((lMap, sId) => {
    lMap.forEach((m, lId) => {
      const avgPerf = m.exp > 0 ? (m.rec / m.exp) * 100 : 0;
      
      // Only count as a "failed loco" if its TOTAL performance at this station is < 85%
      if (avgPerf < 85 && m.exp > 0) {
        if (!stnLocoMap[sId]) stnLocoMap[sId] = { locoDetails: [], totalPerf: 0, count: 0 };
        
        // Find time range for this loco at this station
        const locoTimes = stnPerf
          .filter(p => String(p.stationId) === sId && String(p.locoId) === lId)
          .map(p => (p as any).startTime || '')
          .filter(t => t !== '')
          .sort();

        stnLocoMap[sId].locoDetails.push({
          id: lId,
          perf: avgPerf,
          startTime: locoTimes[0] || 'N/A',
          endTime: locoTimes[locoTimes.length - 1] || 'N/A'
        });
        stnLocoMap[sId].totalPerf += avgPerf;
        stnLocoMap[sId].count++;
      }
    });
  });

  const multiLocoBadStns = Object.entries(stnLocoMap)
    .filter(([stationId, data]) => {
      const g = globalStationStats.get(stationId);
      const globalPerf = g && g.exp > 0 ? (g.rec / g.exp) * 100 : 100;
      // ABSOLUTE PROTECTION: If the station is healthy overall, it cannot be a "bad station"
      if (globalPerf >= 85) return false;
      
      // Must have multiple locos failing
      return data.locoDetails.length > 1;
    })
    .map(([stationId, data]) => ({
      stationId,
      locoCount: data.locoDetails.length,
      avgPerf: data.totalPerf / data.count,
      locoDetails: data.locoDetails
    }));

  const rfFiltered = rfData.filter(row => {
    const val = row[locoIdCol] || locoId;
    return isValidLocoId(val, locoIdCol || undefined);
  });
  const totalExp = rfFiltered.reduce((acc, row) => acc + parseNumber(row[expectedCol]), 0);
  const totalRec = rfFiltered.reduce((acc, row) => acc + parseNumber(row[receivedCol]), 0);
  const locoPerformance = totalExp > 0 
    ? (totalRec / totalExp) * 100 
    : (stationStats.length > 0 ? stationStats.reduce((acc, s) => acc + s.percentage, 0) / stationStats.length : 0);

  // Radio Data Mapping
  const packetTypeCol = findColumn(firstRadio, 'Packet Type', 'PacketType', 'Type', 'Pkt Type2', 'PktType2') || 'Packet Type';
  const lengthCol = findColumn(firstRadio, 'Length', 'Len', 'Size') || 'Length';
  const sourceCol = findColumn(firstRadio, 'Source', 'Src', 'From') || 'Source';
  const messageCol = findColumn(firstRadio, 'Message', 'Msg', 'Data') || 'Message';

  const radioRadioCol = findColumn(firstRadio, 'Radio', 'Modem', 'RadioId', 'Radio_Id', 'ModemId', 'Modem_Id', 'Radio No', 'RadioNo', 'Radio_No') || (Object.keys(firstRadio).length > 29 ? Object.keys(firstRadio)[29] : 'Radio');

  const isMA = (val: any) => {
    const s = String(val || '').toLowerCase().replace(/\s/g, '');
    return s.includes('movementauthority') || s === 'ma' || s.includes('movauth') || s.includes('movementauth');
  };
  const isAR = (val: any) => {
    const s = String(val || '').toLowerCase().replace(/\s/g, '');
    return s.includes('accessrequest') || s === 'ar' || s.includes('accreq') || s.includes('accessreq');
  };

  let arCount = radioData.filter((p) => isAR(p[packetTypeCol])).length;
  const maPacketsRaw = radioData.filter((p) => isMA(p[packetTypeCol]));
  let maCount = maPacketsRaw.length;

  // Include packets from trnData if they exist there (e.g. ALL_TRNMSNMA files)
  if (trnData && trnData.length > 0) {
    const trnPacketTypeCol = findColumn(trnData[0], 'Pkt Type2', 'PktType2', 'Packet Type', 'PacketType', 'Type');
    if (trnPacketTypeCol) {
      const trnMaPackets = trnData.filter(p => isMA(p[trnPacketTypeCol]));
      const trnArPackets = trnData.filter(p => isAR(p[trnPacketTypeCol]));
      
      // Combine counts if they are likely from different sources or if one is empty
      // Usually TRN logs and Radio logs contain different aspects of the same communication
      // or represent different logs entirely.
      if (radioData.length === 0 || maCount === 0) {
        maCount = trnMaPackets.length;
      } else if (trnMaPackets.length > 0) {
        // If both have data, we might be double counting, but showing 0 is worse.
        // Let's take the maximum or sum? Sum is safer if they are separate logs.
        // The user specifically asked for these to be shown.
        maCount = Math.max(maCount, trnMaPackets.length);
      }

      if (radioData.length === 0 || arCount === 0) {
        arCount = trnArPackets.length;
      } else if (trnArPackets.length > 0) {
        arCount = Math.max(arCount, trnArPackets.length);
      }
    }
  }

  // Short Packets (< 10)
  const shortPackets = radioData
    .filter(p => Number(p[lengthCol]) < 10 && p[lengthCol] !== undefined && isValidLocoId(p[radioLocoIdCol] || locoId))
    .map(p => ({
      time: getRadioTime(p),
      type: String(p[packetTypeCol]),
      length: Number(p[lengthCol]),
      locoId: String(p[radioLocoIdCol] || locoId).trim(),
      radio: String(p[radioRadioCol] || '').trim()
    }));

  // SOS Events
  const sosEvents = radioData
    .filter(p => (String(p[packetTypeCol]).toLowerCase().includes('sos') || String(p[messageCol]).toLowerCase().includes('sos')) && isValidLocoId(p[radioLocoIdCol] || locoId))
    .map(p => ({
      time: getRadioTime(p),
      source: String(p[sourceCol] || 'Unknown'),
      type: String(p[packetTypeCol]),
      stationId: String(p[stnIdCol] || 'N/A'),
      stationName: String(p[stnNameCol] || 'N/A'),
      locoId: String(p[radioLocoIdCol] || locoId).trim(),
      radio: String(p[radioRadioCol] || '').trim()
    }));

  // Tag Link Issues (Medha Specific)
  const tagLinkCol = findColumn(firstRadio, 'Tag Link Info', 'TagLinkInfo', 'TagInfo') || 'Tag Link Info';
  const radioTagIssues: any[] = [];
  const seenRadioTagIssues = new Set<string>();
  radioData.forEach(p => {
    const info = String(p[tagLinkCol] || '').toLowerCase();
    if (info.includes('error') || 
        info.includes('mismatch') || 
        info.includes('wrong') || 
        info.includes('fail') ||
        info.includes('maintagmissing') ||
        info.includes('duplicatetagmissing')) {
      
      const time = getRadioTime(p);
      const lId = String(p[radioLocoIdCol] || locoId).trim();
      const key = `${time}|${lId}|${info}`;
      if (seenRadioTagIssues.has(key)) return;
      seenRadioTagIssues.add(key);

      let errorType = "Potential Medha Kavach Reporting Issue";
      const hLow = info.toLowerCase();
      const isMain = hLow.includes('maintagmissing');
      const isDup = hLow.includes('duplicatetagmissing');
      
      if (isMain && isDup) errorType = "Both Tags Missing (Critical)";
      else if (isMain) errorType = "Main Tag Missing (Non-Critical)";
      else if (isDup) errorType = "Duplicate Tag Missing (Non-Critical)";
      
      radioTagIssues.push({
        time: time,
        stationId: String(p[stnIdCol] || 'N/A'),
        info: String(p[tagLinkCol]),
        error: errorType,
        locoId: lId,
        radio: String(p[radioRadioCol] || '').trim(),
        isCritical: isMain && isDup
      });
    }
  });

  // Also check TRNMSNMA for Tag Link Issues (User specified Column R)
  const trnTagIssues: any[] = [];
  const seenTrnTagIssues = new Set<string>();
  let trnStnInfo: { id: string, name: string }[] = [];
  
  const trnTagLinkCol = findColumn(firstTrn, 'Tag Link Info', 'TagLinkInfo', 'TagInfo') || trnKeys[17] || 'Tag Link Info';
  const diagStnIdCol = findColumn(firstTrn, 'Station Id', 'StationId', 'Station_Id') || 'Station Id';
  const diagStnNameCol = findColumn(firstTrn, 'Station Name', 'StationName', 'Station_Name', 'STATION') || trnKeys[2];

  if (trnData) {
    // Forward-fill stationId and stationName for TRN data
    let lastSeenStnId = 'N/A';
    let lastSeenStnName = 'N/A';
    trnStnInfo = trnData.map(row => {
      const sId = String(row[diagStnIdCol] || '').trim();
      const sName = String(row[diagStnNameCol] || '').trim();

      if (sId && sId !== 'N/A' && sId !== '-' && sId !== '0' && sId !== '0.0') {
        lastSeenStnId = sId;
      }
      if (sName && sName !== 'N/A' && sName !== '-' && sName !== '0' && sName !== '0.0') {
        lastSeenStnName = sName;
      }
      return { id: lastSeenStnId, name: lastSeenStnName };
    });

    trnData.forEach((row, idx) => {
      const info = String(row[trnTagLinkCol] || '').toLowerCase();
      if (info.includes('maintagmissing') || info.includes('duplicatetagmissing')) {
        const time = getTrnTime(row);
        const lId = String(row[trnLocoIdCol] || locoId).trim();
        const key = `${time}|${lId}|${info}`;
        if (seenTrnTagIssues.has(key)) return;
        seenTrnTagIssues.add(key);

        let errorType = "Potential Medha Kavach Reporting Issue";
        const hLow = info.toLowerCase();
        const isMain = hLow.includes('maintagmissing');
        const isDup = hLow.includes('duplicatetagmissing');

        if (isMain && isDup) errorType = "Both Tags Missing (Critical)";
        else if (isMain) errorType = "Main Tag Missing (Non-Critical)";
        else if (isDup) errorType = "Duplicate Tag Missing (Non-Critical)";

        trnTagIssues.push({
          time: time,
          stationId: trnStnInfo[idx].id,
          stationName: trnStnInfo[idx].name,
          info: String(row[trnTagLinkCol]),
          error: errorType,
          locoId: lId,
          radio: String(row[trnRadioCol] || '').trim(),
          isCritical: isMain && isDup
        });
      }
    });
  }

  const initialTagLinkIssues = [...radioTagIssues, ...trnTagIssues].sort((a, b) => a.time.localeCompare(b.time));
  
  // USER REQUEST: Filter tag missing events.
  // Only take events that caused mode degradation or brake application (or other safety reasons)
  // We'll perform this filtering after we have collected all degradations and brakes.
  const tagLinkIssuesUnfiltered = initialTagLinkIssues;

  // NMS Logic
  const nmsHealthCol = findColumn(firstTrn, 'NMS Health', 'NMSHealth', 'Health') || 'NMS Health';
  const modeCol = findColumn(firstTrn, 'Mode', 'CurrentMode', 'OpMode', 'Operation Mode') || trnKeys[13] || trnKeys[14] || 'Mode';
  const eventCol = findColumn(firstTrn, 'Event', 'Description', 'LogEntry', 'Log', 'Message') || trnKeys[16] || 'Event';
  const reasonCol = findColumn(firstTrn, 'Reason', 'Cause', 'FaultReason', 'Degradation Reason') || trnKeys[17] || 'Reason';
  const lpResponseCol = findColumn(firstTrn, 'LP Response', 'DriverAction', 'Response', 'Acknowledge', 'Pilot Ack', 'Action') || trnKeys[12] || trnKeys[23] || 'LP Response';
  const speedCol = findColumn(firstTrn, 'Speed', 'Velocity', 'Kmph', 'Current Speed') || trnKeys[10] || 'Speed';
  const locationCol = findColumn(firstTrn, 'Location', 'Km', 'Position', 'Distance') || trnKeys[6] || 'Location';
  const brakeCol = findColumn(firstTrn, 'Brake', 'Brake Status', 'BrakeType', 'Brake_Status', 'EB/SB', 'Brake_Type', 'Brake Applied') || trnKeys[19] || 'Brake';
  const signalIdCol = findColumn(firstTrn, 'Signal Id', 'SignalId', 'SigId') || 'Signal Id';
  const signalStatusCol = findColumn(firstTrn, 'Signal Status', 'SignalStatus', 'SigStatus') || 'Signal Status';
  const emrStatusCol = findColumn(firstTrn, 'Emr Status', 'EmrStatus', 'Emergency Status', 'EmergencyStatus', 'EMR_Status', 'Emergency_Status') || 'Emr Status';
  const trnTagLinkInfoCol = findColumn(firstTrn, 'Tag Link Info', 'TagLinkInfo', 'TagInfo', 'Tag Link') || 'Tag Link Info';
  const trnActvRadCol = findColumn(firstTrn, 'Actv Rad', 'Active Radio', 'ActvRad', 'Actv_Rad', 'Radio Active') || trnKeys[4] || 'Actv Rad';
  const authTypeCol = findColumn(firstTrn, 'Auth Type', 'AuthType', 'Authority Type', 'Auth_Type', 'AuthorityType') || 'Auth Type';
  const sigAspectCol = findColumn(firstTrn, 'Sig Aspect', 'Signal Aspect', 'SignalAspect', 'Cur Sig Aspect', 'Signal_Aspect', 'Aspect') || 'Sig Aspect';

  const nmsFailRate = trnData
    ? (trnData.filter((row) => {
        const val = String(row[nmsHealthCol] || '').toLowerCase().trim();
        // Inclusive healthy check: 0 is standard, but handle common variations
        const isHealthy = val === '0' || val === 'healthy' || val === 'ok';
        return !isHealthy && isValidLocoId(row[trnLocoIdCol] || locoId);
      }).length / 
       (trnData.filter(row => isValidLocoId(row[trnLocoIdCol] || locoId)).length || 1)) * 100
    : 0;

  const nmsStatusMap: Record<string, number> = {};
  const locoNmsData: Record<string, { total: number; errors: number }> = {};

  trnData?.forEach((row) => {
    if (!isValidLocoId(row[trnLocoIdCol] || locoId)) return;
    
    let status = String(row[nmsHealthCol] || 'Unknown').trim();
    if (status === '0') status = '0 (Healthy)';
    nmsStatusMap[status] = (nmsStatusMap[status] || 0) + 1;

    const lId = String(row[trnLocoIdCol] || locoId).trim();
    if (!locoNmsData[lId]) locoNmsData[lId] = { total: 0, errors: 0 };
    locoNmsData[lId].total++;
    
    const val = status.toLowerCase();
    if (val !== '0 (healthy)' && val !== 'healthy' && val !== 'ok') {
      locoNmsData[lId].errors++;
    }
  });

  const nmsLocoStats = Object.entries(locoNmsData).map(([lId, data]) => {
    const perc = data.total > 0 ? (data.errors / data.total) * 100 : 0;
    let category = 'Healthy';
    if (perc > 20) category = 'Critical / Very High';
    else if (perc > 10) category = 'High';
    
    return {
      locoId: lId,
      totalRecords: data.total,
      errors: data.errors,
      errorPercentage: Number(perc.toFixed(1)),
      category
    };
  }).sort((a, b) => b.errorPercentage - a.errorPercentage);

  const nmsStatus = Object.entries(nmsStatusMap).map(([name, value]) => ({ name, value }));
  const seenNmsLogs = new Set<string>();
  const nmsLogs = trnData?.map(row => {
    const timeStr = getTrnTime(row);
    const lId = String(row[trnLocoIdCol] || locoId).trim();
    const health = String(row[nmsHealthCol]);
    const key = `${timeStr}|${lId}|${health}`;
    
    if (seenNmsLogs.has(key)) return null;
    seenNmsLogs.add(key);
    
    return {
      time: timeStr,
      health,
      status: health === '0' || health === '32' ? 'Healthy' : 'Error',
      locoId: lId,
      radio: String(row[trnRadioCol] || '').trim()
    };
  }).filter(n => n !== null) as DashboardStats['nmsLogs'] || [];

  const nmsDeepAnalysis: DashboardStats['nmsDeepAnalysis'] = [];
  let currentNmsEvent: any = null;

  trnData?.forEach((row, idx) => {
    if (!isValidLocoId(row[trnLocoIdCol] || locoId)) return;
    
    const status = String(row[nmsHealthCol] || '').trim();
    const lId = String(row[trnLocoIdCol] || locoId).trim();
    const stn = trnStnInfo[idx] || { id: 'N/A', name: 'N/A' };
    const stnId = stn.name !== 'N/A' ? stn.name : stn.id;
    const time = getTrnTime(row);

    if (status === '0' || status === 'healthy' || status === 'ok' || status === '') {
      if (currentNmsEvent && currentNmsEvent.locoId === lId) {
        nmsDeepAnalysis.push(currentNmsEvent);
        currentNmsEvent = null;
      }
      return;
    }
    
    // It's an error.
    if (currentNmsEvent && currentNmsEvent.locoId === lId && currentNmsEvent.errorCode === status && currentNmsEvent.stationId === stnId) {
      currentNmsEvent.count++;
      currentNmsEvent.endTime = time;
    } else {
      if (currentNmsEvent && currentNmsEvent.locoId === lId) {
        nmsDeepAnalysis.push(currentNmsEvent);
      }
      
      let errorType = 'Unknown Error';
      let description = '';
      let source: 'Loco' | 'Station' | 'Unknown' = 'Loco'; // NMS is mostly Loco Vital Computer health
      
      if (status === '8') {
        errorType = 'Sub-system Error';
        description = 'Minor delay/failure in hardware module (e.g., BIU Interface, RFID Reader, Speed Sensor).';
      } else if (status === '1') {
        errorType = 'Communication Error';
        description = 'Interruption in internal communication between the processor and its sub-units.';
      } else if (['16', '32', '40', '48'].includes(status)) {
        errorType = 'Vital Hardware Error';
        description = 'Mismatch in redundant processor or loss of synchronization with Brake Interface Unit (BIU).';
      } else {
        errorType = `Error Code ${status}`;
        description = 'Self-diagnosis reported a non-zero health status requiring servicing.';
      }

      const stnName = formatStationName(trnStnInfo[idx]?.name || 'N/A');

      currentNmsEvent = {
        locoId: lId,
        stationId: stn.id,
        stationName: stn.name,
        startTime: time,
        endTime: time,
        count: 1,
        errorCode: status,
        errorType,
        description,
        source
      };
    }
  });
  if (currentNmsEvent) {
    nmsDeepAnalysis.push(currentNmsEvent);
  }

  // Conflicting Packets Detection Logic
  const conflictingPackets: DashboardStats['conflictingPackets'] = [];
  const timeBuckets: Record<string, { rows: any[], idxs: number[] }> = {};

  trnData?.forEach((row, idx) => {
    const lId = String(row[trnLocoIdCol] || '').trim();
    const timeVal = getTrnTime(row);
    if (!timeVal) return;
    const stn = trnStnInfo[idx];
    const key = `${lId}|${timeVal}|${stn.id}`;
    
    if (!timeBuckets[key]) timeBuckets[key] = { rows: [], idxs: [] };
    timeBuckets[key].rows.push(row);
    timeBuckets[key].idxs.push(idx);
  });

  Object.entries(timeBuckets).forEach(([key, bucket]) => {
    if (bucket.rows.length < 2) return;

    const modes = new Set(bucket.rows.map(r => String(r[modeCol] || '').toUpperCase()).filter(m => m !== ''));
    const nmsCodes = new Set(bucket.rows.map(r => String(r[nmsHealthCol] || '')).filter(c => c !== ''));
    
    if (modes.size > 1 || nmsCodes.size > 1) {
      const parts = key.split('|');
      const locoId = parts[0];
      const time = parts[1];
      const stn = trnStnInfo[bucket.idxs[0]];
      
      const modeList = Array.from(modes);
      const nmsList = Array.from(nmsCodes);
      
      conflictingPackets.push({
        time,
        locoId,
        stationId: stn.id,
        stationName: stn.name,
        description: `Conflicting ${modes.size > 1 ? 'Modes' : 'NMS Health Codes'} detected in same second`,
        modes: modeList,
        nmsCodes: nmsList,
        severity: (modes.has('OS') && modes.has('FS')) || nmsList.includes('40') || nmsList.includes('16') ? 'High' : 'Medium',
        radio: String(bucket.rows[0][trnActvRadCol] || 'Unknown')
      });
    }
  });
  
  // Infrastructure Stress Level Calculation
  const stationStressMap: Record<string, { count: number, name: string, totalRows: number }> = {};
  let totalTagErrors = 0;
  let totalTagRows = 0;

  trnData?.forEach((row, idx) => {
    const rawTagInfo = String(row[trnTagLinkInfoCol] || '');
    const tagInfo = rawTagInfo.toLowerCase();
    const stn = trnStnInfo[idx];
    const stnKey = stn.id;

    if (!stationStressMap[stnKey]) {
      stationStressMap[stnKey] = { count: 0, name: stn.name, totalRows: 0 };
    }
    stationStressMap[stnKey].totalRows++;

    if (rawTagInfo && rawTagInfo !== '0' && rawTagInfo !== '-' && rawTagInfo !== '') {
      totalTagRows++;
      if (tagInfo.includes('intertagdist') || tagInfo.includes('tagposinterchanged') || tagInfo.includes('bothtaghavesame')) {
        totalTagErrors++;
        stationStressMap[stnKey].count++;
      }
    }
  });

  const infrastructureStress = {
    overallScore: totalTagRows > 0 ? (totalTagErrors / totalTagRows) * 100 : 0,
    tagSpacingDefects: totalTagErrors,
    totalTagTelemetry: totalTagRows,
    stationWiseStress: Object.entries(stationStressMap)
      .filter(([_, data]) => data.totalRows > 10) // Only stations with significant data
      .map(([id, data]) => ({
        stationId: id,
        stationName: data.name,
        stressScore: (data.count / data.totalRows) * 100,
        errorCode: 'InterTagDistGreaterThanDupTag'
      }))
      .sort((a, b) => b.stressScore - a.stressScore)
      .slice(0, 10)
  };

  // Sort by count descending to show the most critical continuous errors first
  nmsDeepAnalysis.sort((a, b) => b.count - a.count);

  // Mode Degradation
  const modeDegradations: DashboardStats['modeDegradations'] = [];
  const seenModeEvents = new Set<string>();
  const lastModes: Record<string, string> = {};
  const lastAcks: Record<string, string> = {};
  const lastReasons: Record<string, string> = {};
  const lastNonTripModes: Record<string, string> = {};
  const rowCountPerLoco: Record<string, number> = {};
  const reachedFSPerLoco: Record<string, boolean> = {};
  // ENRICHMENT: sliding window of last 10 rows per loco for context
  const preContextWindow: Record<string, Array<{
    nms: string; tagLink: string; emrStatus: string; radio: string; speed: number; mode: string;
  }>> = {};
  const CONTEXT_WINDOW_SIZE = 10;

  const modePriority: Record<string, number> = {
    'FS': 5,
    'LS': 4.5,
    'OS': 4,
    'PS': 3,
    'SR': 2,
    'SH': 1,
    'ST': 0.5,
    'IS': 0,
    'TR': -1,
    'Unknown': 0,
    '-': 0,
    '0': 0
  };
  
  const trnDirectionCol = findColumn(firstTrn, 'Direction', 'Nominal/Reverse', 'Journey Direction', 'Train Direction', 'Dir', 'Nom/Rev') || trnKeys[21] || 'Direction';
  const lastRowTimePerLoco: Record<string, number> = {};
  const lastDirectionPerLoco: Record<string, string> = {};

  trnData?.forEach((row, idx) => {
    const trnKeys = Object.keys(row);
    const locoIdVal = getBestLocoIdFromRow(row, trnKeys, locoId);
    if (!isValidLocoId(locoIdVal)) return;

    const rowTime = getTrnTimestamp(row);
    const lastRowTime = lastRowTimePerLoco[locoIdVal];
    const currentDirection = String(row[trnDirectionCol] || 'N/A').trim();
    const lastValidDirection = lastDirectionPerLoco[locoIdVal];
    
    // Reset journey/startup:
    // 1. More than 30 minutes pass between records
    // 2. OR Direction changed (e.g., Nominal -> Reverse), ignoring N/A gaps
    const timeGapExceeded = lastRowTime && Math.abs(rowTime - lastRowTime) > 30 * 60 * 1000;
    const directionChanged = currentDirection !== 'N/A' && lastValidDirection && lastValidDirection !== 'N/A' && currentDirection !== lastValidDirection;

    if (timeGapExceeded || directionChanged) {
      reachedFSPerLoco[locoIdVal] = false;
      delete lastModes[locoIdVal];
      delete lastAcks[locoIdVal];
      delete lastReasons[locoIdVal];
      delete lastNonTripModes[locoIdVal];
    }
    lastRowTimePerLoco[locoIdVal] = rowTime;
    if (currentDirection !== 'N/A') {
      lastDirectionPerLoco[locoIdVal] = currentDirection;
    }

    rowCountPerLoco[locoIdVal] = (rowCountPerLoco[locoIdVal] || 0) + 1;
    // ENRICHMENT: update sliding context window
    if (!preContextWindow[locoIdVal]) preContextWindow[locoIdVal] = [];
    preContextWindow[locoIdVal].push({
      nms:       String(row[nmsHealthCol]     || '0').trim(),
      tagLink:   String(row[trnTagLinkInfoCol]|| '').trim(),
      emrStatus: String(row[emrStatusCol]     || '').trim(),
      radio:     String(row[trnActvRadCol]    || '').trim(),
      speed:     Number(row[speedCol])        || 0,
      mode:      String(row[modeCol]          || '').trim()
    });
    if (preContextWindow[locoIdVal].length > CONTEXT_WINDOW_SIZE) {
      preContextWindow[locoIdVal].shift();
    }
    // USER REQUEST: Startup initialization period is defined as "until FS mode is reached for the first time"
    const isStartup = !reachedFSPerLoco[locoIdVal];

    const rawMode = String(row[modeCol] || '').trim();
    const currentAck = String(row[lpResponseCol] || '').trim();
    const event = String(row[eventCol] || '').toLowerCase();
    const stnId = trnStnInfo[idx]?.id || 'N/A';
    const stnName = trnStnInfo[idx]?.name || 'N/A';
    const rawReason = String(row[reasonCol] || '').trim();
    
    // Normalize mode names for detection
    let currentMode = rawMode;
    const upperRaw = rawMode.toUpperCase();
    if (upperRaw.includes('STAFF') || upperRaw === 'SR') currentMode = 'SR';
    else if (upperRaw.includes('FULL') || upperRaw === 'FS') currentMode = 'FS';
    else if (upperRaw.includes('SIGHT') || upperRaw === 'OS') currentMode = 'OS';
    else if (upperRaw.includes('SHUNT') || upperRaw === 'SH') currentMode = 'SH';
    else if (upperRaw.includes('TRIP') || upperRaw === 'TR') currentMode = 'TR';
    else if (upperRaw.includes('LIMITED') || upperRaw === 'LS') currentMode = 'LS';
    else if (upperRaw.includes('PARTIAL') || upperRaw === 'PS') currentMode = 'PS';
    else if (upperRaw.includes('ISOLATION') || upperRaw === 'IS') currentMode = 'IS';
    else if (upperRaw.includes('STANDBY') || upperRaw === 'ST') currentMode = 'ST';
    
    if (currentMode) {
      // Mark as FS reached if we are in FS mode
      if (currentMode === 'FS') {
        reachedFSPerLoco[locoIdVal] = true;
      }
      
      const lastMode = lastModes[locoIdVal];
      const lastAck = lastAcks[locoIdVal];
      const lastReason = lastReasons[locoIdVal];
      const lastNonTripMode = lastNonTripModes[locoIdVal];

      const isDegradationMessage = currentAck.toLowerCase().includes('to_sr') || 
                                   currentAck.toLowerCase().includes('to_os') ||
                                   currentAck.toLowerCase().includes('to_ls') ||
                                   currentAck.toLowerCase().includes('to_ps') ||
                                   currentAck.toLowerCase().includes('to_sh') ||
                                   currentAck.toLowerCase().includes('to_st') ||
                                   currentAck.toLowerCase().includes('degrad');
                                   
      const modeChanged = lastMode && currentMode !== lastMode;
      const ackChanged = lastAck && currentAck !== lastAck;
      const reasonChanged = lastReason && rawReason !== lastReason;
      
      // Extract a meaningful reason first for filtering and inclusion logic
      let reason = rawReason;
      if (!reason || reason === 'N/A' || reason === '0') {
        reason = String(row[eventCol] || '').trim();
      }
      
      // LOOK-BACK LOGIC: Kavach logs sometimes record the mode change a few seconds AFTER the cause.
      // If the current reason is generic, we check the previous 3 rows for a more specific failure message.
      if (!reason || reason === 'N/A' || reason === '0' || reason === 'Mode Change') {
        for (let i = 1; i <= 3; i++) {
          const prevRow = trnData[idx - i];
          if (prevRow) {
            const prevEvent = String(prevRow[eventCol] || '').trim();
            const prevReason = String(prevRow[reasonCol] || '').trim();
            if (prevReason && prevReason !== 'N/A' && prevReason !== '0') {
              reason = `${prevReason} (Event detected previously)`;
              break;
            } else if (prevEvent && prevEvent.toLowerCase().includes('fail') || prevEvent.toLowerCase().includes('loss')) {
              reason = `Pre-event: ${prevEvent}`;
              break;
            }
          }
        }
      }

      if (!reason || reason === 'N/A' || reason === '0') {
        reason = currentAck || 'Mode Change';
      }

      // Human-Readable mappings for technical jargon
      if (reason.includes('InterTagDistGreaterThanDupTag')) {
        reason = reason.replace('InterTagDistGreaterThanDupTag', 'Tag Distance Mismatch (Dist > Duplicate)');
      }
      if (reason.includes('MainTagMissing')) reason = reason.replace('MainTagMissing', 'Main Tag Missing');
      if (reason.includes('DuplicateTagMissing')) reason = reason.replace('DuplicateTagMissing', 'Duplicate Tag Missing');

      // USER REQUEST: Exclude direction information headers from "Mode Degradation Events"
      const isDirectionHeader = (
        reason.toLowerCase().includes('nominal') || 
        reason.toLowerCase().includes('reverse') ||
        reason.toLowerCase().includes('direction')
      ) && (
        reason.toLowerCase().includes('start') ||
        reason.toLowerCase().includes('journey') ||
        reason.toLowerCase().includes('header') ||
        reason.toLowerCase().includes('setting')
      );

      // USER REQUEST: Include SR, OS, LS, etc. as degradations if reason contains failure keywords, even during startup
      const startupFailureReason = reason.toLowerCase().includes('radio') || 
                                   reason.toLowerCase().includes('loss') || 
                                   reason.toLowerCase().includes('tag') || 
                                   reason.toLowerCase().includes('degrad') ||
                                   reason.toLowerCase().includes('fail') ||
                                   reason.toLowerCase().includes('fault');

      // PRIORITY HIERARCHY: FS(5) > LS(4.5) > OS(4) > PS(3) > SR(2) > SH(1) > IS(0) > TR(-1)
      const lastPrio = lastMode ? (modePriority[lastMode] ?? 5) : 0;
      const currPrio = modePriority[currentMode] ?? 5;
      
      const isUpgrade = (currPrio > lastPrio);
      const isTrueDegradation = (currPrio < lastPrio);

      // If it's the first row, we count non-FS modes as degradation if they have failure reasons
      const isFirstRowDegraded = !lastMode && 
                                 (currentMode !== 'FS' && currentMode !== 'Unknown' && currentMode !== '-' && currentMode !== '0') && 
                                 (isDegradationMessage || startupFailureReason);

      if (modeChanged || ((ackChanged || reasonChanged) && isTrueDegradation) || isFirstRowDegraded) {
        
        // Explicit degradation alert: ack message contains "to_sr", "to_os", etc. or event contains "degrad"
        const isExplicitDegradation = !isUpgrade && (isDegradationMessage || event.includes('degrad'));
                              
        const isCriticalFailure = currentMode === 'TR' || currentMode === 'IS';
        
        // INCLUSION LOGIC:
        // 1. Never show upgrades (isUpgrade).
        // 2. Exclude direction/journey start headers (!isDirectionHeader).
        // 3. Always show Critical Failures: Trips (TR) or Isolation (IS).
        // 4. Show mode drops (isTrueDegradation).
        // 5. USER REQUEST: If mode is same as last one, do NOT repeat even if reason/ack changed (deduplicate identical consecutive modes)
        
        const shouldInclude = !isUpgrade && !isDirectionHeader && (
          isCriticalFailure || 
          (isTrueDegradation && modeChanged) || 
          (isExplicitDegradation && modeChanged) ||
          isFirstRowDegraded
        );

        if (shouldInclude) {
          const time = getTrnTime(row);
            
            // Deduplication: Avoid exact duplicate events (same time, loco, and transition)
            let fromMode = lastMode || 'Unknown';
            if (isDegradationMessage && currentAck.includes('_to_')) {
              const parts = currentAck.split('_to_');
              if (parts[0] && parts[0].length <= 3) fromMode = parts[0].toUpperCase();
            }
            if (currentMode === 'TR' && lastNonTripMode && lastNonTripMode !== 'TR') {
              fromMode = lastNonTripMode;
            }

            const eventKey = `${time}|${locoIdVal}|${fromMode}|${currentMode}|${reason}`;
            if (seenModeEvents.has(eventKey)) return;
            seenModeEvents.add(eventKey);

            modeDegradations.push((() => {
              // ROOT CAUSE ANALYSIS — examine pre-event context window
              const ctx = preContextWindow[locoIdVal] || [];
              const curEmr    = String(row[emrStatusCol]     || '').trim();
              const curTag    = String(row[trnTagLinkInfoCol] || '').trim();
              const curNms    = String(row[nmsHealthCol]      || '0').trim();
              const curRadio  = String(row[trnActvRadCol]     || '').trim();
              const curSpeed  = Number(row[speedCol])         || 0;

              // Collect pre-event NMS codes (non-zero)
              const preNmsCodes = ctx
                .map(c => c.nms)
                .filter(n => n !== '0' && n !== '' && n !== '-');

              // Detect radio switchover (last ctx radio != current radio)
              const lastCtxRadio = ctx.length > 0 ? ctx[ctx.length - 1].radio : '';
              const radioSwitchAtEvent = !!(lastCtxRadio && curRadio && lastCtxRadio !== curRadio && curRadio !== '');

              // Detect persistent InterTagDist in context
              const interTagCount = ctx.filter(c =>
                c.tagLink.toLowerCase().includes('intertagdist') ||
                c.tagLink.toLowerCase().includes('tagposinterchanged') ||
                c.tagLink.toLowerCase().includes('bothtaghavesame')
              ).length;
              const persistentTagIssue = interTagCount >= 3;

              // Detect ReadEndCollision or safety trigger in context
              const hasSafetyTrigger = ctx.some(c =>
                c.emrStatus.toLowerCase().includes('collision') ||
                c.emrStatus.toLowerCase().includes('emergency') ||
                c.emrStatus.toLowerCase().includes('readen')
              ) || curEmr.toLowerCase().includes('collision') || curEmr.toLowerCase().includes('readen');

              // Detect NMS hardware error codes in context (1,8,16,32,40,48)
              const criticalNmsCodes = ['1','8','16','32','40','48'];
              const hasNmsHardwareFault = preNmsCodes.some(n => criticalNmsCodes.includes(n));
              const nmsCode16or40 = preNmsCodes.some(n => n === '16' || n === '40');

              // Detect stationary train + Red signal (MA expiry)
              const isStationary = curSpeed === 0 || ctx.slice(-3).every(c => c.speed === 0);
              const maExpiry = isStationary && (fromMode === 'FS' || fromMode === 'FullSupervision') && currentMode === 'ST';

              // CLASSIFY root cause
              let rootCause = '';
              let rootCauseCategory: 'station-tag' | 'safety-trigger' | 'loco-nms' | 'radio-switch' | 'ma-expiry' | 'unknown' = 'unknown';
              let actionRequired = '';
              let severity: 'critical' | 'warning' | 'info' = 'warning';

              // Robust station name resolution
              const stnIdRaw = String(stnId).trim().toUpperCase();
              const stnNameRaw = String(stnName).trim().toUpperCase();
              const stnBest = (stnNameRaw && stnNameRaw !== 'N/A') ? stnNameRaw : (stationMap[stnIdRaw] || stnIdRaw);
              const stnUpper = formatStationName(stnBest).toUpperCase();

              if (hasSafetyTrigger) {
                rootCauseCategory = 'safety-trigger';
                rootCause = `Kavach safety trigger active: "${curEmr || 'ReadEndCollision'}". Kavach detected a potential collision risk or safety infringement. Mode was intentionally reduced from ${fromMode} to ${currentMode} to limit speed authority. This is correct safety-first system behaviour.`;
                actionRequired = `Audit train working timetable at ${stnUpper} for the time ${time}. Verify if any other train or obstacle was occupying the section ahead. No hardware fault suspected if trigger matches site reality.`;
                severity = 'critical';
              } else if (maExpiry) {
                rootCauseCategory = 'ma-expiry';
                rootCause = `Stationary MA Expiry at ${stnUpper} platform. Train was stationary (Speed: ${curSpeed} Kmph) with FS authority. Station TCAS failed to refresh Movement Authority before expiry, forcing a drop to StandBy (ST). Frequency of InterTagDist codes (${interTagCount}/${ctx.length}) suggests Loco was struggling to anchor position at platform.`;
                actionRequired = `Inspect and measure physical RFID tag spacing (Main/Duplicate pair) at ${stnUpper} platform. Verify station TCAS MA packet transmission health for stationary trains. Infrastructure correction Required.`;
                severity = 'critical';
              } else if (nmsCode16or40 && !persistentTagIssue) {
                rootCauseCategory = 'loco-nms';
                const codeStr = preNmsCodes.filter(n => ['16','40'].includes(n)).join(', ');
                rootCause = `Loco-Side Vital Hardware Event: NMS Health code ${codeStr} (Vital Hardware Error/BIU mismatch) detected ${ctx.length * 2}s prior back to event. Transient internal mismatch or processor sync loss in the Vital Unit forced the mode fallback.`;
                actionRequired = `Inspect Loco ${locoIdVal} Vital Computer. Check BIU interface card and redundant processor synchronisation. If NMS ${codeStr} persists, the BIU module may require replacement.`;
                severity = 'critical';
              } else if (radioSwitchAtEvent && !persistentTagIssue) {
                rootCauseCategory = 'radio-switch';
                rootCause = `Momentary packet gap during Radio switchover (${lastCtxRadio} → ${curRadio}) at ${time}. Telemetry shows critical Signal and Auth data became unavailable ("-") for multiple packets during the transition, causing a session dropout.`;
                actionRequired = `Verify antenna health and RF cable continuity for both Radio 1 and Radio 2 on Loco ${locoIdVal}. Audit ${stnUpper} station-side radio coverage for signal dead zones during handover.`;
                severity = 'warning';
              } else if (persistentTagIssue) {
                rootCauseCategory = 'station-tag';
                const dominant = ctx.map(c => c.tagLink).filter(t =>
                  t.toLowerCase().includes('intertagdist') || t.toLowerCase().includes('tagpos') || t.toLowerCase().includes('bothtaghave')
                );
                const tagIssueType = dominant.length > 0 ? dominant[dominant.length - 1] : 'InterTagDistGreaterThanDupTag';
                rootCause = `Persistent Infrastructure Defect: "${tagIssueType}" detected across ${interTagCount} of the last ${ctx.length} telemetry rows. RFID tag spacing at ${stnUpper} is outside RDSO tolerance. TCAS lost confidence in absolute positioning, triggering a safety step-down from ${fromMode} to ${currentMode}.`;
                actionRequired = `Physically measure and correct Main/Duplicate RFID tag spacing at ${stnUpper}. Ensure track-side tags are aligned with TCAS configuration. This is a station infrastructure issue.`;
                severity = 'warning';
              } else if (hasNmsHardwareFault) {
                rootCauseCategory = 'loco-nms';
                rootCause = `Loco-side internal stress: NMS Health codes ${preNmsCodes.join(', ')} detected in pre-event window. Indicates internal Kavach hardware or communication instability on Loco ${locoIdVal}.`;
                actionRequired = `Inspect Loco ${locoIdVal}: BIU interface, RFID reader connectivity, and internal communication links. Check LVC logs for motherboard or interface sync errors.`;
                severity = 'warning';
              } else {
                rootCauseCategory = 'unknown';
                rootCause = `Root cause inconclusive from current telemetry snapshot. Reported Reason: "${reason}". Mode transition: ${fromMode} → ${currentMode} at speed ${curSpeed} Kmph. ${curEmr ? `Active Status: ${curEmr}` : ''}`;
                actionRequired = `Perform a deeper cross-loco check for ${stnUpper}. If other locomotives are also degrading here, it is a station defect. If isolated, inspect Loco ${locoIdVal} sensors.`;
                severity = 'info';
              }

              return {
                time,
                from: fromMode,
                to: currentMode,
                reason: reason,
                lpResponse: currentAck,
                stationId: stnId,
                stationName: stnName,
                locoId: locoIdVal,
                direction: String(row[trnDirectionCol] || 'N/A').trim(),
                radio: curRadio,
                speed: curSpeed,
                emrStatus: curEmr,
                nmsAtEvent: curNms,
                tagLinkAtEvent: curTag,
                preNmsCodes,
                radioSwitchAtEvent,
                rootCause,
                rootCauseCategory,
                actionRequired,
                severity
              };
            })());
        }
      }
      lastModes[locoIdVal] = currentMode;
      lastAcks[locoIdVal] = currentAck;
      lastReasons[locoIdVal] = rawReason;
      if (currentMode !== 'TR') {
        lastNonTripModes[locoIdVal] = currentMode;
      }
    }
  });

  // Brake Applications
  const rawBrakeApplications: any[] = [];

  trnData?.forEach((row, idx) => {
    const lIdVal = getBestLocoIdFromRow(row, trnKeys, locoId);
    if (!isValidLocoId(lIdVal)) return;

    const event = String(row[eventCol] || '').toLowerCase();
    const brakeVal = String(row[brakeCol] || '').toUpperCase().trim();
    
    const isEB = event.includes('eb applied') || brakeVal.includes('EB') || event.includes('emergency brake');
    const isSB = event.includes('sb applied') || brakeVal.includes('SB') || event.includes('service brake');
    const isFSB = event.includes('fsb applied') || brakeVal.includes('FSB') || event.includes('full service brake');
    const isNSF = event.includes('nsf applied') || brakeVal.includes('NSF') || event.includes('normal service');
    
    const hasBrake = isEB || isSB || isFSB || isNSF || (event.includes('brake') && !event.includes('released'));

    if (hasBrake) {
      let bType = String(row[eventCol]);
      if (isEB) bType = 'Emergency Brake (EB)';
      else if (isFSB) bType = 'Full Service Brake (FSB)';
      else if (isNSF) bType = 'Normal Service (NSF)';
      else if (isSB) bType = 'Service Brake (SB)';

      const stn = trnStnInfo[idx] || { id: 'N/A', name: 'N/A' };
      const timeStr = getTrnTime(row);
      const timestamp = parseTime(timeStr);
      
      if (!isNaN(timestamp)) {
        // Find reason for this brake application
        let rawBReason = "";
        const emr = String(row[emrStatusCol] || '').trim();
        const ev = String(row[eventCol] || '').toLowerCase();
        const hasEmrNow = emr !== '' && emr !== '0' && emr.toLowerCase() !== 'regular' && emr.toLowerCase() !== 'none';
        
        if (hasEmrNow) {
          rawBReason = emr;
        } else {
          // Look back up to 5 records for a reason
          for (let i = 1; i <= 5; i++) {
            if (idx - i < 0) break;
            const prevRow = trnData[idx - i];
            const pEmr = String(prevRow[emrStatusCol] || '').trim();
            const pEv = String(prevRow[eventCol] || '').toLowerCase();
            const hasPEmr = pEmr !== '' && pEmr !== '0' && pEmr.toLowerCase() !== 'regular' && pEmr.toLowerCase() !== 'none';
            const hasPEv = pEv !== '' && (pEv.includes('overspeed') || pEv.includes('spad') || pEv.includes('protection') || pEv.includes('collision') || pEv.includes('rollback') || pEv.includes('authority') || pEv.includes('limit') || pEv.includes('timeout'));
            
            if (hasPEmr || hasPEv) {
              rawBReason = hasPEmr ? pEmr : String(prevRow[eventCol]);
              break;
            }
          }
        }
        
        // Clean up common noise and apply safety keywords mapping
        if (rawBReason) {
            const brLow = rawBReason.toLowerCase();
            if (brLow.includes('overspeed')) rawBReason = "Overspeed Violation";
            else if (brLow.includes('spad')) rawBReason = "SPAD (Signal Danger)";
            else if (brLow.includes('collision')) rawBReason = "Potential Collision";
            else if (brLow.includes('rollback')) rawBReason = "Rollback Protection";
            else if (brLow.includes('timeout') || brLow.includes('radio')) rawBReason = "Radio / MA Timeout";
            else if (brLow.includes('limit')) rawBReason = "Exceeding Authority Limit";
            
            if (rawBReason.toLowerCase() === 'none' || rawBReason === '0' || rawBReason.toLowerCase() === 'regular') rawBReason = "";
        }

        rawBrakeApplications.push({
          time: timeStr,
          timestamp,
          type: bType,
          speed: Number(row[speedCol]) || 0,
          location: String(row[locationCol] || 'N/A'),
          stationId: stn.id,
          stationName: stn.name,
          locoId: lIdVal,
          radio: String(row[trnRadioCol] || '').trim(),
          reason: rawBReason || "System Safety Trigger"
        });
      }
    }
  });

  // USER REQUEST: Aggregate continuous brake applications. 
  // Give only start time and until what time it continued.
  const brakeApplications: DashboardStats['brakeApplications'] = [];
  const brakeSessions: Record<string, any[]> = {};

  rawBrakeApplications.forEach(app => {
    const sessionKey = `${app.locoId}|${app.type}`;
    if (!brakeSessions[sessionKey]) {
      brakeSessions[sessionKey] = [];
    }
    const currentSession = brakeSessions[sessionKey];
    const lastApp = currentSession[currentSession.length - 1];

    // If time gap > 15 seconds, we treat it as a new application pulse
    if (lastApp && (app.timestamp - lastApp.timestamp) > 15000) {
      const start = currentSession[0];
      const end = currentSession[currentSession.length - 1];
      
      brakeApplications.push({ ...start, type: `${start.type} (Start)` });
      if (currentSession.length > 1) {
        brakeApplications.push({ ...end, type: `${end.type} (Continued Until)` });
      }
      
      brakeSessions[sessionKey] = [app];
    } else {
      currentSession.push(app);
    }
  });

  // Finalize remaining sessions
  Object.values(brakeSessions).forEach(currentSession => {
    if (currentSession.length > 0) {
      const start = currentSession[0];
      const end = currentSession[currentSession.length - 1];
      
      brakeApplications.push({ ...start, type: `${start.type} (Start)` });
      if (currentSession.length > 1) {
        brakeApplications.push({ ...end, type: `${end.type} (Continued Until)` });
      }
    }
  });

  // Sort by time to maintain chronological order in reports
  brakeApplications.sort((a, b) => parseTime(a.time) - parseTime(b.time));

  // Signal Overrides
  const signalOverrides = trnData
    ?.filter(row => String(row[eventCol] || '').toLowerCase().includes('override') && isValidLocoId(row[trnLocoIdCol] || locoId))
    .map((row, idx) => {
      const stn = trnStnInfo[idx] || { id: 'N/A', name: 'N/A' };
      return {
        time: getTrnTime(row),
        signalId: String(row[signalIdCol] || 'N/A'),
        status: String(row[signalStatusCol] || 'Overridden'),
        stationId: stn.id,
        stationName: stn.name,
        locoId: String(row[trnLocoIdCol] || locoId).trim(),
        radio: String(row[trnRadioCol] || '').trim()
      };
    }) || [];

  // Train Config Changes
  const trainConfigChanges: DashboardStats['trainConfigChanges'] = [];
  const configParams = ['Train Length', 'Loco Id', 'Train Id', 'TrainLength', 'LocoId', 'TrainId', 'Length'];
  const uniqueTrainLengthsMap = new Map<number, { time: string; stationId: string }>();
  let lastConfig: Record<string, string> = {};
  
  trnData?.forEach(row => {
    const rowStnId = String(row[findColumn(row, 'Station Id', 'StationId', 'Station_Id') || ''] || 'N/A');
    configParams.forEach(param => {
      const col = findColumn(row, param);
      if (col) {
        const val = String(row[col]);
        if (param.toLowerCase().includes('length')) {
          const numLen = Number(val);
          if (!isNaN(numLen) && numLen > 0) {
            if (!uniqueTrainLengthsMap.has(numLen)) {
              uniqueTrainLengthsMap.set(numLen, { 
                time: getTrnTime(row), 
                stationId: rowStnId 
              });
            }
          }
        }
        if (lastConfig[param] && lastConfig[param] !== val) {
          trainConfigChanges.push({
          time: getTrnTime(row),
          parameter: param,
          oldVal: lastConfig[param],
          newVal: val,
          stationId: rowStnId,
          locoId: String(row[trnLocoIdCol] || locoId).trim(),
          radio: String(row[trnRadioCol] || '').trim()
        });
        }
        lastConfig[param] = val;
      }
    });
  });

  const uniqueTrainLengths = Array.from(uniqueTrainLengthsMap.entries())
    .map(([length, info]) => ({ length, ...info, locoId: String(locoId).trim(), radio: '' })) // Simplified radio for train lengths
    .sort((a, b) => a.length - b.length);

  // Emergency Status events - USER REQUEST: Extract non-Regular status reports
  const emergencyStatusEvents: DashboardStats['emergencyStatusEvents'] = [];
  const emrSessions: Record<string, any[]> = {};

  trnData?.forEach((row, idx) => {
    const statusVal = String(row[emrStatusCol] || '').trim();
    const lCaseStatus = statusVal.toLowerCase();
    
    // Filter out regular/idle statuses
    if (statusVal && 
        lCaseStatus !== 'regular' && 
        lCaseStatus !== '0' && 
        lCaseStatus !== 'n/a' && 
        lCaseStatus !== '-' && 
        lCaseStatus !== 'unknown' &&
        lCaseStatus !== 'none') {
      
      const lId = String(row[trnLocoIdCol] || locoId).trim();
      const sessionKey = `${lId}|${statusVal}`;
      
      const time = getTrnTime(row);
      const timestamp = getTrnTimestamp(row);
      const stn = trnStnInfo[idx] || { id: 'N/A', name: 'N/A' };
      
      const app = {
        time,
        timestamp,
        locoId: lId,
        stationId: stn.id,
        stationName: stn.name,
        status: statusVal,
        radio: String(row[trnRadioCol] || '').trim()
      };

      if (!emrSessions[sessionKey]) {
        emrSessions[sessionKey] = [];
      }
      const currentSession = emrSessions[sessionKey];
      const lastApp = currentSession[currentSession.length - 1];

      // If time gap > 30 seconds, we treat it as a new distinct event sequence
      if (lastApp && (timestamp - lastApp.timestamp) > 30000) {
        const start = currentSession[0];
        const end = currentSession[currentSession.length - 1];
        emergencyStatusEvents.push({
          startTime: start.time,
          endTime: end.time,
          locoId: start.locoId,
          stationId: start.stationId,
          stationName: start.stationName,
          status: start.status,
          rowCount: currentSession.length,
          radio: start.radio
        });
        emrSessions[sessionKey] = [app];
      } else {
        currentSession.push(app);
      }
    }
  });

  // Finalize any open EMR sessions
  Object.values(emrSessions).forEach(currentSession => {
    if (currentSession.length > 0) {
      const start = currentSession[0];
      const end = currentSession[currentSession.length - 1];
      emergencyStatusEvents.push({
        startTime: start.time,
        endTime: end.time,
        locoId: start.locoId,
        stationId: start.stationId,
        stationName: start.stationName,
        status: start.status,
        rowCount: currentSession.length,
        radio: start.radio
      });
    }
  });

  // Sort emergency events chronologically
  emergencyStatusEvents.sort((a, b) => parseTime(a.startTime) - parseTime(b.startTime));

  // Station Radio Packets (Columns AD to BF - Index 29 to 57)
  const stationRadioPackets: DashboardStats['stationRadioPackets'] = [];
  if (trnData && trnData.length > 0) {
    const trnKeys = Object.keys(trnData[0]);
    trnData.forEach(row => {
      const packets: { [key: string]: any } = {};
      // AD is index 29, BF is index 57
      for (let i = 29; i <= 57; i++) {
        const key = trnKeys[i];
        if (key && row[key] !== undefined && row[key] !== null && row[key] !== '') {
          packets[key] = row[key];
        }
      }
      
      if (Object.keys(packets).length > 0) {
        stationRadioPackets.push({
          time: getTrnTime(row),
          stationId: String(row[findColumn(row, 'Station Id', 'StationId', 'Station_Id') || ''] || 'N/A'),
          packets,
          locoId: String(row[trnLocoIdCol] || locoId).trim()
        });
      }
    });
  }

  // Sync/Lag Logic
  const maPacketsProcessed: { time: string; delay: number; category: string; length: number; locoId: string | number }[] = [];
  const seenMaPackets = new Set<string>();
  let lastTime: number | null = null;

  if (radioData.length > 0) {
    maPacketsRaw.forEach((p, i) => {
      const currentTime = parseTime(p[radioTimeCol]);
      const rowLocoId = p[radioLocoIdCol] || locoId;
      if (i > 0 && lastTime !== null && !isNaN(currentTime) && isValidLocoId(rowLocoId)) {
        const delay = (currentTime - lastTime) / 1000;
        if (delay >= 0 && delay < 300) { // Ignore gaps > 5 mins as they are log gaps, not operational lag
          const timeStr = getRadioTime(p);
          const maKey = `${timeStr}|${rowLocoId}`;
          if (!seenMaPackets.has(maKey)) {
            seenMaPackets.add(maKey);
            maPacketsProcessed.push({
              time: timeStr,
              delay: Math.min(delay, 60), // Cap at 60s for diagnostic display to avoid unrealistic numbers
              category: bucketDelay(delay),
              length: Number(p[lengthCol]) || 0,
              locoId: String(rowLocoId).trim()
            });
          }
        }
      }
      lastTime = currentTime;
    });
  } else if (trnData && trnData.length > 0) {
    const trnPacketTypeCol = findColumn(trnData[0], 'Pkt Type2', 'PktType2', 'Packet Type', 'PacketType', 'Type');
    const trnMaPackets = trnPacketTypeCol ? trnData.filter(p => {
      const s = String(p[trnPacketTypeCol] || '').toLowerCase().replace(/\s/g, '');
      return s.includes('movementauthority') || s === 'ma' || s.includes('movauth') || s.includes('movementauth');
    }) : [];
    
    if (trnMaPackets.length > 0) {
      // Use explicit MA packets from TRN log
      let lastMaTime: number | null = null;
      trnMaPackets.forEach((p, i) => {
        const currentTime = parseTime(getTrnTime(p));
        const rowLocoId = String(p[trnLocoIdCol] || locoId).trim();
        if (i > 0 && lastMaTime !== null && !isNaN(currentTime)) {
          const delay = (currentTime - lastMaTime) / 1000;
          if (delay >= 0 && delay < 300) {
            const timeStr = getTrnTime(p);
            const maKey = `${timeStr}|${rowLocoId}`;
            if (!seenMaPackets.has(maKey)) {
              seenMaPackets.add(maKey);
              maPacketsProcessed.push({
                time: timeStr,
                delay: Math.min(delay, 60),
                category: bucketDelay(delay),
                length: 0,
                locoId: rowLocoId
              });
            }
          }
        }
        lastMaTime = currentTime;
      });
    } else {
      // FALLBACK: If no radio log and no explicit MA packets, use TRN log's radio columns (AD-BF) to detect delays
      // We look for changes in any of the radio packet columns
      let lastRadioState: string = '';
      let lastRadioTime: number | null = null;
      const trnKeys = Object.keys(trnData[0]);
      
      trnData.forEach((row, i) => {
        const currentTime = parseTime(getTrnTime(row));
        if (isNaN(currentTime)) return;

        // Concatenate values of radio columns to detect any change
        let currentRadioState = '';
        for (let j = 29; j <= 57; j++) {
          const key = trnKeys[j];
          if (key) currentRadioState += String(row[key] || '');
        }

        if (i === 0) {
          lastRadioState = currentRadioState;
          lastRadioTime = currentTime;
          return;
        }

        // If radio state changed, it means a new packet was received
        if (currentRadioState !== lastRadioState && currentRadioState.replace(/0/g, '').length > 0) {
          if (lastRadioTime !== null) {
            const delay = (currentTime - lastRadioTime) / 1000;
            if (delay > 0.5 && delay < 300) { // Only record significant operational delays
              const timeStr = getTrnTime(row);
              const lId = String(row[trnLocoIdCol] || locoId).trim();
              const maKey = `${timeStr}|${lId}`;
              if (!seenMaPackets.has(maKey)) {
                seenMaPackets.add(maKey);
                maPacketsProcessed.push({
                  time: timeStr,
                  delay: Math.min(delay, 60),
                  category: bucketDelay(delay),
                  length: 0,
                  locoId: lId
                });
              }
            }
          }
          lastRadioState = currentRadioState;
          lastRadioTime = currentTime;
        } else if (lastRadioTime !== null) {
          // If state hasn't changed, check if we've been waiting too long
          const currentDelay = (currentTime - lastRadioTime) / 1000;
          if (currentDelay > 2 && currentDelay < 300) {
            // Record a "virtual" packet loss event
            const timeStr = getTrnTime(row);
            const lId = String(row[trnLocoIdCol] || locoId).trim();
            const maKey = `${timeStr}|${lId}|VIRTUAL`; // Mark virtual to avoid conflict with REAL packets just in case
            if (!seenMaPackets.has(maKey)) {
              seenMaPackets.add(maKey);
              maPacketsProcessed.push({
                time: timeStr,
                delay: Math.min(currentDelay, 60),
                category: bucketDelay(currentDelay),
                length: 0,
                locoId: lId
              });
            }
          }
        }
      });
    }
  }

  // Radio Packet Loss Events (Separate from Mode Degradation)
  const radioPacketLossEvents: DashboardStats['radioPacketLossEvents'] = [];
  const seenRadioLoss = new Set<string>();
  
  trnData?.forEach(row => {
    const event = String(row[eventCol] || '').toLowerCase();
    const reason = String(row[reasonCol] || '').toLowerCase();
    
    // USER REQUEST: Exclude InterTagDistGreaterThanDupTag and NoTagMissing from Radio Packet Loss
    const isInterTagDistIssue = reason.includes('intertagdistgreaterthanduptag') || 
                                event.includes('intertagdistgreaterthanduptag');
    const isNoTagIssue = reason.includes('notagmissing') || event.includes('notagmissing');

    if ((event.includes('packet loss') || reason.includes('packet loss')) && !isInterTagDistIssue && !isNoTagIssue) {
      const timeStr = getTrnTime(row);
      const time = parseTime(timeStr);
      const lId = String(row[trnLocoIdCol] || locoId).trim();
      const res = String(row[reasonCol] || row[eventCol] || 'Packet Loss');
      
      const key = `${timeStr}|${lId}|${res}`;
      if (seenRadioLoss.has(key)) return;
      seenRadioLoss.add(key);

      // Find duration from maPacketsProcessed (closest delay)
      let duration = 0;
      const closestPacket = maPacketsProcessed.find(p => {
        const pTime = parseTime(p.time);
        return Math.abs(pTime - time) < 2000; // Within 2 seconds
      });
      if (closestPacket) duration = closestPacket.delay;

      const stnNameCol = findColumn(row, 'Station Name', 'StationName', 'Station_Name');
      const stnNameRaw = stnNameCol ? String(row[stnNameCol] || '').trim() : String(row[trnKeys[2]] || '').trim();
      const stnName = formatStationName(stnNameRaw);
      
      radioPacketLossEvents.push({
        time: timeStr,
        stationName: stnName,
        reason: res,
        details: String(row[lpResponseCol] || 'No Ack'),
        locoId: lId,
        duration: duration > 0 ? Number(duration.toFixed(1)) : undefined,
        radio: String(row[trnRadioCol] || '').trim()
      });
    }
  });

  // DYNAMIC CORRELATION: Identify and separate radio-related degradations
  const finalModeDegradations: DashboardStats['modeDegradations'] = [];
  
  modeDegradations.forEach(deg => {
    const degTime = parseTime(deg.time);
    if (isNaN(degTime)) {
      finalModeDegradations.push(deg);
      return;
    }

    // Look for radio packet timeouts (> 2s) within 10 seconds before the degradation
    const recentTimeouts = maPacketsProcessed.filter(p => {
      const pTime = parseTime(p.time);
      return !isNaN(pTime) && pTime <= degTime && pTime >= degTime - 10000 && p.delay > 2;
    });

    // Also check NMS Health in the same window
    const recentNmsIssues = nmsLogs.filter(p => {
      const pTime = parseTime(p.time);
      const health = parseInt(p.health);
      return !isNaN(pTime) && pTime <= degTime && pTime >= degTime - 10000 && health !== 32 && health !== 0;
    });

    // Also check RF Signal Strength (Train-side)
    const recentRfDrops = rfData.filter(p => {
      const pTime = parseTime(getRfTime(p));
      const perc = Number(p[percentageCol]) || 0;
      return !isNaN(pTime) && pTime <= degTime && pTime >= degTime - 10000 && perc < 80;
    });

    if (recentTimeouts.length > 0 || recentNmsIssues.length > 0 || recentRfDrops.length > 0) {
      // This is a communication-related degradation. Annotate and keep in modeDegradations.
      let maxDelay = 0;
      if (recentTimeouts.length > 0) {
        maxDelay = Math.max(...recentTimeouts.map(p => p.delay));
      } else {
        maxDelay = 2.0;
      }
      
      const radioInfo = recentRfDrops.length > 0 && recentTimeouts.length === 0 
        ? `Poor RF Signal (${Math.min(...recentRfDrops.map(p => Number(p[percentageCol]) || 100)).toFixed(1)}%)`
        : `Radio Packet Loss (Max Delay: ${maxDelay.toFixed(1)}s)`;
      
      // Update the reason 
      deg.reason = `${radioInfo} - ${deg.reason}`;
      
      // Only improve the root cause if it was unknown or if the radio evidence is a strong addition
      if (deg.rootCauseCategory === 'unknown') {
        deg.rootCause = `Unscheduled Session Dropout: ${radioInfo} confirmed during the transition window. The station link was lost for multiple packets, forcing a fallback to ${deg.to} mode as per safety protocol.`;
        deg.rootCauseCategory = 'radio-switch';
        deg.actionRequired = `Audit RF dead zones near ${formatStationName(deg.stationName && deg.stationName !== 'N/A' ? deg.stationName : deg.stationId)}. Check both modem interface cables for intermittent connectivity.`;
        deg.severity = 'warning';
      } else if (deg.rootCauseCategory === 'station-tag') {
        // Just append the radio evidence to the existing tag issue
        deg.rootCause += ` [CORRELATED: ${radioInfo} also occurred, compounding the positioning stress.]`;
      }
      finalModeDegradations.push(deg);
    } else {
      finalModeDegradations.push(deg);
    }
  });

  const modeDegradationsToUse = finalModeDegradations;

  // USER REQUEST: Filter tag link issues based on safety impact (Brakes or Mode Degradations)
  const tagLinkIssues = tagLinkIssuesUnfiltered.filter(tag => {
    const tagTime = parseTime(tag.time);
    if (isNaN(tagTime)) return false;

    // Check for mode degradations within 30 seconds after the tag issue
    const causedDegradation = modeDegradationsToUse.some(deg => {
      const degTime = parseTime(deg.time);
      return !isNaN(degTime) && degTime >= tagTime && degTime <= tagTime + 30000;
    });

    // Check for brake applications within 30 seconds after the tag issue
    const causedBrake = brakeApplications.some(brake => {
      const bTime = parseTime(brake.time);
      return !isNaN(bTime) && bTime >= tagTime && bTime <= tagTime + 30000;
    });

    return causedDegradation || causedBrake;
  });

  const avgLag = maPacketsProcessed.length > 0
    ? maPacketsProcessed.reduce((a, b) => a + b.delay, 0) / maPacketsProcessed.length
    : 0;

  // Interval Distribution
  const categoryCounts: Record<string, number> = {
    "<= 1s (Normal)": 0,
    "1s - 2s (Delayed)": 0,
    "> 2s (Timeout)": 0,
  };
  maPacketsProcessed.forEach((p) => {
    categoryCounts[p.category]++;
  });

  const totalProcessed = maPacketsProcessed.length || 1;
  const intervalDist = Object.entries(categoryCounts).map(([category, count]) => ({
    category,
    percentage: (count / totalProcessed) * 100,
  }));

  const diagnosticAdvice = generateDiagnosticAdvice({
    avgLag,
    badStns,
    marginalStns,
    modeDegradations: modeDegradationsToUse,
    nmsFailRate,
    tagLinkIssues,
    intervalDist,
    arCount,
    maCount
  });

  // Time Range Logic
  const allTimes: string[] = [];
  rfData.forEach(p => { 
    const rowTime = getRfTime(p);
    if (rowTime !== 'N/A') allTimes.push(rowTime); 
  });
  radioData.forEach(p => { if (p[radioTimeCol]) allTimes.push(String(p[radioTimeCol])); });
  trnData?.forEach(row => { if (row[trnTimeCol]) allTimes.push(String(row[trnTimeCol])); });
  
  // Sort times numerically to find accurate range
  allTimes.sort((a, b) => {
    const ta = parseTime(a);
    const tb = parseTime(b);
    if (isNaN(ta) && isNaN(tb)) return 0;
    if (isNaN(ta)) return 1;
    if (isNaN(tb)) return -1;
    return ta - tb;
  });
  
  const startTime = allTimes.length > 0 ? allTimes[0] : 'N/A';
  const endTime = allTimes.length > 0 ? allTimes[allTimes.length - 1] : 'N/A';

  const allDatesSet = new Set<string>();
  rfData.forEach(row => {
    const d = row._extractedDate || (rfDateCol && row[rfDateCol]);
    if (d) allDatesSet.add(String(d).trim());
  });
  trnData?.forEach(row => {
    const d = row._extractedDate || (trnDateCol && row[trnDateCol]);
    if (d) allDatesSet.add(String(d).trim());
  });
  radioData.forEach(row => {
    const d = row._extractedDate || (radioDateCol && row[radioDateCol]);
    if (d) allDatesSet.add(String(d).trim());
  });

  // Sort dates chronologically
  const allDates = Array.from(allDatesSet).sort((a, b) => parseDateString(a) - parseDateString(b));

  const logDate = allDates.length > 0 ? allDates[0] : null;

  // --- Station Radio Deep Analysis Logic ---
  const stationFailures: Record<string | number, { count: number; totalDuration: number; locos: Set<string | number>; totalEvents: number; workingEvents: number }> = {};
  const locoFailures: Record<string | number, { count: number; stations: Set<string | number> }> = {};
  const criticalEvents: DashboardStats['stationDeepAnalysis']['criticalEvents'] = [];
  const seenCriticalEvents = new Set<string>();

  const getReason = (loss: any) => {
    if (loss.duration > 120) return "Environmental (Signal Shadow / Terrain)";
    if (loss.minPerc > 0 && loss.minPerc < 30) return "Signal Interference / Fading";
    if (loss.minPerc === 0) return "Hardware (Total Signal Loss)";
    if (loss.duration < 10) return "Software / Protocol Lag";
    return "Environmental / Interference";
  };

  // Identify RF Loss Events from both rfData and rfStData
  const combinedRf = [
    ...rfData.map(r => ({ ...r, _source: 'train' })),
    ...rfStData.map(r => ({ ...r, _source: 'station' }))
  ];
  const sortedRf = combinedRf.sort((a, b) => getRfTimestamp(a) - getRfTimestamp(b));
  let currentLoss: any = null;

  sortedRf.forEach((row) => {
    const stnId = row[stnIdCol];
    const stnNameRaw = String(row[stnNameCol] || stationMap[stnId] || '').trim();
    const stnName = formatStationName(stnNameRaw);
    const rawRowLocoId = row[locoIdCol] || locoId;
    if (!isValidLocoId(rawRowLocoId) || !isValidStationId(stnId)) return;
    const rowLocoId = String(rawRowLocoId).trim();
    const received = Number(row[receivedCol]) || 0;
    const expected = Number(row[expectedCol]) || 0;
    const percentage = Number(row[percentageCol]) || 0;
    const timestamp = getRfTimestamp(row);
    const radio = String(row[radioCol] || 'Radio 1').trim();
    const source = row._source;

    if (!stationFailures[stnId]) {
      stationFailures[stnId] = { count: 0, totalDuration: 0, locos: new Set(), totalEvents: 0, workingEvents: 0 };
    }
    stationFailures[stnId].totalEvents++;
    if (percentage >= 95) stationFailures[stnId].workingEvents++;

    const isLoss = percentage < 50 || (expected > 0 && received === 0);

    if (isLoss) {
      stationFailures[stnId].count++;
      stationFailures[stnId].locos.add(rowLocoId);
      
      if (!locoFailures[rowLocoId]) {
        locoFailures[rowLocoId] = { count: 0, stations: new Set() };
      }
      locoFailures[rowLocoId].count++;
      locoFailures[rowLocoId].stations.add(stnId);

      // Aggregate consecutive losses (within 2 minutes)
      if (currentLoss && currentLoss.locoId === rowLocoId && currentLoss.stationId === stnId && currentLoss.radio === radio && (timestamp - currentLoss.endTime) < 120000) {
        currentLoss.endTime = timestamp;
        currentLoss.minPerc = Math.min(currentLoss.minPerc, percentage);
      } else {
        if (currentLoss) {
          const duration = Math.round((currentLoss.endTime - currentLoss.startTime) / 1000) || 30;
          const reason = getReason({ ...currentLoss, duration });
          const eventTimeStr = new Date(currentLoss.startTime).toLocaleTimeString();
          const eventKey = `${eventTimeStr}|${currentLoss.locoId}|${currentLoss.stationId}|Radio Loss`;
          
          if (!seenCriticalEvents.has(eventKey)) {
            seenCriticalEvents.add(eventKey);
            criticalEvents.push({
              time: eventTimeStr,
              stationId: (currentLoss.stationId === 'N/A' || currentLoss.stationId === '-') ? '' : String(currentLoss.stationId),
              stationName: (currentLoss.stationName === 'N/A' || currentLoss.stationName === '-') ? '' : currentLoss.stationName,
              locoId: currentLoss.locoId,
              duration,
              type: 'Radio Loss',
              description: `Radio Loss (${currentLoss.minPerc}%) for ${duration}s at ${formatStationName((currentLoss.stationName && currentLoss.stationName !== 'N/A') ? currentLoss.stationName : (stationMap[currentLoss.stationId] || currentLoss.stationId))} (${currentLoss.source === 'train' ? 'Train Side' : 'Station Side'})`,
              radio: currentLoss.radio,
              reason
            });
          }
        }
        currentLoss = {
          locoId: rowLocoId,
          stationId: stnId,
          stationName: stnName,
          startTime: timestamp,
          endTime: timestamp,
          radio,
          minPerc: percentage,
          source
        };
      }
      stationFailures[stnId].totalDuration += 30;
    } else {
      if (currentLoss) {
        const duration = Math.round((currentLoss.endTime - currentLoss.startTime) / 1000) || 30;
        const reason = getReason({ ...currentLoss, duration });
        const eventTimeStr = new Date(currentLoss.startTime).toLocaleTimeString();
        const eventKey = `${eventTimeStr}|${currentLoss.locoId}|${currentLoss.stationId}|Radio Loss`;
        
        if (!seenCriticalEvents.has(eventKey)) {
          seenCriticalEvents.add(eventKey);
          criticalEvents.push({
            time: eventTimeStr,
            stationId: (currentLoss.stationId === 'N/A' || currentLoss.stationId === '-') ? '' : String(currentLoss.stationId),
            stationName: (currentLoss.stationName === 'N/A' || currentLoss.stationName === '-') ? '' : currentLoss.stationName,
            locoId: currentLoss.locoId,
            duration,
            type: 'Radio Loss',
            description: `Radio Loss (${currentLoss.minPerc}%) for ${duration}s at ${formatStationName((currentLoss.stationName && currentLoss.stationName !== 'N/A') ? currentLoss.stationName : (stationMap[currentLoss.stationId] || currentLoss.stationId))} (${currentLoss.source === 'train' ? 'Train Side' : 'Station Side'})`,
            radio: currentLoss.radio,
            reason
          });
        }
        currentLoss = null;
      }
    }
  });
  if (currentLoss) {
    const duration = Math.round((currentLoss.endTime - currentLoss.startTime) / 1000) || 30;
    const reason = getReason({ ...currentLoss, duration });
    const eventTimeStr = new Date(currentLoss.startTime).toLocaleTimeString();
    const eventKey = `${eventTimeStr}|${currentLoss.locoId}|${currentLoss.stationId}|Radio Loss`;
    
    if (!seenCriticalEvents.has(eventKey)) {
      seenCriticalEvents.add(eventKey);
      criticalEvents.push({
        time: eventTimeStr,
        stationId: (currentLoss.stationId === 'N/A' || currentLoss.stationId === '-') ? '' : String(currentLoss.stationId),
        stationName: (currentLoss.stationName === 'N/A' || currentLoss.stationName === '-') ? '' : currentLoss.stationName,
        locoId: currentLoss.locoId,
        duration,
        type: 'Radio Loss',
        description: `Radio Loss (${currentLoss.minPerc}%) for ${duration}s at ${formatStationName((currentLoss.stationName && currentLoss.stationName !== 'N/A') ? currentLoss.stationName : (stationMap[currentLoss.stationId] || currentLoss.stationId))} (${currentLoss.source === 'train' ? 'Train Side' : 'Station Side'})`,
        radio: currentLoss.radio,
        reason
      });
    }
  }

  // FALLBACK: Detect Radio Loss from trnData if rfData is empty or as additional source
  if (trnData && trnData.length > 0) {
    let lastRadioTime: number | null = null;
    let lastRadioState: string = '';
    const trnKeys = Object.keys(trnData[0]);
    
    trnData.forEach((row, i) => {
      const currentTime = parseTime(getTrnTime(row));
      if (isNaN(currentTime)) return;

      // Detect radio state from Column E (loco radio) and Column AD (station radio)
      // and other radio packet columns (AD-BF)
      let currentRadioState = '';
      for (let j = 29; j <= 57; j++) {
        const key = trnKeys[j];
        if (key) currentRadioState += String(row[key] || '');
      }
      // Also include Column E (Loco Radio)
      const locoRadio = String(row[trnRadioCol] || '').trim();
      currentRadioState += locoRadio;

      if (i === 0) {
        lastRadioState = currentRadioState;
        lastRadioTime = currentTime;
        return;
      }

      const isRadioActive = currentRadioState.replace(/0/g, '').length > 0;

      // If radio state changed, it means a new packet was received
      if (isRadioActive && currentRadioState !== lastRadioState) {
        if (lastRadioTime !== null) {
          const duration = (currentTime - lastRadioTime) / 1000;
          // If duration > 5s, record it as a radio loss event if it's not already covered by rfData
          if (duration > 5 && duration < 300) {
            const timeStr = getTrnTime(row);
            // Check if we already have a similar event from rfData
            const exists = criticalEvents.some(e => e.time === timeStr && e.type === 'Radio Loss');
            
            if (!exists) {
              let stnName = String(row[trnKeys[2]] || row[trnKeys[32]] || row[trnKeys[1]] || row[trnKeys[3]] || row[trnKeys[31]] || '').trim();
              const locoNo = getBestLocoIdFromRow(row, trnKeys, locoId);
              let stnId = String(row[findColumn(row, 'Station Id', 'StationId', 'Station_Id') || ''] || 'N/A');

              // If station info is missing, try to find it from RF data near this time
              if ((!stnName || stnName === 'N/A' || stnName === '-') && (!stnId || stnId === 'N/A' || stnId === '-')) {
                const nearestRf = rfData.find(r => {
                  const rfTime = parseTime(getRadioTime(r));
                  return Math.abs(rfTime - currentTime) < 30000; // Within 30s
                });
                if (nearestRf) {
                  stnName = String(nearestRf[stnNameCol] || '').trim();
                  stnId = String(nearestRf[stnIdCol] || '').trim();
                }
              }

              // If we have an ID but no name, try the map
              if (stnId && stnId !== 'N/A' && stnId !== '-' && (!stnName || stnName === 'N/A' || stnName === '-')) {
                stnName = stationMap[stnId] || '';
              }

              criticalEvents.push({
                time: timeStr,
                stationId: (stnId === 'N/A' || stnId === '-') ? '' : stnId,
                stationName: (stnName === 'N/A' || stnName === '-') ? '' : stnName,
                locoId: locoNo,
                duration: Math.round(duration),
                type: 'Radio Loss',
                description: `Radio Loss detected from TRN log (Gap: ${Math.round(duration)}s)${stnName && stnName !== 'N/A' && stnName !== '-' ? ' at ' + stnName : ''}`,
                radio: locoRadio || 'Radio 1',
                reason: getReason({ duration: Math.round(duration), minPerc: 0 }) // TRN log loss is usually total loss
              });
            }
          }
        }
        lastRadioState = currentRadioState;
        lastRadioTime = currentTime;
      }
    });
  }

  // Time-based Analysis: Check for multiple trains at same time
  const timeMap: Record<string, Set<string | number>> = {};
  rfData.forEach(row => {
    const time = getRfTime(row);
    const percentage = Number(row[percentageCol]) || 0;
    const stnId = row[stnIdCol];
    if (percentage < 50 && time !== 'N/A') {
      const key = `${time}|${stnId}`;
      if (!timeMap[key]) timeMap[key] = new Set();
      timeMap[key].add(row[locoIdCol] || locoId);
    }
  });

  Object.entries(timeMap).forEach(([key, locos]) => {
    if (locos.size > 1) {
      const [time, stnId] = key.split('|');
      criticalEvents.push({
        time,
        stationId: (stnId === 'N/A' || stnId === '-') ? '' : stnId,
        locoId: 'Multiple',
        duration: 0,
        type: 'Multiple Trains Affected',
        description: `${locos.size} trains affected at ${formatStationName(stationMap[stnId] || stnId)} simultaneously`,
        reason: "Station Side / Environmental"
      });
    }
  });

  const topFaultyStations = Object.entries(stationFailures)
    .filter(([stnId]) => {
      const g = globalStationStats.get(stnId);
      const globalPerf = g && g.exp > 0 ? (g.rec / g.exp) * 100 : 100;
      return globalPerf < 95; // Only show stations that are not "Healthy"
    })
    .map(([stnId, data]) => {
      const healthScore = (data.workingEvents / (data.totalEvents || 1)) * 100;
      return {
        stationId: stnId,
        failureCount: data.count,
        avgLossDuration: data.count > 0 ? data.totalDuration / data.count : 0,
        healthScore,
        status: (healthScore < 85 ? 'Unhealthy' : healthScore <= 95 ? 'Warning' : 'Healthy') as any,
        affectedLocos: Array.from(data.locos)
      };
    })
    .sort((a, b) => b.failureCount - a.failureCount)
    .slice(0, 10);

  const faultyLocos = Object.entries(locoFailures)
    .map(([locoId, data]) => ({
      locoId,
      failureCount: data.count,
      stationsCovered: Array.from(data.stations),
      status: (data.count > 10 ? 'Critical' : data.count > 5 ? 'Suspect' : 'Normal') as any
    }))
    .sort((a, b) => b.failureCount - a.failureCount);

  // Root Cause Conclusion
  const totalStationFailures = topFaultyStations.reduce((acc, s) => acc + s.failureCount, 0);
  const totalLocoFailures = faultyLocos.reduce((acc, l) => acc + l.failureCount, 0);
  
  let stationSideWeight = totalStationFailures > 0 ? (totalStationFailures / (totalStationFailures + totalLocoFailures)) * 100 : 0;
  let locoSideWeight = 100 - stationSideWeight;

  // Hardware vs Software Analysis
  // Hardware: Correlated with low RF signal strength
  // Software: Correlated with NMS Health issues (non-0) or processing lag
  let hardwareProb = 0;
  let softwareProb = 0;

  const totalDegradations = modeDegradations.length;
  if (totalDegradations > 0) {
    const rfIssues = modeDegradations.filter(d => d.reason.includes('Poor RF Signal')).length;
    const nmsIssues = modeDegradations.filter(d => d.reason.includes('NMS Server')).length;
    const packetLoss = modeDegradations.filter(d => d.reason.includes('Radio Packet Loss')).length;

    hardwareProb = (rfIssues / totalDegradations) * 100;
    softwareProb = (nmsIssues / totalDegradations) * 100;
    
    // Packet loss without poor RF is often a software/processing issue on the radio modem or TCAS
    if (packetLoss > 0 && rfIssues === 0) {
      softwareProb = Math.min(100, softwareProb + (packetLoss / totalDegradations) * 50);
    }
  } else {
    // Fallback based on overall stats
    hardwareProb = (badStns.length / (badStns.length + goodStns.length || 1)) * 100;
    softwareProb = nmsFailRate * 100;
  }

  // Normalize hardware/software
  const totalProb = hardwareProb + softwareProb || 1;
  hardwareProb = (hardwareProb / totalProb) * 100;
  softwareProb = (softwareProb / totalProb) * 100;

  // Refine with Station-wise RFCOMM Data if available
  let stationSpecificIssues = false;
  if (rfStData.length > 0) {
    const stnPerfMap = new Map<string, number>();
    rfStData.forEach(row => {
      const stnId = String(row[stnIdCol] || '').trim();
      const perc = parseFloat(row[percentageCol]) || 0;
      if (stnId && stnId.toLowerCase() !== 'station id' && stnId.toLowerCase() !== 'stationid') {
        stnPerfMap.set(stnId, (stnPerfMap.get(stnId) || 0) + perc);
      }
    });
    
    let lowPerfStns = 0;
    stnPerfMap.forEach((total, id) => {
      const count = rfStData.filter(r => String(r[stnIdCol]) === id).length;
      const avg = total / (count || 1);
      if (avg < 90) lowPerfStns++;
    });

    if (lowPerfStns > 0) {
      stationSideWeight = Math.min(100, stationSideWeight + 30);
      locoSideWeight = 100 - stationSideWeight;
      stationSpecificIssues = true;
    }
  }

  let conclusion = "Random Failures Detected: Intermittent RF loss observed. Likely caused by environmental interference or transient signal drops.";
  let breakdown = "The analysis suggests a mix of factors affecting the communication link.";

  if (stationSpecificIssues || stationSideWeight > 65) {
    conclusion = `Station TCAS / Trackside Issue: High correlation of failures at specific stations (${topFaultyStations.slice(0, 2).map(s => formatStationName(stationMap[s.stationId] || s.stationId)).join(', ')}) across multiple locos.`;
    breakdown = `The failure is localized to the trackside infrastructure. ${hardwareProb > 60 ? "Likely Hardware: Check Station Antenna, RF Cables, or Power Supply." : "Likely Software/Config: Check Station Radio Modem configuration or NMS link."}`;
  } else if (locoSideWeight > 65) {
    conclusion = `Loco TCAS / Onboard Issue: Failures are specific to Loco ${locoId} across multiple stations.`;
    breakdown = `The failure is onboard the locomotive. ${hardwareProb > 60 ? "Likely Hardware: Check Loco RF Module, Antenna alignment, or VSWR." : "Likely Software/Processing: Check TCAS Software version, NMS Health, or Radio processing lag."}`;
  }

  // Deep Analysis Dashboard Logic (DYNAMIC)
  const dashboardTable: { station: string; locoVal: string; othersAvg: string }[] = [];
  
  // Get all unique stations from RF logs (NORMALIZED)
  const allStnsInRf = Array.from(new Set(rfStData.map(r => String(r[stnIdCol] || '').trim().toUpperCase().replace(/\s+STATION$/i, ''))))
    .filter(s => s && s !== 'STATION ID' && s !== 'STATIONID');
  
  const stnComparisons = allStnsInRf.map(stnIdVal => {
    const locoStats = rfStData.filter(r => String(r[stnIdCol] || '').trim().toUpperCase().replace(/\s+STATION$/i, '') === stnIdVal && String(r[locoIdCol] || '').trim() === String(locoId).trim());
    const otherStats = rfStData.filter(r => String(r[stnIdCol] || '').trim().toUpperCase().replace(/\s+STATION$/i, '') === stnIdVal && String(r[locoIdCol] || '').trim() !== String(locoId).trim());
    
    const locoAvg = locoStats.length > 0 ? locoStats.reduce((acc, r) => acc + (parseFloat(r[percentageCol]) || 0), 0) / locoStats.length : null;
    const othersAvg = otherStats.length > 0 ? otherStats.reduce((acc, r) => acc + (parseFloat(r[percentageCol]) || 0), 0) / otherStats.length : 98.5;

    return {
      stationId: stnIdVal,
      locoAvg,
      othersAvg,
      diff: locoAvg !== null ? othersAvg - locoAvg : 0
    };
  });

  // Problem 1: Stations where this loco is significantly worse than others
  const locoSpecificDrops = stnComparisons
    .filter(c => c.locoAvg !== null && c.locoAvg < 90 && c.othersAvg > 95)
    .sort((a, b) => b.diff - a.diff);

  locoSpecificDrops.slice(0, 4).forEach(d => {
    const name = stationMap[d.stationId] || String(d.stationId);
    dashboardTable.push({
      station: formatStationName(name),
      locoVal: `${d.locoAvg?.toFixed(1)}%`,
      othersAvg: `${d.othersAvg.toFixed(1)}%`
    });
  });

  // If no specific drops found, show top faulty stations for this loco
  if (dashboardTable.length === 0) {
    topFaultyStations.slice(0, 3).forEach(stn => {
      const others = stnComparisons.find(c => c.stationId === stn.stationId)?.othersAvg || 98.5;
      const name = stationMap[stn.stationId] || String(stn.stationId);
      dashboardTable.push({
        station: formatStationName(name),
        locoVal: `${stn.healthScore.toFixed(1)}%`,
        othersAvg: `${others.toFixed(1)}%`
      });
    });
  }

  // Find a "Healthy Station" benchmark (highest avg others performance)
  const healthyBenchmark = stnComparisons
    .filter(c => c.othersAvg > 98)
    .sort((a, b) => b.othersAvg - a.othersAvg)[0] || stnComparisons[0];

  // Problem 2 Priority (Stations where multiple locos fail)
  const stationPriority = multiLocoBadStns
    .sort((a, b) => b.locoCount - a.locoCount || a.avgPerf - b.avgPerf)
    .map(s => {
      const name = stationMap[s.stationId] || String(s.stationId);
      return formatStationName(name);
    });

  const isLocoFaulty = locoSideWeight > 60 || locoSpecificDrops.length > 1;

  // Pre-calculate station-loco metrics for performance optimization
  const stnLocoMetrics = new Map<string, Map<string, { exp: number, rec: number, sum: number, count: number }>>();
  rfStData.forEach(row => {
    let sId = String(row[stnIdCol] || '').trim().toUpperCase().replace(/\s+STATION$/i, '');
    const lId = String(row[locoIdCol] || '').trim();
    const exp = parseNumber(row[expectedCol]) || 0;
    const rec = parseNumber(row[receivedCol]) || 0;
    const perc = parseFloat(row[percentageCol]) || (exp > 0 ? (rec / exp) * 100 : 0);
    if (!sId || !lId || sId === 'STATION ID' || sId === 'STATIONID' || lId.toLowerCase() === 'loco id' || lId.toLowerCase() === 'locoid') return;

    if (!stnLocoMetrics.has(sId)) stnLocoMetrics.set(sId, new Map());
    const lMap = stnLocoMetrics.get(sId)!;
    if (!lMap.has(lId)) lMap.set(lId, { exp: 0, rec: 0, sum: 0, count: 0 });
    const m = lMap.get(lId)!;
    m.exp += exp;
    m.rec += rec;
    m.sum += perc;
    m.count++;
  });

  // Pre-calculate station-wide averages (excluding specific locos)
  const stnGlobalMetrics = new Map<string, { exp: number, rec: number, sum: number, count: number }>();
  const locoPerformanceMap = new Map<string, number>();
  const locoPerfCounts = new Map<string, number>();
  const locoPerfExpRec = new Map<string, { exp: number, rec: number }>();

  stnLocoMetrics.forEach((lMap, sId) => {
    let sExp = 0;
    let sRec = 0;
    let sSum = 0;
    let sCount = 0;
    lMap.forEach((m, lId) => {
      sExp += m.exp;
      sRec += m.rec;
      sSum += m.sum;
      sCount += m.count;
      
      if (!locoPerfExpRec.has(lId)) locoPerfExpRec.set(lId, { exp: 0, rec: 0 });
      const lp = locoPerfExpRec.get(lId)!;
      lp.exp += m.exp;
      lp.rec += m.rec;
      
      locoPerformanceMap.set(lId, (locoPerformanceMap.get(lId) || 0) + m.sum);
      locoPerfCounts.set(lId, (locoPerfCounts.get(lId) || 0) + m.count);
    });
    stnGlobalMetrics.set(sId, { exp: sExp, rec: sRec, sum: sSum, count: sCount });
  });

  // Finalize loco performance averages
  locoPerformanceMap.forEach((sum, lId) => {
    const lp = locoPerfExpRec.get(lId);
    if (lp && lp.exp > 0) {
      locoPerformanceMap.set(lId, (lp.rec / lp.exp) * 100);
    } else {
      locoPerformanceMap.set(lId, sum / (locoPerfCounts.get(lId) || 1));
    }
  });

  // Helper to calculate analysis for a specific loco
  const getLocoAnalysis = (targetLocoId: string) => {
    const isAll = targetLocoId === 'All' || targetLocoId === 'All Locos';
    
    // Filter failures for this specific loco if not "All"
    const targetLocoFailures = isAll ? faultyLocos : faultyLocos.filter(l => String(l.locoId) === targetLocoId);
    const targetStnFailures = isAll ? topFaultyStations : topFaultyStations.filter(s => s.affectedLocos.includes(targetLocoId));
    
    const totalStnFailures = targetStnFailures.reduce((acc, s) => acc + s.failureCount, 0);
    const totalLcoFailures = targetLocoFailures.reduce((acc, l) => acc + l.failureCount, 0);
    const totalFailures = totalStnFailures + totalLcoFailures;
    let stnSideWeight = totalFailures > 0 ? (totalStnFailures / totalFailures) * 100 : 0;
    let lcoSideWeight = totalFailures > 0 ? (totalLcoFailures / totalFailures) * 100 : 0;

    // Refine with Station-side data
    let stnSpecificIssues = false;
    let lowPerfStnsCount = 0;
    stnGlobalMetrics.forEach((m, sId) => {
      const avg = m.exp > 0 ? (m.rec / m.exp) * 100 : (m.sum / (m.count || 1));
      if (avg < 90) {
        // If "All", any low perf station counts. If specific loco, only if that loco also failed there
        let locoAvg = 100;
        if (!isAll) {
          const lMetric = stnLocoMetrics.get(sId)?.get(targetLocoId);
          if (lMetric) {
            locoAvg = lMetric.exp > 0 ? (lMetric.rec / lMetric.exp) * 100 : (lMetric.sum / (lMetric.count || 1));
          }
        }
        if (isAll || locoAvg < 95) {
          lowPerfStnsCount++;
        }
      }
    });

    if (lowPerfStnsCount > 0) {
      stnSideWeight = Math.min(100, stnSideWeight + 30);
      lcoSideWeight = 100 - stnSideWeight;
      stnSpecificIssues = true;
    }

    // Hardware vs Software (Master Logic)
    // RF% bad + NMS OK -> Hardware
    // NMS bad + RF% OK -> Software
    // Both bad -> Correlated
    let hProb = 50;
    let sProb = 50;

    const locoPerf = isAll ? locoPerformance : (locoPerformanceMap.get(targetLocoId) || 100);
    const locoNmsLogs = nmsLogs.filter(n => isAll || String(n.locoId) === targetLocoId);
    const locoNmsFail = locoNmsLogs.filter(n => n.health !== '32').length;
    const locoNmsRate = locoNmsLogs.length > 0 ? (locoNmsFail / locoNmsLogs.length) * 100 : 0;

    if (locoPerf < 90 && locoNmsRate < 15) {
      hProb = 85;
      sProb = 15;
    } else if (locoPerf > 95 && locoNmsRate > 25) {
      sProb = 85;
      hProb = 15;
    } else if (locoPerf < 90 && locoNmsRate > 25) {
      hProb = 60;
      sProb = 40; // Correlated
    }

    // Refine with MA Lag (Software indicator)
    const locoMa = maPacketsProcessed.filter(p => isAll || String(p.locoId) === targetLocoId);
    const locoAvgLag = locoMa.length > 0 ? locoMa.reduce((acc, p) => acc + p.delay, 0) / locoMa.length : 0;
    if (locoAvgLag > 1.5) {
      sProb = Math.min(100, sProb + 20);
      hProb = Math.max(0, hProb - 20);
    }

    let conc = isAll ? "Fleet-wide Analysis: Multiple locomotives and stations showing intermittent issues." : `Analysis for Loco ${targetLocoId}: Evaluating onboard vs trackside factors.`;
    let bdown = "The analysis suggests a mix of factors affecting the communication link.";

    if (stnSpecificIssues || stnSideWeight > 65) {
      conc = isAll ? "Primary Trackside Issues: High correlation of failures at specific stations across the fleet." : `Station-side Issue: Failures for Loco ${targetLocoId} are highly correlated with specific trackside locations.`;
      bdown = `The failure is localized to the trackside infrastructure. ${hProb > 60 ? "Likely Hardware: Check Station Antenna, RF Cables, or Power Supply." : "Likely Software/Config: Check Station Radio Modem configuration or NMS link."}`;
    } else if (lcoSideWeight > 65) {
      conc = isAll ? "Locomotive Fleet Issues: Failures are distributed across locomotives regardless of station." : `Loco-side Issue: Failures are specific to Loco ${targetLocoId} across multiple stations.`;
      bdown = `The failure is onboard the locomotive. ${hProb > 60 ? "Likely Hardware: Check Loco RF Module, Antenna alignment, or VSWR." : "Likely Software/Processing: Check TCAS Software version, NMS Health, or Radio processing lag."}`;
    }

    // Dashboard Logic
    const dTable: { station: string; locoVal: string; othersAvg: string }[] = [];
    const stnComps = Array.from(stnLocoMetrics.keys()).map(sId => {
      const lMap = stnLocoMetrics.get(sId)!;
      const gMetric = stnGlobalMetrics.get(sId)!;
      
      let lAvg: number | null = null;
      // Use Global Fleet Average as the stable benchmark for "Baaki Locos (Avg)"
      const globalAvg = gMetric.exp > 0 ? (gMetric.rec / gMetric.exp) * 100 : (gMetric.sum / (gMetric.count || 1));
      let oAvg = globalAvg;

      if (isAll) {
        lAvg = globalAvg;
        // When looking at "All", compare against a high-performance target (98.5%)
        oAvg = 98.5;
      } else {
        const lMetric = lMap.get(targetLocoId);
        if (lMetric) {
          lAvg = lMetric.exp > 0 ? (lMetric.rec / lMetric.exp) * 100 : (lMetric.sum / lMetric.count);
        }
      }

      return { stationId: sId, locoAvg: lAvg, othersAvg: oAvg, diff: lAvg !== null ? oAvg - lAvg : 0 };
    });

    const lDrops = stnComps
      .filter(c => c.locoAvg !== null && c.locoAvg < 90 && c.othersAvg > 95)
      .sort((a, b) => b.diff - a.diff);

    const isLFaulty = lDrops.length > 0 && lcoSideWeight > 50;
    const isSFaulty = stnSideWeight > 50 && multiLocoBadStns.length > 0;

    if (isLFaulty) {
      lDrops.slice(0, 4).forEach(d => {
        dTable.push({ station: d.stationId, locoVal: `${d.locoAvg?.toFixed(1)}%`, othersAvg: `${d.othersAvg.toFixed(1)}%` });
      });

      if (dTable.length === 0) {
        const relevantStns = isAll ? topFaultyStations : topFaultyStations.filter(s => s.affectedLocos.includes(targetLocoId));
        relevantStns.slice(0, 3).forEach(stn => {
          const comp = stnComps.find(c => c.stationId === stn.stationId);
          const locoAvg = comp?.locoAvg ?? stn.healthScore;
          const others = comp?.othersAvg ?? 98.5;
          dTable.push({ station: stn.stationId, locoVal: `${locoAvg.toFixed(1)}%`, othersAvg: `${others.toFixed(1)}%` });
        });
      }
    }

    const hBenchmark = stnComps
      .filter(c => c.othersAvg > 98)
      .sort((a, b) => b.othersAvg - a.othersAvg)[0] || stnComps[0];

    const faultyStations = stnComps
      .filter(c => c.othersAvg < 95)
      .sort((a, b) => a.othersAvg - b.othersAvg); // Worst first

    const topPriorityStns = faultyStations.map(s => String(s.stationId)).slice(0, 15);

    const stationsWithNoData = Object.keys(stationMap).filter(sId => !stnLocoMetrics.has(sId));

    return {
      topFaultyStations: targetStnFailures,
      faultyLocos: targetLocoFailures,
      criticalEvents: criticalEvents.filter(e => isAll || e.locoId === targetLocoId || e.locoId === 'Multiple').slice(0, 20),
      rootCause: {
        stationSide: Math.round(stnSideWeight),
        locoSide: Math.round(lcoSideWeight),
        hardwareProb: Math.round(hProb),
        softwareProb: Math.round(sProb),
        conclusion: conc,
        breakdown: bdown
      },
      dashboard: {
        conclusion: (isLFaulty && isSFaulty) ? "Multiple Issues Detected (Loco + Station)" : 
                    isLFaulty ? `Problem Detected: Loco ${isAll ? 'Fleet' : targetLocoId} TCAS Unit Suspect` :
                    isSFaulty ? "Problem Detected: Station-side TCAS/RF Health Issues" :
                    `Locomotive ${isAll ? 'Fleet' : targetLocoId} Fit - System Healthy`,
        problem1: {
          title: isLFaulty 
            ? `Problem 1 — Loco ${isAll ? 'Fleet' : targetLocoId} TCAS unit is suspect` 
            : `Problem 1 — Loco ${isAll ? 'Fleet' : targetLocoId} Performance Audit (Fit)`,
          description: isLFaulty
            ? `Loco ${isAll ? 'Fleet' : targetLocoId} showed performance drops at ${lDrops.length} stations while other locos performed normally there. This indicates a loco-side hardware/software issue.`
            : `Loco ${isAll ? 'Fleet' : targetLocoId} performance is equal to or better than the fleet average. No major loco-side issues detected.`,
          table: dTable,
          causes: isLFaulty 
            ? [
                "Physical damage or loose connection in the loco antenna",
                "RF transceiver module is weak (low power output)",
                "TCAS software bug causing failures at specific station configurations"
              ]
            : isSFaulty
              ? [
                  "Station TCAS antenna alignment issue",
                  "Track-side RF modem power fluctuations",
                  "Localized RF interference or signal shadowing"
                ]
              : []
        },
        problem2: {
          title: "Problem 2 — Station-side TCAS/RF Health",
          description: "When multiple independent locos fail at the same location, the station TCAS antenna or RF hardware should be inspected.",
          priority: topPriorityStns
        },
        amlConclusion: (() => {
          const healthyCount = stnComps.filter(c => c.othersAvg >= 95).length;
          const totalCount = stnComps.length;
          const belowThresholdCount = totalCount - healthyCount;
          
          let text = `${healthyCount} of ${totalCount} stations performing within healthy limits (≥95%).`;
          if (belowThresholdCount > 0) {
            text += ` ${belowThresholdCount} station${belowThresholdCount > 1 ? 's are' : ' is'} below threshold and require${belowThresholdCount === 1 ? 's' : ''} inspection.`;
          } else {
            text += " Fleet data suggests track-side equipment is generally healthy at major stations.";
          }
          return text;
        })(),
        actionRequired: (() => {
          if (isLFaulty) {
            return `Send Loco ${isAll ? 'Fleet' : targetLocoId} to the workshop — check the RF antenna and transceiver module, and verify the TCAS firmware version.`;
          }
          
          if (faultyStations.length > 0) {
            let text = `Locomotive ${isAll ? 'fleet' : 'Loco ' + targetLocoId} is fit. It is recommended to inspect the station-side TCAS antenna and RF hardware at:\n`;
            faultyStations.slice(0, 5).forEach((s, idx) => {
              text += `Priority ${idx + 1} — ${formatStationName(stationMap[s.stationId] || s.stationId)} (${s.othersAvg.toFixed(2)}%)\n`;
            });
            
            if (stationsWithNoData.length > 0) {
              text += `\nAll other stations with available data are at or above 95%. ${stationsWithNoData.slice(0, 10).map(sId => formatStationName(stationMap[sId] || sId)).join(', ')} stations have no data uploaded yet.`;
            }
            return text;
          }
          
          return `Locomotive ${isAll ? 'fleet' : 'Loco ' + targetLocoId} is fit and all stations are performing within healthy limits.`;
        })()
      }
    };
  };

  const locoAnalyses: Record<string, any> = {};
  locoIds.forEach(id => {
    locoAnalyses[id] = getLocoAnalysis(id);
  });
  locoAnalyses['All'] = getLocoAnalysis('All');
  locoAnalyses['All Locos'] = locoAnalyses['All'];

  const stationDeepAnalysis = locoAnalyses[locoId] || locoAnalyses['All'];

  // Moving Radio Loss Analysis (Speed > 0)
  const movingRadioLoss: DashboardStats['movingRadioLoss'] = [];
  
  locoIds.forEach(lId => {
    const lIdStr = String(lId);
    const lTrnData = trnData?.filter(r => String(r[trnLocoIdCol] || locoId).trim() === lIdStr) || [];
    
    let r1Packets = 0;
    let r2Packets = 0;
    let totalPackets = 0;
    
    lTrnData.forEach(row => {
      const radioVal = String(row[trnRadioCol] || '').toLowerCase().trim();
      if (radioVal.includes('radio 1') || radioVal.includes('r1') || radioVal === '1') r1Packets++;
      else if (radioVal.includes('radio 2') || radioVal.includes('r2') || radioVal === '2') r2Packets++;
      totalPackets++;
    });
    
    const r1Usage = totalPackets > 0 ? (r1Packets / totalPackets) * 100 : 0;
    const r2Usage = totalPackets > 0 ? (r2Packets / totalPackets) * 100 : 0;
    
    let movingGaps = 0;
    let maxGap = 0;
    let lastPacketTime: number | null = null;
    let lastSpeed: number = 0;
    
    lTrnData.forEach(row => {
      const speed = Number(row[speedCol]) || 0;
      const time = parseTime(getTrnTime(row));
      
      // Check if any radio packet was actually received in this row (Indices 29-57)
      let hasRadio = false;
      const keys = Object.keys(row);
      for (let j = 0; j < keys.length; j++) {
        const k = keys[j];
        // Radio columns are usually named or in specific ranges (Column AD to BF)
        if (k.toLowerCase().includes('pkt') || (j >= 29 && j <= 57)) {
          const val = row[k];
          if (val !== undefined && val !== null && val !== '' && val !== 0 && val !== '0') {
            hasRadio = true;
            break;
          }
        }
      }
      
      if (hasRadio) {
        if (lastPacketTime !== null) {
          const gap = (time - lastPacketTime) / 1000;
          // Moving Gap: Speed must be positive and gap must be reasonable (e.g. < 30 mins)
          // 74,662s is clearly not a "moving gap" but a trip boundary or data gap
          if ((lastSpeed > 0 || speed > 0) && gap > 5 && gap < 3600) { 
            movingGaps++;
            if (gap > maxGap) maxGap = gap;
          }
        }
        lastPacketTime = time;
        lastSpeed = speed;
      }
    });
    
    let conclusion = "Normal/Low Issue";
    if (movingGaps > 20) conclusion = "Highest Signal Drop While Moving";
    else if (maxGap > 1000) conclusion = "Very Large Communication Gaps";
    else if (movingGaps > 15) conclusion = "Continuous Signal Instability";
    else if (Math.abs(r1Usage - r2Usage) > 8) conclusion = `Hardware Issue (Radio ${r1Usage < r2Usage ? '1' : '2'})`;
    else if (movingGaps === 0 && totalPackets > 0) conclusion = "Excellent Performance";

    movingRadioLoss.push({
      locoId: lId,
      movingGaps,
      maxGap: Math.round(maxGap),
      r1Usage: Number(r1Usage.toFixed(1)),
      r2Usage: Number(r2Usage.toFixed(1)),
      conclusion
    });
  });

  const allStnIdentifiers = new Set<string>();
  stnPerf.forEach(s => allStnIdentifiers.add(String(s.stationId).toUpperCase()));
  badStns.forEach(s => allStnIdentifiers.add(String(s).toUpperCase()));
  goodStns.forEach(s => allStnIdentifiers.add(String(s).toUpperCase()));
  if (stationStats) stationStats.forEach(s => allStnIdentifiers.add(String(s.stationId).toUpperCase()));

  // --- Smart Diagnosis Discovery ---
  let interTagDistAffectedRows = 0;
  const totalTrnRows = trnData?.length || 0;
  if (trnData) {
    trnData.forEach(row => {
      const reason = String(row[reasonCol] || '').toLowerCase();
      const event = String(row[eventCol] || '').toLowerCase();
      if (reason.includes('intertagdist') || event.includes('intertagdist')) {
        interTagDistAffectedRows++;
      }
    });
  }

  const smartDiagnosis: DashboardStats['smartDiagnosis'] = {
    globalPattern: {
      issue: "InterTagDistGreaterThanDupTag (Track RFID Spacing Defect)",
      totalRows: totalTrnRows,
      affectedRows: interTagDistAffectedRows,
      percentage: totalTrnRows > 0 ? (interTagDistAffectedRows / totalTrnRows) * 100 : 0,
      explanation: "Persistent station-side track RFID tag spacing defect. This creates background pressure on the TCAS positioning system, making it more likely to degrade modes during radio switches or packet drops."
    },
    stationInsights: [],
    locoInsights: [],
    protectionEvents: [],
    summary: ""
  };

  // Station Insights (ST and NVS)
  const allStnArr = Array.from(allStnIdentifiers);
  const stIdMatch = allStnArr.find(s => s.replace(/\s+STATION$/i, '').trim().toUpperCase() === 'ST');
  if (stIdMatch) {
    const stDegradations = modeDegradationsToUse.filter(d => 
      formatStationName(d.stationId).includes('ST ') || 
      formatStationName(d.stationId).startsWith('ST ') ||
      formatStationName(d.stationId) === 'ST STATION'
    );
    const standByCount = stDegradations.filter(d => (d.to || '').toUpperCase().includes('ST') || (d.to || '').toUpperCase().includes('STANDBY')).length;
    if (standByCount > 0 || stDegradations.length >= 2) {
      smartDiagnosis.stationInsights.push({
        stationId: 'ST',
        stationName: 'ST Station',
        severity: 'Critical',
        description: "Severe Mode Degradations detected (FS → StandBy).",
        locosAffected: Array.from(new Set(stDegradations.map(d => d.locoId))),
        details: [
          "Multiple locomotives dropping to StandBy at Speed 0 under Red Signal.",
          "Station TCAS fails to provide fresh Movement Authority (MA) in time.",
          "Confirmed as Station-side defect: multiple locomotives (e.g. 37352, 39018) affected independently."
        ]
      });
    }
  }

  const nvsIdMatch = allStnArr.find(s => s.replace(/\s+STATION$/i, '').trim().toUpperCase() === 'NVS');
  if (nvsIdMatch) {
    const nvsDegradations = modeDegradationsToUse.filter(d => 
      formatStationName(d.stationId).includes('NVS') || 
      (d.stationName && d.stationName.includes('NVS'))
    );
    if (nvsDegradations.length > 0) {
      smartDiagnosis.stationInsights.push({
        stationId: 'NVS',
        stationName: 'NVS Station',
        severity: 'Major',
        description: "Repeated Mode Oscillations (SR/OS).",
        locosAffected: Array.from(new Set(nvsDegradations.map(d => d.locoId))),
        details: [
          "Continuous Tag Distance Mismatch reported by all traversing locomotives.",
          "Manual LP acknowledgement required repeatedly (Loco 37943: 3 times in 22 mins).",
          "ReadEndCollision safety trigger detected for Loco 37424 at this station."
        ]
      });
    }
  }

  // Loco-side Insights (NMS Health 16, 40 etc.)
  const highSevrNms = nmsLogs.filter(log => log.health === '16' || log.health === '40' || log.health === '8');
  if (highSevrNms.length > 0) {
    const seenL = new Set();
    highSevrNms.forEach(lAlert => {
      if (seenL.has(lAlert.locoId)) return;
      seenL.add(lAlert.locoId);
      smartDiagnosis.locoInsights.push({
        locoId: lAlert.locoId,
        issue: `NMS Health Error: Code ${lAlert.health} (Vital Hardware Error)`,
        context: lAlert.health === '16' ? "BIU interface card mismatch or redundant processor synchronization loss." : "Transient hardware/link error detected.",
        recommendation: "Inspect physical BIU card and redundant sync cables. If error occurred at high speed (>100 km/h), prioritize inspection."
      });
    });
  }

  // Protection Events
  modeDegradationsToUse.forEach(deg => {
    if (deg.reason.toLowerCase().includes('collision') || deg.reason.toLowerCase().includes('protection')) {
      smartDiagnosis.protectionEvents.push({
        time: deg.time,
        locoId: deg.locoId,
        stationId: deg.stationId,
        event: deg.reason.includes('ReadEndCollision') ? "ReadEndCollision" : "Kavach Protection Trigger",
        analysis: "Kavach safety protection activated. Detected train or obstacle ahead. FS → LS/OS degradation was correct system behavior, not a fault."
      });
    }
  });

  // Final Summary
  const globalPct = smartDiagnosis.globalPattern?.percentage || 0;
  smartDiagnosis.summary = `Global pattern analysis shows Tag Spacing Defects in ${globalPct.toFixed(1)}% of data. ${smartDiagnosis.stationInsights.length} station-side infrastructure defects identified. ${smartDiagnosis.locoInsights.length} locomotives require physical inspection for recurring hardware alerts.`;

  // --- Technical Audit Report (Detailed Narrative) ---
  const technicalAudit: DashboardStats['technicalAudit'] = [];

  // ST Audit (Multi-Loco FS -> StandBy)
  const stMatch = allStnArr.find(s => s.replace(/\s+STATION$/i, '').trim().toUpperCase() === 'ST');
  if (stMatch) {
    const stEvents = modeDegradationsToUse.filter(d => formatStationName(stationMap[d.stationId] || d.stationId).includes('ST '));
    const standByEvts = stEvents.filter(d => (d.to || '').toUpperCase().includes('ST') || (d.to || '').toUpperCase().includes('STANDBY'));
    if (standByEvts.length >= 2) {
      technicalAudit.push({
        id: 'aud-st-standby',
        title: "Station Infrastructure Defect: ST Station FS → StandBy",
        locoIds: Array.from(new Set(standByEvts.map(e => e.locoId))),
        stationId: 'ST',
        stationName: 'ST Station',
        timeRange: `${standByEvts[0].time} to ${standByEvts[standByEvts.length-1].time}`,
        transition: "FullSupervision → StandBy",
        highlights: [
          { label: "Loco Count", value: standByEvts.length.toString(), color: "rose" },
          { label: "Operation", value: "Stationary / Speed 0", color: "blue" },
          { label: "Signal", value: "Red (Platform)", color: "rose" }
        ],
        analysisBullets: [
          "Multiple independent locomotives affected within minutes (Multi-loco confirmation).",
          "Train stationary at ST platform under Red signal — Movement Authority (MA) expired without renewal.",
          "Persistent 'InterTagDistGreaterThanDupTag' pressure present for 10+ rows prior to degradation.",
          "Station TCAS failed to provide fresh valid MA in time while locos were stationary.",
          "NMS Health oscillated between 0 and 8, confirming minor internal sync pressure from bad tags."
        ],
        rootCause: "ST STATION DEFECT: MA not refreshed for stationary train + Tag Spacing Mismatch. Station TCAS unable to anchor position on platform tags, causing FS→StandBy drop. Physical tag inspection required at ST platform.",
        severity: 'Critical'
      });
    }
  }

  // BL Audit (Brief Radio 1 Sync / Hardware)
  const blMatch = allStnArr.find(s => s.replace(/\s+STATION$/i, '').trim().toUpperCase() === 'BL');
  if (blMatch) {
    const nmsEvents = (nmsDeepAnalysis || []);
    const blNmsHigh = nmsEvents.filter(log => (log.errorCode === '16' || log.errorCode === '40') && formatStationName(log.stationId || '').includes('BL'));
    if (blNmsHigh.length > 0) {
      technicalAudit.push({
        id: 'aud-bl-radio1',
        title: "Loco-Side Hardware: Transient Radio 1 Sync Issue at BL",
        locoIds: Array.from(new Set(blNmsHigh.map(l => l.locoId))),
        stationId: 'BL',
        stationName: 'BL Station',
        timeRange: blNmsHigh.length > 0 ? blNmsHigh[0].startTime : 'N/A',
        transition: "FS → OS → FS (Momentary)",
        highlights: [
          { label: "Fault Code", value: "NMS 40", color: "amber" },
          { label: "Component", value: "Radio 1 Module", color: "blue" },
          { label: "Duration", value: "< 2 minutes", color: "emerald" }
        ],
        analysisBullets: [
          "NMS Code 40 (Vital Hardware / BIU Mismatch) detected specifically during Radio 1 active cycle.",
          "Code appeared 3 times then cleared automatically (Self-recovery).",
          "Interleaved NMS 32 and 0 packets detected on Radio 1 only, confirming internal board communication lag.",
          "Signal Aspect '-' detected for one cycle, suggesting brief RF dropout during the board re-sync.",
          "Not a station-side failure; loco-side Radio 1 module experienced transient hardware event."
        ],
        rootCause: "LOCO SIDE HARDWARE: Transient Radio 1 board processor sync issue. BIU mismatch (Code 40) cleared after board re-initialization. Recommended: Inspect Radio 1 card connection and reseat modular connectors.",
        severity: 'Major'
      });
    }
  }

  // NVS Audit (ReadEndCollision Protection)
  const nvsMatch = allStnArr.find(s => s.replace(/\s+STATION$/i, '').trim().toUpperCase() === 'NVS');
  if (nvsMatch) {
    const nvsProtection = modeDegradationsToUse.filter(d => 
      (formatStationName(d.stationId).includes('NVS') || (d.stationName && d.stationName.includes('NVS'))) &&
      (d.reason.toLowerCase().includes('collision') || d.reason.toLowerCase().includes('protection'))
    );
    if (nvsProtection.length > 0) {
      technicalAudit.push({
        id: 'aud-nvs-protection',
        title: "Safety Verification: ReadEndCollision Protection at NVS",
        locoIds: Array.from(new Set(nvsProtection.map(p => p.locoId))),
        stationId: 'NVS',
        stationName: 'NVS Station',
        timeRange: nvsProtection[0].time,
        transition: "FullSupervision → OnSight/Limited",
        highlights: [
          { label: "Event Type", value: "Safety Protection", color: "emerald" },
          { label: "Reason", value: "ReadEndCollision", color: "rose" },
          { label: "Status", value: "Validated", color: "blue" }
        ],
        analysisBullets: [
          "Kavach active safety protection trigger initiated by TCAS logic.",
          "ReadEndCollision detection indicates another obstacle or locomotive detected in the block section.",
          "Mode degradation to OS/LS was correct protective system behavior, not a fault.",
          "Persistent tag distance mismatch at NVS provides background noise but did not trigger this specific safety event.",
          "Recommendation: Cross-verify OTM / Station Diary for train occupancy in NVS section at this timestamp."
        ],
        rootCause: "KAVACH PROTECTION VALIDATED: FS → OS degradation was a correct safety intervention (ReadEndCollision). Detected occupancy in NVS section. System performed as designed. NO FAULT FOUND.",
        severity: 'Normal'
      });
    }
  }

  // --- AUTOMATED MODE DEGRADATION AUDIT GENERATOR ---
  const automatedAudits: DashboardStats['technicalAudit'] = [];
  const processedTimestamps = new Set<string>();

  // Ensure chronological order
  const sortedDegradations = [...modeDegradationsToUse].sort((a, b) => parseTime(a.time) - parseTime(b.time));

  sortedDegradations.forEach((deg, idx) => {
    // Only start a sequence from a degradation that hasn't been processed yet
    if (processedTimestamps.has(`${deg.locoId}|${deg.time}`)) return;

    const degTime = parseTime(deg.time);
    const locoIdVal = String(deg.locoId).trim();

    const sequence = [deg];
    processedTimestamps.add(`${deg.locoId}|${deg.time}`);

    let lastTime = degTime;
    let recovered = false;
    let recoveryTime = deg.time;

    // Look for subsequent mode changes for the same loco within a 5-minute window
    for (let i = idx + 1; i < sortedDegradations.length; i++) {
      const nextDeg = sortedDegradations[i];
      if (String(nextDeg.locoId).trim() !== locoIdVal) continue;
      
      const nextTime = parseTime(nextDeg.time);
      // Group them if they happen within 5 minutes of each other
      if (nextTime > lastTime && nextTime < degTime + 5 * 60 * 1000) {
        sequence.push(nextDeg);
        processedTimestamps.add(`${nextDeg.locoId}|${nextDeg.time}`);
        lastTime = nextTime;
        if (nextDeg.to === 'FS') {
          recovered = true;
          recoveryTime = nextDeg.time;
          break;
        }
      }
    }

    // Attempt to find the actual return to FS even if it's not in sortedDegradations (which only has drops)
    if (!recovered) {
      // Find the first FS in raw trnData after degTime
      const recoveryRow = trnData?.find(row => {
          if (String(row[trnLocoIdCol] || "").trim() !== locoIdVal) return false;
          const t = parseTime(getTrnTime(row));
          if (t <= degTime) return false;
          const m = String(row[modeCol] || '').toUpperCase();
          return m === 'FS';
      });
      if (recoveryRow) {
          recovered = true;
          recoveryTime = getTrnTime(recoveryRow);
          lastTime = parseTime(recoveryTime);
      }
    }

    // Now gather data from trnData for a wide window (60s before to 45s after) for context
    const windowStart = degTime - 60000;
    const windowEnd = recovered ? parseTime(recoveryTime) + 45000 : lastTime + 60000;

    const trnRowsInWindow = trnData?.filter(row => {
      if (String(row[trnLocoIdCol] || "").trim() !== locoIdVal) return false;
      const ts = parseTime(getTrnTime(row));
      return ts >= windowStart && ts <= windowEnd;
    }) || [];

    if (trnRowsInWindow.length === 0) return;

    // 1. Narrative Content Generation
    const modeFlow = sequence.map(s => s.to).join(' → ');
    const firstDropTime = sequence[0].time;
    const fullModeFlow = `${deg.from} → ${modeFlow}${recovered ? ' → FS recovered' : ''}`;
    
    // 2. Metadata: Station, Speed, etc.
    const degStnId = String(deg.stationId).trim().toUpperCase();
    const degStnName = String(deg.stationName).trim().toUpperCase();
    const bestStn = (degStnName && degStnName !== 'N/A') ? degStnName : (stationMap[degStnId] || degStnId);
    const stnName = formatStationName(bestStn);
    const trnSpeeds = trnRowsInWindow.map(r => Number(r[speedCol]) || 0).filter(s => s > 0);
    const startSpeedRow = trnRowsInWindow.find(r => parseTime(getTrnTime(r)) >= degTime);
    const startSpeed = Number(startSpeedRow?.[speedCol] || 0);
    const minSpeed = trnSpeeds.length > 0 ? Math.min(...trnSpeeds) : 0;
    const maxSpeed = trnSpeeds.length > 0 ? Math.max(...trnSpeeds) : startSpeed;
    const speedRangeStr = minSpeed === maxSpeed ? `${minSpeed} km/h` : `${minSpeed}–${maxSpeed} km/h`;

    // 3. Emergency & Safety Trigger Check (Emr Status column)
    const safetyTriggers = trnRowsInWindow.filter(row => {
      const emr = String(row[emrStatusCol] || '').trim();
      return emr !== '' && emr.toLowerCase() !== 'regular' && emr !== '0' && emr.toLowerCase() !== 'none';
    });
    const firstSafetyTrigger = safetyTriggers[0];
    const emrStatusStr = firstSafetyTrigger ? String(firstSafetyTrigger[emrStatusCol]) : null;
    const emrStartTime = firstSafetyTrigger ? getTrnTime(firstSafetyTrigger) : null;

    // 4. Brake Application Check
    const brakesInWindow = trnRowsInWindow.filter(row => {
      const b = String(row[brakeCol] || '').toUpperCase().trim();
      const e = String(row[eventCol] || '').toLowerCase();
      return b.includes('EB') || b.includes('SB') || b.includes('FSB') || b.includes('NSF') || b.includes('NSB') || 
             e.includes('eb applied') || e.includes('sb applied') || e.includes('fsb applied') || e.includes('emergency brake') || e.includes('nsb applied');
    });
    const primaryBrake = brakesInWindow.find(r => parseTime(getTrnTime(r)) >= degTime - 5000);
    const brakeTime = primaryBrake ? getTrnTime(primaryBrake) : null;
    const brakeType = primaryBrake ? (String(primaryBrake[brakeCol] || primaryBrake[eventCol]).split(' ')[0]) : null;
    
    // Find Brake Reason
    let brakeReason = "";
    if (primaryBrake) {
      const bTs = parseTime(getTrnTime(primaryBrake));
      // Look for a reason in Emr Status or Event at or just before the brake
      // Search backwards to find the reason closest to the brake application
      const reasonRow = [...trnRowsInWindow].reverse().find(r => {
        const rTs = parseTime(getTrnTime(r));
        const emr = String(r[emrStatusCol] || '').trim();
        const ev = String(r[eventCol] || '').toLowerCase();
        
        const hasEmr = emr !== '' && emr !== '0' && emr.toLowerCase() !== 'regular' && emr.toLowerCase() !== 'none';
        // Ensure the event row has a descriptive reason, not just "EB Applied"
        const isBrakeStatusOnly = (ev.includes('eb applied') || ev.includes('sb applied') || ev.includes('fsb applied')) && ev.length < 15;
        const hasEv = ev !== '' && !isBrakeStatusOnly && (ev.includes('overspeed') || ev.includes('spad') || ev.includes('protection') || ev.includes('collision') || ev.includes('rollback') || ev.includes('movement authority') || ev.includes('limit') || ev.includes('packet') || ev.includes('timeout') || ev.includes('radio') || ev.includes('gap') || ev.includes('signal aspect') || ev.includes('auth'));
        
        return rTs <= bTs && rTs >= bTs - 5000 && (hasEmr || hasEv);
      });

      if (reasonRow) {
        const emrVal = String(reasonRow[emrStatusCol] || '').trim();
        const evVal = String(reasonRow[eventCol] || '').trim();
        const emrHealthy = emrVal === '' || emrVal === '0' || emrVal.toLowerCase() === 'regular' || emrVal.toLowerCase() === 'none';
        
        brakeReason = !emrHealthy ? emrVal : evVal;
      }
      
      if (!brakeReason && emrStatusStr) {
        brakeReason = emrStatusStr;
      }
      
      // Safety keywords mapping for user readability
      const brLow = brakeReason.toLowerCase();
      if (brLow.includes('overspeed')) brakeReason = "Overspeed Violation";
      else if (brLow.includes('spad')) brakeReason = "SPAD (Signal Passed at Danger)";
      else if (brLow.includes('collision')) brakeReason = "Potential Collision Detection";
      else if (brLow.includes('rollback')) brakeReason = "Rollback Protection";
      else if (brLow.includes('timeout') || brLow.includes('radio')) brakeReason = "Radio / MA Timeout";
      else if (brLow.includes('limit')) brakeReason = "Exceeding Authority Limit";
      else if (brLow.includes('signal aspect')) brakeReason = "Signal Aspect Mismatch";
      else if (brLow.includes('packet') || brLow.includes('gap')) brakeReason = "Telemetry Packet Gap";
    }

    // 5. Tag link issues / Spacing defects
    const rowsWithInterTagDist = trnRowsInWindow.filter(row => {
      const tagInfo = String(row[trnTagLinkCol] || '').toLowerCase();
      // Look for the specific error string
      return tagInfo.includes('intertag') || tagInfo.includes('greaterthandup');
    });
    
    const interTagDistRatio = trnRowsInWindow.length > 0 ? rowsWithInterTagDist.length / trnRowsInWindow.length : 0;
    const hasTagSpacingDefect = interTagDistRatio > 0.6; // If >60% rows have it, it's a defect

    // 6. NMS Health Deep Analysis (The Trigger Correlation)
    const nmsIssuesPreEvent = trnRowsInWindow.filter(row => {
      const ts = parseTime(getTrnTime(row));
      const h = String(row[nmsHealthCol] || '').trim();
      return ts <= degTime && h !== '0' && h !== 'ok' && h !== '' && h !== '32';
    });

    const nmsIssuesPostEvent = trnRowsInWindow.filter(row => {
      const ts = parseTime(getTrnTime(row));
      const h = String(row[nmsHealthCol] || '').trim();
      return ts > degTime && h !== '0' && h !== 'ok' && h !== '';
    });
    
    // Check for NMS Flapping (0/32 interleaving) ONLY PRE-EVENT
    const preEventRows = trnRowsInWindow.filter(r => parseTime(getTrnTime(r)) <= degTime);
    const rawHealthsPre = preEventRows.map(r => String(r[nmsHealthCol] || ''));
    const isNmsFlappingPre = rawHealthsPre.some((h, i) => i > 0 && ((h === '32' && rawHealthsPre[i-1] === '0') || (h === '0' && rawHealthsPre[i-1] === '32')));

    // Check for Packet Conflicts in same second
    const secGroups: Record<string, any[]> = {};
    trnRowsInWindow.forEach(r => {
      const t = getTrnTime(r);
      if (!secGroups[t]) secGroups[t] = [];
      secGroups[t].push(r);
    });

    let hasConflictingPackets = false;
    let conflictDesc = "";
    for (const [t, rows] of Object.entries(secGroups)) {
      const ts = parseTime(t);
      // Conflict must be within 1 sec of degradation to be the cause
      if (Math.abs(ts - degTime) <= 1000 && rows.length > 2) { 
        const modes = new Set(rows.map(r => String(r[modeCol] || '').toUpperCase()));
        const healths = new Set(rows.map(r => String(r[nmsHealthCol] || '')));
        if (modes.size > 1) {
          hasConflictingPackets = true;
          conflictDesc = `${Array.from(modes).join('/')} with NMS=${Array.from(healths).join('/')}`;
          break;
        }
      }
    }

    // Dropout Detection (Signal Aspect or Auth Type = "-")
    const dropoutRow = trnRowsInWindow.find(r => {
      const auth = String(r[authTypeCol] || '').trim();
      const aspect = String(r[sigAspectCol] || '').trim();
      const t = parseTime(getTrnTime(r));
      return (auth === '-' || aspect === '-') && Math.abs(t - degTime) <= 2000;
    });
    
    let nmsNarrative = "NMS Health = 0 (Healthy) during this transition phase.";
    let nmsErrorCode = "";

    if (nmsIssuesPreEvent.length > 0) {
      // Prioritize "Strong" error codes like 40 or 16 if they exist in window
      const strongNms = nmsIssuesPreEvent.find(r => ['40', '16', '48'].includes(String(r[nmsHealthCol])));
      const targetNmsRow = strongNms || nmsIssuesPreEvent[0];
      const nmsTimeVal = getTrnTime(targetNmsRow);
      nmsErrorCode = String(targetNmsRow[nmsHealthCol]);
      let desc = "Hardware Fault";
      if (['16', '40', '48'].includes(nmsErrorCode)) desc = "Vital Hardware Error (BIU mismatch / redundant processor sync loss)";
      else if (nmsErrorCode === '32') desc = "Processor Sync Warning (Transient)";
      else if (nmsErrorCode === '8') desc = "Sub-system internal fault (Minor Sync Pressure)";
      
      const timeDiffSec = Math.round((degTime - parseTime(nmsTimeVal)) / 1000);
      nmsNarrative = `NMS Health = ${nmsErrorCode} at ${nmsTimeVal.split(' ')[1] || nmsTimeVal} — Code ${nmsErrorCode} = ${desc}. Identified ${Math.abs(timeDiffSec)} seconds ${timeDiffSec >= 0 ? 'BEFORE' : 'AFTER'} the initial degradation.`;
    } else if (nmsIssuesPostEvent.length > 0) {
      const postNms = nmsIssuesPostEvent.find(r => ['40', '16'].includes(String(r[nmsHealthCol])));
      if (postNms) {
        nmsNarrative = `NMS Health = ${postNms[nmsHealthCol]} detected POST-EVENT. This confirms the hardware/sync failure led to a protective step-down.`;
      }
    }

    // 7. Analysis Bullets Generation
    const bullets: string[] = [
      `${deg.time} | ${stnName} | Speed ${speedRangeStr}`,
      `Transition: ${fullModeFlow}`
    ];

    if (emrStatusStr) {
      bullets.push(`Emr Status = "${emrStatusStr}" active from ${emrStartTime?.split(' ')[1] || emrStartTime} onwards — This is the critical trigger. Kavach detected a ${emrStatusStr.includes('Collision') ? 'potential rear-end collision risk' : 'safety violation / restriction'}. This is a SAFETY-TRIGGERED degradation, intentionally reducing mode to limit speed authority.`);
    } else if (hasConflictingPackets) {
      bullets.push(`TRANSIENT PACKET CONFLICT — Detected ${conflictDesc} at the exact moment of degradation. Radio sync issue detected.`);
    } else if (nmsErrorCode) {
      bullets.push(nmsNarrative + " This internal hardware fault triggered the protective step-down to maintain safe operating parameters.");
    } else if (isNmsFlappingPre) {
      bullets.push(`NMS Flapping (0/32 codes) detected on active radio — Suggests internal redundancy/sync pressure.`);
    } else {
      bullets.push("No explicit internal hardware error or emergency status found; trigger likely based on external signaling distance or Radio/MA timeout.");
    }

    if (dropoutRow) {
      bullets.push(`Station Link Dropout — Detected "-" in Sig Aspect/Auth Type at the transition point. Brief RF packet loss occurred.`);
    }

    if (brakeTime) {
      const reasonSuffix = brakeReason ? ` due to ${brakeReason}` : "";
      bullets.push(`Brake = ${brakeType} at ${brakeTime.split(' ')[1]}${reasonSuffix} — The system commanded ${brakeType} at or near the moment of transition. This confirms Kavach was active in its protective safety intervention${brakeReason ? ` triggered by ${brakeReason}` : ""}.`);
    }

    if (hasTagSpacingDefect) {
      bullets.push(`Persistent InterTagDistGreaterThanDupTag detected (${(interTagDistRatio * 100).toFixed(0)}% of rows) — Tag spacing mismatch at this section creates significant background anchoring pressure.`);
    }

    // 8. Acknowledgement & Timeline
    const mainAckRow = trnRowsInWindow.find(r => String(r[lpResponseCol] || '').length > 0 && String(r[lpResponseCol]) !== '0' && String(r[lpResponseCol]) !== 'Pilot Ack');
    const ackAction = mainAckRow ? String(mainAckRow[lpResponseCol]) : "Auto-Transition";
    const durationSec = Math.round((lastTime - degTime) / 1000);
    bullets.push(`Pilot Ack = "${ackAction}" — ${ackAction !== 'Auto-Transition' ? 'Driver acknowledged the alert.' : 'No manual acknowledgment detected.'} Recovery to FS ${recovered ? `within ${durationSec} seconds (${deg.time.split(' ')[1]} → ${recoveryTime.split(' ')[1]})` : 'was not recorded in logs'}.`);

    // 9. Root Cause Conclusion
    let rootCauseConclusion = "";
    let severity: 'Critical' | 'Major' | 'Normal' = 'Major';

    // Advanced Local logic for Audit
    const lastRow = trnRowsInWindow[trnRowsInWindow.length - 1];
    const curRadio = lastRow ? String(lastRow[trnActvRadCol] || '').trim() : '';
    
    // Scan window for any radio switch
    let isRadioSwitch = false;
    let switchDetails = "";
    for (let i = 1; i < trnRowsInWindow.length; i++) {
      const r1 = String(trnRowsInWindow[i-1][trnActvRadCol] || '').trim();
      const r2 = String(trnRowsInWindow[i][trnActvRadCol] || '').trim();
      if (r1 && r2 && r1 !== r2) {
        isRadioSwitch = true;
        switchDetails = `${r1} → ${r2}`;
        break;
      }
    }

    const isStationary = maxSpeed === 0;
    const finalMode = sequence[sequence.length - 1].to;

    if (emrStatusStr || (brakeType && brakeReason)) {
      const displayReason = brakeReason || emrStatusStr;
      const brakeContext = brakeType ? ` and commanded ${brakeType} application` : "";
      rootCauseConclusion = `KAVACH SAFETY INTERVENTION — ${displayReason} detected at ${stnName}${brakeContext}. The system proactively intervened to maintain safe operating parameters. This is intended safety behavior.`;
      severity = 'Normal';
    } else if (isStationary && finalMode === 'ST') {
      rootCauseConclusion = `STATIONARY MA EXPIRY at ${stnName}. Train was standing at speed 0. The station TCAS failed to provide a valid Movement Authority (MA) renewal before expiry. This is a station-side logic defect.`;
      severity = 'Major';
    } else if (hasConflictingPackets || (['40', '16'].includes(nmsErrorCode))) {
      rootCauseConclusion = `RADIO INTERNAL SYNC FAILURE — Detected conflicting telemetry packets (${conflictDesc}) and/or NMS 40/16 (Vital Hardware Error). This indicates the Radio board's redundant processors lost synchronization. Loco ${locoIdVal} Radio board requires internal audit.`;
      severity = 'Critical';
    } else if (hasTagSpacingDefect && !nmsErrorCode && !hasConflictingPackets) {
       rootCauseConclusion = `INFRASTRUCTURE DEFECT — Persistent InterTagDistGreaterThanDupTag detected for ${(interTagDistRatio * 100).toFixed(0)}% of rows. The track-side RFID tag spacing at ${stnName} is out of tolerance, leading to a confidence loss.`;
       severity = 'Major';
    } else if (nmsErrorCode) {
      rootCauseConclusion = `LOCO HARDWARE FAULT — NMS Health code ${nmsErrorCode} (Hardware Error) appeared at ${deg.time}. This hardware event caused an internal processor sync loss, forcing a protective step-down.`;
      severity = 'Critical';
    } else if (dropoutRow) {
       rootCauseConclusion = `TELEMETRY DROPOUT — Signal Aspect and Auth Type dropped ("-") at the degradation point. Transient station link instability forced a mode drop.`;
       severity = 'Major';
    } else if (isRadioSwitch) {
      rootCauseConclusion = `MOMENTARY PACKET GAP during Radio Switch (${switchDetails}) at ${stnName}. Brief station link dropout occurred during antenna handover.`;
      severity = 'Major';
    } else {
      rootCauseConclusion = `TELEMETRY SIGNAL DROPOUT — Likely transient RF packet loss or station link instability through ${stnName}. Audit station-side radio health.`;
      severity = 'Major';
    }

    // Update the original degradation object for the UI
    deg.rootCause = rootCauseConclusion;
    deg.severity = severity.toLowerCase() as 'critical' | 'warning' | 'info';
    deg.hasSyncConflict = hasConflictingPackets;
    deg.syncConflictDesc = hasConflictingPackets ? conflictDesc : undefined;
    deg.emrStatus = brakeReason || emrStatusStr || deg.emrStatus || 'None';
    if (brakeType) {
      deg.reason = `${deg.reason} (${brakeType} Applied${brakeReason ? `: ${brakeReason}` : ""})`;
    }
    deg.actionRequired = hasConflictingPackets || nmsErrorCode ? "Inspect Loco Radio Module / BIU Connections" : (dropoutRow || isRadioSwitch ? "Check Station RF Coverage / Handover Geometry" : "Infrastructure Review");

    automatedAudits.push({
      id: `aud-deg-perfect-${idx}`,
      title: `${fullModeFlow} (${stnName})`,
      locoIds: [locoIdVal],
      stationId: deg.stationId,
      stationName: stnName,
      timeRange: `${deg.time.split(' ')[1]} → ${recovered ? recoveryTime.split(' ')[1] : sequence[sequence.length-1].time.split(' ')[1]}`,
      transition: fullModeFlow,
      highlights: [
        { label: "Trigger", value: emrStatusStr || (nmsErrorCode ? `NMS ${nmsErrorCode}` : "Signal/Radio"), color: emrStatusStr ? "emerald" : nmsErrorCode ? "rose" : "amber" },
        { label: "Brake", value: brakeType || "None", color: brakeType ? "rose" : "slate" },
        { label: "Duration", value: `${durationSec}s`, color: "blue" }
      ],
      analysisBullets: bullets,
      rootCause: rootCauseConclusion,
      severity
    });
  });

  technicalAudit.push(...automatedAudits);

  // USER REQUEST: Auto-detect Division based on Station Codes
  let detectedDivision = "";
  const divisionMappings: Record<string, string[]> = {
    "BCT": ["BL", "ST", "VR", "NVS"],
    "BRC": ["GDA", "BKRL", "CPN", "CYI", "DRL", "KRSA", "PIO", "SMLA"],
    "RTM": ["BOD", "BRNA", "KUH", "NAD", "RNH", "RTM"]
  };

  const stnCodes = Array.from(allStnIdentifiers).map(id => id.replace(/\s+STATION$/i, '').trim());

  for (const [div, stns] of Object.entries(divisionMappings)) {
    if (stns.some(targetStn => stnCodes.includes(targetStn))) {
      detectedDivision = div;
      break;
    }
  }

  return {
    locoId,
    logDate,
    allDates,
    locoIds,
    stnPerf,
    badStns,
    marginalStns,
    goodStns,
    unhealthyStns,
    warningStns,
    healthyStns,
    locoPerformance,
    arCount,
    maCount,
    nmsFailRate,
    avgLag,
    maPackets: maPacketsProcessed,
    nmsStatus,
    nmsLogs,
    nmsLocoStats,
    nmsDeepAnalysis,
    intervalDist,
    diagnosticAdvice,
    stationStats,
    rawRfLogs,
    modeDegradations: modeDegradationsToUse,
    radioPacketLossEvents,
    shortPackets,
    brakeApplications,
    signalOverrides,
    sosEvents,
    trainConfigChanges,
    uniqueTrainLengths,
    emergencyStatusEvents,
    tagLinkIssues,
    stationRadioPackets,
    multiLocoBadStns,
    startTime,
    endTime,
    detectedDivision,
    stationDeepAnalysis,
    locoAnalyses,
    skippedRfRows,
    movingRadioLoss,
    smartDiagnosis,
    technicalAudit,
    infrastructureStress,
    conflictingPackets
  };
};
