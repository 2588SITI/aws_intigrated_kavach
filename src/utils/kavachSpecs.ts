export interface KavachSpecification {
  title: string;
  version: string;
  effectiveDate: string;
  authority: string;
  amendmentNo: number;
  annexures: Record<string, {
    title: string;
    description: string;
    amendment: string;
    parameters?: Array<{ name: string; desc: string; defaultValue: string; limitRange: string; unit: string }>;
    dataFormats?: Array<{ field: string; bits: string; defaultValue?: string; desc: string }>;
    protocols?: Array<{ msgType: string; value: string; purpose: string }>;
  }>;
  rfProtocol: {
    frameCycleMs: number;
    slotsCount: number;
    slotWidthMs: number;
    channelRangeMHz: string;
    sofs: {
      radio1: string;
      radio2: string;
    };
    formulas: {
      lDoubtOver: string;
      lDoubtUnder: string;
    };
  };
  packetTypes: Record<string, { code: string; desc: string }>;
  locoModes: Record<string, { code: string; desc: string; fallbackRule?: string }>;
  emergencyStatus: Record<string, { code: string; desc: string }>;
  tagLinkInfo: Record<string, { code: string; desc: string }>;
  radioFailureFallbacks: {
    withProfile: string;
    withoutProfile: string;
    ackTimeoutSec: number;
    failAction: string;
    stalePacketBlankSec: number;
  };
}

export const KAVACH_OFFICIAL_SPECS: KavachSpecification = {
  title: "Specification of Kavach (The Indian Railway ATP) - System Requirements & Communication Protocols",
  version: "RDSO/SPN/196/2020 Version 4.0",
  effectiveDate: "27.06.2024",
  authority: "Signal & Telecom Directorate, RDSO, Ministry of Railways",
  amendmentNo: 10,
  annexures: {
    "Annexure-A1": {
      title: "Mode Transitions, SOS & MA Handling",
      description: "Controls the state transitions of Onboard Kavach based on trackside telemetry, radio status, and driver input. Outlines automatic brake interventions for SPAD or collision hazard situations.",
      amendment: "Amdt-6",
      protocols: [
        { msgType: "SB to SR", value: "Condition <7, -p4-", purpose: "Power on, cab changed, or self-test success." },
        { msgType: "SR to FS", value: "Condition <17, 23, 30, 39, 40, 44, 85, p5-", purpose: "On sight of valid track profile, MA, and healthy radio." },
        { msgType: "FS to LS", value: "Condition <81, -p8-", purpose: "Safe degraded mode after active radio failure if track profile remains valid (< 3000m)." },
        { msgType: "FS to TR", value: "Condition <44", purpose: "SPAD or Authority overrun triggers Emergency Brake to trip." }
      ]
    },
    "Annexure-A2": {
      title: "Onboard KAVACH Configurable Parameters",
      description: "Defines runtime thresholds, timeouts, and brake application parameters stored inside the Onboard Kavach Vital Computer (2-out-of-2 architecture).",
      amendment: "Amdt-3",
      parameters: [
        { name: "Speed Margin - Warning", desc: "Speed beyond permitted speed after which visual/audio warning triggers", defaultValue: "2", limitRange: "0 to 10", unit: "kmph" },
        { name: "Speed Margin - NSB", desc: "Speed beyond permitted speed after which Normal Service Brake is applied", defaultValue: "5", limitRange: "5 to 10", unit: "kmph" },
        { name: "Speed Margin - FSB", desc: "Speed beyond permitted speed after which Full Service Brake is applied", defaultValue: "3", limitRange: "2 to 10", unit: "kmph" },
        { name: "Speed Margin - EB", desc: "Speed beyond permitted speed after which Emergency Brake is triggered", defaultValue: "2", limitRange: "2 to 15", unit: "kmph" },
        { name: "LP Reaction Time", desc: "Loco pilot reaction time margin inside absolute or automatic block before auto intervention", defaultValue: "8", limitRange: "4 to 30", unit: "seconds" },
        { name: "Communication Timeout (Absolute Block)", desc: "Time with missing radio packets before transiting to LS/SR mode", defaultValue: "30", limitRange: "6 to 120", unit: "seconds" },
        { name: "Communication Timeout (Automatic Block)", desc: "Time with missing radio packets before transiting to LS/SR mode", defaultValue: "10", limitRange: "6 to 120", unit: "seconds" },
        { name: "Wheel roll-away / roll-back distance", desc: "Brake trigger threshold for uncommanded reverse movement", defaultValue: "10", limitRange: "5 to 30", unit: "meters" },
        { name: "Rear End collision margin", desc: "Bumper separation threshold for suburban track section and EMUs", defaultValue: "100", limitRange: "100 to 500", unit: "meters" }
      ]
    },
    "Annexure-A3": {
      title: "Stationary KAVACH Configurable Parameters",
      description: "Defines station boundary, OHE line details, shunting limit, and heartbeat settings managed by the SVK.",
      amendment: "Amdt-3",
      parameters: [
        { name: "Station Traffic Capacity", desc: "Max number of active locomotives Stationary Kavach can monitor", defaultValue: "44", limitRange: "1 to 44", unit: "locos" },
        { name: "Kmax channel delay (Tmax)", desc: "Max permissible delay for a RaSTA package transmission", defaultValue: "1800", limitRange: "100 to 3000", unit: "ms" },
        { name: "Heartbeat Interval (Th)", desc: "Cyclic time interval for Stationary-to-Stationary heartbeats", defaultValue: "300", limitRange: "100 to 1000", unit: "ms" },
        { name: "Heartbeat timeout (Tseq)", desc: "Max age of message before dropping package connection", defaultValue: "100", limitRange: "10 to 500", unit: "ms" },
        { name: "Signal flickering timeout", desc: "Hold time for exit signal aspect transition to avoid hunting", defaultValue: "2000", limitRange: "2000 to 10000", unit: "msec" },
        { name: "Track circuit failure declaring", desc: "Interval of track circuit mismatch before marking block as occupied", defaultValue: "180", limitRange: "30 to 300", unit: "seconds" }
      ]
    },
    "Annexure-D": {
      title: "RFID Tag Data Format",
      description: "Bit layout and encoding of the track-side 128-bit passive tag. Read by Onboard reader antenna to verify absolute position and TIN.",
      amendment: "Amdt-7",
      dataFormats: [
        { field: "Type of Tag", bits: "X3 - X0 (4 bits)", defaultValue: "9 (Normal)", desc: "Defines function: 9=Normal, 10=LC Gate, 11=Adjacent Line, 12=Adjustment/Junction." },
        { field: "KAVACH Version", bits: "X5 - X4 (2 bits)", defaultValue: "1 (Spec 4.0)", desc: "0=Spec 3.2, 1=Spec 4.0, 2-3=Reserved." },
        { field: "Unique ID of RFID Tag set", bits: "X15 - X6 (10 bits)", desc: "Differentiates tag pairs. Range 1 to 1023." },
        { field: "Absolute Location", bits: "X38 - X16 (23 bits)", desc: "Geographical center offset location in meters. Range 0 to 8,388,607m." },
        { field: "TIN (Nominal direction)", bits: "X46 - X39 (8 bits)", desc: "Track Identification Number in Nominal (increasing location)." },
        { field: "TIN (Reverse direction)", bits: "X54 - X47 (8 bits)", desc: "Track Identification Number in Reverse (decreasing location)." },
        { field: "Tag Placement Details", bits: "Y30 - Y27 (4 bits)", desc: "0000=In-line, 0001=Signal standard nominal, 0011=Turnout, 1000=Dead Stop." },
        { field: "Tag Duplication Status", bits: "Y31 (1 bit)", desc: "0=Main Tag, 1=Duplicate Tag." },
        { field: "CRC Checksum", bits: "Y63 - Y34 (30 bits)", desc: "Cyclic Redundancy Checksum using CRC-30 polynomial." }
      ]
    },
    "Annexure-G": {
      title: "Network Monitoring System (NMS) Protocol",
      description: "Dictates the hex packets generated for centralized Intelligent NMS diagnostic telemetry over secure GPRS/LTE networks.",
      amendment: "Amdt-4",
      protocols: [
        { msgType: "0x11", value: "Stationary KAVACH info Message", purpose: "Transmits Station Regular, Access Authority or Emergency packets." },
        { msgType: "0x12", value: "Loco KAVACH Position Message", purpose: "Transmits cyclic (2s) position reports, speed, and active communication channel." },
        { msgType: "0x17", value: "Stationary KAVACH Health Message", purpose: "Relays hardware temperature, power supply voltage, and radio status." },
        { msgType: "0x18", value: "Onboard KAVACH Health Packet", purpose: "Relays RSSI levels, GPS visibility, RFID reader health, and BIU telemetry." },
        { msgType: "0x19", value: "KAVACH Fault Message", purpose: "Provides list of up to 10 active diagnostic faults (e.g. system processor mismatched error)." }
      ]
    },
    "Annexure-P": {
      title: "S-KAVACH to S-KAVACH Inter-station Handover",
      description: "Governs the communication protocol between adjacent stations to safely hand over control of moving trains at boundary zones.",
      amendment: "Amdt-4",
      protocols: [
        { msgType: "0x0101", value: "Command PDI-Version Check", purpose: "Primary partner requests connection initialization and protocol alignment." },
        { msgType: "0x0103", value: "Heart Beat Message", purpose: "Periodic (0.1s to 2s) signal to verify connection live state." },
        { msgType: "0x0104", value: "Train Handover Request", purpose: "Transmits upcoming train's speed, length, and position to target station." },
        { msgType: "0x0105", value: "Train RRI Message", purpose: "Accepting station replies with active route and extended Movement Authority." },
        { msgType: "0x0106", value: "Train Taken Over Message", purpose: "Finalizes the handover as the train successfully passes the border RFID limit." }
      ]
    }
  },
  rfProtocol: {
    frameCycleMs: 2000,
    slotsCount: 70,
    slotWidthMs: 22.5,
    channelRangeMHz: "406 MHz to 470 MHz",
    sofs: {
      radio1: "0xF1 0xA5 0xC3",
      radio2: "0xF2 0xA5 0xC3"
    },
    formulas: {
      lDoubtOver: "Over-reading + 5m Location Accuracy of RFID Tag + 5% Odometer Error + Reader Offset in Rear End (ROR). This is used for safe-side target distance supervision (PSR, TSR, Linking, Rear-End Protection).",
      lDoubtUnder: "Under-reading + 5m Location Accuracy of RFID Tag + 5% Odometer Error + Reader Offset in Front End (RORF). This is used for discarding passed tags (PSR, TSR, Linking, Head-on Protection)."
    }
  },
  packetTypes: {
    "1001": { code: "1001", desc: "Station to Onboard Regular Packet" },
    "1010": { code: "1010", desc: "Onboard to Station Regular Packet" },
    "1011": { code: "1011", desc: "Access Authority Packet" },
    "1100": { code: "1100", desc: "Additional Emergency Packet" },
    "1101": { code: "1101", desc: "Onboard Access Request" }
  },
  locoModes: {
    "0001": { code: "0001", desc: "Stand_By (SB)", fallbackRule: "Default post-power-on or cabin change stop state." },
    "0010": { code: "0010", desc: "Staff_Responsible_Mode (SR)", fallbackRule: "Used when route/track profile unknown, or post radio failure with no active profile." },
    "0011": { code: "0011", desc: "Limited_Supervision (LS)", fallbackRule: "Safe degraded mode after active radio failure if track profile remains valid (< 3000m)." },
    "0100": { code: "0100", desc: "Full_Supervision (FS)", fallbackRule: "Normal mission execution with active MA and valid communication." },
    "0105": { code: "0101", desc: "Override (OVRD)", fallbackRule: "Temporary override to pass red signal under authorized conditions." },
    "0110": { code: "0110", desc: "On_Sight (OS)", fallbackRule: "Permits driving into potentially occupied block sections with strict speed supervision." },
    "0111": { code: "0111", desc: "Trip (TR)", fallbackRule: "Emergency halt state triggered upon SPAD or authority breach." }
  },
  emergencyStatus: {
    "000": { code: "000", desc: "No Emergency - Regular Operation" },
    "001": { code: "001", desc: "Side Collision (Unusual Stoppage)" },
    "010": { code: "010", desc: "SoS Call active" },
    "011": { code: "011", desc: "Roll Back Detected" },
    "100": { code: "100", desc: "Head-On Collision hazard" },
    "101": { code: "101", desc: "Rear-End Collision hazard" },
    "110": { code: "110", desc: "Parting SoS triggered" }
  },
  tagLinkInfo: {
    "000": { code: "000", desc: "No Tag missing (Normal linking)" },
    "001": { code: "001", desc: "Duplicate Tag missing" },
    "010": { code: "010", desc: "Main Tag missing" },
    "011": { code: "011", desc: "Both Tags missing" },
    "100": { code: "100", desc: "Tag position interchanged" },
    "101": { code: "101", desc: "Both Tags have same location info" },
    "110": { code: "110", desc: "Intertag distance less than DIST_DUP_TAG" },
    "111": { code: "111", desc: "Intertag distance greater than DIST_DUP_TAG (Critical Tag Spacing Defect)" }
  },
  radioFailureFallbacks: {
    withProfile: "Limited Supervision (LS) mode with driver acknowledgement required.",
    withoutProfile: "Staff Responsible (SR) mode.",
    ackTimeoutSec: 15,
    failAction: "Automatic application of Service Brakes (SB) via Brake Interface Unit (BIU).",
    stalePacketBlankSec: 6
  }
};

// 70 TDMA Time slots list representing Amendment-10 diagram
export interface TdmaSlot {
  id: string;
  name: string;
  pStart: number;
  widthMs: number;
  assignedTo: string;
  color: string;
  desc: string;
}

export const TDMA_SLOTS_AMDT10: TdmaSlot[] = [
  { id: "P1", name: "P1 (Reserve)", pStart: 0, widthMs: 22.5, assignedTo: "Reserved for system stabilization", color: "amber", desc: "Kept as reserve slot." },
  { id: "P2", name: "P2 (M-1)", pStart: 45, widthMs: 22.5, assignedTo: "Stationary KAVACH & Onboard", color: "emerald", desc: "First active communication slot. Starts exactly at 45ms from frame cycle start." },
  { id: "P3", name: "P3 (M-2)", pStart: 72.5, widthMs: 22.5, assignedTo: "Stationary KAVACH & Onboard", color: "emerald", desc: "Regular communication slot." },
  { id: "P4", name: "P4 (M-3)", pStart: 100, widthMs: 22.5, assignedTo: "Stationary KAVACH & Onboard", color: "emerald", desc: "Regular communication slot." },
  { id: "P43", name: "P43 (M-42)", pStart: 1172.5, widthMs: 22.5, assignedTo: "Stationary KAVACH & Onboard", color: "emerald", desc: "Regular communication slot." },
  { id: "P44", name: "P44 (M-43)", pStart: 1200, widthMs: 22.5, assignedTo: "Stationary KAVACH & Onboard", color: "emerald", desc: "Regular communication slot." },
  { id: "P45", name: "P45 (M-44)", pStart: 1227.5, widthMs: 22.5, assignedTo: "Stationary KAVACH & Onboard", color: "emerald", desc: "Last slot of first half block." },
  { id: "P46", name: "P46 (Reserve)", pStart: 1250, widthMs: 22.5, assignedTo: "Reserved for system stabilization", color: "amber", desc: "Middle reserve slot separator." },
  { id: "P47", name: "P47 (MBS-1)", pStart: 1320, widthMs: 22.5, assignedTo: "Onboard KAVACH Block Section Broadcast", color: "blue", desc: "Block section broadcast starts at 1320ms." },
  { id: "P48", name: "P48 (MBS-2)", pStart: 1347.5, widthMs: 22.5, assignedTo: "Onboard KAVACH Block Section Broadcast", color: "blue", desc: "Block section broadcast." },
  { id: "P53", name: "P53 (ME-1)", pStart: 1485, widthMs: 22.5, assignedTo: "Mobile Emergency (f0)", color: "rose", desc: "Mobile emergency transmission on frequency f0." },
  { id: "P54", name: "P54 (ME-2)", pStart: 1512.5, widthMs: 22.5, assignedTo: "Mobile Emergency (f0)", color: "rose", desc: "Mobile emergency transmission on frequency f0." },
  { id: "P55", name: "P55 (SE-1)", pStart: 1540, widthMs: 22.5, assignedTo: "Stationary Emergency (f0)", color: "orange", desc: "Stationary emergency (SoS) broadcast on frequency f0." },
  { id: "P56", name: "P56 (SE-2)", pStart: 1567.5, widthMs: 22.5, assignedTo: "Stationary Emergency (f0)", color: "orange", desc: "Stationary emergency (SoS) broadcast on frequency f0." },
  { id: "P57", name: "P57 (STS-1)", pStart: 1595, widthMs: 22.5, assignedTo: "Access Authority Broadcast (f0)", color: "purple", desc: "Stationary slot for transmitting Access Authority on frequency f0." },
  { id: "P58", name: "P58 (STS-2)", pStart: 1622.5, widthMs: 22.5, assignedTo: "Access Authority Broadcast (f0)", color: "purple", desc: "Stationary slot for transmitting Access Authority." },
  { id: "P59", name: "P59 (MBS-7)", pStart: 1650, widthMs: 22.5, assignedTo: "Onboard KAVACH Block Section Broadcast", color: "blue", desc: "Block section broadcast." },
  { id: "P60", name: "P60 (MBS-8)", pStart: 1677.5, widthMs: 22.5, assignedTo: "Onboard KAVACH Block Section Broadcast", color: "blue", desc: "Block section broadcast." },
  { id: "P65", name: "P65 (ME-3)", pStart: 1815, widthMs: 22.5, assignedTo: "Mobile Emergency (f0)", color: "rose", desc: "Mobile emergency transmission on frequency f0." },
  { id: "P66", name: "P66 (ME-4)", pStart: 1842.5, widthMs: 22.5, assignedTo: "Mobile Emergency (f0)", color: "rose", desc: "Mobile emergency transmission on frequency f0." },
  { id: "P67", name: "P67 (SE-3)", pStart: 1870, widthMs: 22.5, assignedTo: "Stationary Emergency (f0)", color: "orange", desc: "Stationary emergency (SoS) broadcast on frequency f0." },
  { id: "P68", name: "P68 (SE-4)", pStart: 1897.5, widthMs: 22.5, assignedTo: "Stationary Emergency (f0)", color: "orange", desc: "Stationary emergency (SoS) broadcast on frequency f0." },
  { id: "P69", name: "P69 (STS-3)", pStart: 1925, widthMs: 22.5, assignedTo: "Access Authority Broadcast (f0)", color: "purple", desc: "Stationary slot for transmitting Access Authority on frequency f0." },
  { id: "P70", name: "P70 (STS-4)", pStart: 1952.5, widthMs: 22.5, assignedTo: "Access Authority Broadcast (f0)", color: "purple", desc: "Stationary slot for transmitting Access Authority. Ends cycle." }
];
