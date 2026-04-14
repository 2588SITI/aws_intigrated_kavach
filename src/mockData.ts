import { Train, Alert, SignalHealth } from './types';

export const mockTrains: Train[] = [
  {
    id: 'T-101',
    name: 'Rajdhani Express (12301)',
    speed: 125,
    maxSpeed: 130,
    location: 'Near Kanpur Central',
    status: 'Normal',
    signalStrength: 98,
    lastUpdate: '2026-04-14T10:40:00Z',
  },
  {
    id: 'T-202',
    name: 'Shatabdi Express (12002)',
    speed: 110,
    maxSpeed: 150,
    location: 'Approaching Agra Cantt',
    status: 'Normal',
    signalStrength: 95,
    lastUpdate: '2026-04-14T10:40:05Z',
  },
  {
    id: 'T-303',
    name: 'Vande Bharat (22436)',
    speed: 155,
    maxSpeed: 160,
    location: 'Departing New Delhi',
    status: 'Warning',
    signalStrength: 72,
    lastUpdate: '2026-04-14T10:40:10Z',
  },
  {
    id: 'T-404',
    name: 'Duronto Express (12260)',
    speed: 0,
    maxSpeed: 130,
    location: 'Howrah Junction (Platform 8)',
    status: 'Stopped',
    signalStrength: 100,
    lastUpdate: '2026-04-14T10:40:15Z',
  },
];

export const mockAlerts: Alert[] = [
  {
    id: 'A-001',
    timestamp: '2026-04-14T10:35:00Z',
    severity: 'High',
    message: 'Signal Overlap detected at Block Section B-14',
    trainId: 'T-303',
  },
  {
    id: 'A-002',
    timestamp: '2026-04-14T10:38:00Z',
    severity: 'Medium',
    message: 'Brake Pressure variance detected in Rear Power Car',
    trainId: 'T-101',
  },
  {
    id: 'A-003',
    timestamp: '2026-04-14T10:39:00Z',
    severity: 'Low',
    message: 'Communication latency increased to 150ms',
  },
];

export const mockSignalHealth: SignalHealth[] = [
  { station: 'New Delhi', health: 99, latency: 12, status: 'Online' },
  { station: 'Kanpur Central', health: 94, latency: 24, status: 'Online' },
  { station: 'Agra Cantt', health: 82, latency: 45, status: 'Degraded' },
  { station: 'Howrah Junction', health: 100, latency: 8, status: 'Online' },
];
