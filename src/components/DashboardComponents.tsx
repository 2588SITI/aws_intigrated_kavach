import React from 'react';
import { Train, Alert, SignalHealth } from '../types';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from './ui/card';
import { Badge } from './ui/badge';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from './ui/table';
import { ScrollArea } from './ui/scroll-area';
import { 
  Activity, 
  AlertTriangle, 
  Train as TrainIcon, 
  Wifi, 
  Clock, 
  MapPin, 
  ShieldCheck,
  Zap,
  Signal
} from 'lucide-react';
import { 
  LineChart, 
  Line, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  AreaChart,
  Area
} from 'recharts';
import { motion } from 'motion/react';

const chartData = [
  { time: '10:00', health: 98, latency: 15 },
  { time: '10:05', health: 97, latency: 18 },
  { time: '10:10', health: 99, latency: 12 },
  { time: '10:15', health: 95, latency: 25 },
  { time: '10:20', health: 92, latency: 35 },
  { time: '10:25', health: 96, latency: 20 },
  { time: '10:30', health: 98, latency: 14 },
  { time: '10:35', health: 99, latency: 10 },
];

export const DashboardHeader = () => (
  <header className="flex items-center justify-between p-6 border-b bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/60 sticky top-0 z-50">
    <div className="flex items-center gap-3">
      <div className="p-2 bg-primary rounded-lg">
        <ShieldCheck className="w-6 h-6 text-primary-foreground" />
      </div>
      <div>
        <h1 className="text-xl font-bold tracking-tight">KAVACH DIAGNOSTIC</h1>
        <p className="text-xs text-muted-foreground font-mono uppercase tracking-widest">Automatic Train Protection System</p>
      </div>
    </div>
    <div className="flex items-center gap-4">
      <div className="flex items-center gap-2 px-3 py-1 bg-green-500/10 text-green-500 rounded-full border border-green-500/20">
        <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse" />
        <span className="text-xs font-medium">SYSTEM OPERATIONAL</span>
      </div>
      <div className="text-right hidden md:block">
        <p className="text-xs font-medium">{new Date().toLocaleDateString()}</p>
        <p className="text-xs text-muted-foreground font-mono">{new Date().toLocaleTimeString()}</p>
      </div>
    </div>
  </header>
);

export const StatsGrid = ({ trains, alerts }: { trains: Train[], alerts: Alert[] }) => (
  <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4 p-6">
    <Card className="relative overflow-hidden">
      <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
        <CardTitle className="text-sm font-medium">Active Trains</CardTitle>
        <TrainIcon className="h-4 w-4 text-muted-foreground" />
      </CardHeader>
      <CardContent>
        <div className="text-2xl font-bold">{trains.filter(t => t.status !== 'Stopped').length}</div>
        <p className="text-xs text-muted-foreground">
          Total {trains.length} registered units
        </p>
      </CardContent>
      <div className="absolute bottom-0 left-0 w-full h-1 bg-primary/20" />
    </Card>
    <Card className="relative overflow-hidden">
      <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
        <CardTitle className="text-sm font-medium">System Health</CardTitle>
        <Activity className="h-4 w-4 text-muted-foreground" />
      </CardHeader>
      <CardContent>
        <div className="text-2xl font-bold">98.4%</div>
        <p className="text-xs text-muted-foreground">
          +0.2% from last hour
        </p>
      </CardContent>
      <div className="absolute bottom-0 left-0 w-full h-1 bg-green-500/20" />
    </Card>
    <Card className="relative overflow-hidden">
      <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
        <CardTitle className="text-sm font-medium">Active Alerts</CardTitle>
        <AlertTriangle className="h-4 w-4 text-destructive" />
      </CardHeader>
      <CardContent>
        <div className="text-2xl font-bold text-destructive">{alerts.length}</div>
        <p className="text-xs text-muted-foreground">
          {alerts.filter(a => a.severity === 'High').length} high severity
        </p>
      </CardContent>
      <div className="absolute bottom-0 left-0 w-full h-1 bg-destructive/20" />
    </Card>
    <Card className="relative overflow-hidden">
      <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
        <CardTitle className="text-sm font-medium">Signal Sync</CardTitle>
        <Signal className="h-4 w-4 text-muted-foreground" />
      </CardHeader>
      <CardContent>
        <div className="text-2xl font-bold">12ms</div>
        <p className="text-xs text-muted-foreground">
          Avg. network latency
        </p>
      </CardContent>
      <div className="absolute bottom-0 left-0 w-full h-1 bg-blue-500/20" />
    </Card>
  </div>
);

export const TrainStatusTable = ({ trains }: { trains: Train[] }) => (
  <Card className="col-span-4">
    <CardHeader>
      <CardTitle>Train Telemetry</CardTitle>
      <CardDescription>Real-time monitoring of active locomotive units.</CardDescription>
    </CardHeader>
    <CardContent>
      <Table>
        <TableHeader>
          <TableRow>
            <TableHead>Train ID</TableHead>
            <TableHead>Name</TableHead>
            <TableHead>Speed (km/h)</TableHead>
            <TableHead>Location</TableHead>
            <TableHead>Status</TableHead>
            <TableHead className="text-right">Signal</TableHead>
          </TableRow>
        </TableHeader>
        <TableBody>
          {trains.map((train) => (
            <TableRow key={train.id}>
              <TableCell className="font-mono font-medium">{train.id}</TableCell>
              <TableCell>{train.name}</TableCell>
              <TableCell>
                <div className="flex items-center gap-2">
                  <span className="font-mono">{train.speed}</span>
                  <div className="w-16 h-1 bg-muted rounded-full overflow-hidden">
                    <div 
                      className="h-full bg-primary" 
                      style={{ width: `${(train.speed / train.maxSpeed) * 100}%` }} 
                    />
                  </div>
                </div>
              </TableCell>
              <TableCell className="text-muted-foreground flex items-center gap-1">
                <MapPin className="w-3 h-3" />
                {train.location}
              </TableCell>
              <TableCell>
                <Badge variant={
                  train.status === 'Normal' ? 'default' : 
                  train.status === 'Warning' ? 'outline' : 
                  train.status === 'Stopped' ? 'secondary' : 'destructive'
                }>
                  {train.status}
                </Badge>
              </TableCell>
              <TableCell className="text-right">
                <div className="flex items-center justify-end gap-1">
                  <Wifi className={`w-4 h-4 ${train.signalStrength > 80 ? 'text-green-500' : 'text-yellow-500'}`} />
                  <span className="font-mono text-xs">{train.signalStrength}%</span>
                </div>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </CardContent>
  </Card>
);

export const AlertsPanel = ({ alerts }: { alerts: Alert[] }) => (
  <Card className="col-span-1">
    <CardHeader>
      <CardTitle className="flex items-center gap-2">
        <Zap className="w-4 h-4 text-yellow-500" />
        System Alerts
      </CardTitle>
    </CardHeader>
    <CardContent>
      <ScrollArea className="h-[400px] pr-4">
        <div className="space-y-4">
          {alerts.map((alert) => (
            <div key={alert.id} className="p-3 rounded-lg border bg-muted/50 relative overflow-hidden">
              <div className={`absolute left-0 top-0 w-1 h-full ${
                alert.severity === 'High' ? 'bg-destructive' : 
                alert.severity === 'Medium' ? 'bg-yellow-500' : 'bg-blue-500'
              }`} />
              <div className="flex justify-between items-start mb-1">
                <Badge variant="outline" className="text-[10px] h-4 px-1">
                  {alert.id}
                </Badge>
                <span className="text-[10px] text-muted-foreground flex items-center gap-1">
                  <Clock className="w-2 h-2" />
                  {new Date(alert.timestamp).toLocaleTimeString()}
                </span>
              </div>
              <p className="text-sm font-medium leading-tight mb-1">{alert.message}</p>
              {alert.trainId && (
                <p className="text-[10px] text-muted-foreground font-mono">
                  AFFECTED UNIT: {alert.trainId}
                </p>
              )}
            </div>
          ))}
        </div>
      </ScrollArea>
    </CardContent>
  </Card>
);

export const PerformanceChart = () => (
  <Card className="col-span-3">
    <CardHeader>
      <CardTitle>Network Performance</CardTitle>
      <CardDescription>System latency and signal health over time.</CardDescription>
    </CardHeader>
    <CardContent>
      <div className="h-[300px] w-full">
        <ResponsiveContainer width="100%" height="100%">
          <AreaChart data={chartData}>
            <defs>
              <linearGradient id="colorHealth" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="hsl(var(--primary))" stopOpacity={0.3}/>
                <stop offset="95%" stopColor="hsl(var(--primary))" stopOpacity={0}/>
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="hsl(var(--muted))" />
            <XAxis 
              dataKey="time" 
              stroke="hsl(var(--muted-foreground))" 
              fontSize={12}
              tickLine={false}
              axisLine={false}
            />
            <YAxis 
              stroke="hsl(var(--muted-foreground))" 
              fontSize={12}
              tickLine={false}
              axisLine={false}
              tickFormatter={(value) => `${value}%`}
            />
            <Tooltip 
              contentStyle={{ 
                backgroundColor: 'hsl(var(--background))', 
                border: '1px solid hsl(var(--border))',
                borderRadius: '8px'
              }}
            />
            <Area 
              type="monotone" 
              dataKey="health" 
              stroke="hsl(var(--primary))" 
              fillOpacity={1} 
              fill="url(#colorHealth)" 
              strokeWidth={2}
            />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </CardContent>
  </Card>
);
