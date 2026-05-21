'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { motion } from 'framer-motion';
import { BarChart3, DollarSign, FileSpreadsheet, Filter, Target, TrendingUp, Upload } from 'lucide-react';
import { BarChart, Bar, LineChart, Line, CartesianGrid, XAxis, YAxis, Tooltip, Legend, ResponsiveContainer, ComposedChart, Area } from 'recharts';

type Row = Record<string, any>;
type WorkbookData = Record<string, Row[]>;

const DEFAULT_FILE = '/Marketplace Performance Report.xlsx';

const money = (n: number) =>
  new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(n || 0);
const num = (n: number) => new Intl.NumberFormat('en-US', { maximumFractionDigits: 1 }).format(n || 0);
const pct = (n: number) => `${((n || 0) * 100).toFixed(1)}%`;

const clean = (v: any): number => {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = Number(v.replace(/[$,%]/g, ''));
    return Number.isNaN(n) ? 0 : n;
  }
  return 0;
};

const normalizeHeader = (header: string): string => {
  return header.toLowerCase().trim().replace(/[\s\-_]+/g, ' ');
};

async function parseWorkbook(buffer: ArrayBuffer): Promise<WorkbookData> {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const parsed: WorkbookData = {};

  workbook.SheetNames.forEach((name) => {
    let rows = XLSX.utils.sheet_to_json<Row>(workbook.Sheets[name], { defval: '' });

    // Normalize headers
    rows = rows.map((row) => {
      const normalized: Row = {};
      Object.entries(row).forEach(([key, value]) => {
        normalized[normalizeHeader(key)] = value;
      });
      return normalized;
    });

    // Filter empty rows
    rows = rows.filter((r) => Object.values(r).some((v) => v !== '' && v !== null && v !== undefined));
    parsed[name] = rows;
  });

  return parsed;
}

function findColumn(row: Row, patterns: string[]): any {
  const keys = Object.keys(row);
  for (const pattern of patterns) {
    const key = keys.find((k) => k.includes(pattern.toLowerCase()));
    if (key !== undefined) return row[key];
  }
  return null;
}

function KPICard({ title, value, subtitle, icon: Icon }: any) {
  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="rounded-3xl border border-slate-200 bg-white p-6 shadow-sm"
    >
      <div className="flex items-start justify-between gap-4">
        <div className="flex-1">
          <p className="text-sm font-medium text-slate-500">{title}</p>
          <h3 className="mt-2 text-3xl font-semibold tracking-tight text-slate-950">{value}</h3>
          {subtitle && <p className="mt-2 text-sm text-slate-600">{subtitle}</p>}
        </div>
        {Icon && <div className="rounded-2xl bg-blue-50 p-3 text-blue-700"><Icon size={24} /></div>}
      </div>
    </motion.div>
  );
}

function Section({ title, desc, children }: any) {
  return (
    <section className="rounded-3xl border border-slate-200 bg-white p-6 shadow-sm md:p-8">
      <div className="mb-6">
        <h2 className="text-2xl font-semibold tracking-tight text-slate-950">{title}</h2>
        <p className="mt-1 max-w-3xl text-sm text-slate-500">{desc}</p>
      </div>
      {children}
    </section>
  );
}

export default function Dashboard() {
  const [workbook, setWorkbook] = useState<WorkbookData>({});
  const [fileName, setFileName] = useState('Marketplace Performance Report.xlsx');
  const [coach, setCoach] = useState('All Coaches');
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function loadDefault() {
      try {
        const response = await fetch(DEFAULT_FILE);
        const buffer = await response.arrayBuffer();
        const data = await parseWorkbook(buffer);
        setWorkbook(data);
      } catch (err) {
        console.error('Failed to load default file:', err);
      } finally {
        setLoading(false);
      }
    }
    loadDefault();
  }, []);

  const sheets = useMemo(() => {
    return {
      weeklyCohort: workbook['Weekly Cohort'] || [],
      coachWeekly: workbook['Coach Weekly'] || [],
      coachMonthly: workbook['Coach Monthly'] || [],
      overallWeekly: workbook['Overall Weekly'] || [],
      overallRevenue: workbook['Overall Revenue'] || [],
      engagements: workbook['Engagements'] || [],
      sdr: workbook['SDR'] || [],
      cohortLeadDetail: workbook['Cohort Lead detail'] || [],
      unattributedCalls: workbook['Unattributed intro calls'] || []
    };
  }, [workbook]);

  const sum = (rows: Row[], patterns: string[]): number => {
    return rows.reduce((acc, r) => acc + clean(findColumn(r, patterns)), 0);
  };

  // Primary metrics from Weekly Cohort (source of truth)
  const metrics = useMemo(() => {
    const leads = sum(sheets.weeklyCohort, ['leads', 'form submission', 'lead']);
    const booked = sum(sheets.weeklyCohort, ['booked', 'intro calls', 'intro']);
    const paid = sum(sheets.weeklyCohort, ['paid', 'paid engagement', 'conversion']);
    const adjusted = sum(sheets.weeklyCohort, ['adjusted']);
    const revenue = sum(sheets.overallRevenue, ['take home', 'current', 'revenue']) || sum(sheets.engagements, ['paid amount', 'paid']);

    return {
      leads,
      booked,
      paid,
      adjusted,
      revenue,
      bookingRate: leads > 0 ? booked / leads : 0,
      paidRate: leads > 0 ? paid / leads : 0,
      introClose: adjusted > 0 ? paid / adjusted : 0,
      rpl: leads > 0 ? revenue / leads : 0
    };
  }, [sheets]);

  // Coach metrics from Coach Weekly
  const coachMetrics = useMemo(() => {
    const map: Record<string, any> = {};

    sheets.coachWeekly.forEach((row) => {
      const coachName = findColumn(row, ['coach', 'name']) || 'Unknown';
      if (!map[coachName]) {
        map[coachName] = { coach: coachName, leads: 0, booked: 0, adjusted: 0, paid: 0, revenue: 0 };
      }
      map[coachName].leads += clean(findColumn(row, ['leads', 'form submission', 'lead']));
      map[coachName].booked += clean(findColumn(row, ['booked', 'intro calls']));
      map[coachName].adjusted += clean(findColumn(row, ['adjusted', 'intro']));
      map[coachName].paid += clean(findColumn(row, ['paid', 'engagement']));
    });

    // Add revenue from Overall Revenue
    sheets.overallRevenue.forEach((row) => {
      const coachName = findColumn(row, ['coach', 'name']) || 'Unknown';
      if (map[coachName]) {
        map[coachName].revenue += clean(findColumn(row, ['take home', 'current', 'revenue', 'paid']));
      }
    });

    return Object.values(map)
      .map((c: any) => ({
        ...c,
        paidRate: c.leads > 0 ? c.paid / c.leads : 0,
        bookingRate: c.leads > 0 ? c.booked / c.leads : 0,
        adjustedRate: c.adjusted > 0 ? c.paid / c.adjusted : 0
      }))
      .sort((a: any, b: any) => b.paidRate - a.paidRate);
  }, [sheets]);

  const coaches = useMemo(() => {
    return ['All Coaches', ...coachMetrics.map((c) => c.coach)];
  }, [coachMetrics]);

  const filteredCoaches = coach === 'All Coaches' ? coachMetrics : coachMetrics.filter((c) => c.coach === coach);

  async function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    try {
      setFileName(file.name);
      const buffer = await file.arrayBuffer();
      const data = await parseWorkbook(buffer);
      setWorkbook(data);
      setCoach('All Coaches');
    } catch (err) {
      console.error('Upload failed:', err);
    } finally {
      setLoading(false);
    }
  }

  if (loading) {
    return (
      <main className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 flex items-center justify-center px-4">
        <div className="text-center">
          <div className="animate-spin h-12 w-12 border-4 border-blue-200 border-t-blue-700 rounded-full mx-auto mb-4"></div>
          <p className="text-slate-600">Loading dashboard...</p>
        </div>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 px-4 py-6 md:px-8">
      <div className="mx-auto max-w-7xl space-y-6">
        {/* Header */}
        <header className="rounded-3xl border border-slate-200 bg-white p-6 shadow-sm md:p-8">
          <div className="mb-6 flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="mb-4 inline-flex items-center gap-2 rounded-full bg-blue-50 px-4 py-2 text-sm font-semibold text-blue-700">
                <BarChart3 size={16} /> STS Marketplace Intelligence
              </div>
              <h1 className="text-4xl font-semibold tracking-tight text-slate-950 md:text-5xl">Marketplace Dashboard</h1>
              <p className="mt-3 max-w-3xl text-slate-600">Real-time analysis powered by Weekly Cohort data as source of truth.</p>
            </div>
            <label className="flex cursor-pointer items-center justify-center gap-2 rounded-2xl bg-slate-950 px-4 py-3 font-semibold text-white hover:bg-slate-800">
              <Upload size={16} /> Upload Report
              <input type="file" accept=".xlsx" onChange={handleUpload} className="hidden" />
            </label>
          </div>
          <div className="rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-600">
            <FileSpreadsheet className="mr-2 inline" size={16} /> {fileName}
          </div>
        </header>

        {/* Filter */}
        <div className="sticky top-4 z-20 flex flex-wrap items-center gap-3 rounded-3xl border border-slate-200 bg-white/90 p-4 shadow-sm backdrop-blur">
          <Filter size={16} />
          <select value={coach} onChange={(e) => setCoach(e.target.value)} className="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm outline-none">
            {coaches.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
        </div>

        {/* KPIs */}
        <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
          <KPICard title="Total Leads" value={num(metrics.leads)} subtitle="From Weekly Cohort" icon={Target} />
          <KPICard title="Booked Intros" value={num(metrics.booked)} subtitle={`${pct(metrics.bookingRate)} booking rate`} icon={TrendingUp} />
          <KPICard title="Paid Engagements" value={num(metrics.paid)} subtitle={`${pct(metrics.paidRate)} conversion`} icon={TrendingUp} />
          <KPICard title="Net Revenue" value={money(metrics.revenue)} subtitle={`${money(metrics.rpl)} per lead`} icon={DollarSign} />
        </div>

        {/* Executive Summary */}
        <Section title="Executive Summary" desc="Key insights based on Weekly Cohort metrics.">
          <div className="grid gap-4 md:grid-cols-2">
            <div className="rounded-2xl bg-blue-50 p-4">
              <p className="text-sm font-medium text-blue-700">Booking Efficiency</p>
              <p className="mt-1 text-2xl font-bold text-blue-900">{pct(metrics.bookingRate)}</p>
              <p className="mt-1 text-xs text-blue-600">of leads booked intro calls</p>
            </div>
            <div className="rounded-2xl bg-green-50 p-4">
              <p className="text-sm font-medium text-green-700">Paid Conversion</p>
              <p className="mt-1 text-2xl font-bold text-green-900">{pct(metrics.paidRate)}</p>
              <p className="mt-1 text-xs text-green-600">lead-to-paid conversion rate</p>
            </div>
            <div className="rounded-2xl bg-purple-50 p-4">
              <p className="text-sm font-medium text-purple-700">Intro Close Rate</p>
              <p className="mt-1 text-2xl font-bold text-purple-900">{pct(metrics.introClose)}</p>
              <p className="mt-1 text-xs text-purple-600">from adjusted intro calls</p>
            </div>
            <div className="rounded-2xl bg-amber-50 p-4">
              <p className="text-sm font-medium text-amber-700">Revenue Per Lead</p>
              <p className="mt-1 text-2xl font-bold text-amber-900">{money(metrics.rpl)}</p>
              <p className="mt-1 text-xs text-amber-600">average per lead</p>
            </div>
          </div>
        </Section>

        {/* Coach Leaderboard */}
        {coachMetrics.length > 0 && (
          <Section title="Coach Performance Leaderboard" desc="Ranked by paid engagement conversion rate (Coach Weekly data).">
            <div className="overflow-x-auto rounded-3xl border border-slate-200">
              <table className="w-full min-w-[1000px] text-left text-sm">
                <thead className="bg-slate-100 text-slate-700">
                  <tr>
                    {['Coach', 'Leads', 'Booked', 'Adjusted', 'Paid', 'Paid Rate', 'Book Rate', 'Intro Close'].map((h) => (
                      <th key={h} className="px-4 py-3 font-semibold">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {filteredCoaches.slice(0, 20).map((c, idx) => (
                    <tr key={idx} className="hover:bg-slate-50">
                      <td className="px-4 py-3 font-medium text-slate-900">{c.coach}</td>
                      <td className="px-4 py-3">{num(c.leads)}</td>
                      <td className="px-4 py-3">{num(c.booked)}</td>
                      <td className="px-4 py-3">{num(c.adjusted)}</td>
                      <td className="px-4 py-3 font-semibold">{num(c.paid)}</td>
                      <td className="px-4 py-3 font-semibold text-green-700">{pct(c.paidRate)}</td>
                      <td className="px-4 py-3">{pct(c.bookingRate)}</td>
                      <td className="px-4 py-3">{pct(c.adjustedRate)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Section>
        )}

        {/* Charts */}
        {coachMetrics.length > 0 && (
          <div className="grid gap-6 xl:grid-cols-2">
            <Section title="Paid Conversion Rate by Coach" desc="Who converts leads to paid engagements most effectively.">
              <div className="h-96">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={coachMetrics.slice(0, 12)} layout="vertical" margin={{ left: 120 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis type="number" tickFormatter={(v) => pct(v)} />
                    <YAxis type="category" dataKey="coach" width={110} />
                    <Tooltip formatter={(v: any) => pct(v)} />
                    <Bar dataKey="paidRate" radius={[0, 8, 8, 0]} fill="#10b981" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Section>

            <Section title="Booking Rate by Coach" desc="Who books intro calls most effectively from leads.">
              <div className="h-96">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={coachMetrics.slice(0, 12)} layout="vertical" margin={{ left: 120 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis type="number" tickFormatter={(v) => pct(v)} />
                    <YAxis type="category" dataKey="coach" width={110} />
                    <Tooltip formatter={(v: any) => pct(v)} />
                    <Bar dataKey="bookingRate" radius={[0, 8, 8, 0]} fill="#3b82f6" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Section>
          </div>
        )}

        {/* Data Validation */}
        <Section title="Data Validation" desc="Workbook sheets detected and rows parsed. Weekly Cohort is primary data source.">
          <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-3">
            {Object.entries(sheets).map(([name, data]) => (
              <div key={name} className={`rounded-2xl p-4 ${name === 'weeklyCohort' ? 'bg-blue-50 border-2 border-blue-200' : 'bg-slate-50'}`}>
                <p className={`text-sm font-semibold ${name === 'weeklyCohort' ? 'text-blue-700' : 'text-slate-700'}`}>
                  {name === 'weeklyCohort' ? '⭐ ' : ''}{name.replace(/([A-Z])/g, ' $1').trim()}
                </p>
                <p className="mt-1 text-lg font-bold text-slate-900">{data.length} rows</p>
              </div>
            ))}
          </div>
        </Section>
      </div>
    </main>
  );
}
