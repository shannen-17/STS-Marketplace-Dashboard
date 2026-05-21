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

const tabs = [
  { id: 'weekly', label: 'Weekly Analysis', icon: '📅' },
  { id: 'monthly', label: 'Monthly Analysis', icon: '📊' },
  { id: 'revenue', label: 'Overall Revenue', icon: '💰' },
  { id: 'sdr', label: 'SDR Performance', icon: '📞' }
];

export default function Dashboard() {
  const [workbook, setWorkbook] = useState<WorkbookData>({});
  const [fileName, setFileName] = useState('Marketplace Performance Report.xlsx');
  const [coach, setCoach] = useState('All Coaches');
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('weekly');

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

  // Weekly metrics from Weekly Cohort
  const weeklyMetrics = useMemo(() => {
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

  // Monthly metrics from Coach Monthly
  const monthlyMetrics = useMemo(() => {
    const leads = sum(sheets.coachMonthly, ['leads', 'form submission', 'lead']);
    const booked = sum(sheets.coachMonthly, ['booked', 'intro calls', 'intro']);
    const paid = sum(sheets.coachMonthly, ['paid', 'paid engagement', 'conversion']);
    const revenue = sum(sheets.coachMonthly, ['revenue', 'take home', 'paid amount']);

    return {
      leads,
      booked,
      paid,
      revenue,
      bookingRate: leads > 0 ? booked / leads : 0,
      paidRate: leads > 0 ? paid / leads : 0,
      rpl: leads > 0 ? revenue / leads : 0
    };
  }, [sheets]);

  // Revenue metrics from Overall Revenue
  const revenueMetrics = useMemo(() => {
    const totalRevenue = sum(sheets.overallRevenue, ['take home', 'current', 'revenue', 'paid']);
    const billedAmount = sum(sheets.overallRevenue, ['billed', 'invoice']);
    const numCoaches = new Set(sheets.overallRevenue.map((r) => findColumn(r, ['coach', 'name']))).size;
    const avgPerCoach = numCoaches > 0 ? totalRevenue / numCoaches : 0;

    return {
      totalRevenue,
      billedAmount,
      numCoaches,
      avgPerCoach
    };
  }, [sheets]);

  // SDR metrics from SDR sheet
  const sdrMetrics = useMemo(() => {
    const booked = sum(sheets.sdr, ['booked', 'calls booked', 'intro']);
    const showUp = sum(sheets.sdr, ['show up', 'showed', 'attended']);
    const paid = sum(sheets.sdr, ['paid', 'paid conversion']);
    const noShow = booked - showUp;
    const showUpRate = booked > 0 ? showUp / booked : 0;
    const paidRate = showUp > 0 ? paid / showUp : 0;

    return {
      booked,
      showUp,
      paid,
      noShow,
      showUpRate,
      paidRate
    };
  }, [sheets]);

  // Get coach metrics based on active tab
  const getCoachMetrics = useMemo(() => {
    return () => {
      const map: Record<string, any> = {};

      if (activeTab === 'weekly') {
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
      } else if (activeTab === 'monthly') {
        sheets.coachMonthly.forEach((row) => {
          const coachName = findColumn(row, ['coach', 'name']) || 'Unknown';
          if (!map[coachName]) {
            map[coachName] = { coach: coachName, leads: 0, booked: 0, paid: 0, revenue: 0 };
          }
          map[coachName].leads += clean(findColumn(row, ['leads', 'form submission', 'lead']));
          map[coachName].booked += clean(findColumn(row, ['booked', 'intro calls']));
          map[coachName].paid += clean(findColumn(row, ['paid', 'engagement']));
          map[coachName].revenue += clean(findColumn(row, ['revenue', 'take home', 'paid amount']));
        });
      } else if (activeTab === 'revenue') {
        sheets.overallRevenue.forEach((row) => {
          const coachName = findColumn(row, ['coach', 'name']) || 'Unknown';
          if (!map[coachName]) {
            map[coachName] = { coach: coachName, revenue: 0, billed: 0 };
          }
          map[coachName].revenue += clean(findColumn(row, ['take home', 'current', 'revenue', 'paid']));
          map[coachName].billed += clean(findColumn(row, ['billed', 'invoice']));
        });
      } else if (activeTab === 'sdr') {
        sheets.sdr.forEach((row) => {
          const coachName = findColumn(row, ['coach', 'name', 'week']) || 'Unknown';
          if (!map[coachName]) {
            map[coachName] = { coach: coachName, booked: 0, showUp: 0, paid: 0 };
          }
          map[coachName].booked += clean(findColumn(row, ['booked', 'calls booked']));
          map[coachName].showUp += clean(findColumn(row, ['show up', 'attended']));
          map[coachName].paid += clean(findColumn(row, ['paid', 'conversion']));
        });
      }

      // Add revenue data to other tabs
      if (activeTab !== 'revenue' && activeTab !== 'sdr') {
        sheets.overallRevenue.forEach((row) => {
          const coachName = findColumn(row, ['coach', 'name']) || 'Unknown';
          if (map[coachName]) {
            map[coachName].revenue = (map[coachName].revenue || 0) + clean(findColumn(row, ['take home', 'current', 'revenue', 'paid']));
          }
        });
      }

      return Object.values(map)
        .map((c: any) => ({
          ...c,
          paidRate: c.leads > 0 ? c.paid / c.leads : c.showUp > 0 ? c.paid / c.showUp : 0,
          bookingRate: c.leads > 0 ? c.booked / c.leads : 0,
          adjustedRate: c.adjusted > 0 ? c.paid / c.adjusted : 0,
          showUpRate: c.booked > 0 ? c.showUp / c.booked : 0
        }))
        .sort((a: any, b: any) => (b.paidRate || b.revenue || 0) - (a.paidRate || a.revenue || 0));
    };
  }, [sheets, activeTab]);

  const coachMetrics = useMemo(() => getCoachMetrics(), [getCoachMetrics]);

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

  const getMetricsForTab = () => {
    if (activeTab === 'weekly') return weeklyMetrics;
    if (activeTab === 'monthly') return monthlyMetrics;
    if (activeTab === 'revenue') return revenueMetrics;
    if (activeTab === 'sdr') return sdrMetrics;
  };

  const metrics = getMetricsForTab();

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
              <p className="mt-3 max-w-3xl text-slate-600">Real-time analysis across weekly, monthly, revenue, and SDR performance.</p>
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

        {/* Tab Navigation */}
        <div className="flex flex-wrap gap-2 rounded-3xl border border-slate-200 bg-white p-2 shadow-sm">
          {tabs.map((tab) => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`flex-1 min-w-[140px] px-4 py-3 rounded-2xl font-semibold transition-all ${
                activeTab === tab.id
                  ? 'bg-slate-950 text-white'
                  : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
              }`}
            >
              {tab.icon} {tab.label}
            </button>
          ))}
        </div>

        {/* Filter */}
        <div className="sticky top-4 z-20 flex flex-wrap items-center gap-3 rounded-3xl border border-slate-200 bg-white/90 p-4 shadow-sm backdrop-blur">
          <Filter size={16} />
          <select value={coach} onChange={(e) => setCoach(e.target.value)} className="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm outline-none">
            {coaches.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
        </div>

        {/* KPIs - Dynamic based on tab */}
        {activeTab === 'weekly' && metrics && (
          <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
            <KPICard title="Total Leads" value={num(metrics.leads)} subtitle="From Weekly Cohort" icon={Target} />
            <KPICard title="Booked Intros" value={num(metrics.booked)} subtitle={`${pct(metrics.bookingRate)} booking rate`} icon={TrendingUp} />
            <KPICard title="Paid Engagements" value={num(metrics.paid)} subtitle={`${pct(metrics.paidRate)} conversion`} icon={TrendingUp} />
            <KPICard title="Net Revenue" value={money(metrics.revenue)} subtitle={`${money(metrics.rpl)} per lead`} icon={DollarSign} />
          </div>
        )}

        {activeTab === 'monthly' && metrics && (
          <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
            <KPICard title="Monthly Leads" value={num(metrics.leads)} subtitle="From Coach Monthly" icon={Target} />
            <KPICard title="Booked Intros" value={num(metrics.booked)} subtitle={`${pct(metrics.bookingRate)} booking rate`} icon={TrendingUp} />
            <KPICard title="Paid Engagements" value={num(metrics.paid)} subtitle={`${pct(metrics.paidRate)} conversion`} icon={TrendingUp} />
            <KPICard title="Revenue Per Lead" value={money(metrics.rpl)} subtitle="Monthly average" icon={DollarSign} />
          </div>
        )}

        {activeTab === 'revenue' && metrics && (
          <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
            <KPICard title="Total Revenue" value={money(metrics.totalRevenue)} subtitle="From Overall Revenue" icon={DollarSign} />
            <KPICard title="Billed Amount" value={money(metrics.billedAmount)} subtitle="Total invoiced" icon={DollarSign} />
            <KPICard title="Active Coaches" value={metrics.numCoaches} subtitle="Unique coaches" icon={Target} />
            <KPICard title="Avg per Coach" value={money(metrics.avgPerCoach)} subtitle="Revenue average" icon={DollarSign} />
          </div>
        )}

        {activeTab === 'sdr' && metrics && (
          <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
            <KPICard title="Calls Booked" value={num(metrics.booked)} subtitle="SDR bookings" icon={TrendingUp} />
            <KPICard title="Show-up Rate" value={pct(metrics.showUpRate)} subtitle={`${num(metrics.showUp)} showed up`} icon={TrendingUp} />
            <KPICard title="Paid Conversions" value={num(metrics.paid)} subtitle={`${pct(metrics.paidRate)} from show-ups`} icon={TrendingUp} />
            <KPICard title="No-shows" value={num(metrics.noShow)} subtitle={`${pct(1 - metrics.showUpRate)} didn't show`} icon={Target} />
          </div>
        )}

        {/* Coach Leaderboard */}
        {coachMetrics.length > 0 && (
          <Section title={`Coach Performance${activeTab === 'revenue' ? ' - Revenue' : activeTab === 'sdr' ? ' - SDR' : ''}`} desc={`Ranked by ${activeTab === 'sdr' ? 'show-up rate' : 'paid conversion rate'}.`}>
            <div className="overflow-x-auto rounded-3xl border border-slate-200">
              <table className="w-full min-w-[1000px] text-left text-sm">
                <thead className="bg-slate-100 text-slate-700">
                  <tr>
                    {activeTab === 'revenue' ? (
                      <>
                        <th className="px-4 py-3 font-semibold">Coach</th>
                        <th className="px-4 py-3 font-semibold">Revenue</th>
                        <th className="px-4 py-3 font-semibold">Billed</th>
                      </>
                    ) : activeTab === 'sdr' ? (
                      <>
                        <th className="px-4 py-3 font-semibold">Week/Coach</th>
                        <th className="px-4 py-3 font-semibold">Booked</th>
                        <th className="px-4 py-3 font-semibold">Show-up</th>
                        <th className="px-4 py-3 font-semibold">Show-up Rate</th>
                        <th className="px-4 py-3 font-semibold">Paid</th>
                      </>
                    ) : (
                      <>
                        <th className="px-4 py-3 font-semibold">Coach</th>
                        <th className="px-4 py-3 font-semibold">Leads</th>
                        <th className="px-4 py-3 font-semibold">Booked</th>
                        <th className="px-4 py-3 font-semibold">Paid</th>
                        <th className="px-4 py-3 font-semibold">Paid Rate</th>
                        <th className="px-4 py-3 font-semibold">Revenue</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {filteredCoaches.slice(0, 20).map((c, idx) => (
                    <tr key={idx} className="hover:bg-slate-50">
                      <td className="px-4 py-3 font-medium text-slate-900">{c.coach}</td>
                      {activeTab === 'revenue' ? (
                        <>
                          <td className="px-4 py-3 font-semibold text-green-700">{money(c.revenue)}</td>
                          <td className="px-4 py-3">{money(c.billed)}</td>
                        </>
                      ) : activeTab === 'sdr' ? (
                        <>
                          <td className="px-4 py-3">{num(c.booked)}</td>
                          <td className="px-4 py-3">{num(c.showUp)}</td>
                          <td className="px-4 py-3 font-semibold text-blue-700">{pct(c.showUpRate)}</td>
                          <td className="px-4 py-3">{num(c.paid)}</td>
                        </>
                      ) : (
                        <>
                          <td className="px-4 py-3">{num(c.leads)}</td>
                          <td className="px-4 py-3">{num(c.booked)}</td>
                          <td className="px-4 py-3 font-semibold">{num(c.paid)}</td>
                          <td className="px-4 py-3 font-semibold text-green-700">{pct(c.paidRate)}</td>
                          <td className="px-4 py-3">{money(c.revenue)}</td>
                        </>
                      )}
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
            {activeTab === 'sdr' ? (
              <>
                <Section title="Show-up Rate by Coach/Week" desc="Intro call attendance rate.">
                  <div className="h-96">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={coachMetrics.slice(0, 12)} layout="vertical" margin={{ left: 120 }}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis type="number" tickFormatter={(v) => pct(v)} />
                        <YAxis type="category" dataKey="coach" width={110} />
                        <Tooltip formatter={(v: any) => pct(v)} />
                        <Bar dataKey="showUpRate" radius={[0, 8, 8, 0]} fill="#3b82f6" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </Section>
                <Section title="Paid Conversions from Show-ups" desc="Show-up to paid rate.">
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
              </>
            ) : activeTab === 'revenue' ? (
              <>
                <Section title="Revenue by Coach" desc="Total revenue generated.">
                  <div className="h-96">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={coachMetrics.slice(0, 12)} layout="vertical" margin={{ left: 120 }}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis type="number" tickFormatter={(v) => money(v)} />
                        <YAxis type="category" dataKey="coach" width={110} />
                        <Tooltip formatter={(v: any) => money(v)} />
                        <Bar dataKey="revenue" radius={[0, 8, 8, 0]} fill="#f59e0b" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </Section>
                <Section title="Billed vs Revenue" desc="Invoice vs actual revenue received.">
                  <div className="h-96">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={coachMetrics.slice(0, 12)}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis dataKey="coach" />
                        <YAxis tickFormatter={(v) => money(v)} />
                        <Tooltip formatter={(v: any) => money(v)} />
                        <Legend />
                        <Bar dataKey="billed" fill="#8b5cf6" />
                        <Bar dataKey="revenue" fill="#10b981" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </Section>
              </>
            ) : (
              <>
                <Section title="Paid Conversion Rate" desc="Who converts leads to paid engagements most effectively.">
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

                <Section title="Booking Rate" desc="Who books intro calls most effectively from leads.">
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
              </>
            )}
          </div>
        )}

        {/* Data Validation */}
        <Section title="Data Validation" desc="Workbook sheets detected and rows parsed.">
          <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-3">
            {Object.entries(sheets).map(([name, data]) => (
              <div
                key={name}
                className={`rounded-2xl p-4 ${
                  (activeTab === 'weekly' && name === 'weeklyCohort') ||
                  (activeTab === 'monthly' && name === 'coachMonthly') ||
                  (activeTab === 'revenue' && name === 'overallRevenue') ||
                  (activeTab === 'sdr' && name === 'sdr')
                    ? 'bg-blue-50 border-2 border-blue-200'
                    : 'bg-slate-50'
                }`}
              >
                <p
                  className={`text-sm font-semibold ${
                    (activeTab === 'weekly' && name === 'weeklyCohort') ||
                    (activeTab === 'monthly' && name === 'coachMonthly') ||
                    (activeTab === 'revenue' && name === 'overallRevenue') ||
                    (activeTab === 'sdr' && name === 'sdr')
                      ? 'text-blue-700'
                      : 'text-slate-700'
                  }`}
                >
                  {((activeTab === 'weekly' && name === 'weeklyCohort') ||
                    (activeTab === 'monthly' && name === 'coachMonthly') ||
                    (activeTab === 'revenue' && name === 'overallRevenue') ||
                    (activeTab === 'sdr' && name === 'sdr')) && '⭐ '}
                  {name.replace(/([A-Z])/g, ' $1').trim()}
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
