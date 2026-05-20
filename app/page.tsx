"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { motion } from "framer-motion";
import { AlertTriangle, ArrowUpRight, BarChart3, DollarSign, FileSpreadsheet, Filter, LineChart as LineIcon, RefreshCw, Target, TrendingUp, Upload } from "lucide-react";
import { Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, ComposedChart, Funnel, FunnelChart, LabelList, Legend, Line, LineChart, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";

type Row = Record<string, any>;
type WorkbookData = Record<string, Row[]>;

const DEFAULT_FILE = "/Marketplace Performance Report.xlsx";
const money = (n: any) => new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(Number(n || 0));
const num = (n: any) => new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 }).format(Number(n || 0));
const pct = (n: any) => `${((Number(n || 0)) * 100).toFixed(1)}%`;
const clean = (v: any) => typeof v === "number" && !Number.isNaN(v) ? v : Number(String(v ?? "").replace(/[$,%]/g, "")) || 0;
const rows = (wb: WorkbookData, sheet: string) => wb[sheet] || [];
const sum = (items: Row[], key: string) => items.reduce((a, r) => a + clean(r[key]), 0);
const avg = (items: Row[], key: string) => items.length ? sum(items, key) / items.length : 0;
const uniq = (arr: string[]) => Array.from(new Set(arr.filter(Boolean))).sort();

function parseWorkbook(buffer: ArrayBuffer): WorkbookData {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const parsed: WorkbookData = {};
  workbook.SheetNames.forEach((name) => {
    let json = XLSX.utils.sheet_to_json<Row>(workbook.Sheets[name], { defval: "" });
    if (name === "Overall Revenue" && json.length && !json[0]["Period"]) {
      const raw = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets[name], { header: 1, defval: "" });
      const headerIndex = raw.findIndex((r) => r.includes("Period") && r.includes("Coach"));
      if (headerIndex >= 0) {
        const headers = raw[headerIndex];
        json = raw.slice(headerIndex + 1).filter((r) => r.some(Boolean)).map((r) => Object.fromEntries(headers.map((h: string, i: number) => [h || `Column ${i}`, r[i] ?? ""])));
      }
    }
    parsed[name] = json.filter((r) => Object.values(r).some((v) => v !== "" && v !== null && v !== undefined));
  });
  return parsed;
}

function Kpi({ title, value, sub, icon: Icon, tone = "blue" }: any) {
  return <motion.div initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} className="kpi">
    <div className="flex items-start justify-between gap-4">
      <div><p className="text-sm font-medium text-slate-500">{title}</p><h3 className="mt-2 text-3xl font-semibold tracking-tight text-slate-950">{value}</h3><p className="mt-2 text-sm leading-5 text-slate-500">{sub}</p></div>
      <div className={`rounded-2xl p-3 ${tone === "gold" ? "bg-amber-50 text-amber-700" : tone === "green" ? "bg-emerald-50 text-emerald-700" : "bg-blue-50 text-blue-700"}`}><Icon size={20}/></div>
    </div>
  </motion.div>;
}

function Section({ title, desc, children }: any) {
  return <section className="card p-6 md:p-8"><div className="mb-6 flex flex-col gap-2 md:flex-row md:items-end md:justify-between"><div><h2 className="text-2xl font-semibold tracking-tight text-slate-950">{title}</h2><p className="mt-1 max-w-3xl text-sm leading-6 text-slate-500">{desc}</p></div></div>{children}</section>;
}

export default function Dashboard() {
  const [data, setData] = useState<WorkbookData>({});
  const [fileName, setFileName] = useState("Marketplace Performance Report.xlsx");
  const [coach, setCoach] = useState("All Coaches");

  useEffect(() => { fetch(DEFAULT_FILE).then(r => r.arrayBuffer()).then(buf => setData(parseWorkbook(buf))); }, []);

  const coachWeekly = rows(data, "Coach Weekly");
  const coachMonthly = rows(data, "Coach Monthly");
  const marketplace = rows(data, "Marketplace Report").filter(r => r.Date || r.Month);
  const engagements = rows(data, "Engagements");
  const sdr = rows(data, "SDR");
  const dataSweep = rows(data, "Weekly DataSweep");
  const overallRevenue = rows(data, "Overall Revenue");
  const coaches = useMemo(() => ["All Coaches", ...uniq([...coachWeekly, ...coachMonthly, ...overallRevenue].map(r => String(r.Coach || "")))], [data]);
  const filteredWeekly = coach === "All Coaches" ? coachWeekly : coachWeekly.filter(r => r.Coach === coach);

  const totals = useMemo(() => {
    const revenueRows = marketplace.filter(r => r.Date);
    const leads = sum(coachWeekly, "Leads Given") || sum(marketplace, "Form Submission");
    const booked = sum(coachWeekly, "Booked Intro Calls") || sum(marketplace, "Booked Intro Call");
    const paid = sum(coachWeekly, "Paid Engagements") || sum(marketplace, "Converted");
    const adjusted = sum(coachWeekly, "Adjusted Intro Calls");
    const revenue = sum(marketplace, "Net Revenue") || sum(engagements, "Paid Amount");
    return { leads, booked, paid, adjusted, revenue, bookingRate: booked / Math.max(leads, 1), paidRate: paid / Math.max(leads, 1), introClose: paid / Math.max(adjusted || booked, 1), rpl: revenue / Math.max(leads, 1), avgDays: avg(dataSweep, "Avg Days to Engagement"), projected12: sum(dataSweep, "Projected 12-Month Value") };
  }, [data]);

  const coachRank = useMemo(() => {
    const map: Record<string, any> = {};
    coachWeekly.forEach(r => {
      const c = r.Coach || "Unknown";
      map[c] ||= { coach: c, leads: 0, booked: 0, adjusted: 0, paid: 0, started: 0, revenue: 0 };
      map[c].leads += clean(r["Leads Given"]); map[c].booked += clean(r["Booked Intro Calls"]); map[c].adjusted += clean(r["Adjusted Intro Calls"]); map[c].paid += clean(r["Paid Engagements"]); map[c].started += clean(r["Engagements Started"]);
    });
    overallRevenue.forEach(r => { if (map[r.Coach]) map[r.Coach].revenue += clean(r["Current Take-Home 12 mo"]) || clean(r["Paid Detail Total"]); });
    return Object.values(map).map((r: any) => ({ ...r, bookingRate: r.booked / Math.max(r.leads, 1), paidRate: r.paid / Math.max(r.leads, 1), closeRate: r.paid / Math.max(r.adjusted || r.booked, 1), score: Math.round((r.paid / Math.max(r.leads, 1) * 55 + r.booked / Math.max(r.leads, 1) * 25 + r.paid * 2) * 100) })).sort((a: any, b: any) => b.paidRate - a.paidRate);
  }, [data]);

  const funnel = [
    { name: "Landing Hits", value: sum(marketplace, "Landing Page Hits") },
    { name: "Forms", value: sum(marketplace, "Form Submission") },
    { name: "Intro Calls", value: sum(marketplace, "Booked Intro Call") || totals.booked },
    { name: "Converted", value: sum(marketplace, "Converted") || totals.paid },
    { name: "Transactions", value: sum(marketplace, "# of Transaction") }
  ].filter(x => x.value > 0);

  const insights = [
    `Marketplace generated ${money(totals.revenue)} from ${num(totals.leads)} leads, or ${money(totals.rpl)} per lead.`,
    `The current paid engagement rate is ${pct(totals.paidRate)}. This is the core marketplace efficiency metric to improve.`,
    `Average conversion timing is ${num(totals.avgDays)} days, which means some lead cohorts may need follow-up beyond the first week.`,
    `Projected 12-month value from observed paid conversions is ${money(totals.projected12)} based on the Weekly DataSweep sheet.`
  ];

  async function upload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]; if (!file) return;
    setFileName(file.name); setData(parseWorkbook(await file.arrayBuffer()));
  }

  return <main className="min-h-screen bg-[radial-gradient(circle_at_top_left,#EAF4FF,transparent_34%),linear-gradient(135deg,#f8fafc,#fff8ef)] px-4 py-6 md:px-8">
    <div className="mx-auto max-w-7xl space-y-6">
      <header className="flex flex-col gap-4 rounded-[2rem] border border-white/70 bg-white/70 p-6 shadow-sm backdrop-blur md:flex-row md:items-center md:justify-between">
        <div><div className="pill mb-4 inline-flex items-center gap-2"><BarChart3 size={14}/> STS Marketplace Intelligence</div><h1 className="text-4xl font-semibold tracking-tight text-slate-950 md:text-5xl">Marketplace Performance Dashboard</h1><p className="mt-3 max-w-3xl text-slate-600">Executive view of coach performance, funnel movement, SDR impact, revenue, and cohort quality from the uploaded workbook.</p></div>
        <div className="flex flex-col gap-3 md:min-w-80"><label className="flex cursor-pointer items-center justify-center gap-2 rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white shadow-sm hover:bg-slate-800"><Upload size={16}/> Upload new report<input type="file" accept=".xlsx,.xls" className="hidden" onChange={upload}/></label><div className="rounded-2xl border border-slate-200 bg-white px-4 py-3 text-xs text-slate-500"><FileSpreadsheet className="mr-2 inline" size={14}/> {fileName}</div></div>
      </header>

      <div className="sticky top-3 z-20 flex flex-wrap items-center gap-3 rounded-3xl border border-slate-200 bg-white/85 p-3 shadow-sm backdrop-blur"><Filter size={16}/><select value={coach} onChange={(e) => setCoach(e.target.value)} className="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm outline-none">{coaches.map(c => <option key={c}>{c}</option>)}</select><span className="pill">{coachWeekly.length} weekly coach rows</span><span className="pill">{engagements.length} engagements</span><span className="pill">{dataSweep.length} cohorts</span></div>

      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
        <Kpi title="Total Leads" value={num(totals.leads)} sub="Lead volume from Coach Weekly/Form submissions." icon={Target}/>
        <Kpi title="Booked Intro Calls" value={num(totals.booked)} sub={`${pct(totals.bookingRate)} of leads booked an intro call.`} icon={LineIcon}/>
        <Kpi title="Paid Engagements" value={num(totals.paid)} sub={`${pct(totals.paidRate)} lead-to-paid conversion.`} icon={TrendingUp} tone="green"/>
        <Kpi title="Net Revenue" value={money(totals.revenue)} sub={`${money(totals.rpl)} revenue per lead.`} icon={DollarSign} tone="gold"/>
      </div>

      <Section title="Executive Interpretation" desc="Automatically generated readout of the whole report based on the workbook totals and derived marketplace metrics.">
        <div className="grid gap-4 md:grid-cols-2">{insights.map((i, idx) => <div key={i} className="rounded-3xl border border-slate-200 bg-slate-50 p-5"><div className="mb-3 flex h-9 w-9 items-center justify-center rounded-2xl bg-white shadow-sm"><ArrowUpRight size={16}/></div><p className="text-sm leading-6 text-slate-700">{i}</p></div>)}</div>
      </Section>

      <div className="grid gap-6 xl:grid-cols-2">
        <Section title="Marketplace Funnel" desc="Landing page traffic through revenue-producing marketplace actions.">
          <div className="h-80"><ResponsiveContainer><FunnelChart><Tooltip formatter={(v:any)=>num(v)}/><Funnel dataKey="value" data={funnel} isAnimationActive><LabelList position="right" fill="#334155" stroke="none" dataKey="name" /></Funnel></FunnelChart></ResponsiveContainer></div>
        </Section>
        <Section title="Weekly Marketplace Revenue" desc="Net revenue and converted clients by reporting week.">
          <div className="h-80"><ResponsiveContainer><ComposedChart data={marketplace.filter(r => r.Date)}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="Date" fontSize={11}/><YAxis yAxisId="left" fontSize={11}/><YAxis yAxisId="right" orientation="right" fontSize={11}/><Tooltip formatter={(v:any,n:any)=>String(n).includes("Revenue") ? money(v) : num(v)}/><Legend/><Bar yAxisId="right" dataKey="Converted" radius={[8,8,0,0]}/><Area yAxisId="left" type="monotone" dataKey="Net Revenue" fillOpacity={0.15}/></ComposedChart></ResponsiveContainer></div>
        </Section>
      </div>

      <Section title="Coach Performance Analytics" desc="Leaderboard based on leads, intro calls, paid engagements, conversion rates, and available revenue signals.">
        <div className="overflow-hidden rounded-3xl border border-slate-200"><table className="w-full min-w-[900px] text-left text-sm"><thead className="bg-slate-950 text-white"><tr>{["Coach","Leads","Booked","Paid","Booking Rate","Paid Rate","Intro Close","Score"].map(h=><th className="px-4 py-3 font-medium" key={h}>{h}</th>)}</tr></thead><tbody className="divide-y divide-slate-100 bg-white">{coachRank.slice(0,12).map((r:any)=><tr key={r.coach} className="hover:bg-slate-50"><td className="px-4 py-3 font-semibold text-slate-900">{r.coach}</td><td className="px-4 py-3">{num(r.leads)}</td><td className="px-4 py-3">{num(r.booked)}</td><td className="px-4 py-3">{num(r.paid)}</td><td className="px-4 py-3">{pct(r.bookingRate)}</td><td className="px-4 py-3">{pct(r.paidRate)}</td><td className="px-4 py-3">{pct(r.closeRate)}</td><td className="px-4 py-3"><span className="rounded-full bg-blue-50 px-3 py-1 font-semibold text-blue-700">{r.score}</span></td></tr>)}</tbody></table></div>
      </Section>

      <div className="grid gap-6 xl:grid-cols-2">
        <Section title="Coach Conversion Comparison" desc="Paid conversion rate by coach. Use this to spot who is turning lead flow into paid engagements.">
          <div className="h-96"><ResponsiveContainer><BarChart data={coachRank.slice(0,10)} layout="vertical" margin={{ left: 90 }}><CartesianGrid strokeDasharray="3 3"/><XAxis type="number" tickFormatter={(v)=>pct(v)} fontSize={11}/><YAxis type="category" dataKey="coach" fontSize={11}/><Tooltip formatter={(v:any)=>pct(v)}/><Bar dataKey="paidRate" radius={[0,8,8,0]}/></BarChart></ResponsiveContainer></div>
        </Section>
        <Section title="Cohort Quality" desc="Weekly DataSweep shows delayed conversion behavior, paid conversion, and projected 12-month value per lead.">
          <div className="h-96"><ResponsiveContainer><LineChart data={dataSweep}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="Lead Week" fontSize={11}/><YAxis fontSize={11}/><Tooltip formatter={(v:any,n:any)=>String(n).includes("Rate") ? pct(v) : money(v)}/><Legend/><Line type="monotone" dataKey="Paid Conversion Rate" strokeWidth={3}/><Line type="monotone" dataKey="Projected 12-Month Value per Lead" strokeWidth={3}/></LineChart></ResponsiveContainer></div>
        </Section>
      </div>

      <div className="grid gap-6 xl:grid-cols-2">
        <Section title="SDR Performance" desc="Stephen SDR rows interpreted as intro booking, show-up, and paid conversion activity.">
          <div className="h-72"><ResponsiveContainer><BarChart data={sdr}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="Week" fontSize={11}/><YAxis fontSize={11}/><Tooltip/><Legend/><Bar dataKey="Leads Given" radius={[8,8,0,0]}/><Bar dataKey="Booked Intro Calls" radius={[8,8,0,0]}/><Bar dataKey="Paid Engagements" radius={[8,8,0,0]}/></BarChart></ResponsiveContainer></div>
          <p className="mt-4 rounded-2xl bg-amber-50 p-4 text-sm leading-6 text-amber-900"><AlertTriangle className="mr-2 inline" size={16}/> SDR interpretation: compare booked calls against adjusted intro calls and paid engagements. High booked volume with low paid conversion can point to lead quality, follow-up, or onboarding friction.</p>
        </Section>
        <Section title="Engagement Revenue Status" desc="Paid vs unpaid engagement status based on the Engagements sheet.">
          <div className="h-72"><ResponsiveContainer><PieChart><Tooltip/><Pie data={[{name:"Paid", value: engagements.filter(r=>String(r["Paid Status"]).toLowerCase().includes("paid") && !String(r["Paid Status"]).toLowerCase().includes("unpaid")).length},{name:"Unpaid", value: engagements.filter(r=>String(r["Paid Status"]).toLowerCase().includes("unpaid")).length}]} dataKey="value" nameKey="name" outerRadius={105} label>{[0,1].map(i=><Cell key={i}/>)}</Pie></PieChart></ResponsiveContainer></div>
          <div className="grid gap-3 md:grid-cols-2"><Kpi title="Billed Amount" value={money(sum(engagements,"Billed Amount"))} sub="Total billed from engagement records." icon={DollarSign}/><Kpi title="Paid Amount" value={money(sum(engagements,"Paid Amount"))} sub="Collected revenue from engagement records." icon={RefreshCw} tone="green"/></div>
        </Section>
      </div>

      <Section title="Recent Engagement Activity" desc="A quick operational feed for new engagement starts, status, and payment amounts.">
        <div className="overflow-hidden rounded-3xl border border-slate-200"><table className="w-full min-w-[1000px] text-left text-sm"><thead className="bg-slate-100 text-slate-700"><tr>{["Week","Start Date","Coach","Client","Engagement","Status","Billed","Paid"].map(h=><th className="px-4 py-3 font-semibold" key={h}>{h}</th>)}</tr></thead><tbody className="divide-y divide-slate-100 bg-white">{engagements.slice(0,15).map((r,i)=><tr key={i}><td className="px-4 py-3">{String(r.Week || "")}</td><td className="px-4 py-3">{String(r["Engagement Start Date"] || "")}</td><td className="px-4 py-3 font-medium">{r.Coach}</td><td className="px-4 py-3">{r.Client}</td><td className="px-4 py-3 text-slate-500">{r.Engagement}</td><td className="px-4 py-3">{r["Paid Status"]}</td><td className="px-4 py-3">{money(r["Billed Amount"])}</td><td className="px-4 py-3">{money(r["Paid Amount"])}</td></tr>)}</tbody></table></div>
      </Section>
    </div>
  </main>;
}