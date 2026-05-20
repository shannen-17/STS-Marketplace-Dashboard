import { Row, WorkbookData, KPIMetrics, CoachMetrics } from '@/lib/types';
import { getSheet } from '@/lib/parsers/workbookParser';

const clean = (v: any): number => {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const num = Number(v.replace(/[$,%]/g, ''));
    return Number.isNaN(num) ? 0 : num;
  }
  return 0;
};

const sum = (rows: Row[], key: string | null): number => {
  if (!key) return 0;
  return rows.reduce((acc, row) => acc + clean(row[key]), 0);
};

const findBestMatch = (row: Row, patterns: string[]): any => {
  const keys = Object.keys(row);
  for (const pattern of patterns) {
    const key = keys.find((k) => k.toLowerCase().includes(pattern.toLowerCase()));
    if (key) return row[key];
  }
  return null;
};

const findKey = (rows: Row[], patterns: string[]): string | null => {
  if (rows.length === 0) return null;
  const headers = Object.keys(rows[0]);
  for (const pattern of patterns) {
    const key = headers.find((h) => h.toLowerCase().includes(pattern.toLowerCase()));
    if (key) return key;
  }
  return null;
};

export function calculateGlobalKPIs(data: WorkbookData): KPIMetrics {
  const coachWeekly = getSheet(data, 'Coach Weekly');
  const engagements = getSheet(data, 'Engagements');
  const overallRevenue = getSheet(data, 'Overall Revenue');

  const leadsGiven = sum(coachWeekly, findKey(coachWeekly, ['leads given', 'form submission']));
  const bookedIntroCalls = sum(coachWeekly, findKey(coachWeekly, ['booked intro', 'intro calls booked']));
  const adjustedIntroCalls = sum(coachWeekly, findKey(coachWeekly, ['adjusted intro']));
  const paidEngagements = sum(coachWeekly, findKey(coachWeekly, ['paid engagement', 'paid engagements']));
  const revenue = sum(overallRevenue, findKey(overallRevenue, ['revenue', 'net revenue'])) ||
                 sum(engagements, findKey(engagements, ['paid amount']));
  const billedAmount = sum(engagements, findKey(engagements, ['billed amount']));
  const paidAmount = sum(engagements, findKey(engagements, ['paid amount']));

  const introCallRate = leadsGiven > 0 ? bookedIntroCalls / leadsGiven : 0;
  const paidEngagementRate = leadsGiven > 0 ? paidEngagements / leadsGiven : 0;
  const paidConversionFromIntro = adjustedIntroCalls > 0 ? paidEngagements / adjustedIntroCalls : 0;

  return {
    leadsGiven: Math.round(leadsGiven),
    bookedIntroCalls: Math.round(bookedIntroCalls),
    cancelledCalls: 0,
    noShowCalls: 0,
    notFoundZoom: 0,
    adjustedIntroCalls: Math.round(adjustedIntroCalls),
    engagementsStarted: 0,
    paidEngagements: Math.round(paidEngagements),
    unpaidEngagements: 0,
    introCallRate,
    showUpRate: 0,
    paidEngagementRate,
    paidConversionFromIntro,
    revenue: Math.round(revenue),
    paidAmount: Math.round(paidAmount),
    billedAmount: Math.round(billedAmount),
    currentTakeHome: Math.round(paidAmount),
    revenuePerLead: leadsGiven > 0 ? revenue / leadsGiven : 0,
    revenuePerCoach: 0
  };
}

export function calculateCoachMetrics(data: WorkbookData): CoachMetrics[] {
  const coachWeekly = getSheet(data, 'Coach Weekly');
  const overallRevenue = getSheet(data, 'Overall Revenue');

  const coachMap: Record<string, any> = {};

  coachWeekly.forEach((row) => {
    const coach = findBestMatch(row, ['coach', 'name']) || 'Unknown';
    if (!coachMap[coach]) {
      coachMap[coach] = {
        coach,
        leadsGiven: 0,
        bookedIntroCalls: 0,
        adjustedIntroCalls: 0,
        paidEngagements: 0,
        revenue: 0
      };
    }
    coachMap[coach].leadsGiven += clean(findBestMatch(row, ['leads given', 'form']));
    coachMap[coach].bookedIntroCalls += clean(findBestMatch(row, ['booked intro']));
    coachMap[coach].adjustedIntroCalls += clean(findBestMatch(row, ['adjusted intro']));
    coachMap[coach].paidEngagements += clean(findBestMatch(row, ['paid engagement']));
  });

  overallRevenue.forEach((row) => {
    const coach = findBestMatch(row, ['coach', 'name']) || 'Unknown';
    if (coachMap[coach]) {
      coachMap[coach].revenue += clean(findBestMatch(row, ['take home', 'revenue', 'paid']));
    }
  });

  return Object.values(coachMap)
    .map((m: any) => ({
      ...m,
      cancelledCalls: 0,
      noShowCalls: 0,
      notFoundZoom: 0,
      engagementsStarted: 0,
      unpaidEngagements: 0,
      introCallRate: m.leadsGiven > 0 ? m.bookedIntroCalls / m.leadsGiven : 0,
      showUpRate: 0,
      paidEngagementRate: m.leadsGiven > 0 ? m.paidEngagements / m.leadsGiven : 0,
      paidConversionFromIntro: m.adjustedIntroCalls > 0 ? m.paidEngagements / m.adjustedIntroCalls : 0,
      paidAmount: m.revenue,
      billedAmount: 0,
      currentTakeHome: m.revenue,
      revenuePerLead: m.leadsGiven > 0 ? m.revenue / m.leadsGiven : 0,
      revenuePerCoach: m.revenue,
      score: Math.round(
        ((m.paidEngagements / Math.max(m.leadsGiven, 1)) * 55 +
        (m.bookedIntroCalls / Math.max(m.leadsGiven, 1)) * 25 +
        m.paidEngagements * 2) * 100
      )
    }))
    .sort((a, b) => b.paidEngagementRate - a.paidEngagementRate);
}
