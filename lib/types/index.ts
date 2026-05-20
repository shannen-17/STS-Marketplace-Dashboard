export type Row = Record<string, any>;
export type WorkbookData = Record<string, Row[]>;

export interface ParsedWorkbook {
  sheets: SheetInfo[];
  data: WorkbookData;
  metadata: WorkbookMetadata;
}

export interface SheetInfo {
  name: string;
  rowCount: number;
  headers: string[];
  sample: Row[];
}

export interface WorkbookMetadata {
  parsedAt: Date;
  fileName: string;
  totalSheets: number;
  validSheets: SheetInfo[];
  errors: string[];
}

export interface KPIMetrics {
  leadsGiven: number;
  bookedIntroCalls: number;
  cancelledCalls: number;
  noShowCalls: number;
  notFoundZoom: number;
  adjustedIntroCalls: number;
  engagementsStarted: number;
  paidEngagements: number;
  unpaidEngagements: number;
  introCallRate: number;
  showUpRate: number;
  paidEngagementRate: number;
  paidConversionFromIntro: number;
  revenue: number;
  paidAmount: number;
  billedAmount: number;
  currentTakeHome: number;
  revenuePerLead: number;
  revenuePerCoach: number;
}

export interface CoachMetrics extends KPIMetrics {
  coach: string;
  score: number;
}

export interface CohortMetrics {
  cohortWeek: string;
  leadsGenerated: number;
  delayedConversions: number;
  conversionRate: number;
  projectedValue: number;
  avgDaysToConversion: number;
}

export interface ExecutiveInsight {
  category: string;
  text: string;
  severity: 'positive' | 'neutral' | 'warning';
  metric?: string;
  value?: string;
}
