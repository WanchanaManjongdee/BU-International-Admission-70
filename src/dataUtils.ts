import * as XLSX from 'xlsx';
import { ApplicationRecord, sample50Applications } from './sampleData';

export interface DashboardStats {
  totalApplications: number;
  byAdmissionOwner: { owner: string; displayName: string; count: number; percentage: number; isUnassigned: boolean }[];
  byOwner: { owner: string; count: number; percentage: number }[];
  byStage: { stage: string; count: number; percentage: number }[];
  byMonth: { month: string; label: string; count: number; stageBreakdown: Record<string, number> }[];
  reviewAging: {
    totalReviewDocs: number;
    lessThan1Week: { count: number; percentage: number; list: ApplicationRecord[] };
    twoToFourWeeks: { count: number; percentage: number; list: ApplicationRecord[] };
    moreThan4Weeks: { count: number; percentage: number; list: ApplicationRecord[] };
  };
  confirmedCount: number;
  invoicedCount: number;
  waitLanguageCount: number;
  closedLostCount: number;
  unassignedAdmissionCount: number;
}

// Reference date for calculating aging: use snapshot date of HubSpot CRM export or max date
export const DEFAULT_REFERENCE_DATE = new Date('2026-11-10T23:59:59');

export function parseCSVText(csvText: string, referenceDate: Date = DEFAULT_REFERENCE_DATE): ApplicationRecord[] {
  const lines = csvText.split(/\r?\n/).filter(line => line.trim().length > 0);
  if (lines.length <= 1) return [];

  // Parse header
  const headers = parseCSVLine(lines[0]).map(h => h.trim().replace(/^"|"$/g, ''));
  
  // Find column indices with exact priority to avoid collision between 'Close Date' and 'Create Date'
  const idIdx = headers.findIndex(h => /^record id$/i.test(h) || /^id$/i.test(h)) !== -1
    ? headers.findIndex(h => /^record id$/i.test(h) || /^id$/i.test(h))
    : headers.findIndex(h => /record id/i.test(h) || /id/i.test(h));

  const nameIdx = headers.findIndex(h => /application name/i.test(h) || /student name/i.test(h) || /name/i.test(h));
  const stageIdx = headers.findIndex(h => /application stage/i.test(h) || /^stage$/i.test(h));
  const closeIdx = headers.findIndex(h => /close\s*date/i.test(h));
  
  // Application owner vs Admission owner
  const ownerIdx = headers.findIndex(h => /application\s*owner/i.test(h)) !== -1
    ? headers.findIndex(h => /application\s*owner/i.test(h))
    : headers.findIndex(h => /owner/i.test(h) && !/admission/i.test(h));
  const admissionIdx = headers.findIndex(h => /admission\s*owner/i.test(h));

  // Create date must NOT match close date or last activity date
  const createIdx = headers.findIndex(h => /create\s*date/i.test(h)) !== -1
    ? headers.findIndex(h => /create\s*date/i.test(h))
    : headers.findIndex(h => /create/i.test(h) || (/date/i.test(h) && !/close|last/i.test(h)));

  const appYearIdx = headers.findIndex(h => /application\s*year/i.test(h));
  const termIdx = headers.findIndex(h => /intake\s*term/i.test(h));
  const actIdx = headers.findIndex(h => /last activity/i.test(h));
  const intakeYearIdx = headers.findIndex(h => /intake year/i.test(h));

  const records: ApplicationRecord[] = [];

  for (let i = 1; i < lines.length; i++) {
    const rawLine = lines[i];
    if (!rawLine.trim()) continue;
    const cols = parseCSVLine(rawLine).map(c => c.trim().replace(/^"|"$/g, ''));
    if (cols.length < 3) continue;

    const rawNameCol = (nameIdx >= 0 && cols[nameIdx]) ? cols[nameIdx] : 'Unknown';
    // Often format is "First Last email@example.com"
    let name = rawNameCol;
    let email = '';
    const emailMatch = rawNameCol.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailMatch) {
      email = emailMatch[1];
      name = rawNameCol.replace(emailMatch[0], '').trim();
    }

    const stage = (stageIdx >= 0 && cols[stageIdx]) ? cols[stageIdx].trim() : 'Unknown';
    const owner = (ownerIdx >= 0 && cols[ownerIdx]) ? cols[ownerIdx].trim() : 'Unassigned';
    const createDate = (createIdx >= 0 && cols[createIdx]) ? cols[createIdx].trim() : '';
    const applicationYear = (appYearIdx >= 0 && cols[appYearIdx]) ? cols[appYearIdx].trim() : '';
    const intakeYearApp = (intakeYearIdx >= 0 && cols[intakeYearIdx]) ? cols[intakeYearIdx].trim() : '';

    const rec: ApplicationRecord = {
      id: (idIdx >= 0 && cols[idIdx]) ? cols[idIdx] : String(i),
      name: name || 'Applicant #' + i,
      email: email,
      stage: stage,
      closeDate: (closeIdx >= 0 && cols[closeIdx]) ? cols[closeIdx] : '',
      owner: owner || 'Unassigned',
      admissionOwner: (admissionIdx >= 0 && cols[admissionIdx]) ? cols[admissionIdx] : '',
      createDate: createDate,
      applicationYear: applicationYear,
      intakeTerm: (termIdx >= 0 && cols[termIdx]) ? cols[termIdx] : '',
      lastActivityDate: (actIdx >= 0 && cols[actIdx]) ? cols[actIdx] : '',
      intakeYearApp: intakeYearApp,
    };

    // Calculate aging
    calculateRecordAging(rec, referenceDate);
    records.push(rec);
  }

  return records;
}

export function parseExcelFile(arrayBuffer: ArrayBuffer, referenceDate: Date = DEFAULT_REFERENCE_DATE): ApplicationRecord[] {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const rawJson: Record<string, unknown>[] = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

  return rawJson.map((row, index) => {
    // Find keys flexibly with optional exclusion
    const keys = Object.keys(row);
    const findVal = (regex: RegExp, excludeRegex?: RegExp): string => {
      const match = keys.find(k => regex.test(k) && (!excludeRegex || !excludeRegex.test(k)));
      return match ? String(row[match] || '').trim() : '';
    };

    const rawName = findVal(/application name|student name|name|ชื่อ/i);
    let name = rawName;
    let email = '';
    const emailMatch = rawName.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailMatch) {
      email = emailMatch[1];
      name = rawName.replace(emailMatch[0], '').trim();
    }
    if (!email) {
      email = findVal(/email|อีเมล/i);
    }

    const rec: ApplicationRecord = {
      id: findVal(/record id|id|รหัส/i) || String(index + 1),
      name: name || `Applicant #${index + 1}`,
      email: email,
      stage: findVal(/application stage|stage|สถานะ|สเตจ/i) || 'Unknown',
      closeDate: findVal(/close date/i),
      owner: findVal(/application owner/i) || findVal(/owner|ผู้ดูแล|ผู้รับผิดชอบ/i, /admission/i) || 'Unassigned',
      admissionOwner: findVal(/admission owner|admission/i),
      createDate: findVal(/create\s*date|created|วันที่สมัคร/i) || findVal(/วันที่|date/i, /close|last|ปิด/i),
      applicationYear: findVal(/application year|year|ปี/i, /intake/i),
      intakeTerm: findVal(/intake term|term|เทอม/i),
      lastActivityDate: findVal(/last activity/i),
      intakeYearApp: findVal(/intake year/i),
    };

    calculateRecordAging(rec, referenceDate);
    return rec;
  });
}

function parseCSVLine(text: string): string[] {
  const result: string[] = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (inQuotes && text[i + 1] === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (c === ',' && !inQuotes) {
      result.push(cur);
      cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur);
  return result;
}

export function calculateRecordAging(record: ApplicationRecord, referenceDate: Date) {
  if (!record.createDate) return;
  const created = new Date(record.createDate.replace(' ', 'T'));
  if (isNaN(created.getTime())) return;

  const diffMs = referenceDate.getTime() - created.getTime();
  const diffDays = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
  record.agingDays = diffDays;

  if (diffDays <= 7) {
    record.agingBucket = 'less_1w';
  } else if (diffDays <= 28) {
    record.agingBucket = '2_to_4w';
  } else {
    record.agingBucket = 'more_4w';
  }
}

export function getAdmissionOwnerDisplayName(rawOwner: string): string {
  const trimmed = (rawOwner || '').trim();
  if (!trimmed || trimmed.toLowerCase() === 'unassigned') {
    return 'ยังไม่ระบุ (Unassigned)';
  }
  const emailMap: Record<string, string> = {
    'sarawut.p@bu.ac.th': 'Sarawut P. (sarawut.p@bu.ac.th)',
    'krittin.c@bu.ac.th': 'Krittin C. (krittin.c@bu.ac.th)',
    'pattama.m@bu.ac.th': 'Pattama M. (pattama.m@bu.ac.th)',
    'raweewan.m@bu.ac.th': 'Raweewan M. (raweewan.m@bu.ac.th)',
    'janpim.s@bu.ac.th': 'Janpim S. (janpim.s@bu.ac.th)',
    'piyaporn.s@bu.ac.th': 'Piyaporn S. (piyaporn.s@bu.ac.th)',
    'thitinan.n@bu.ac.th': 'Thitinan N. (thitinan.n@bu.ac.th)',
    'wanchana.m@bu.ac.th': 'Wanchana M. (wanchana.m@bu.ac.th)',
  };
  return emailMap[trimmed.toLowerCase()] || trimmed;
}

export function computeDashboardStats(records: ApplicationRecord[]): DashboardStats {
  const total = records.length;

  // By Admission Owner
  const admissionOwnerCounts: Record<string, number> = {};
  // By Owner (Application Owner)
  const ownerCounts: Record<string, number> = {};
  // By Stage
  const stageCounts: Record<string, number> = {};
  // By Month (Format: YYYY-MM)
  const monthMap: Record<string, { count: number; stageBreakdown: Record<string, number> }> = {};

  const reviewAgingBuckets = {
    lessThan1Week: [] as ApplicationRecord[],
    twoToFourWeeks: [] as ApplicationRecord[],
    moreThan4Weeks: [] as ApplicationRecord[],
  };

  let confirmedCount = 0;
  let invoicedCount = 0;
  let waitLanguageCount = 0;
  let closedLostCount = 0;
  let unassignedAdmissionCount = 0;

  records.forEach(rec => {
    // Admission Owner count
    const admOwnerRaw = (rec.admissionOwner || '').trim();
    const admKey = admOwnerRaw || 'Unassigned';
    if (!admOwnerRaw || admOwnerRaw.toLowerCase() === 'unassigned') {
      unassignedAdmissionCount++;
    }
    admissionOwnerCounts[admKey] = (admissionOwnerCounts[admKey] || 0) + 1;

    // Application Owner count
    const ownerName = rec.owner || 'Unassigned';
    ownerCounts[ownerName] = (ownerCounts[ownerName] || 0) + 1;

    // Stage count
    const stageName = rec.stage || 'Unknown';
    stageCounts[stageName] = (stageCounts[stageName] || 0) + 1;

    if (stageName === 'Confirmed Payment') confirmedCount++;
    if (stageName === 'Invoiced') invoicedCount++;
    if (stageName === 'Wait for Language Result') waitLanguageCount++;
    if (stageName === 'Closed Lost') closedLostCount++;

    // Month extraction
    let monthKey = 'Unknown';
    if (rec.createDate) {
      const match = rec.createDate.match(/^(\d{4}-\d{2})/);
      if (match) {
        monthKey = match[1];
      }
    }
    if (!monthMap[monthKey]) {
      monthMap[monthKey] = { count: 0, stageBreakdown: {} };
    }
    monthMap[monthKey].count++;
    monthMap[monthKey].stageBreakdown[stageName] = (monthMap[monthKey].stageBreakdown[stageName] || 0) + 1;

    // Check review document aging
    if (stageName === 'Review Document') {
      if (rec.agingBucket === 'less_1w') {
        reviewAgingBuckets.lessThan1Week.push(rec);
      } else if (rec.agingBucket === '2_to_4w') {
        reviewAgingBuckets.twoToFourWeeks.push(rec);
      } else {
        reviewAgingBuckets.moreThan4Weeks.push(rec);
      }
    }
  });

  const byAdmissionOwner = Object.entries(admissionOwnerCounts)
    .map(([owner, count]) => ({
      owner,
      displayName: getAdmissionOwnerDisplayName(owner),
      count,
      percentage: total > 0 ? (count / total) * 100 : 0,
      isUnassigned: owner === 'Unassigned' || !owner,
    }))
    .sort((a, b) => {
      // Put assigned ones first sorted by count, or sort strictly by count
      return b.count - a.count;
    });

  const byOwner = Object.entries(ownerCounts)
    .map(([owner, count]) => ({
      owner,
      count,
      percentage: total > 0 ? (count / total) * 100 : 0,
    }))
    .sort((a, b) => b.count - a.count);

  const byStage = Object.entries(stageCounts)
    .map(([stage, count]) => ({
      stage,
      count,
      percentage: total > 0 ? (count / total) * 100 : 0,
    }))
    .sort((a, b) => b.count - a.count);

  // Month sorting
  const thaiMonths: Record<string, string> = {
    '01': 'ม.ค.', '02': 'ก.พ.', '03': 'มี.ค.', '04': 'เม.ย.',
    '05': 'พ.ค.', '06': 'มิ.ย.', '07': 'ก.ค.', '08': 'ส.ค.',
    '09': 'ก.ย.', '10': 'ต.ค.', '11': 'พ.ย.', '12': 'ธ.ค.',
  };

  const byMonth = Object.keys(monthMap)
    .filter(k => k !== 'Unknown')
    .sort()
    .map(m => {
      const parts = m.split('-');
      const y = parseInt(parts[0], 10);
      const thaiYear = y + 543;
      const mStr = thaiMonths[parts[1]] || parts[1];
      const label = `${mStr} ${thaiYear}`;
      return {
        month: m,
        label,
        count: monthMap[m].count,
        stageBreakdown: monthMap[m].stageBreakdown,
      };
    });

  const totalReviews =
    reviewAgingBuckets.lessThan1Week.length +
    reviewAgingBuckets.twoToFourWeeks.length +
    reviewAgingBuckets.moreThan4Weeks.length;

  return {
    totalApplications: total,
    byAdmissionOwner,
    byOwner,
    byStage,
    byMonth,
    reviewAging: {
      totalReviewDocs: totalReviews,
      lessThan1Week: {
        count: reviewAgingBuckets.lessThan1Week.length,
        percentage: totalReviews > 0 ? (reviewAgingBuckets.lessThan1Week.length / totalReviews) * 100 : 0,
        list: reviewAgingBuckets.lessThan1Week,
      },
      twoToFourWeeks: {
        count: reviewAgingBuckets.twoToFourWeeks.length,
        percentage: totalReviews > 0 ? (reviewAgingBuckets.twoToFourWeeks.length / totalReviews) * 100 : 0,
        list: reviewAgingBuckets.twoToFourWeeks,
      },
      moreThan4Weeks: {
        count: reviewAgingBuckets.moreThan4Weeks.length,
        percentage: totalReviews > 0 ? (reviewAgingBuckets.moreThan4Weeks.length / totalReviews) * 100 : 0,
        list: reviewAgingBuckets.moreThan4Weeks,
      },
    },
    confirmedCount,
    invoicedCount,
    waitLanguageCount,
    closedLostCount,
    unassignedAdmissionCount,
  };
}

// Anonymize personal student data (PDPA Compliance)
export function maskStudentName(name: string, isMasked: boolean = true): string {
  if (!isMasked || !name) return name;
  const parts = name.split(/\s+/).filter(Boolean);
  if (parts.length === 0) return 'Applicant-***';
  return parts.map(part => {
    if (part.length <= 2) return part.charAt(0) + '*';
    return part.charAt(0) + '*'.repeat(Math.min(part.length - 2, 4)) + part.charAt(part.length - 1);
  }).join(' ');
}

export function maskEmail(email: string, isMasked: boolean = true): string {
  if (!isMasked || !email) return email;
  const atIdx = email.indexOf('@');
  if (atIdx === -1) return '***';
  const local = email.substring(0, atIdx);
  const domain = email.substring(atIdx);
  const maskedLocal = local.length <= 2 ? local.charAt(0) + '***' : local.charAt(0) + '***' + local.charAt(local.length - 1);
  return maskedLocal + domain;
}

// Export records to CSV
export function exportToCSV(records: ApplicationRecord[], filename = 'bu-applications-export.csv') {
  const headers = ['Record ID', 'Application Name', 'Application Stage', 'Application Owner', 'Create Date', 'Aging Days', 'Intake Year'];
  const rows = records.map(r => [
    `"${r.id}"`,
    `"${r.name}"`,
    `"${r.stage}"`,
    `"${r.owner}"`,
    `"${r.createDate}"`,
    `"${r.agingDays ?? ''}"`,
    `"${r.intakeYearApp || r.applicationYear || ''}"`,
  ]);

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
