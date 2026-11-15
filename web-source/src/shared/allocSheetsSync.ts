import {
  DEFAULT_OFFICER_PROFILE,
  createDefaultTemplatesMap,
  type AllocationCategory,
  type AllocationDatabase,
  type CategoryRate,
  type ChairmanEntry,
  type ChallanRecord,
  type DCAllocation,
  type DOBook,
  type DOIssue,
  type DealerSuballocation,
  type FireServiceEntry,
  type FoodDealer,
  type MinistryAllocation,
  type OMSDealer,
  type ProjectEntry,
  type UNOSuballocation,
} from '../types/allocation';
import type { DriveSpreadsheetFile } from '../types/grain';
import { bnToEnDigits, fiscalYearLabel, isChallanOnOrBeforeDo, isSamePartyName, memosLikelyMatch, namesFuzzyMatch, namesLikelyMatch, normMemoNo, parseMonthYear, parseNum, round2, round3, toBnDigits, toIsoDateStr, uid } from './bnUtils';
import { callAndroidNativeHttp, callGasAction } from './gasBridge';
import {
  GoogleAuthExpiredError,
  PublicSheetAccessError,
  extractSpreadsheetId,
  getPublicSpreadsheetMetadata,
  getSpreadsheetMetadata,
  listUserSpreadsheets,
  parseCsvToRows,
  type SpreadsheetMetadata,
} from './googleSheetsSync';

const SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

export const ALLOC_SHEET_TABS = {
  ministry: {
    name: 'Ministry_Allocation',
    aliases: [
      'ministry_allocation',
      'ministryallocation',
      'ministry_alloc',
      'ministry_entry',
      'ministry',
      'allocation',
      'ministry_allocations',
      'ministryallocations',
      'ministry_entries',
      'ministry_list',
      'allocations',
      'মন্ত্রণালয়_বরাদ্দ',
      'মন্ত্রণালয়_বরাদ্দ_এন্ট্রি',
      'মন্ত্রণালয়ের_বরাদ্দ',
      'মন্ত্রণালয়',
      'বরাদ্দ',
    ],
    headers: ['category', 'memoNo', 'memoDate', 'allocationQty', 'grain', 'phase', 'comment', 'createdAt', 'id'],
  },
  dc: {
    name: 'DC_Allocation',
    aliases: [
      'dc_allocation',
      'dcallocation',
      'dc_alloc',
      'dc_entry',
      'district_allocation',
      'dc',
      'dc_allocations',
      'dcallocations',
      'dc_entries',
      'district_allocations',
      'জেলা_প্রশাসক',
      'জেলা_প্রশাসক_বরাদ্দ',
      'ডিসি_বরাদ্দ',
      'ডিসি_বরাদ্দ_এন্ট্রি',
      'ডিসি',
    ],
    headers: ['category', 'ministryMemo', 'dcMemo', 'dcDate', 'allocationQty', 'createdAt', 'id'],
  },
  dealerSub: {
    name: 'Dealer_Suballocation',
    aliases: [
      'dealer_suballocation',
      'dealersuballocation',
      'dealer_allocation',
      'dealer_entry',
      'suballocation',
      'sub_allocation',
      'dealer_sub',
      'ucf_suballocation',
      'ucf_allocation',
      'dealer_suballocations',
      'dealersuballocations',
      'dealer_allocations',
      'suballocations',
      'sub_allocations',
      'ucf_suballocations',
      'ucf_allocations',
      'ডিলার_উপবরাদ্দ',
      'ইউসিএফ_উপবরাদ্দ',
      'উপবরাদ্দ',
    ],
    headers: ['category', 'sourceMemo', 'unoMemo', 'date', 'dealerName', 'unionName', 'salesCenter', 'cardCount', 'unusedCards', 'allocationQty', 'totalAmount', 'taxAmount', 'createdAt', 'id'],
  },
  grProjects: {
    name: 'GR_Project',
    aliases: [
      'gr_project',
      'grproject',
      'gr_projects',
      'grprojects',
      'জিআর_প্রকল্প',
      'জি_আর_প্রকল্প',
      'জিআরপ্রকল্প',
      'জি_আর_প্রকল্প_তালিকা',
    ],
    headers: ['category', 'dcMemo', 'projectNo', 'projectName', 'presidentName', 'unionName', 'allocationQty', 'createdAt', 'id'],
  },
  kabikhaProjects: {
    name: 'Kabikha_Project',
    aliases: [
      'kabikha_project',
      'kabikhaproject',
      'kabikha_projects',
      'kabikhaprojects',
      'কাবিখা_প্রকল্প',
      'কাবিখাপ্রকল্প',
      'কাবিখা_প্রকল্প_তালিকা',
    ],
    headers: ['category', 'dcMemo', 'projectNo', 'projectName', 'presidentName', 'unionName', 'allocationQty', 'createdAt', 'id'],
  },
  unoSub: {
    name: 'UNO_Suballocation',
    aliases: [
      'uno_suballocation',
      'unosuballocation',
      'uno_allocation',
      'unoallocation',
      'uno_sub',
      'uno_entry',
      'upazila_suballocation',
      'upazila_allocation',
      'uno',
      'relief_suballocation',
      'uno_suballocations',
      'unosuballocations',
      'uno_allocations',
      'unoallocations',
      'uno_entries',
      'upazila_suballocations',
      'upazila_allocations',
      'relief_suballocations',
      'ইউএনও_উপবরাদ্দ',
      'ইউএনও_উপবরাদ্দ_এন্ট্রি',
      'ইউএনও_বরাদ্দ',
      'ত্রাণ_উপবরাদ্দ',
      'ত্রাণ_বরাদ্দ',
    ],
    headers: ['category', 'ministryMemo', 'dcMemo', 'unoMemo', 'unoDate', 'unionName', 'recipientName', 'projectNo', 'projectName', 'allocationQty', 'createdAt', 'id'],
  },
  doBooks: {
    name: 'DO_Book',
    aliases: ['do_book', 'dobook', 'do_books', 'dobooks', 'ডিও_বই'],
    headers: ['bookNo', 'startPage', 'endPage', 'totalCount', 'createdAt', 'id'],
  },
  doIssue: {
    name: 'DO_Issue',
    aliases: ['do_issue', 'doissue', 'do_entry', 'do_list', 'do_register', 'do_issues', 'do_entries', 'ডিও_ইস্যু', 'ডিও_এন্ট্রি', 'ডিও'],
    headers: ['doNo', 'issueDate', 'category', 'unionName', 'recipientName', 'salesCenter', 'allocationQty', 'ministryMemo', 'dcMemo', 'unoMemo', 'projectNo', 'dealerAllocationId', 'challanNo', 'challanDate', 'challanAmount', 'taxChallanNo', 'taxChallanDate', 'taxChallanAmount', 'requiredAmount', 'requiredTaxAmount', 'createdAt', 'id'],
  },
  challans: {
    name: 'Challan',
    aliases: [
      'challan',
      'challan_entry',
      'challan_register',
      'treasury_challan',
      'challans',
      'treasury_challans',
      'চালান',
      'ট্রেজারি_চালান',
      'চালান_এন্ট্রি',
      'চালান_রেজিস্টার',
    ],
    headers: ['challanNo', 'challanDate', 'category', 'dealerName', 'challanType', 'bankName', 'doNo', 'recordedAmount', 'depositedAmount', 'note', 'createdAt', 'dealerAllocationId', 'doDate', 'allocationQty', 'id'],
  },
  foodDealers: {
    name: 'Food_Dealer',
    aliases: ['food_dealer', 'fooddealer', 'dealer_food', 'food_dealers', 'dealers_food', 'খাদ্য_বান্ধব_ডিলার', 'খাদ্যবান্ধব_ডিলার'],
    headers: ['id', 'dealerName', 'fatherName', 'nid', 'mobile', 'unionName', 'salesCenter', 'cardCount', 'totalQty', 'poNo', 'poDate', 'bankName', 'amount', 'supervisorName', 'supervisorDesignation', 'supervisorMobile'],
  },
  omsDealers: {
    name: 'OMS_Dealer',
    aliases: ['oms_dealer', 'omsdealer', 'dealer_oms', 'oms_dealers', 'dealers_oms', 'ওএমএস_ডিলার'],
    headers: ['id', 'dealerName', 'fatherName', 'nid', 'mobile', 'areaType', 'unionName', 'municipality', 'salesCenter', 'poNo', 'poDate', 'bankName', 'amount', 'supervisorName', 'supervisorDesignation', 'supervisorMobile'],
  },
  chairmen: {
    name: 'Chairman',
    aliases: ['chairman', 'chairmen', 'চেয়ারম্যান', 'ইউপি_চেয়ারম্যান'],
    headers: ['id', 'unionName', 'chairmanName', 'mobile', 'supervisorName', 'supervisorMobile'],
  },
  fireServices: {
    name: 'Fire_Service',
    aliases: ['fire_service', 'fireservice', 'ep_office', 'fire_services', 'fireservices', 'ep_offices', 'ফায়ার_সার্ভিস'],
    headers: ['id', 'officeName', 'divisionName', 'riceRate', 'wheatRate'],
  },
  rates: {
    name: 'Category_Rate',
    aliases: ['category_rate', 'rate', 'price_rate', 'category_rates', 'rates', 'price_rates', 'দর', 'একক_মূল্য'],
    headers: ['id', 'category', 'unitPrice', 'taxRate'],
  },
  settings: {
    name: 'Alloc_Settings',
    aliases: ['alloc_settings', 'alloc_setting', 'settings', 'report_settings', 'config', 'সেটিংস'],
    headers: ['key', 'value'],
  },
} as const;

type AllocationTabKey = keyof typeof ALLOC_SHEET_TABS;

function normKey(s: unknown): string {
  return String(s ?? '')
    .normalize('NFC')
    .replace(/\u09AF\u09BC/g, '\u09DF')
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, '');
}

function normColKey(s: unknown): string {
  return String(s ?? '')
    .normalize('NFC')
    .replace(/\u09AF\u09BC/g, '\u09DF')
    .trim()
    .toLowerCase()
    .replace(/[\s_.\-/():]+/g, '');
}

/**
 * Detect whether a sheet table predominantly uses US M/D/YYYY dates (e.g. 7/19/2026, 8/27/2026)
 * vs D/M/YYYY dates. Returns true if MDY markers appear strictly without DMY markers.
 */
function detectPreferMdy(rows: unknown[][]): boolean {
  let mdyHits = 0;
  let dmyHits = 0;
  const maxRows = Math.min(rows.length, 250);
  for (let i = 1; i < maxRows; i++) {
    const r = rows[i] || [];
    for (const cell of r) {
      if (typeof cell !== 'string') continue;
      const s = bnToEnDigits(cell).trim();
      const slash = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
      if (!slash) continue;
      const a = Number(slash[1]);
      const b = Number(slash[2]);
      if (b > 12 && a >= 1 && a <= 12) mdyHits += 1;
      else if (a > 12 && b >= 1 && b <= 12) dmyHits += 1;
    }
  }
  return mdyHits > dmyHits;
}

function parseDateStr(val: unknown, preferMdy = false): string {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number' && val > 20000 && val < 80000) {
    const ms = Math.round((val - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  const s = bnToEnDigits(String(val)).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const slash = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (slash) {
    const a = Number(slash[1]);
    const b = Number(slash[2]);
    const y = slash[3];
    let mon: number;
    let day: number;
    if (a > 12) {
      day = a;
      mon = b;
    } else if (b > 12) {
      mon = a;
      day = b;
    } else if (preferMdy) {
      mon = a;
      day = b;
    } else {
      day = a;
      mon = b;
    }
    return `${y}-${String(mon).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return s;
}

function normalizeCategoryName(raw: unknown, hintText?: string): AllocationCategory {
  const s = String(raw ?? '')
    .normalize('NFC')
    .replace(/\s+/g, ' ')
    .trim();
  const compact = s.replace(/\s+/g, '').toLowerCase();
  const valid: AllocationCategory[] = [
    'খাদ্য বান্ধব',
    'ওএমএস সাধারণ',
    'ওএমএস বিশেষ',
    'ওএমএস টিসিবি',
    'ইপি চাল',
    'ইপি গম',
    'ভিজিএফ মৎস্য',
    'ভিজিএফ ঈদ',
    'জি আর',
    'কাবিখা সাধারণ চাল',
    'কাবিখা সাধারণ গম',
    'কাবিখা বিশেষ চাল',
    'কাবিখা বিশেষ গম',
    'ভিডাব্লিউভি',
  ];
  const hit = valid.find((v) => v === s || v.replace(/\s+/g, '') === compact);
  if (hit) return hit;

  const classify = (txt: string): AllocationCategory | null => {
    const n = txt.normalize('NFC').trim();
    if (!n) return null;
    const low = n.toLowerCase();
    const cmp = low.replace(/\s+/g, '');

    if ((n.includes('খাদ্য') && n.includes('বান্ধব')) || low.includes('food friendly') || cmp.includes('khadya')) {
      return 'খাদ্য বান্ধব';
    }
    if (n.includes('টিসিবি') || low.includes('tcb')) return 'ওএমএস টিসিবি';
    if (n.includes('বিশেষ') && (n.includes('ওএমএস') || low.includes('oms'))) return 'ওএমএস বিশেষ';
    if (n.includes('ওএমএস') || low.includes('oms')) return 'ওএমএস সাধারণ';
    if ((n.includes('ইপি') || n.includes('ফায়ার') || low.includes('fire') || cmp === 'ep' || low.includes('ep ') || low.includes('ep-')) && (n.includes('গম') || low.includes('wheat'))) {
      return 'ইপি গম';
    }
    if (n.includes('ইপি') || n.includes('ফায়ার') || low.includes('fire') || cmp === 'ep' || low.includes('ep ') || low.includes('ep-')) {
      return 'ইপি চাল';
    }
    if (n.includes('মৎস্য') || low.includes('fish')) return 'ভিজিএফ মৎস্য';
    if (n.includes('ঈদ') || low.includes('eid')) return 'ভিজিএফ ঈদ';
    if (cmp === 'জিআর' || (n.includes('জি') && n.includes('আর')) || cmp === 'gr' || /\bgr\b/i.test(n)) {
      return 'জি আর';
    }
    if (n.includes('কাবিখা') && n.includes('বিশেষ') && (n.includes('গম') || low.includes('wheat'))) {
      return 'কাবিখা বিশেষ গম';
    }
    if (n.includes('কাবিখা') && n.includes('বিশেষ')) return 'কাবিখা বিশেষ চাল';
    if (n.includes('কাবিখা') && (n.includes('গম') || low.includes('wheat'))) return 'কাবিখা সাধারণ গম';
    if (n.includes('কাবিখা') || low.includes('kabikha')) return 'কাবিখা সাধারণ চাল';
    const noDots = low.replace(/[.\-_\s]+/g, '');
    if (
      n.includes('ভিডাব্লিউ') ||
      n.includes('ভিডব্লিউ') ||
      n.includes('ভিজিডি') ||
      n.includes('ভিজিডব্লিউ') ||
      n.includes('জিডব্লিউ') ||
      n.includes('জি ডব্লিউ') ||
      noDots.includes('vwb') ||
      noDots.includes('vgd') ||
      noDots.includes('vwd') ||
      noDots.includes('gwb')
    ) {
      return 'ভিডাব্লিউভি';
    }
    return null;
  };

  return classify(s) || (hintText ? classify(hintText) : null) || 'খাদ্য বান্ধব';
}

const CATEGORY_FIELD_KEYS = [
  'category',
  'khath',
  'khat',
  'খাত',
  'খাতেরনাম',
  'কর্মসূচি',
  'কর্মসূচিরনাম',
  'program',
  'scheme',
  'sector',
  'head',
  'type',
];

function makeUniqueId(rawId: string, prefix: string, index: number, seenIds: Set<string>): string {
  const base = rawId.trim() || `${prefix}-${index}`;
  if (!seenIds.has(base)) {
    seenIds.add(base);
    return base;
  }
  let suffix = 1;
  while (seenIds.has(`${base}-${suffix}`)) suffix++;
  const unique = `${base}-${suffix}`;
  seenIds.add(unique);
  return unique;
}

const HEADER_HINT_TOKENS = [
  'id', 'category', 'khath', 'khat', 'খাত', 'খাতেরনাম', 'কর্মসূচি', 'program', 'scheme', 'sector', 'head',
  'memo', 'স্মারক', 'date', 'তারিখ',
  'qty', 'quantity', 'allocationqty', 'পরিমাণ', 'dealer', 'ডিলার', 'union', 'ইউনিয়ন',
  'dono', 'ডিও', 'challan', 'চালান', 'bookno', 'বইনং', 'amount', 'টাকা',
  'recipient', 'প্রাপক', 'project', 'প্রকল্প', 'bank', 'ব্যাংক', 'chairman', 'চেয়ারম্যান',
];

function findHeaderRowIndex(rows: unknown[][]): number {
  const limit = Math.min(rows.length, 8);
  let bestIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < limit; i++) {
    const r = rows[i] || [];
    const nonEmpty = r.filter((c) => String(c ?? '').trim() !== '');
    if (nonEmpty.length < 2) continue;
    let score = nonEmpty.length;
    for (const cell of nonEmpty) {
      const nk = normKey(cell);
      if (!nk) continue;
      if (HEADER_HINT_TOKENS.some((tok) => nk === tok || nk.includes(tok))) {
        score += 5;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }
  return bestIdx;
}

function mapRowsByHeader(rows: unknown[][]): Array<Record<string, unknown>> {
  if (!rows || rows.length < 2) return [];
  const headerIdx = findHeaderRowIndex(rows);
  if (headerIdx >= rows.length - 1) return [];
  const rawHeaders = (rows[headerIdx] || []).map((h) => String(h ?? '').trim());
  const normHeaders = rawHeaders.map((h) => normKey(h));
  const out: Array<Record<string, unknown>> = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const hasAny = r.some((c) => String(c ?? '').trim() !== '');
    if (!hasAny) continue;
    const obj: Record<string, unknown> = {};
    normHeaders.forEach((nh, idx) => {
      if (nh) obj[nh] = r[idx];
      if (rawHeaders[idx]) obj[rawHeaders[idx]] = r[idx];
    });
    out.push(obj);
  }
  return out;
}

function pickField(obj: Record<string, unknown>, keys: string[]): string {
  for (const k of keys) {
    const nk = normKey(k);
    if (obj[nk] !== undefined && obj[nk] !== null && String(obj[nk]).trim() !== '') {
      return String(obj[nk]).trim();
    }
    if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== '') {
      return String(obj[k]).trim();
    }
  }
  return '';
}

function pickNum(obj: Record<string, unknown>, keys: string[]): number {
  const s = pickField(obj, keys);
  return s ? parseNum(s) : 0;
}

const LEGACY_COMBINED_PROJECTS_TAB = {
  name: 'Project',
  aliases: ['project', 'projects', 'প্রকল্প', 'প্রকল্প_তালিকা'],
} as const;

type AllocationRawTabKey = keyof typeof ALLOC_SHEET_TABS | 'legacyProjects';

function isExactOnlyTabKey(key: AllocationRawTabKey): boolean {
  return key === 'grProjects' || key === 'kabikhaProjects' || key === 'legacyProjects';
}

function isSummaryOrReportTabNorm(norm: string): boolean {
  return (
    norm.includes('dashboard') ||
    norm.includes('summary') ||
    norm.includes('সারসংক্ষেপ') ||
    norm.includes('ড্যাশবোর্ড') ||
    norm.includes('report') ||
    norm.includes('রিপোর্ট') ||
    norm.includes('instructions') ||
    norm.includes('নির্দেশনা')
  );
}

function isIncompatiblePartialTab(tabKey: string | undefined, normTitle: string): boolean {
  if (!tabKey) return false;
  if (isSummaryOrReportTabNorm(normTitle)) return true;
  if (tabKey === 'dealerSub') {
    return (
      normTitle.includes('fooddealer') ||
      normTitle.includes('omsdealer') ||
      normTitle.includes('খাদ্যবান্ধবডিলার') ||
      normTitle.includes('ওএমএসডিলার') ||
      normTitle.includes('uno') ||
      normTitle.includes('upazila') ||
      normTitle.includes('relief') ||
      normTitle.includes('ইউএনও') ||
      normTitle.includes('উপজেলা') ||
      normTitle.includes('ত্রাণ') ||
      normTitle.includes('ministry') ||
      normTitle.includes('dc') ||
      normTitle.includes('মন্ত্রণাল') ||
      normTitle.includes('ডিসি') ||
      normTitle.includes('dobook') ||
      normTitle.includes('doissue') ||
      normTitle.includes('challan') ||
      normTitle.includes('চালান') ||
      normTitle.includes('project') ||
      normTitle.includes('প্রকল্প')
    );
  }
  if (tabKey === 'foodDealers') return normTitle.includes('oms') || normTitle.includes('sub') || normTitle.includes('allocation');
  if (tabKey === 'omsDealers') return normTitle.includes('food') || normTitle.includes('খাদ্য') || normTitle.includes('sub') || normTitle.includes('allocation');
  if (tabKey === 'unoSub') {
    return (
      normTitle.includes('dealer') ||
      normTitle.includes('ucf') ||
      normTitle.includes('ডিলার') ||
      normTitle.includes('ইউসিএফ') ||
      normTitle.includes('dobook') ||
      normTitle.includes('doissue') ||
      normTitle.includes('ministry') ||
      normTitle.includes('dc') ||
      normTitle.includes('মন্ত্রণাল') ||
      normTitle.includes('ডিসি') ||
      normTitle.includes('challan') ||
      normTitle.includes('চালান') ||
      normTitle.includes('project') ||
      normTitle.includes('প্রকল্প')
    );
  }
  if (tabKey === 'doBooks') return normTitle.includes('issue') || normTitle.includes('ইস্যু') || normTitle.includes('ministry') || normTitle.includes('dc') || normTitle.includes('uno') || normTitle.includes('dealer') || normTitle.includes('challan');
  if (tabKey === 'doIssue') return normTitle.includes('book') || normTitle.includes('বই') || normTitle.includes('ministry') || normTitle.includes('dc') || normTitle.includes('uno') || normTitle.includes('dealer') || normTitle.includes('challan');
  if (tabKey === 'challans') return normTitle.includes('ministry') || normTitle.includes('dc') || normTitle.includes('uno') || normTitle.includes('dealer') || normTitle.includes('dobook') || normTitle.includes('doissue') || normTitle.includes('project');
  if (tabKey === 'ministry') {
    return (
      normTitle.includes('dc') ||
      normTitle.includes('uno') ||
      normTitle.includes('dealer') ||
      normTitle.includes('challan') ||
      normTitle.includes('dobook') ||
      normTitle.includes('doissue') ||
      normTitle.includes('project') ||
      normTitle.includes('প্রকল্প') ||
      normTitle.includes('ডিসি') ||
      normTitle.includes('ইউএনও') ||
      normTitle.includes('উপবরাদ্দ')
    );
  }
  if (tabKey === 'dc') {
    return (
      normTitle.includes('ministry') ||
      normTitle.includes('uno') ||
      normTitle.includes('dealer') ||
      normTitle.includes('challan') ||
      normTitle.includes('dobook') ||
      normTitle.includes('doissue') ||
      normTitle.includes('project') ||
      normTitle.includes('প্রকল্প') ||
      normTitle.includes('মন্ত্রণাল') ||
      normTitle.includes('ইউএনও') ||
      normTitle.includes('উপবরাদ্দ')
    );
  }
  if (tabKey === 'chairmen' || tabKey === 'fireServices' || tabKey === 'rates' || tabKey === 'settings') {
    return normTitle.includes('allocation') || normTitle.includes('uno') || normTitle.includes('dealer') || normTitle.includes('dobook') || normTitle.includes('doissue') || normTitle.includes('challan');
  }
  return false;
}

function findMatchingSheetTitle(
  titles: string[],
  primary: string,
  aliases: readonly string[],
  allowPartial = true,
  tabKey?: string
): string | undefined {
  const normTitles = titles.map((t) => ({ raw: t, norm: normKey(t) }));
  const targets = [primary, ...aliases].map(normKey);
  for (const target of targets) {
    const exact = normTitles.find((t) => t.norm === target);
    if (exact) return exact.raw;
  }
  if (!allowPartial) return undefined;
  for (const target of targets) {
    if (target.length < 3) continue;
    const part = normTitles.find(
      (t) =>
        !isIncompatiblePartialTab(tabKey, t.norm) &&
        (t.norm.includes(target) || target.includes(t.norm))
    );
    if (part) return part.raw;
  }
  return undefined;
}

function findAllMatchingSheetTitlesForWrite(
  titles: string[],
  primary: string,
  aliases: readonly string[],
  allowPartial = true,
  tabKey?: string
): string[] {
  const normTitles = titles.map((t) => ({ raw: t, norm: normKey(t) }));
  const targets: string[] = [];
  for (const rawT of [primary, ...aliases]) {
    const nt = normKey(rawT);
    if (nt && !targets.includes(nt)) targets.push(nt);
  }
  const exactMatches: string[] = [];
  for (const target of targets) {
    for (const nt of normTitles) {
      if (nt.norm === target && !exactMatches.includes(nt.raw)) {
        exactMatches.push(nt.raw);
      }
    }
  }
  if (exactMatches.length > 0) return exactMatches;
  if (!allowPartial) return [];
  const partialMatches: string[] = [];
  for (const nt of normTitles) {
    if (isIncompatiblePartialTab(tabKey, nt.norm)) continue;
    const hit = targets.some((target) => target.length >= 3 && (nt.norm.includes(target) || target.includes(nt.norm)));
    if (hit && !partialMatches.includes(nt.raw)) {
      partialMatches.push(nt.raw);
    }
  }
  return partialMatches;
}

function isRedundantPluralDuplicate(primaryTitle: string, candidateTitle: string): boolean {
  const p = normKey(primaryTitle);
  const c = normKey(candidateTitle);
  if (!p || !c || p === c) return false;
  if (p === 'chairman' && c === 'chairmen') return true;
  if (!p.endsWith('s') && c === `${p}s`) return true;
  return false;
}

function resolveSheetWritePlan(
  titles: string[],
  primary: string,
  aliases: readonly string[],
  allowPartial = true,
  tabKey?: string
): { writeTitles: string[]; duplicatePluralTitles: string[] } {
  const matched = findAllMatchingSheetTitlesForWrite(titles, primary, aliases, allowPartial, tabKey);
  if (matched.length <= 1) {
    return { writeTitles: matched, duplicatePluralTitles: [] };
  }
  const canonicalPrimary = matched[0];
  const writeTitles: string[] = [canonicalPrimary];
  const duplicatePluralTitles: string[] = [];
  for (let i = 1; i < matched.length; i++) {
    const cand = matched[i];
    if (isRedundantPluralDuplicate(canonicalPrimary, cand)) {
      duplicatePluralTitles.push(cand);
    } else {
      writeTitles.push(cand);
    }
  }
  return { writeTitles, duplicatePluralTitles };
}

async function googleFetch<T>(
  url: string,
  accessToken: string,
  init: RequestInit = {}
): Promise<T> {
  let res: Response;
  try {
    res = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        ...(init.body ? { 'Content-Type': 'application/json' } : {}),
        ...(init.headers || {}),
      },
    });
  } catch (err) {
    const raw = err instanceof Error ? err.message : '';
    if (!raw || /failed to fetch/i.test(raw)) {
      throw new Error(
        'Google Sheets-এর সাথে নেটওয়ার্ক সংযোগ ব্যর্থ হয়েছে (Failed to fetch) — ইন্টারনেট যাচাই করে পেজটি Reload দিন অথবা সেটিংসের ৪ নম্বর ঘরে Apps Script URL দিয়ে চেষ্টা করুন।'
      );
    }
    throw err;
  }
  if (res.status === 401) throw new GoogleAuthExpiredError();
  const txt = await res.text();
  let json: Record<string, unknown> = {};
  try {
    json = txt ? (JSON.parse(txt) as Record<string, unknown>) : {};
  } catch {
    throw new Error(`Google Sheets HTTP ${res.status}`);
  }
  if (!res.ok) {
    const e = json.error as { message?: string } | undefined;
    throw new Error(e?.message || `Google API ত্রুটি (HTTP ${res.status})`);
  }
  return json as T;
}

async function fetchPublicUrlRaw(url: string): Promise<{ ok: boolean; text: string }> {
  if (
    typeof window !== 'undefined' &&
    typeof (window as unknown as { AndroidBlobBridge?: { httpRequest?: unknown } }).AndroidBlobBridge
      ?.httpRequest === 'function'
  ) {
    const r = await callAndroidNativeHttp('GET', url);
    if (r.status === 401 || r.status === 403 || (r.text || '').includes('accounts.google.com/ServiceLogin')) {
      throw new PublicSheetAccessError();
    }
    return { ok: Boolean(r.ok && r.text && !r.text.trim().startsWith('<')), text: r.text || '' };
  }

  const res = await fetch('/api/public-sheets/fetch', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ url }),
  });
  const data = (await res.json().catch(() => ({}))) as {
    ok?: boolean;
    code?: string;
    text?: string;
  };
  if (data.code === 'PRIVATE_SHEET') throw new PublicSheetAccessError();
  return {
    ok: Boolean(data.ok && data.text && !data.text.trim().startsWith('<')),
    text: String(data.text || ''),
  };
}

async function fetchPublicCsvSheet(
  spreadsheetId: string,
  gid?: string,
  sheetName?: string
): Promise<unknown[][]> {
  if (gid) {
    const exportUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
      spreadsheetId
    )}/export?format=csv&gid=${encodeURIComponent(gid)}`;
    const r = await fetchPublicUrlRaw(exportUrl);
    if (r.ok) return parseCsvToRows(r.text);
  }
  if (sheetName) {
    const gvizUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
      spreadsheetId
    )}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
    const r = await fetchPublicUrlRaw(gvizUrl);
    if (r.ok) return parseCsvToRows(r.text);
  }
  return [];
}

function buildAllocationDbFromRawTables(
  tables: Partial<Record<AllocationRawTabKey, unknown[][]>>,
  matchedTabs: Set<AllocationRawTabKey>,
  currentDb?: AllocationDatabase
): AllocationDatabase {
  const workbookPreferMdy = Object.values(tables).some((t) => (t ? detectPreferMdy(t) : false));

  const mdyOf = (key: AllocationRawTabKey) => {
    const t = tables[key] || [];
    return t.length > 1 ? detectPreferMdy(t) || workbookPreferMdy : workbookPreferMdy;
  };

  const ministryRows = mapRowsByHeader(tables.ministry || []);
  const dcRows = mapRowsByHeader(tables.dc || []);
  const dealerSubRows = mapRowsByHeader(tables.dealerSub || []);
  const grProjectRows = mapRowsByHeader(tables.grProjects || []);
  const kabikhaProjectRows = mapRowsByHeader(tables.kabikhaProjects || []);
  const legacyProjectRows = mapRowsByHeader(tables.legacyProjects || []);
  const unoSubRows = mapRowsByHeader(tables.unoSub || []);
  const doBookRows = mapRowsByHeader(tables.doBooks || []);
  const doIssueRows = mapRowsByHeader(tables.doIssue || []);
  const foodDealerRows = mapRowsByHeader(tables.foodDealers || []);
  const omsDealerRows = mapRowsByHeader(tables.omsDealers || []);
  const chairmanRows = mapRowsByHeader(tables.chairmen || []);
  const fireRows = mapRowsByHeader(tables.fireServices || []);
  const rateRows = mapRowsByHeader(tables.rates || []);

  const settingsMap: Record<string, string> = {};
  for (let i = 1; i < (tables.settings?.length || 0); i++) {
    const r = tables.settings?.[i] || [];
    const k = String(r[0] ?? '').trim();
    if (k) settingsMap[k] = String(r[1] ?? '');
  }

  const seenIds = new Set<string>();

  const minMdy = mdyOf('ministry');
  const ministryAllocations: MinistryAllocation[] = ministryRows
    .map((o, i) => {
      const memoNo = pickField(o, ['memono', 'memo', 'ministrymemo', 'স্মারক', 'স্মারকনং']);
      const comment = pickField(o, ['comment', 'note', 'remarks', 'মন্তব্য']);
      return {
        id: makeUniqueId(pickField(o, ['id']), 'min-s', i, seenIds),
        category: normalizeCategoryName(pickField(o, CATEGORY_FIELD_KEYS), `${memoNo} ${comment}`),
        memoNo,
        memoDate: parseDateStr(pickField(o, ['memodate', 'date', 'তারিখ', 'স্মারকতারিখ', 'createdat']), minMdy),
        allocationQty: pickNum(o, ['allocationqty', 'qty', 'quantity', 'পরিমাণ', 'বরাদ্দ']),
        grain: (pickField(o, ['grain', 'শস্য', 'পণ্য']) || 'চাল') as 'চাল' | 'গম' | 'আটা',
        phase: pickNum(o, ['phase', 'পর্যায়']) || undefined,
        comment,
        createdAt: pickField(o, ['createdat']) || new Date().toISOString(),
      };
    })
    .filter((x) => x.memoNo || x.allocationQty > 0);

  const dcMdy = mdyOf('dc');
  const dcAllocations: DCAllocation[] = dcRows
    .map((o, i) => {
      const ministryMemo = pickField(o, ['ministrymemo', 'sourcememo', 'মন্ত্রণালয়স্মারক']);
      const dcMemo = pickField(o, ['dcmemo', 'ডিসিস্মারক', 'স্মারক']);
      return {
        id: makeUniqueId(pickField(o, ['id']), 'dc-s', i, seenIds),
        category: normalizeCategoryName(pickField(o, CATEGORY_FIELD_KEYS), `${ministryMemo} ${dcMemo}`),
        ministryMemo,
        dcMemo,
        dcMemoDate: parseDateStr(pickField(o, ['dcdate', 'dcmemodate', 'date', 'তারিখ', 'createdat']), dcMdy),
        allocationQty: pickNum(o, ['allocationqty', 'qty', 'quantity', 'পরিমাণ', 'বরাদ্দ']),
        createdAt: pickField(o, ['createdat']) || new Date().toISOString(),
      };
    })
    .filter((x) => x.dcMemo || x.allocationQty > 0);

  const subMdy = mdyOf('dealerSub');
  const dealerSuballocations: DealerSuballocation[] = dealerSubRows
    .map((o, i) => {
      const dealerName = pickField(o, ['dealername', 'recipientname', 'dealer', 'ডিলার', 'ডিলারেরনাম', 'প্রাপক']);
      const sourceMemo = pickField(o, ['sourcememo', 'ministrymemo', 'dcmemo', 'স্মারক', 'উৎসস্মারক']);
      const grainHint = pickField(o, ['grain', 'product', 'item', 'শস্য', 'পণ্য']);
      return {
        id: makeUniqueId(pickField(o, ['id']), 'ds-s', i, seenIds),
        category: normalizeCategoryName(pickField(o, CATEGORY_FIELD_KEYS), `${dealerName} ${sourceMemo} ${grainHint}`),
        sourceMemo,
        unoMemo: pickField(o, ['unomemo', 'ucfmemo', 'sourcememo', 'dcmemo', 'ইউসিএফস্মারক', 'উপবরাদ্দস্মারক']),
        unoDate: parseDateStr(pickField(o, ['unodate', 'date', 'তারিখ', 'createdat']), subMdy),
        dealerName,
        unionName: pickField(o, ['unionname', 'municipality', 'area', 'union', 'ইউনিয়ন']),
        salesCenter: pickField(o, ['salescenter', 'center', 'বিক্রয়কেন্দ্র']),
        cardCount: pickNum(o, ['cardcount', 'cards', 'কার্ড', 'কার্ডসংখ্যা']),
        unusedCards: pickNum(o, ['unusedcards', 'অবিক্রিতকার্ড']),
        allocationQty: pickNum(o, ['allocationqty', 'qty', 'quantity', 'পরিমাণ', 'উপবরাদ্দ']),
        totalAmount: pickNum(o, ['totalamount', 'amount', 'টাকা', 'মোটটাকা']),
        taxAmount: pickNum(o, ['taxamount', 'tax', 'আয়কর']),
        createdAt: pickField(o, ['createdat']) || new Date().toISOString(),
      };
    })
    .filter((x) => x.dealerName || x.allocationQty > 0);

  const parseProjectGroup = (
    rows: Array<Record<string, unknown>>,
    tabKind: 'gr' | 'kabikha' | 'any'
  ): ProjectEntry[] =>
    rows
      .map((o, i) => {
        const projectName = pickField(o, ['projectname', 'প্রকল্পেরনাম', 'প্রকল্পনাম', 'প্রনাম', 'নাম']);
        const rawCat = pickField(o, CATEGORY_FIELD_KEYS);
        const grainCell = pickField(o, ['grain', 'শস্য', 'পণ্য']);
        let resolvedCategory: AllocationCategory;
        if (tabKind === 'gr') {
          resolvedCategory = rawCat
            ? normalizeCategoryName(rawCat, `${projectName} জি আর`)
            : 'জি আর';
        } else if (tabKind === 'kabikha') {
          const catSource =
            rawCat && !rawCat.includes('কাবিখা') && !rawCat.toLowerCase().includes('kabikha') && !rawCat.includes('জি')
              ? `কাবিখা ${rawCat}`
              : rawCat;
          resolvedCategory = catSource
            ? normalizeCategoryName(catSource, `${projectName} ${grainCell} কাবিখা`)
            : grainCell.includes('গম')
              ? 'কাবিখা সাধারণ গম'
              : 'কাবিখা সাধারণ চাল';
        } else {
          resolvedCategory = normalizeCategoryName(rawCat, `${projectName} ${grainCell}`);
        }
        return {
          id: makeUniqueId(pickField(o, ['id']), tabKind === 'gr' ? 'prj-gr' : tabKind === 'kabikha' ? 'prj-kb' : 'prj-s', i, seenIds),
          category: resolvedCategory,
          dcMemo: pickField(o, ['dcmemo', 'ডিসিস্মারক']),
          projectNo: pickField(o, ['projectno', 'প্রকল্পনং', 'প্রনং', 'প্রকল্পনম্বর']),
          projectName,
          presidentName: pickField(o, [
            'presidentname',
            'president',
            'secretaryname',
            'secretary',
            'সভাপতি',
            'সম্পাদক',
            'প্রকল্পসভাপতি',
            'সভাপতিসম্পাদক',
          ]),
          unionName: pickField(o, ['unionname', 'ইউনিয়ন']),
          allocationQty: pickNum(o, ['allocationqty', 'qty', 'quantity', 'বরাদ্দ']),
          createdAt: pickField(o, ['createdat']) || new Date().toISOString(),
        };
      })
      .filter((x) => x.projectNo || x.projectName);

  const hasSplitProjectTabs =
    matchedTabs.has('grProjects') || matchedTabs.has('kabikhaProjects');
  const rawMergedProjects: ProjectEntry[] = [
    ...parseProjectGroup(grProjectRows, 'gr'),
    ...parseProjectGroup(kabikhaProjectRows, 'kabikha'),
    ...(hasSplitProjectTabs ? [] : parseProjectGroup(legacyProjectRows, 'any')),
  ];
  const seenProjectKeys = new Set<string>();
  const projects: ProjectEntry[] = [];
  for (const p of rawMergedProjects) {
    const key = `${normKey(p.category)}::${normKey(p.dcMemo)}::${normKey(p.projectNo)}::${normKey(p.projectName)}`;
    if (seenProjectKeys.has(key)) continue;
    seenProjectKeys.add(key);
    projects.push(p);
  }

  const unoMdy = mdyOf('unoSub');
  const unoSuballocations: UNOSuballocation[] = unoSubRows
    .map((o, i) => {
      const recipientName = pickField(o, [
        'recipientname',
        'chairmanname',
        'presidentname',
        'president',
        'secretaryname',
        'secretary',
        'প্রাপক',
        'সভাপতি',
        'সম্পাদক',
        'প্রকল্পসভাপতি',
        'সভাপতিসম্পাদক',
      ]);
      const projectName = pickField(o, ['projectname', 'প্রকল্পেরনাম', 'প্রকল্পনাম', 'প্রনাম']);
      return {
        id: makeUniqueId(pickField(o, ['id']), 'uno-s', i, seenIds),
        category: normalizeCategoryName(pickField(o, CATEGORY_FIELD_KEYS), `${recipientName} ${projectName}`),
        ministryMemo: pickField(o, [
          'ministrymemo',
          'sourcememo',
          'মন্ত্রণালয়স্মারক',
          'উৎসস্মারক',
          'বরাদ্দস্মারক',
        ]),
        dcMemo: pickField(o, ['dcmemo', 'ডিসিস্মারক']),
        unoMemo: pickField(o, ['unomemo', 'ইউএনওস্মারক', 'স্মারক']),
        unoDate: parseDateStr(pickField(o, ['unodate', 'date', 'তারিখ', 'createdat']), unoMdy),
        unionName: pickField(o, ['unionname', 'ইউনিয়ন']),
        recipientName,
        projectNo: pickField(o, ['projectno', 'প্রকল্পনং', 'প্রনং', 'প্রকল্পনম্বর']),
        projectName,
        allocationQty: pickNum(o, ['allocationqty', 'qty', 'quantity', 'পরিমাণ']),
        createdAt: pickField(o, ['createdat']) || new Date().toISOString(),
      };
    })
    .filter((x) => x.recipientName || x.allocationQty > 0);

  const doBooks: DOBook[] = doBookRows
    .map((o, i) => {
      const startPage = Math.round(pickNum(o, ['startpage', 'start', 'শুরুপাতা']));
      const endPage = Math.round(pickNum(o, ['endpage', 'end', 'শেষপাতা']));
      return {
        id: makeUniqueId(pickField(o, ['id']), 'book-s', i, seenIds),
        bookNo: Math.round(pickNum(o, ['bookno', 'book', 'বইনং'])),
        startPage,
        endPage,
        totalCount: Math.round(pickNum(o, ['totalcount', 'total', 'মোটডিও'])) || Math.max(0, endPage - startPage + 1),
        createdAt: pickField(o, ['createdat']) || new Date().toISOString(),
      };
    })
    .filter((x) => x.bookNo > 0 && x.endPage >= x.startPage);

  const doMdy = mdyOf('doIssue');
  const doIssues: DOIssue[] = doIssueRows
    .map((o, i) => {
      const recipientName = pickField(o, ['recipientname', 'dealername', 'dealer', 'প্রাপক', 'ডিলার', 'প্রতিষ্ঠান']);
      const ministryMemo = pickField(o, ['ministrymemo', 'sourcememo', 'মন্ত্রণালয়স্মারক', 'উৎসস্মারক', 'বরাদ্দস্মারক']);
      const grainHint = pickField(o, ['grain', 'product', 'item', 'শস্য', 'পণ্য']);
      const dealerAllocationId = pickField(o, ['dealerallocationid', 'suballocationid', 'allocationid']) || undefined;
      return {
        id: makeUniqueId(pickField(o, ['id']), 'doi-s', i, seenIds),
        doNo: pickField(o, ['dono', 'donumber', 'do', 'ডিওনং', 'ডিওনম্বর', 'ডিও']),
        issueDate: parseDateStr(pickField(o, ['issuedate', 'date', 'dodate', 'createdat', 'তারিখ', 'ইস্যুতারিখ']), doMdy),
        category: normalizeCategoryName(pickField(o, CATEGORY_FIELD_KEYS), `${recipientName} ${ministryMemo} ${grainHint}`),
        unionName: pickField(o, ['unionname', 'union', 'area', 'municipality', 'ইউনিয়ন', 'এলাকা']),
        recipientName,
        salesCenter: pickField(o, ['salescenter', 'center', 'বিক্রয়কেন্দ্র']),
        allocationQty: pickNum(o, ['quantity', 'allocationqty', 'qty', 'doqty', 'পরিমাণ', 'বরাদ্দ']),
        ministryMemo,
        dcMemo: pickField(o, ['dcmemo', 'ডিসিস্মারক']),
        unoMemo: pickField(o, ['unomemo', 'ucfmemo', 'suballocationmemo', 'memono', 'memo', 'ইউসিএফস্মারক', 'ইউএনওস্মারক', 'উপবরাদ্দস্মারক', 'স্মারক']),
        projectNo: pickField(o, ['projectno', 'প্রকল্পনং']),
        dealerAllocationId,
        challanNo: pickField(o, ['challanno', 'mainchallanno', 'মূলচালান', 'মূলচালাননং', 'চালাননং']),
        challanDate: parseDateStr(pickField(o, ['challandate', 'mainchallandate', 'চালানেরতারিখ']), doMdy),
        challanAmount: pickNum(o, ['challanamount', 'mainchallanamount', 'depositedamount', 'amount', 'totalamount', 'মূলচালানেরটাকা', 'জমাকৃতটাকা', 'মোটটাকা', 'টাকা']) || undefined,
        taxChallanNo: pickField(o, ['taxchallanno', 'incometaxchallanno', 'আয়করচালান', 'আয়করচালাননং']),
        taxChallanDate: parseDateStr(pickField(o, ['taxchallandate', 'incometaxdate', 'আয়করতারিখ']), doMdy),
        taxChallanAmount: pickNum(o, ['taxchallanamount', 'incometaxamount', 'taxamount', 'আয়করটাকা', 'আয়করেরপরিমাণ']) || undefined,
        requiredAmount: pickNum(o, ['requiredamount', 'recordedamount']) || undefined,
        requiredTaxAmount: pickNum(o, ['requiredtaxamount']) || undefined,
        createdAt: pickField(o, ['createdat']) || new Date().toISOString(),
      };
    })
    .filter((x) => x.doNo);

  const challans = tables.challans ? parseChallanRowsFromTable(tables.challans, workbookPreferMdy) : [];

  const dedupByKey = <T>(items: T[], getKey: (item: T) => string): T[] => {
    const seen = new Set<string>();
    const out: T[] = [];
    for (const item of items) {
      const k = getKey(item);
      if (seen.has(k)) continue;
      seen.add(k);
      out.push(item);
    }
    return out;
  };

  const foodMdy = mdyOf('foodDealers');
  const foodDealers: FoodDealer[] = dedupByKey(
    foodDealerRows
      .map((o, i) => ({
        id: makeUniqueId(pickField(o, ['id']), 'fd-s', i, seenIds),
        dealerName: pickField(o, ['dealername', 'dealer', 'ডিলারেরনাম', 'ডিলার', 'নাম', 'name']),
        fatherName: pickField(o, ['fathername', 'পিতারনাম']),
        nid: pickField(o, ['nid', 'এনআইডি']),
        mobile: pickField(o, ['mobile', 'phone', 'মোবাইল']),
        unionName: pickField(o, ['unionname', 'union', 'ইউনিয়নেরনাম', 'ইউনিয়ন']),
        salesCenter: pickField(o, ['salescenter', 'center', 'বিক্রয়কেন্দ্র', 'কেন্দ্র']),
        cardCount: pickNum(o, ['cardcount', 'cards', 'কার্ড', 'কার্ডসংখ্যা']),
        totalQty: pickNum(o, ['totalqty', 'মোটপরিমাণ', 'পরিমাণ']) || Number((pickNum(o, ['cardcount', 'cards', 'কার্ড', 'কার্ডসংখ্যা']) * 0.03).toFixed(3)),
        poNo: pickField(o, ['pono', 'পেঅর্ডারনং']),
        poDate: parseDateStr(pickField(o, ['podate', 'পেঅর্ডারতারিখ']), foodMdy),
        bankName: pickField(o, ['bankname', 'ব্যাংক', 'ব্যাংকেরনাম']),
        amount: pickNum(o, ['amount', 'টাকারপরিমাণ']),
        supervisorName: pickField(o, ['supervisorname', 'তদারকিকর্মকর্তা']),
        supervisorDesignation: pickField(o, ['supervisordesignation', 'পদবী', 'পদবি']),
        supervisorMobile: pickField(o, ['supervisormobile']),
      }))
      .filter((x) => x.dealerName),
    (d) => `${normKey(d.dealerName)}::${normKey(d.unionName)}`
  );

  const omsMdy = mdyOf('omsDealers');
  const omsDealers: OMSDealer[] = dedupByKey(
    omsDealerRows
      .map((o, i) => ({
        id: makeUniqueId(pickField(o, ['id']), 'od-s', i, seenIds),
        dealerName: pickField(o, ['dealername', 'dealer', 'ডিলারেরনাম', 'ডিলার', 'নাম', 'name']),
        fatherName: pickField(o, ['fathername', 'পিতারনাম']),
        nid: pickField(o, ['nid', 'এনআইডি']),
        mobile: pickField(o, ['mobile', 'phone', 'মোবাইল']),
        areaType: (pickField(o, ['areatype', 'এলাকারধরন']) === 'পৌরসভা' ? 'পৌরসভা' : 'ইউনিয়ন') as 'ইউনিয়ন' | 'পৌরসভা',
        unionName: pickField(o, ['unionname', 'union', 'ইউনিয়নেরনাম', 'ইউনিয়ন']),
        municipality: pickField(o, ['municipality', 'পৌরসভা']),
        salesCenter: pickField(o, ['salescenter', 'center', 'বিক্রয়কেন্দ্র', 'কেন্দ্র']),
        poNo: pickField(o, ['pono', 'পেঅর্ডারনং']),
        poDate: parseDateStr(pickField(o, ['podate']), omsMdy),
        bankName: pickField(o, ['bankname', 'ব্যাংক', 'ব্যাংকেরনাম']),
        amount: pickNum(o, ['amount', 'টাকারপরিমাণ']),
        supervisorName: pickField(o, ['supervisorname', 'তদারকিকর্মকর্তা']),
        supervisorDesignation: pickField(o, ['supervisordesignation', 'পদবী', 'পদবি']),
        supervisorMobile: pickField(o, ['supervisormobile']),
      }))
      .filter((x) => x.dealerName),
    (d) => `${normKey(d.dealerName)}::${normKey(d.unionName || d.municipality)}`
  );

  const chairmen: ChairmanEntry[] = dedupByKey(
    chairmanRows
      .map((o, i) => ({
        id: makeUniqueId(pickField(o, ['id']), 'ch-s', i, seenIds),
        unionName: pickField(o, ['unionname', 'union', 'ইউনিয়নেরনাম', 'ইউনিয়ন', 'ইউনিয়ন/পৌরসভা', 'এলাকা']),
        chairmanName: pickField(o, ['chairmanname', 'chairman', 'চেয়ারম্যানেরনাম', 'চেয়ারম্যান', 'চেয়ারম্যান/প্রাপক', 'প্রাপক', 'নাম', 'name']),
        mobile: pickField(o, ['mobile', 'phone', 'মোবাইল']),
        supervisorName: pickField(o, ['supervisorname', 'তদারকিকর্মকর্তা']),
        supervisorMobile: pickField(o, ['supervisormobile']),
      }))
      .filter((x) => x.chairmanName),
    (c) => `${normKey(c.unionName)}::${normKey(c.chairmanName)}`
  );

  const fireServices: FireServiceEntry[] = dedupByKey(
    fireRows
      .map((o, i) => ({
        id: makeUniqueId(pickField(o, ['id']), 'fs-s', i, seenIds),
        officeName: pickField(o, ['officename', 'office', 'অফিস']),
        divisionName: pickField(o, ['divisionname', 'division', 'বিভাগীয়কার্যালয়']),
        riceRate: pickNum(o, ['ricerate']),
        wheatRate: pickNum(o, ['wheatrate']),
      }))
      .filter((x) => x.officeName),
    (f) => normKey(f.officeName)
  );

  const rates: CategoryRate[] = dedupByKey(
    rateRows.map((o, i) => ({
      id: makeUniqueId(pickField(o, ['id']), 'rt-s', i, seenIds),
      category: normalizeCategoryName(pickField(o, CATEGORY_FIELD_KEYS)),
      unitPrice: pickNum(o, ['unitprice', 'price', 'rate', 'এককমূল্য', 'দর']),
      taxRate: pickNum(o, ['taxrate', 'tax', 'আয়কর']),
    })),
    (r) => r.category
  );

  const readChunkedSetting = (key: string): string => {
    const chunkKeys = Object.keys(settingsMap)
      .filter((k) => k.startsWith(`${key}__chunk_`))
      .sort((a, b) => {
        const na = Number(a.slice(key.length + 8)) || 0;
        const nb = Number(b.slice(key.length + 8)) || 0;
        return na - nb;
      });
    if (chunkKeys.length > 0) {
      return chunkKeys.map((ck) => settingsMap[ck] || '').join('');
    }
    const direct = settingsMap[key] || '';
    return direct.startsWith('__CHUNKED__:') ? '' : direct;
  };

  let officer = currentDb?.officer || { ...DEFAULT_OFFICER_PROFILE };
  const rawOfficerSetting = readChunkedSetting('officer');
  if (rawOfficerSetting) {
    try {
      officer = { ...officer, ...(JSON.parse(rawOfficerSetting) as Partial<typeof officer>) };
    } catch {
      // ignore invalid JSON
    }
  }
  let templates = currentDb?.templates || createDefaultTemplatesMap(officer);
  const rawTemplatesSetting = readChunkedSetting('templates');
  if (rawTemplatesSetting) {
    try {
      templates = { ...templates, ...(JSON.parse(rawTemplatesSetting) as typeof templates) };
    } catch {
      // ignore
    }
  }
  let reportMemos = currentDb?.reportMemos || {};
  const rawReportMemosSetting = readChunkedSetting('reportMemos');
  if (rawReportMemosSetting) {
    try {
      reportMemos = {
        ...reportMemos,
        ...(JSON.parse(rawReportMemosSetting) as Record<string, { no: string; date: string }>),
      };
    } catch {
      // ignore
    }
  }

  const isRealAllocationSheet = matchedTabs.size >= 2;
  const resolveTab = <T>(key: keyof typeof ALLOC_SHEET_TABS, parsed: T[], fallback: T[], keepFallbackWhenMissing = false): T[] => {
    if (matchedTabs.has(key)) return parsed;
    if (isRealAllocationSheet && !keepFallbackWhenMissing) return [];
    return parsed.length > 0 ? parsed : fallback;
  };

  const anyProjectTabMatched =
    matchedTabs.has('grProjects') ||
    matchedTabs.has('kabikhaProjects') ||
    matchedTabs.has('legacyProjects');

  const built: AllocationDatabase = {
    ministryAllocations: resolveTab('ministry', ministryAllocations, currentDb?.ministryAllocations || []),
    dcAllocations: resolveTab('dc', dcAllocations, currentDb?.dcAllocations || []),
    dealerSuballocations: resolveTab('dealerSub', dealerSuballocations, currentDb?.dealerSuballocations || []),
    projects: anyProjectTabMatched
      ? projects
      : isRealAllocationSheet
        ? []
        : projects.length > 0
          ? projects
          : currentDb?.projects || [],
    unoSuballocations: resolveTab('unoSub', unoSuballocations, currentDb?.unoSuballocations || []),
    doBooks: resolveTab('doBooks', doBooks, currentDb?.doBooks || []),
    doIssues: resolveTab('doIssue', doIssues, currentDb?.doIssues || []),
    challans: resolveTab('challans', challans, currentDb?.challans || [], true),
    foodDealers: resolveTab('foodDealers', foodDealers, currentDb?.foodDealers || [], true),
    omsDealers: resolveTab('omsDealers', omsDealers, currentDb?.omsDealers || [], true),
    chairmen: resolveTab('chairmen', chairmen, currentDb?.chairmen || [], true),
    fireServices: resolveTab('fireServices', fireServices, currentDb?.fireServices || [], true),
    rates: resolveTab('rates', rates, currentDb?.rates || [], true),
    officer,
    templates,
    reportMemos,
  };

  return reconcileChallansAndDOs(built);
}

/**
 * Links external Challan rows (including dealerAllocationId) with DealerSuballocations
 * and enriches DOIssues that lack embedded challan metadata.
 */
export function reconcileChallansAndDOs(db: AllocationDatabase): AllocationDatabase {
  const rateMap = new Map(db.rates.map((r) => [r.category, r]));

  const getEffectiveUnitPrice = (cat: AllocationCategory): number => {
    const r = rateMap.get(cat);
    if (r && r.unitPrice > 0) return r.unitPrice;
    return cat === 'খাদ্য বান্ধব' ? 13000 : 0;
  };

  const getEffectiveTaxRate = (cat: AllocationCategory): number => {
    const r = rateMap.get(cat);
    if (r && r.taxRate > 0) return r.taxRate;
    return cat === 'খাদ্য বান্ধব' ? 200 : 0;
  };

  const monthKey = (iso?: string): string => {
    const s = toIsoDateStr(iso);
    return s ? s.slice(0, 7) : '';
  };

  const enrichedSubs: DealerSuballocation[] = db.dealerSuballocations.map((s) => {
    const unitPrice = getEffectiveUnitPrice(s.category);
    const taxRate = getEffectiveTaxRate(s.category);
    const totalAmount = s.totalAmount > 0 ? s.totalAmount : round2(s.allocationQty * unitPrice);
    const taxAmount = s.taxAmount > 0 ? s.taxAmount : round2(s.allocationQty * taxRate);
    if (totalAmount === s.totalAmount && taxAmount === s.taxAmount) return s;
    return { ...s, totalAmount, taxAmount };
  });

  const subById = new Map<string, DealerSuballocation>();
  for (const s of enrichedSubs) {
    if (s.id) subById.set(s.id.trim(), s);
  }

  // Strictly deduplicate challans: once a challan number (per type) exists, it cannot be used twice
  const dedupMap = new Map<string, ChallanRecord>();
  for (const c of db.challans) {
    const normChallan = normMemoNo(c.challanNo);
    const dedupKey = normChallan ? `${normChallan}::${c.challanType}` : `id::${c.id}`;
    const existing = dedupMap.get(dedupKey);
    if (!existing) {
      dedupMap.set(dedupKey, c);
    } else {
      const existingIso = toIsoDateStr(existing.challanDate) || '9999-99-99';
      const newIso = toIsoDateStr(c.challanDate) || '9999-99-99';
      const pickNew =
        newIso < existingIso ||
        (newIso === existingIso &&
          ((c.recordedAmount > 0 && existing.recordedAmount <= 0) ||
            (Boolean(c.allocationQty) && !existing.allocationQty) ||
            (Boolean(c.dealerAllocationId) && !existing.dealerAllocationId) ||
            c.depositedAmount > existing.depositedAmount));
      if (pickNew) dedupMap.set(dedupKey, c);
    }
  }
  const uniqueChallans = Array.from(dedupMap.values());

  // Resolve consistent same-category & same-dealer suballocation metadata (never cross-category / cross-dealer)
  const preNormalizedChallans: ChallanRecord[] = uniqueChallans.map((c) => {
    const rawSub = c.dealerAllocationId ? subById.get(c.dealerAllocationId.trim()) : undefined;
    const validSub =
      rawSub &&
      (!c.category || rawSub.category === c.category) &&
      (!c.dealerName || c.dealerName === '—' || isSamePartyName(c.dealerName, rawSub.dealerName))
        ? rawSub
        : undefined;
    const resolvedCat = c.category || validSub?.category || c.category;
    const resolvedDealer =
      c.dealerName && c.dealerName !== '—'
        ? c.dealerName
        : validSub?.dealerName || c.dealerName || '';
    const resolvedSubId = validSub ? c.dealerAllocationId : undefined;
    if (
      resolvedCat !== c.category ||
      resolvedDealer !== c.dealerName ||
      resolvedSubId !== c.dealerAllocationId
    ) {
      return {
        ...c,
        category: resolvedCat,
        dealerName: resolvedDealer,
        dealerAllocationId: resolvedSubId,
      };
    }
    return c;
  });

  // Sort DOs chronologically (earliest issueDate, then numeric doNo, then createdAt)
  const sortedDos = [...db.doIssues].sort((a, b) => {
    const da = toIsoDateStr(a.issueDate) || '9999-99-99';
    const dbDate = toIsoDateStr(b.issueDate) || '9999-99-99';
    if (da !== dbDate) return da.localeCompare(dbDate);
    const noCmp = bnToEnDigits(a.doNo || '').localeCompare(bnToEnDigits(b.doNo || ''), undefined, {
      numeric: true,
    });
    if (noCmp !== 0) return noCmp;
    return (a.createdAt || '').localeCompare(b.createdAt || '');
  });

  const usedChallanIds = new Set<string>();
  const usedChallanNoType = new Set<string>();
  const matchedDoByChallanId = new Map<string, DOIssue>();
  const matchedMainByDoId = new Map<string, ChallanRecord>();
  const matchedTaxByDoId = new Map<string, ChallanRecord>();

  // Strict 1-to-1 chronological matching:
  // 1) Same category (never cross-category)
  // 2) Same dealer / institution / person (never cross-dealer)
  // 3) Challan date <= DO issue date
  // 4) Earliest challan -> earliest DO; once used, never reused
  for (const type of ['মূল চালান', 'আয়কর চালান'] as const) {
    for (const d of sortedDos) {
      const dIso = toIsoDateStr(d.issueDate);
      const targetDoChallanNo = type === 'মূল চালান' ? d.challanNo : d.taxChallanNo;

      const candidates = preNormalizedChallans.filter((c) => {
        if (c.challanType !== type) return false;
        if (usedChallanIds.has(c.id)) return false;
        const chKey = normMemoNo(c.challanNo);
        if (chKey && usedChallanNoType.has(`${chKey}::${type}`)) return false;

        // Rule: Same category strictly
        if (c.category !== d.category) return false;

        // Rule: Same dealer / institution / person strictly
        if (!c.dealerName || c.dealerName === '—' || !d.recipientName) return false;
        if (!isSamePartyName(c.dealerName, d.recipientName)) return false;

        // Rule: Challan date must be on or before DO issue date
        const cIso = toIsoDateStr(c.challanDate);
        if (cIso && dIso) {
          if (cIso > dIso) return false;
        } else if (!cIso && dIso) {
          // Undated challan only eligible if explicitly linked to this DO/sub
          const explicit =
            (c.doNo && d.doNo && normMemoNo(c.doNo) === normMemoNo(d.doNo)) ||
            (d.dealerAllocationId &&
              c.dealerAllocationId &&
              normMemoNo(d.dealerAllocationId) === normMemoNo(c.dealerAllocationId)) ||
            (targetDoChallanNo && chKey && chKey === normMemoNo(targetDoChallanNo));
          if (!explicit) return false;
        }

        return true;
      });

      if (candidates.length === 0) continue;

      candidates.sort((a, b) => {
        const aIso = toIsoDateStr(a.challanDate);
        const bIso = toIsoDateStr(b.challanDate);
        if (Boolean(aIso) !== Boolean(bIso)) return aIso ? -1 : 1;
        // Earliest challan ALWAYS wins for earliest DO
        if (aIso && bIso && aIso !== bIso) return aIso.localeCompare(bIso);

        // Same-day tie-breakers: explicit link, then quantity proximity, then challanNo numeric
        const aExplicit =
          (a.doNo && d.doNo && normMemoNo(a.doNo) === normMemoNo(d.doNo)) ||
          (targetDoChallanNo && a.challanNo && normMemoNo(a.challanNo) === normMemoNo(targetDoChallanNo)) ||
          (d.dealerAllocationId && a.dealerAllocationId && normMemoNo(a.dealerAllocationId) === normMemoNo(d.dealerAllocationId))
            ? 0
            : 1;
        const bExplicit =
          (b.doNo && d.doNo && normMemoNo(b.doNo) === normMemoNo(d.doNo)) ||
          (targetDoChallanNo && b.challanNo && normMemoNo(b.challanNo) === normMemoNo(targetDoChallanNo)) ||
          (d.dealerAllocationId && b.dealerAllocationId && normMemoNo(b.dealerAllocationId) === normMemoNo(d.dealerAllocationId))
            ? 0
            : 1;
        if (aExplicit !== bExplicit) return aExplicit - bExplicit;

        if (d.allocationQty > 0) {
          const qa = a.allocationQty ? Math.abs(a.allocationQty - d.allocationQty) : 999;
          const qb = b.allocationQty ? Math.abs(b.allocationQty - d.allocationQty) : 999;
          if (Math.abs(qa - qb) > 0.01) return qa - qb;
        }

        const noCmp = bnToEnDigits(a.challanNo || '').localeCompare(
          bnToEnDigits(b.challanNo || ''),
          undefined,
          { numeric: true }
        );
        if (noCmp !== 0) return noCmp;
        return (a.createdAt || '').localeCompare(b.createdAt || '');
      });

      const best = candidates[0];
      usedChallanIds.add(best.id);
      const bestKey = normMemoNo(best.challanNo);
      if (bestKey) usedChallanNoType.add(`${bestKey}::${type}`);
      matchedDoByChallanId.set(best.id, d);
      if (type === 'মূল চালান') matchedMainByDoId.set(d.id, best);
      else matchedTaxByDoId.set(d.id, best);
    }
  }

  const doNoExistsSet = new Set(db.doIssues.map((d) => normMemoNo(d.doNo)).filter(Boolean));
  const allChallanNoSetByType = {
    main: new Set(
      preNormalizedChallans
        .filter((c) => c.challanType === 'মূল চালান' && c.challanNo)
        .map((c) => normMemoNo(c.challanNo))
    ),
    tax: new Set(
      preNormalizedChallans
        .filter((c) => c.challanType === 'আয়কর চালান' && c.challanNo)
        .map((c) => normMemoNo(c.challanNo))
    ),
  };

  const enrichedChallans: ChallanRecord[] = preNormalizedChallans.map((c) => {
    const linkedSub = c.dealerAllocationId ? subById.get(c.dealerAllocationId.trim()) : undefined;
    const validLinkedSub =
      linkedSub &&
      linkedSub.category === c.category &&
      (!c.dealerName || isSamePartyName(c.dealerName, linkedSub.dealerName))
        ? linkedSub
        : undefined;
    const matchingDo = matchedDoByChallanId.get(c.id);

    let resolvedDoNo: string | undefined;
    let resolvedDoDate: string | undefined;
    if (matchingDo) {
      resolvedDoNo = matchingDo.doNo;
      resolvedDoDate = matchingDo.issueDate;
    } else if (c.doNo) {
      // If DO exists in db.doIssues and didn't match c (wrong dealer/category/date or already used), clear stale DO link
      if (doNoExistsSet.has(normMemoNo(c.doNo))) {
        resolvedDoNo = undefined;
        resolvedDoDate = undefined;
      } else if (c.doDate && c.challanDate && !isChallanOnOrBeforeDo(c.challanDate, c.doDate, false)) {
        resolvedDoNo = undefined;
        resolvedDoDate = undefined;
      } else {
        resolvedDoNo = c.doNo;
        resolvedDoDate = c.doDate;
      }
    }

    if (c.recordedAmount > 0) {
      if (resolvedDoNo !== c.doNo || resolvedDoDate !== c.doDate) {
        return { ...c, doNo: resolvedDoNo, doDate: resolvedDoDate };
      }
      return c;
    }

    let rec = 0;
    if (validLinkedSub) {
      rec =
        c.challanType === 'আয়কর চালান'
          ? validLinkedSub.taxAmount ||
            round2((validLinkedSub.allocationQty || c.allocationQty || 0) * getEffectiveTaxRate(c.category))
          : validLinkedSub.totalAmount ||
            round2((validLinkedSub.allocationQty || c.allocationQty || 0) * getEffectiveUnitPrice(c.category));
    } else if (matchingDo) {
      const matchedSub = enrichedSubs.find(
        (s) =>
          s.category === matchingDo.category &&
          normMemoNo(s.unoMemo) === normMemoNo(matchingDo.unoMemo) &&
          isSamePartyName(s.dealerName, matchingDo.recipientName)
      );
      rec =
        c.challanType === 'আয়কর চালান'
          ? matchingDo.requiredTaxAmount ||
            matchedSub?.taxAmount ||
            round2(
              (matchedSub?.allocationQty || c.allocationQty || matchingDo.allocationQty || 0) *
                getEffectiveTaxRate(c.category)
            )
          : matchingDo.requiredAmount ||
            matchedSub?.totalAmount ||
            round2(
              (matchedSub?.allocationQty || c.allocationQty || matchingDo.allocationQty || 0) *
                getEffectiveUnitPrice(c.category)
            );
    } else {
      // Same category + same dealer suballocations ONLY (never cross-category / cross-dealer)
      const refDate = c.challanDate || '';
      const refMonth = monthKey(refDate);
      const sameCatDealerSubs = c.dealerName
        ? enrichedSubs.filter(
            (s) => s.category === c.category && isSamePartyName(s.dealerName, c.dealerName)
          )
        : [];
      const monthSub =
        (refMonth ? sameCatDealerSubs.find((s) => monthKey(s.unoDate) === refMonth) : undefined) ||
        (c.allocationQty && c.allocationQty > 0
          ? sameCatDealerSubs.find((s) => Math.abs((s.allocationQty || 0) - c.allocationQty!) < 0.02)
          : undefined) ||
        (c.depositedAmount > 0 && sameCatDealerSubs.length > 1
          ? sameCatDealerSubs.find(
              (s) =>
                (c.challanType === 'মূল চালান' &&
                  s.totalAmount > 0 &&
                  Math.abs(s.totalAmount - c.depositedAmount) <= Math.max(500, s.totalAmount * 0.05)) ||
                (c.challanType === 'আয়কর চালান' &&
                  s.taxAmount > 0 &&
                  Math.abs(s.taxAmount - c.depositedAmount) <= Math.max(50, s.taxAmount * 0.1))
            )
          : undefined);
      const chosenSub = monthSub || (sameCatDealerSubs.length === 1 && !(c.allocationQty && c.allocationQty > 0) ? sameCatDealerSubs[0] : undefined);
      if (chosenSub) {
        rec =
          c.challanType === 'আয়কর চালান'
            ? chosenSub.taxAmount ||
              round2((chosenSub.allocationQty || c.allocationQty || 0) * getEffectiveTaxRate(c.category))
            : chosenSub.totalAmount ||
              round2((chosenSub.allocationQty || c.allocationQty || 0) * getEffectiveUnitPrice(c.category));
      } else if (c.allocationQty && c.allocationQty > 0) {
        rec =
          c.challanType === 'আয়কর চালান'
            ? round2(c.allocationQty * getEffectiveTaxRate(c.category))
            : round2(c.allocationQty * getEffectiveUnitPrice(c.category));
      }
    }

    if (rec > 0 || resolvedDoNo !== c.doNo || resolvedDoDate !== c.doDate) {
      return {
        ...c,
        doNo: resolvedDoNo,
        doDate: resolvedDoDate,
        recordedAmount: rec > 0 ? round2(rec) : c.recordedAmount,
      };
    }
    return c;
  });

  // Track inline-only challans so even legacy DO rows without ChallanRecord cannot reuse a challan
  const usedInlineMainNos = new Set<string>(usedChallanNoType);
  const enrichedDoMap = new Map<string, DOIssue>();

  for (const d of sortedDos) {
    const mainMatch = matchedMainByDoId.get(d.id);
    const taxMatch = matchedTaxByDoId.get(d.id);

    let nextMainNo = mainMatch?.challanNo;
    let nextMainDate = mainMatch?.challanDate;
    let nextMainAmt = mainMatch?.depositedAmount;
    if (!mainMatch && d.challanNo) {
      const mKey = normMemoNo(d.challanNo);
      const existsInChallans = mKey && allChallanNoSetByType.main.has(mKey);
      const dateOk = isChallanOnOrBeforeDo(d.challanDate, d.issueDate, false);
      if (!existsInChallans && dateOk && (!mKey || !usedInlineMainNos.has(`${mKey}::মূল চালান`))) {
        nextMainNo = d.challanNo;
        nextMainDate = d.challanDate;
        nextMainAmt = d.challanAmount;
        if (mKey) usedInlineMainNos.add(`${mKey}::মূল চালান`);
      }
    }

    let nextTaxNo = taxMatch?.challanNo;
    let nextTaxDate = taxMatch?.challanDate;
    let nextTaxAmt = taxMatch?.depositedAmount;
    if (!taxMatch && d.taxChallanNo) {
      const tKey = normMemoNo(d.taxChallanNo);
      const existsInChallans = tKey && allChallanNoSetByType.tax.has(tKey);
      const dateOk = isChallanOnOrBeforeDo(d.taxChallanDate, d.issueDate, false);
      if (!existsInChallans && dateOk && (!tKey || !usedInlineMainNos.has(`${tKey}::আয়কর চালান`))) {
        nextTaxNo = d.taxChallanNo;
        nextTaxDate = d.taxChallanDate;
        nextTaxAmt = d.taxChallanAmount;
        if (tKey) usedInlineMainNos.add(`${tKey}::আয়কর চালান`);
      }
    }

    const resolvedReqMain =
      (d.requiredAmount && d.requiredAmount > 0 ? d.requiredAmount : undefined) ??
      (mainMatch?.recordedAmount && mainMatch.recordedAmount > 0 ? mainMatch.recordedAmount : undefined) ??
      (d.allocationQty > 0 ? round2(d.allocationQty * getEffectiveUnitPrice(d.category)) || undefined : undefined);

    const resolvedReqTax =
      (d.requiredTaxAmount && d.requiredTaxAmount > 0 ? d.requiredTaxAmount : undefined) ??
      (taxMatch?.recordedAmount && taxMatch.recordedAmount > 0 ? taxMatch.recordedAmount : undefined) ??
      (d.allocationQty > 0 ? round2(d.allocationQty * getEffectiveTaxRate(d.category)) || undefined : undefined);

    enrichedDoMap.set(d.id, {
      ...d,
      challanNo: nextMainNo,
      challanDate: nextMainDate,
      challanAmount: nextMainNo || (mainMatch && mainMatch.depositedAmount > 0) ? nextMainAmt : undefined,
      taxChallanNo: nextTaxNo,
      taxChallanDate: nextTaxDate,
      taxChallanAmount: nextTaxNo || (taxMatch && taxMatch.depositedAmount > 0) ? nextTaxAmt : undefined,
      requiredAmount: resolvedReqMain,
      requiredTaxAmount: resolvedReqTax,
    });
  }

  const enrichedDos: DOIssue[] = db.doIssues.map((d) => enrichedDoMap.get(d.id) || d);

  const vwbMinistryMemos = db.ministryAllocations.filter((m) => m.category === 'ভিডাব্লিউভি');
  const enrichedUnoSubs: UNOSuballocation[] = db.unoSuballocations.map((u) => {
    if (u.category !== 'ভিডাব্লিউভি' || u.ministryMemo.trim()) return u;
    const fromDc = u.dcMemo.trim()
      ? vwbMinistryMemos.find((m) => memosLikelyMatch(m.memoNo, u.dcMemo))
      : undefined;
    if (fromDc) {
      return { ...u, ministryMemo: fromDc.memoNo };
    }
    if (vwbMinistryMemos.length === 1) {
      return { ...u, ministryMemo: vwbMinistryMemos[0].memoNo };
    }
    return u;
  });

  return {
    ...db,
    unoSuballocations: enrichedUnoSubs,
    dealerSuballocations: enrichedSubs,
    challans: enrichedChallans,
    doIssues: enrichedDos,
  };
}

export async function fetchAllocationFromPublicSheet(
  rawSpreadsheetId: string,
  currentDb?: AllocationDatabase
): Promise<{ db: AllocationDatabase; title: string; tabsFound: string[]; matchedTabNames: string[] }> {
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  if (!spreadsheetId) throw new Error('Google Sheet ID বা লিংক দিন।');

  const { meta, gidMap } = await getPublicSpreadsheetMetadata(spreadsheetId);
  const rawTables: Partial<Record<AllocationRawTabKey, unknown[][]>> = {};
  const matchedTabs = new Set<AllocationRawTabKey>();
  const matchedTabNames: string[] = [];

  const entries: Array<[AllocationRawTabKey, { name: string; aliases: readonly string[] }]> = [
    ...(Object.entries(ALLOC_SHEET_TABS) as Array<
      [keyof typeof ALLOC_SHEET_TABS, (typeof ALLOC_SHEET_TABS)[keyof typeof ALLOC_SHEET_TABS]]
    >),
    ['legacyProjects', LEGACY_COMBINED_PROJECTS_TAB],
  ];

  await Promise.all(
    entries.map(async ([key, def]) => {
      const candidates = findAllMatchingSheetTitlesForWrite(
        meta.sheetTitles,
        def.name,
        def.aliases,
        !isExactOnlyTabKey(key),
        key
      );
      if (candidates.length === 0 && meta.sheetTitles.length === 0 && key !== 'legacyProjects') {
        const rows = await fetchPublicCsvSheet(spreadsheetId, undefined, def.name);
        if (rows.length > 0) {
          matchedTabs.add(key);
          matchedTabNames.push(def.name);
          rawTables[key] = rows;
        }
        return;
      }
      if (candidates.length === 0) return;

      const primaryName = candidates[0];
      let chosenName = primaryName;
      let chosenRows = await fetchPublicCsvSheet(spreadsheetId, gidMap[primaryName], primaryName);
      if (chosenRows.length <= 1 && candidates.length > 1) {
        for (const altName of candidates.slice(1)) {
          const altRows = await fetchPublicCsvSheet(spreadsheetId, gidMap[altName], altName);
          if (altRows.length > 1) {
            chosenName = altName;
            chosenRows = altRows;
            break;
          }
        }
      }
      matchedTabs.add(key);
      matchedTabNames.push(chosenName);
      rawTables[key] = chosenRows;
    })
  );

  if (matchedTabs.size === 0 && meta.sheetTitles.length > 0) {
    throw new Error(
      `"${meta.title}" শিটে বরাদ্দ/ডিও-এর কাঙ্ক্ষিত ট্যাব পাওয়া যায়নি (পাওয়া ট্যাব: ${meta.sheetTitles.slice(0, 6).join(', ')}).`
    );
  }

  return {
    db: buildAllocationDbFromRawTables(rawTables, matchedTabs, currentDb),
    title: meta.title,
    tabsFound: meta.sheetTitles,
    matchedTabNames,
  };
}

export function countAllocationDatabaseRecords(db: AllocationDatabase): number {
  return (
    db.ministryAllocations.length +
    db.dcAllocations.length +
    db.dealerSuballocations.length +
    db.projects.length +
    db.unoSuballocations.length +
    db.doBooks.length +
    db.doIssues.length +
    db.challans.length +
    db.foodDealers.length +
    db.omsDealers.length +
    db.chairmen.length +
    db.fireServices.length +
    db.rates.length
  );
}

export async function fetchAllocationFromAppsScript(
  rawAppsScriptUrl: string,
  rawSpreadsheetId: string,
  currentDb?: AllocationDatabase
): Promise<{
  db: AllocationDatabase;
  title: string;
  spreadsheetId: string;
  tabsFound: string[];
  matchedTabNames: string[];
}> {
  const cleanUrl = rawAppsScriptUrl.trim();
  if (!cleanUrl || (!cleanUrl.startsWith('https://') && !cleanUrl.startsWith('http://'))) {
    throw new Error('সঠিক Apps Script Web App URL (https://script.google.com/...) দিন।');
  }
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  const res = await callGasAction<{
    title?: string;
    spreadsheetId?: string;
    sheetTitles?: string[];
    tables?: Record<string, unknown[][]>;
  }>(cleanUrl, 'readAllTabs', { spreadsheetId });

  if (!res.ok || !res.data?.tables) {
    throw new Error(
      res.error ||
        'Apps Script থেকে শিট পড়া যায়নি। Code.gs-এ নতুন আপডেট করা কোড বসিয়ে Version: New version দিয়ে Deploy করুন।'
    );
  }

  const tablesMap = res.data.tables || {};
  const sheetTitles = Array.isArray(res.data.sheetTitles)
    ? res.data.sheetTitles
    : Object.keys(tablesMap);
  const sheetTitle = String(res.data.title || 'Google Sheet');
  const resolvedSpreadsheetId = String(res.data.spreadsheetId || spreadsheetId || '');

  const entries: Array<[AllocationRawTabKey, { name: string; aliases: readonly string[] }]> = [
    ...(Object.entries(ALLOC_SHEET_TABS) as Array<
      [keyof typeof ALLOC_SHEET_TABS, (typeof ALLOC_SHEET_TABS)[keyof typeof ALLOC_SHEET_TABS]]
    >),
    ['legacyProjects', LEGACY_COMBINED_PROJECTS_TAB],
  ];

  const rawTables: Partial<Record<AllocationRawTabKey, unknown[][]>> = {};
  const matchedTabs = new Set<AllocationRawTabKey>();
  const matchedTabNames: string[] = [];

  for (const [key, def] of entries) {
    const candidates = findAllMatchingSheetTitlesForWrite(
      sheetTitles,
      def.name,
      def.aliases,
      !isExactOnlyTabKey(key),
      key
    );
    if (candidates.length === 0) continue;

    let chosenName = candidates[0];
    let chosenRows = Array.isArray(tablesMap[chosenName]) ? tablesMap[chosenName] : [];
    if (chosenRows.length <= 1 && candidates.length > 1) {
      for (const altName of candidates.slice(1)) {
        const altRows = Array.isArray(tablesMap[altName]) ? tablesMap[altName] : [];
        if (altRows.length > 1) {
          chosenName = altName;
          chosenRows = altRows;
          break;
        }
      }
    }
    matchedTabs.add(key);
    matchedTabNames.push(chosenName);
    rawTables[key] = chosenRows;
  }

  if (matchedTabs.size === 0 && sheetTitles.length > 0) {
    throw new Error(
      `"${sheetTitle}" শিটে বরাদ্দ/ডিও-এর কাঙ্ক্ষিত ট্যাব পাওয়া যায়নি (পাওয়া ট্যাব: ${sheetTitles.slice(0, 6).join(', ')}).`
    );
  }

  return {
    db: buildAllocationDbFromRawTables(rawTables, matchedTabs, currentDb),
    title: sheetTitle,
    spreadsheetId: resolvedSpreadsheetId,
    tabsFound: sheetTitles,
    matchedTabNames,
  };
}

function parseChallanRowsFromTable(rows: unknown[][], workbookPreferMdy = false): ChallanRecord[] {
  const challanRows = mapRowsByHeader(rows);
  const preferMdy = detectPreferMdy(rows) || workbookPreferMdy;
  const out: ChallanRecord[] = [];

  challanRows.forEach((r, i) => {
    const dealerName = pickField(r, [
      'Dealer_Name', 'dealerName', 'dealer', 'Dealer',
      'Recipient_Name', 'recipientName', 'recipient', 'partyName', 'party', 'name',
      'ডিলার', 'ডিলারের নাম', 'প্রাপক', 'প্রাপকের নাম', 'প্রতিষ্ঠান', 'নাম',
    ]);
    const dealerAllocationId = pickField(r, [
      'dealerAllocationId', 'dealer_allocation_id', 'suballocationId', 'allocationId',
    ]);
    const note = pickField(r, ['Note', 'note', 'remarks', 'comment', 'মন্তব্য']);
    const baseCategory = normalizeCategoryName(
      pickField(r, CATEGORY_FIELD_KEYS),
      `${dealerName} ${note}`
    );
    const bankName = pickField(r, ['Bank_Name', 'bankName', 'bank', 'Bank', 'branch', 'ব্যাংক', 'ব্যাংকের নাম', 'ব্যাংক ও শাখা']);
    const doNo = pickField(r, ['DO_No', 'doNo', 'doNumber', 'ডিও নং', 'ডিও নম্বর', 'ডিও']);
    const doDate = parseDateStr(
      pickField(r, ['DO_Date', 'doDate', 'ডিও তারিখ']),
      preferMdy
    ) || undefined;
    const createdAt = pickField(r, ['Created_At', 'createdAt', 'timestamp']) || new Date().toISOString();

    const mainNo = pickField(r, [
      'Challan_No', 'challanNo', 'Challan Number', 'ChallanNumber', 'Challan',
      'Main_Challan_No', 'mainChallanNo', 'trxId', 'voucherNo',
      'চালান নং', 'চালান নম্বর', 'চালান', 'মূল চালান নং', 'মূল চালান', 'ট্রেজারি চালান নং',
    ]);
    const taxNoExplicit = pickField(r, [
      'Tax_Challan_No', 'taxChallanNo', 'IncomeTaxChallanNo', 'incomeTaxChallanNo',
      'আয়কর চালান নং', 'আয়কর চালান', 'ভ্যাট/আয়কর চালান', 'ট্যাক্স চালান',
    ]);

    const mainDate = parseDateStr(
      pickField(r, ['Challan_Date', 'challanDate', 'mainChallanDate', 'date', 'Date', 'entryDate', 'timestamp', 'তারিখ', 'চালানের তারিখ', 'জমার তারিখ']),
      preferMdy
    );
    const taxDate =
      parseDateStr(
        pickField(r, [
          'Tax_Challan_Date', 'taxChallanDate', 'IncomeTaxDate', 'incomeTaxDate', 'income_tax_date',
          'আয়কর তারিখ', 'আয়কর চালানের তারিখ',
        ]),
        preferMdy
      ) || mainDate;

    const rawType = pickField(r, ['Challan_Type', 'challanType', 'type', 'Type', 'kind', 'ধরন', 'চালানের ধরন']);
    const rate = pickNum(r, ['rate', 'unitPrice', 'দর']);
    const qty = pickNum(r, ['quantity', 'qty', 'allocationQty', 'পরিমাণ']);
    const recMainRaw = pickNum(r, [
      'Recorded_Amount', 'recordedAmount', 'requiredAmount', 'payableAmount', 'targetAmount',
      'রেকর্ড অনুযায়ী টাকা', 'রেকর্ড টাকা', 'রেকর্ড', 'ধার্যকৃত টাকা', 'প্রদেয় টাকা',
    ]);
    const recMain = recMainRaw > 0 ? recMainRaw : rate > 0 && qty > 0 ? round2(rate * qty) : 0;

    const depMain = pickNum(r, [
      'Deposited_Amount', 'depositedAmount', 'paidAmount', 'depositAmount', 'challanAmount',
      'amount', 'Amount', 'totalAmount', 'মূল চালান টাকা', 'জমাকৃত টাকা', 'জমার পরিমাণ', 'জমা', 'মোট টাকা', 'টাকা',
    ]);

    const recTax = pickNum(r, ['Required_Tax_Amount', 'requiredTaxAmount', 'recordedTaxAmount', 'রেকর্ড আয়কর']);
    const depTax = pickNum(r, [
      'Tax_Challan_Amount', 'taxChallanAmount', 'IncomeTaxAmount', 'incomeTaxAmount', 'income_tax_amount',
      'taxAmount', 'depositedTaxAmount', 'আয়কর টাকা', 'আয়করের পরিমাণ',
    ]);

    const parsedQty = qty > 0 ? round3(qty) : undefined;

    // Case A: Row has BOTH explicit main + tax challan columns -> split cleanly into 2 records
    if (taxNoExplicit && (mainNo || depMain > 0)) {
      if (mainNo || dealerName || dealerAllocationId) {
        out.push({
          id: pickField(r, ['ID', 'id']) || `ch-ext-m-${i}-${uid('c')}`,
          challanNo: mainNo,
          challanDate: mainDate,
          category: baseCategory,
          dealerName,
          dealerAllocationId: dealerAllocationId || undefined,
          challanType: 'মূল চালান',
          bankName,
          doNo,
          doDate,
          allocationQty: parsedQty,
          recordedAmount: recMain,
          depositedAmount: depMain,
          note,
          createdAt,
        });
      }
      out.push({
        id: `ch-ext-t-${i}-${uid('t')}`,
        challanNo: taxNoExplicit,
        challanDate: taxDate,
        category: baseCategory,
        dealerName,
        dealerAllocationId: dealerAllocationId || undefined,
        challanType: 'আয়কর চালান',
        bankName,
        doNo,
        doDate,
        allocationQty: parsedQty,
        recordedAmount: recTax,
        depositedAmount: depTax,
        note,
        createdAt,
      });
      return;
    }

    // Case B: Single challan per row
    const finalNo = mainNo || taxNoExplicit;
    if (!finalNo && !dealerName && !dealerAllocationId) return;
    const isTax =
      rawType.includes('আয়কর') ||
      rawType.toLowerCase().includes('tax') ||
      rawType.toLowerCase().includes('vat') ||
      (!mainNo && Boolean(taxNoExplicit));

    out.push({
      id: pickField(r, ['ID', 'id']) || `ch-ext-${i}-${uid('c')}`,
      challanNo: finalNo,
      challanDate: isTax ? taxDate : mainDate,
      category: baseCategory,
      dealerName,
      dealerAllocationId: dealerAllocationId || undefined,
      challanType: isTax ? 'আয়কর চালান' : 'মূল চালান',
      bankName,
      doNo,
      doDate,
      allocationQty: parsedQty,
      recordedAmount: isTax ? recTax || recMain : recMain,
      depositedAmount: isTax ? depTax || depMain : depMain,
      note,
      createdAt,
    });
  });

  return out;
}

export async function fetchExternalChallansFromSheet(
  rawSpreadsheetId: string,
  accessToken?: string | null
): Promise<{ challans: ChallanRecord[]; title: string; tabUsed?: string }> {
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  if (!spreadsheetId) throw new Error('চালান Google Sheet ID বা লিংক দিন।');

  if (accessToken) {
    const meta = await getSpreadsheetMetadata(accessToken, spreadsheetId);
    const primaryMatch = findMatchingSheetTitle(
      meta.sheetTitles,
      ALLOC_SHEET_TABS.challans.name,
      ALLOC_SHEET_TABS.challans.aliases
    );
    const candidates = primaryMatch
      ? [primaryMatch, ...meta.sheetTitles.filter((t) => t !== primaryMatch)]
      : meta.sheetTitles;

    for (const tab of candidates.slice(0, 6)) {
      const escaped = tab.replace(/'/g, "''");
      const resp = await googleFetch<{ valueRanges?: Array<{ values?: unknown[][] }> }>(
        `${SHEETS_API_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchGet?ranges=${encodeURIComponent(`'${escaped}'!A1:Z`)}&majorDimension=ROWS&valueRenderOption=UNFORMATTED_VALUE`,
        accessToken
      );
      const parsed = parseChallanRowsFromTable(resp.valueRanges?.[0]?.values || []);
      if (parsed.length > 0) {
        return { challans: parsed, title: meta.title, tabUsed: tab };
      }
    }
    return { challans: [], title: meta.title };
  }

  const { meta, gidMap } = await getPublicSpreadsheetMetadata(spreadsheetId);
  const primaryMatch = findMatchingSheetTitle(
    meta.sheetTitles,
    ALLOC_SHEET_TABS.challans.name,
    ALLOC_SHEET_TABS.challans.aliases
  );
  const candidates =
    meta.sheetTitles.length > 0
      ? primaryMatch
        ? [primaryMatch, ...meta.sheetTitles.filter((t) => t !== primaryMatch)]
        : meta.sheetTitles
      : [ALLOC_SHEET_TABS.challans.name, 'Challan', 'Sheet1', 'Form Responses 1'];

  for (const tab of candidates.slice(0, 6)) {
    const rows = await fetchPublicCsvSheet(spreadsheetId, gidMap[tab], tab);
    const parsed = parseChallanRowsFromTable(rows);
    if (parsed.length > 0) {
      return { challans: parsed, title: meta.title, tabUsed: tab };
    }
  }
  return { challans: [], title: meta.title };
}

export async function fetchExternalChallansFromAppsScript(
  rawUrl: string
): Promise<{ challans: ChallanRecord[]; sourceUrl: string }> {
  const url = rawUrl.trim();
  if (!url) throw new Error('চালান ওয়েব অ্যাপ URL দিন।');

  const isNativeOrLocal =
    typeof window !== 'undefined' &&
    window.location.protocol === 'file:';

  let json: unknown = null;
  if (!isNativeOrLocal) {
    const resp = await fetch('/api/gas-proxy', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url, action: 'getAppData' }),
    });
    const body = (await resp.json()) as { ok?: boolean; data?: unknown; error?: string };
    if (!body?.ok) {
      throw new Error(body?.error || 'চালান Apps Script থেকে ডেটা পাওয়া যায়নি।');
    }
    json = body.data;
  } else {
    const u = new URL(url);
    u.searchParams.set('api', '1');
    u.searchParams.set('action', 'getAppData');
    const resp = await fetch(u.toString());
    json = await resp.json();
  }

  const root = (json && typeof json === 'object' ? json : {}) as Record<string, unknown>;
  const rawList = Array.isArray(json)
    ? json
    : Array.isArray(root.challans)
      ? root.challans
      : Array.isArray(root.data)
        ? root.data
        : Array.isArray(root.rows)
          ? root.rows
          : [];

  if (rawList.length === 0) {
    return { challans: [], sourceUrl: url };
  }

  let table: unknown[][];
  if (Array.isArray(rawList[0])) {
    table = rawList as unknown[][];
  } else {
    const firstObj = rawList[0] as Record<string, unknown>;
    const headers = Object.keys(firstObj);
    table = [
      headers,
      ...rawList.map((item) => {
        const obj = (item || {}) as Record<string, unknown>;
        return headers.map((h) => obj[h] ?? '');
      }),
    ];
  }

  return {
    challans: parseChallanRowsFromTable(table),
    sourceUrl: url,
  };
}

export async function fetchAllocationFromOAuthSheet(
  accessToken: string,
  rawSpreadsheetId: string,
  currentDb?: AllocationDatabase
): Promise<{ db: AllocationDatabase; title: string; tabsFound: string[]; matchedTabNames: string[] }> {
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  if (!spreadsheetId) throw new Error('Google Sheet ID বা লিংক দিন।');

  const meta = await getSpreadsheetMetadata(accessToken, spreadsheetId);
  const entries: Array<[AllocationRawTabKey, { name: string; aliases: readonly string[] }]> = [
    ...(Object.entries(ALLOC_SHEET_TABS) as Array<
      [keyof typeof ALLOC_SHEET_TABS, (typeof ALLOC_SHEET_TABS)[keyof typeof ALLOC_SHEET_TABS]]
    >),
    ['legacyProjects', LEGACY_COMBINED_PROJECTS_TAB],
  ];

  const matchedCandidatesMap = new Map<AllocationRawTabKey, string[]>();
  const matchedTabs = new Set<AllocationRawTabKey>();
  const matchedTabNames: string[] = [];
  const ranges: string[] = [];
  const rangeRefs: Array<{ key: AllocationRawTabKey; title: string }> = [];
  for (const [key, def] of entries) {
    const candidates = findAllMatchingSheetTitlesForWrite(
      meta.sheetTitles,
      def.name,
      def.aliases,
      !isExactOnlyTabKey(key),
      key
    );
    if (candidates.length > 0) {
      matchedCandidatesMap.set(key, candidates);
      matchedTabs.add(key);
      matchedTabNames.push(candidates[0]);
      for (const cand of candidates.slice(0, 2)) {
        const escaped = cand.replace(/'/g, "''");
        ranges.push(`'${escaped}'!A1:Z`);
        rangeRefs.push({ key, title: cand });
      }
    }
  }

  if (matchedTabs.size === 0 && meta.sheetTitles.length > 0) {
    throw new Error(
      `"${meta.title}" শিটে বরাদ্দ/ডিও-এর কাঙ্ক্ষিত ট্যাব পাওয়া যায়নি (পাওয়া ট্যাব: ${meta.sheetTitles.slice(0, 6).join(', ')}).`
    );
  }

  const rawTables: Partial<Record<AllocationRawTabKey, unknown[][]>> = {};
  if (ranges.length > 0) {
    const params = new URLSearchParams();
    ranges.forEach((r) => params.append('ranges', r));
    params.set('majorDimension', 'ROWS');
    params.set('valueRenderOption', 'UNFORMATTED_VALUE');
    const resp = await googleFetch<{ valueRanges?: Array<{ values?: unknown[][] }> }>(
      `${SHEETS_API_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchGet?${params.toString()}`,
      accessToken
    );
    const byKey = new Map<AllocationRawTabKey, Array<{ title: string; values: unknown[][] }>>();
    (resp.valueRanges || []).forEach((vr, idx) => {
      const ref = rangeRefs[idx];
      if (!ref) return;
      const list = byKey.get(ref.key) || [];
      list.push({ title: ref.title, values: vr.values || [] });
      byKey.set(ref.key, list);
    });
    for (const [key, list] of byKey.entries()) {
      const primaryEntry = list[0];
      if (!primaryEntry) continue;
      if (primaryEntry.values.length <= 1 && list.length > 1 && list[1].values.length > 1) {
        rawTables[key] = list[1].values;
        const idx = matchedTabNames.indexOf(primaryEntry.title);
        if (idx >= 0) matchedTabNames[idx] = list[1].title;
      } else {
        rawTables[key] = primaryEntry.values;
      }
    }
  }

  return {
    db: buildAllocationDbFromRawTables(rawTables, matchedTabs, currentDb),
    title: meta.title,
    tabsFound: meta.sheetTitles,
    matchedTabNames,
  };
}

export async function autoDetectAllocationSpreadsheets(accessToken: string): Promise<{
  files: DriveSpreadsheetFile[];
  main?: { file: DriveSpreadsheetFile; meta: SpreadsheetMetadata };
  challan?: { file: DriveSpreadsheetFile; meta: SpreadsheetMetadata };
}> {
  const files = await listUserSpreadsheets(accessToken);
  let main: { file: DriveSpreadsheetFile; meta: SpreadsheetMetadata } | undefined;
  let challan: { file: DriveSpreadsheetFile; meta: SpreadsheetMetadata } | undefined;

  const sample = files.slice(0, 15);
  for (let i = 0; i < sample.length; i += 5) {
    const chunk = sample.slice(i, i + 5);
    const metas = await Promise.all(
      chunk.map(async (file) => {
        try {
          const meta = await getSpreadsheetMetadata(accessToken, file.id);
          return { file, meta };
        } catch {
          return null;
        }
      })
    );
    for (const item of metas) {
      if (!item) continue;
      const titles = item.meta.sheetTitles;
      const hasMin = Boolean(findMatchingSheetTitle(titles, ALLOC_SHEET_TABS.ministry.name, ALLOC_SHEET_TABS.ministry.aliases));
      const hasSub = Boolean(findMatchingSheetTitle(titles, ALLOC_SHEET_TABS.dealerSub.name, ALLOC_SHEET_TABS.dealerSub.aliases));
      const hasDo = Boolean(findMatchingSheetTitle(titles, ALLOC_SHEET_TABS.doIssue.name, ALLOC_SHEET_TABS.doIssue.aliases));
      const hasChallan = Boolean(findMatchingSheetTitle(titles, ALLOC_SHEET_TABS.challans.name, ALLOC_SHEET_TABS.challans.aliases));
      const looksChallanOnly = hasChallan && !hasMin && !hasSub && !hasDo;

      if (!main && (hasMin || hasSub || hasDo)) {
        main = item;
      }
      if (!challan && (looksChallanOnly || (hasChallan && item.file.id !== main?.file.id))) {
        challan = item;
      }
    }
    if (main && challan) break;
  }

  return { files, main, challan };
}

const OLD_APP_ID_FIRST_HEADERS: Partial<Record<AllocationTabKey, string[]>> = {
  ministry: ['id', 'category', 'memoNo', 'memoDate', 'allocationQty', 'grain', 'phase', 'comment', 'createdAt'],
  dc: ['id', 'category', 'ministryMemo', 'dcMemo', 'dcDate', 'allocationQty', 'createdAt'],
  dealerSub: ['id', 'category', 'sourceMemo', 'unoMemo', 'date', 'dealerName', 'unionName', 'salesCenter', 'cardCount', 'unusedCards', 'allocationQty', 'totalAmount', 'taxAmount', 'createdAt'],
  grProjects: ['id', 'category', 'dcMemo', 'projectNo', 'projectName', 'presidentName', 'unionName', 'allocationQty', 'createdAt'],
  kabikhaProjects: ['id', 'category', 'dcMemo', 'projectNo', 'projectName', 'presidentName', 'unionName', 'allocationQty', 'createdAt'],
  unoSub: ['id', 'category', 'ministryMemo', 'dcMemo', 'unoMemo', 'unoDate', 'unionName', 'recipientName', 'projectNo', 'projectName', 'allocationQty', 'createdAt'],
  doBooks: ['id', 'bookNo', 'startPage', 'endPage', 'totalCount', 'createdAt'],
  challans: ['id', 'challanNo', 'challanDate', 'category', 'dealerName', 'challanType', 'bankName', 'doNo', 'recordedAmount', 'depositedAmount', 'note', 'createdAt', 'dealerAllocationId', 'doDate', 'allocationQty'],
};

const TAB_COLUMN_ALIASES: Record<AllocationTabKey, Record<string, string[]>> = {
  ministry: {
    id: ['id', 'rowid', 'recordid', 'ministryid', 'allocid', 'uid'],
    category: ['category', 'khath', 'khat', 'khatname', 'categoryname', 'খাত', 'খাতেরনাম', 'কর্মসূচি', 'কর্মসূচিরনাম', 'program', 'scheme', 'sector', 'head', 'type'],
    memoNo: ['memono', 'memo', 'ministrymemo', 'orderno', 'স্মারক', 'স্মারকনং', 'মন্ত্রণালয়স্মারক', 'আদেশনং'],
    memoDate: ['memodate', 'date', 'orderdate', 'entrydate', 'তারিখ', 'স্মারকতারিখ', 'স্মারকেরতারিখ'],
    allocationQty: ['allocationqty', 'qty', 'quantity', 'amountmt', 'mt', 'পরিমাণ', 'পরিমাণমেটন', 'বরাদ্দ', 'বরাদ্দেরপরিমাণ', 'বরাদ্দপরিমাণ'],
    grain: ['grain', 'commodity', 'item', 'product', 'foodgrain', 'শস্য', 'পণ্য', 'খাদ্যশস্য'],
    phase: ['phase', 'step', 'round', 'কিস্তি', 'পর্যায়', 'দফা'],
    comment: ['comment', 'note', 'notes', 'remarks', 'remark', 'description', 'মন্তব্য', 'বিবরণ'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
  },
  dc: {
    id: ['id', 'rowid', 'recordid', 'dcid', 'uid'],
    category: ['category', 'khath', 'khat', 'khatname', 'categoryname', 'খাত', 'খাতেরনাম', 'কর্মসূচি', 'কর্মসূচিরনাম', 'program', 'scheme', 'sector', 'head', 'type'],
    ministryMemo: ['ministrymemo', 'sourcememo', 'parentmemo', 'minmemo', 'মন্ত্রণালয়স্মারক', 'মন্ত্রণালয়েরস্মারক', 'উৎসস্মারক'],
    dcMemo: ['dcmemo', 'memono', 'memo', 'ডিসিস্মারক', 'জেলাপ্রশাসকস্মারক', 'স্মারক', 'স্মারকনং'],
    dcDate: ['dcdate', 'dcmemodate', 'memodate', 'date', 'orderdate', 'তারিখ', 'স্মারকতারিখ', 'ডিসিতারিক', 'ডিসিতারিখ'],
    allocationQty: ['allocationqty', 'qty', 'quantity', 'amountmt', 'mt', 'পরিমাণ', 'পরিমাণমেটন', 'বরাদ্দ', 'বরাদ্দপরিমাণ', 'বরাদ্দেরপরিমাণ'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
  },
  dealerSub: {
    id: ['id', 'rowid', 'recordid', 'subid', 'dealerallocationid', 'uid'],
    category: ['category', 'khath', 'khat', 'khatname', 'categoryname', 'খাত', 'খাতেরনাম', 'কর্মসূচি', 'program', 'scheme', 'sector', 'head', 'type'],
    sourceMemo: ['sourcememo', 'ministrymemo', 'dcmemo', 'parentmemo', 'উৎসস্মারক', 'মন্ত্রণালয়স্মারক', 'ডিসিস্মারক'],
    unoMemo: ['unomemo', 'ucfmemo', 'submemo', 'memono', 'memo', 'ইউসিএফস্মারক', 'উপবরাদ্দস্মারক', 'স্মারক'],
    date: ['date', 'unodate', 'ucfdate', 'subdate', 'memodate', 'তারিখ', 'উপবরাদ্দতারিখ'],
    dealerName: ['dealername', 'dealer', 'recipientname', 'recipient', 'officename', 'ডিলার', 'ডিলারেরনাম', 'প্রাপক'],
    unionName: ['unionname', 'union', 'municipality', 'area', 'ইউনিয়ন', 'এলাকা'],
    salesCenter: ['salescenter', 'center', 'distributioncenter', 'বিক্রয়কেন্দ্র', 'কেন্দ্র'],
    cardCount: ['cardcount', 'cards', 'totalcards', 'কার্ড', 'কার্ডসংখ্যা'],
    unusedCards: ['unusedcards', 'unsoldcards', 'অবিক্রিতকার্ড'],
    allocationQty: ['allocationqty', 'qty', 'quantity', 'subqty', 'mt', 'পরিমাণ', 'পরিমাণমেটন', 'উপবরাদ্দ', 'বরাদ্দ'],
    totalAmount: ['totalamount', 'amount', 'price', 'টাকা', 'মোটটাকা'],
    taxAmount: ['taxamount', 'tax', 'incometax', 'আয়কর'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
  },
  grProjects: {
    id: ['id', 'rowid', 'recordid', 'projectid', 'uid'],
    category: ['category', 'khath', 'khat', 'খাত', 'কর্মসূচি', 'program', 'scheme', 'type'],
    dcMemo: ['dcmemo', 'memono', 'memo', 'ডিসিস্মারক', 'স্মারক'],
    projectNo: ['projectno', 'projno', 'pno', 'প্রকল্পনং', 'প্রনং', 'প্রকল্পনম্বর'],
    projectName: ['projectname', 'projname', 'name', 'প্রকল্পেরনাম', 'প্রকল্পনাম', 'প্রনাম'],
    presidentName: ['presidentname', 'president', 'secretaryname', 'secretary', 'সভাপতি', 'সভাপতিরনাম', 'সম্পাদক', 'সভাপতিসম্পাদক'],
    unionName: ['unionname', 'union', 'ইউনিয়ন'],
    allocationQty: ['allocationqty', 'qty', 'quantity', 'mt', 'পরিমাণ', 'বরাদ্দ'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
  },
  kabikhaProjects: {
    id: ['id', 'rowid', 'recordid', 'projectid', 'uid'],
    category: ['category', 'khath', 'khat', 'খাত', 'কর্মসূচি', 'program', 'scheme', 'type'],
    dcMemo: ['dcmemo', 'memono', 'memo', 'ডিসিস্মারক', 'স্মারক'],
    projectNo: ['projectno', 'projno', 'pno', 'প্রকল্পনং', 'প্রনং', 'প্রকল্পনম্বর'],
    projectName: ['projectname', 'projname', 'name', 'প্রকল্পেরনাম', 'প্রকল্পনাম', 'প্রনাম'],
    presidentName: ['presidentname', 'president', 'secretaryname', 'secretary', 'সভাপতি', 'সভাপতিরনাম', 'সম্পাদক', 'সভাপতিসম্পাদক'],
    unionName: ['unionname', 'union', 'ইউনিয়ন'],
    allocationQty: ['allocationqty', 'qty', 'quantity', 'mt', 'পরিমাণ', 'বরাদ্দ'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
  },
  unoSub: {
    id: ['id', 'rowid', 'recordid', 'unoid', 'subid', 'uid'],
    category: ['category', 'khath', 'khat', 'khatname', 'categoryname', 'খাত', 'খাতেরনাম', 'কর্মসূচি', 'program', 'scheme', 'sector', 'head', 'type'],
    ministryMemo: ['ministrymemo', 'sourcememo', 'minmemo', 'মন্ত্রণালয়স্মারক', 'উৎসস্মারক', 'বরাদ্দস্মারক'],
    dcMemo: ['dcmemo', 'ডিসিস্মারক', 'জেলাপ্রশাসকস্মারক'],
    unoMemo: ['unomemo', 'submemo', 'memono', 'memo', 'ইউএনওস্মারক', 'উপবরাদ্দস্মারক', 'স্মারক', 'স্মারকনং'],
    unoDate: ['unodate', 'date', 'subdate', 'memodate', 'তারিখ', 'উপবরাদ্দতারিখ'],
    unionName: ['unionname', 'union', 'area', 'ইউনিয়ন'],
    recipientName: ['recipientname', 'recipient', 'chairmanname', 'chairman', 'presidentname', 'president', 'প্রাপক', 'চেয়ারম্যান', 'সভাপতি'],
    projectNo: ['projectno', 'projno', 'প্রকল্পনং', 'প্রনং', 'প্রকল্পনম্বর'],
    projectName: ['projectname', 'projname', 'প্রকল্পেরনাম', 'প্রকল্পনাম', 'প্রনাম'],
    allocationQty: ['allocationqty', 'qty', 'quantity', 'subqty', 'mt', 'পরিমাণ', 'পরিমাণমেটন', 'উপবরাদ্দ', 'বরাদ্দ'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
  },
  doBooks: {
    id: ['id', 'rowid', 'bookid', 'uid'],
    bookNo: ['bookno', 'book', 'বইনং'],
    startPage: ['startpage', 'start', 'শুরুপাতা'],
    endPage: ['endpage', 'end', 'শেষপাতা'],
    totalCount: ['totalcount', 'total', 'মোটডিও'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
  },
  doIssue: {
    doNo: ['dono', 'donumber', 'do', 'ডিওনং', 'ডিওনম্বর', 'ডিও'],
    issueDate: ['issuedate', 'dodate', 'date', 'তারিখ', 'ইস্যুতারিখ'],
    category: ['category', 'khath', 'khat', 'খাত', 'খাতেরনাম', 'কর্মসূচি', 'program', 'scheme', 'sector', 'head', 'type'],
    unionName: ['unionname', 'union', 'area', 'municipality', 'ইউনিয়ন', 'এলাকা'],
    recipientName: ['recipientname', 'dealername', 'dealer', 'প্রাপক', 'ডিলার', 'প্রতিষ্ঠান'],
    salesCenter: ['salescenter', 'center', 'বিক্রয়কেন্দ্র'],
    allocationQty: ['allocationqty', 'quantity', 'qty', 'doqty', 'mt', 'পরিমাণ', 'পরিমাণমেটন', 'বরাদ্দ'],
    ministryMemo: ['ministrymemo', 'sourcememo', 'মন্ত্রণালয়স্মারক', 'উৎসস্মারক'],
    dcMemo: ['dcmemo', 'ডিসিস্মারক'],
    unoMemo: ['unomemo', 'ucfmemo', 'suballocationmemo', 'memono', 'memo', 'ইউসিএফস্মারক', 'ইউএনওস্মারক', 'উপবরাদ্দস্মারক'],
    projectNo: ['projectno', 'প্রকল্পনং'],
    dealerAllocationId: ['dealerallocationid', 'suballocationid', 'allocationid'],
    challanNo: ['challanno', 'mainchallanno', 'মূলচালান', 'মূলচালাননং', 'চালাননং'],
    challanDate: ['challandate', 'mainchallandate', 'চালানেরতারিখ'],
    challanAmount: ['challanamount', 'mainchallanamount', 'depositedamount', 'মোটটাকা', 'জমাকৃতটাকা'],
    taxChallanNo: ['taxchallanno', 'incometaxchallanno', 'আয়করচালান', 'আয়করচালাননং'],
    taxChallanDate: ['taxchallandate', 'incometaxdate', 'আয়করতারিখ'],
    taxChallanAmount: ['taxchallanamount', 'incometaxamount', 'taxamount', 'আয়করটাকা'],
    requiredAmount: ['requiredamount', 'recordedamount'],
    requiredTaxAmount: ['requiredtaxamount'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
    id: ['id', 'rowid', 'doid', 'uid'],
  },
  challans: {
    id: ['id', 'rowid', 'challanid', 'uid'],
    challanNo: ['challanno', 'challannumber', 'mainchallanno', 'চালাননং', 'চালাননম্বর', 'চালান'],
    challanDate: ['challandate', 'mainchallandate', 'date', 'তারিখ', 'চালানেরতারিখ'],
    category: ['category', 'khath', 'khat', 'খাত', 'কর্মসূচি', 'program', 'scheme'],
    dealerName: ['dealername', 'dealer', 'recipientname', 'ডিলার', 'ডিলারেরনাম', 'প্রাপক'],
    challanType: ['challantype', 'type', 'ধরন', 'চালানেরধরন'],
    bankName: ['bankname', 'bank', 'ব্যাংক', 'ব্যাংকেরনাম'],
    doNo: ['dono', 'donumber', 'ডিওনং'],
    recordedAmount: ['recordedamount', 'requiredamount', 'রেকর্ডটাকা'],
    depositedAmount: ['depositedamount', 'challanamount', 'amount', 'জমাকৃতটাকা'],
    note: ['note', 'remarks', 'comment', 'মন্তব্য'],
    createdAt: ['createdat', 'timestamp', 'updatedat', 'created'],
    dealerAllocationId: ['dealerallocationid', 'suballocationid', 'allocationid'],
    doDate: ['dodate', 'ডিওতারিখ'],
    allocationQty: ['allocationqty', 'quantity', 'qty', 'পরিমাণ'],
  },
  foodDealers: {
    id: ['id', 'rowid', 'uid'],
    dealerName: ['dealername', 'dealer', 'ডিলারেরনাম', 'ডিলার', 'নাম'],
    fatherName: ['fathername', 'পিতারনাম'],
    nid: ['nid', 'এনআইডি'],
    mobile: ['mobile', 'phone', 'মোবাইল'],
    unionName: ['unionname', 'union', 'ইউনিয়নেরনাম', 'ইউনিয়ন'],
    salesCenter: ['salescenter', 'center', 'বিক্রয়কেন্দ্র'],
    cardCount: ['cardcount', 'cards', 'কার্ড', 'কার্ডসংখ্যা'],
    totalQty: ['totalqty', 'পরিমাণ', 'মোটপরিমাণ'],
    poNo: ['pono', 'পেঅর্ডারনং'],
    poDate: ['podate', 'পেঅর্ডারতারিখ'],
    bankName: ['bankname', 'bank', 'ব্যাংক'],
    amount: ['amount', 'টাকারপরিমাণ'],
    supervisorName: ['supervisorname', 'তদারকিকর্মকর্তা'],
    supervisorDesignation: ['supervisordesignation', 'পদবি', 'পদবী'],
    supervisorMobile: ['supervisormobile'],
  },
  omsDealers: {
    id: ['id', 'rowid', 'uid'],
    dealerName: ['dealername', 'dealer', 'ডিলারেরনাম', 'ডিলার', 'নাম'],
    fatherName: ['fathername', 'পিতারনাম'],
    nid: ['nid', 'এনআইডি'],
    mobile: ['mobile', 'phone', 'মোবাইল'],
    areaType: ['areatype', 'এলাকারধরন'],
    unionName: ['unionname', 'union', 'ইউনিয়নেরনাম', 'ইউনিয়ন'],
    municipality: ['municipality', 'পৌরসভা'],
    salesCenter: ['salescenter', 'center', 'বিক্রয়কেন্দ্র'],
    poNo: ['pono', 'পেঅর্ডারনং'],
    poDate: ['podate', 'পেঅর্ডারতারিখ'],
    bankName: ['bankname', 'bank', 'ব্যাংক'],
    amount: ['amount', 'টাকারপরিমাণ'],
    supervisorName: ['supervisorname', 'তদারকিকর্মকর্তা'],
    supervisorDesignation: ['supervisordesignation', 'পদবি', 'পদবী'],
    supervisorMobile: ['supervisormobile'],
  },
  chairmen: {
    id: ['id', 'rowid', 'uid'],
    unionName: ['unionname', 'union', 'ইউনিয়নেরনাম', 'ইউনিয়ন'],
    chairmanName: ['chairmanname', 'chairman', 'চেয়ারম্যানেরনাম', 'চেয়ারম্যান', 'প্রাপক', 'নাম'],
    mobile: ['mobile', 'phone', 'মোবাইল'],
    supervisorName: ['supervisorname', 'তদারকিকর্মকর্তা'],
    supervisorMobile: ['supervisormobile'],
  },
  fireServices: {
    id: ['id', 'rowid', 'uid'],
    officeName: ['officename', 'office', 'অফিস'],
    divisionName: ['divisionname', 'division', 'বিভাগীয়কার্যালয়'],
    riceRate: ['ricerate'],
    wheatRate: ['wheatrate'],
  },
  rates: {
    id: ['id', 'rowid', 'uid'],
    category: ['category', 'khath', 'খাত'],
    unitPrice: ['unitprice', 'price', 'rate', 'এককমূল্য', 'দর'],
    taxRate: ['taxrate', 'tax', 'আয়কর'],
  },
  settings: {
    key: ['key'],
    value: ['value'],
  },
};

export interface BuiltSheetTabPayload {
  key: AllocationTabKey;
  name: string;
  aliases: readonly string[];
  exactOnly: boolean;
  dateFieldKey?: string;
  columnAliases: Record<string, string[]>;
  grainValues?: string[];
  rows: Array<Array<string | number>>;
}

function toCleanIsoDate(val: unknown): string {
  const s = String(val ?? '').trim();
  if (!s) return '';
  const parsed = parseDateStr(s);
  return /^\d{4}-\d{2}-\d{2}$/.test(parsed) ? parsed : s;
}

function resolveGrainForCategory(
  cat: AllocationCategory,
  ministryAllocs: MinistryAllocation[],
  minMemo?: string
): string {
  if (minMemo) {
    const nt = normMemoNo(minMemo);
    const hit = ministryAllocs.find(
      (m) => m.category === cat && (m.memoNo === minMemo || (nt && normMemoNo(m.memoNo) === nt))
    );
    if (hit?.grain) return hit.grain;
  }
  const low = cat.toLowerCase();
  if (cat.includes('গম') || low.includes('wheat')) return 'গম';
  if (cat.includes('আটা') || low.includes('atta')) return 'আটা';
  return 'চাল';
}

function computeFiscalYearFormatted(dateStr: string, sampleStyle = ''): string {
  const my = parseMonthYear(dateStr);
  const now = new Date();
  const yr = my ? my.year : now.getFullYear();
  const mo0 = my ? my.month : now.getMonth();
  const fullBn = fiscalYearLabel(yr, mo0);
  const fullEn = bnToEnDigits(fullBn);
  const m = fullEn.match(/^(\d{4})-(\d{4})$/);
  if (!m) return fullEn;
  const startYr = m[1];
  const endYr = m[2];
  if (/[০-৯]/.test(sampleStyle)) {
    return /^\d{4}-\d{2}$/.test(bnToEnDigits(sampleStyle.trim()))
      ? toBnDigits(`${startYr}-${endYr.slice(2)}`)
      : fullBn;
  }
  if (/^\d{4}-\d{2}$/.test(sampleStyle.trim())) {
    return `${startYr}-${endYr.slice(2)}`;
  }
  return fullEn;
}

function projectPayloadRowsForExistingSheet(
  payload: BuiltSheetTabPayload,
  existingGrid: unknown[][]
): { startRow: number; rows: Array<Array<string | number>> } {
  if (payload.key === 'settings' || !existingGrid || existingGrid.length === 0) {
    return { startRow: 1, rows: payload.rows };
  }

  const incomingHeaders = (payload.rows[0] || []).map((h) => String(h));
  const aliasToIncomingIdx = new Map<string, number>();
  incomingHeaders.forEach((h, idx) => {
    const nh = normColKey(h);
    if (nh && !aliasToIncomingIdx.has(nh)) aliasToIncomingIdx.set(nh, idx);
    const aliases = payload.columnAliases[h] || payload.columnAliases[nh] || [];
    for (const a of aliases) {
      const na = normColKey(a);
      if (na && !aliasToIncomingIdx.has(na)) aliasToIncomingIdx.set(na, idx);
    }
  });

  // Detect header row index within top 8 rows
  const scanLimit = Math.min(existingGrid.length, 8);
  let headerRowIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < scanLimit; i++) {
    const r = existingGrid[i] || [];
    const nonEmpty = r.filter((c) => String(c ?? '').trim() !== '');
    if (nonEmpty.length < 2) continue;
    let score = nonEmpty.length;
    for (const cell of nonEmpty) {
      const nk = normColKey(cell);
      if (!nk) continue;
      if (aliasToIncomingIdx.has(nk)) score += 6;
      else if (HEADER_HINT_TOKENS.some((tok) => nk === tok || nk.includes(tok))) score += 5;
    }
    if (score > bestScore) {
      bestScore = score;
      headerRowIdx = i;
    }
  }

  if (bestScore < 6) {
    return { startRow: 1, rows: payload.rows };
  }

  const rawHeaderRow = [...(existingGrid[headerRowIdx] || [])];
  while (rawHeaderRow.length > 0 && String(rawHeaderRow[rawHeaderRow.length - 1] ?? '').trim() === '') {
    rawHeaderRow.pop();
  }
  if (rawHeaderRow.length < 2) {
    return { startRow: 1, rows: payload.rows };
  }

  const existingHeaders = rawHeaderRow.map((h) => String(h ?? '').trim());
  const oldDefault = OLD_APP_ID_FIRST_HEADERS[payload.key];
  if (
    oldDefault &&
    existingHeaders.length === oldDefault.length &&
    existingHeaders.every((h, idx) => normColKey(h) === normColKey(oldDefault[idx]))
  ) {
    // Upgrade legacy app-generated id-first header layout to id-last (matching DO_Issue)
    return { startRow: headerRowIdx + 1, rows: payload.rows };
  }

  const idIncomingIdx = incomingHeaders.findIndex((h) => normColKey(h) === 'id');
  const dateIncomingIdx = payload.dateFieldKey
    ? incomingHeaders.findIndex((h) => normColKey(h) === normColKey(payload.dateFieldKey))
    : -1;
  const hasIdCol = existingHeaders.some((h) => aliasToIncomingIdx.get(normColKey(h)) === idIncomingIdx);
  const targetHeaders = hasIdCol || idIncomingIdx < 0 ? [...existingHeaders] : [...existingHeaders, 'id'];

  // Gather existing data rows for preserving custom unknown columns, FY style, and status/constant defaults
  const oldRowsById = new Map<string, unknown[]>();
  const existingIdCol = existingHeaders.findIndex((h) => aliasToIncomingIdx.get(normColKey(h)) === idIncomingIdx);
  let fySampleStyle = '';
  const colModeMap = new Map<number, string>();
  const fyColTokens = ['fiscalyear', 'fy', 'session', 'অর্থবছর'];
  const slColTokens = ['sl', 'slno', 'serial', 'serialno', 'ক্রমিক', 'ক্রম'];
  const grainColTokens = ['grain', 'commodity', 'item', 'foodgrain', 'শস্য', 'পণ্য', 'খাদ্যশস্য'];
  const modeAllowedTokens = ['status', 'state', 'অবস্থা', 'upazila', 'উপজেলা', 'district', 'জেলা', 'office', 'division', 'fiscalyear', 'fy', 'অর্থবছর'];

  const oldDataRows = existingGrid.slice(headerRowIdx + 1);
  if (existingIdCol >= 0) {
    for (const r of oldDataRows) {
      const idVal = String(r[existingIdCol] ?? '').trim();
      if (idVal) oldRowsById.set(idVal, r);
    }
  }

  existingHeaders.forEach((h, cIdx) => {
    const nh = normColKey(h);
    if (fyColTokens.includes(nh) && !fySampleStyle) {
      for (const r of oldDataRows) {
        const v = String(r[cIdx] ?? '').trim();
        if (v) {
          fySampleStyle = v;
          break;
        }
      }
    }
    if (modeAllowedTokens.some((tok) => nh === tok || nh.includes(tok))) {
      const counts = new Map<string, number>();
      for (const r of oldDataRows) {
        const v = String(r[cIdx] ?? '').trim();
        if (!v) continue;
        counts.set(v, (counts.get(v) || 0) + 1);
      }
      let bestVal = '';
      let bestCount = 0;
      for (const [val, cnt] of counts.entries()) {
        if (cnt > bestCount) {
          bestCount = cnt;
          bestVal = val;
        }
      }
      if (bestVal) colModeMap.set(cIdx, bestVal);
    }
  });

  const projectedRows: Array<Array<string | number>> = [targetHeaders];
  for (let rIdx = 1; rIdx < payload.rows.length; rIdx++) {
    const srcRow = payload.rows[rIdx];
    const rowId = idIncomingIdx >= 0 ? String(srcRow[idIncomingIdx] ?? '').trim() : '';
    const oldRow = rowId ? oldRowsById.get(rowId) : undefined;
    const rowDateStr = dateIncomingIdx >= 0 ? String(srcRow[dateIncomingIdx] ?? '') : '';

    const outRow: Array<string | number> = targetHeaders.map((th, colIdx) => {
      const nh = normColKey(th);
      const srcIdx = aliasToIncomingIdx.get(nh);
      if (srcIdx !== undefined) {
        return srcRow[srcIdx] ?? '';
      }
      if (slColTokens.includes(nh)) {
        return rIdx;
      }
      if (fyColTokens.includes(nh)) {
        const oldFy = oldRow && colIdx < existingHeaders.length ? String(oldRow[colIdx] ?? '').trim() : '';
        return oldFy || computeFiscalYearFormatted(rowDateStr, fySampleStyle);
      }
      if (grainColTokens.includes(nh) && payload.grainValues?.[rIdx - 1]) {
        return payload.grainValues[rIdx - 1];
      }
      if (oldRow && colIdx < existingHeaders.length) {
        const oldVal = oldRow[colIdx];
        if (oldVal !== undefined && oldVal !== null && String(oldVal).trim() !== '') {
          return typeof oldVal === 'number' ? oldVal : String(oldVal);
        }
      }
      return colModeMap.get(colIdx) ?? '';
    });
    projectedRows.push(outRow);
  }

  return { startRow: headerRowIdx + 1, rows: projectedRows };
}

const ALLOC_MAIN_APPS_SCRIPT_BODY = `function normTabName_(s) {
  return String(s || '').normalize('NFC').replace(/\\u09AF\\u09BC/g, '\\u09DF').trim().toLowerCase().replace(/[\\s_-]+/g, '');
}

function normColKey_(s) {
  return String(s || '').normalize('NFC').replace(/\\u09AF\\u09BC/g, '\\u09DF').trim().toLowerCase().replace(/[\\s_.\\-/():]+/g, '');
}

function isSummaryOrReportTab_(norm) {
  return norm.indexOf('dashboard') >= 0 || norm.indexOf('summary') >= 0 ||
    norm.indexOf('সারসংক্ষেপ') >= 0 || norm.indexOf('ড্যাশবোর্ড') >= 0 ||
    norm.indexOf('report') >= 0 || norm.indexOf('রিপোর্ট') >= 0 ||
    norm.indexOf('instructions') >= 0 || norm.indexOf('নির্দেশনা') >= 0;
}

function isIncompatiblePartialTab_(tabKey, normTitle) {
  if (!tabKey) return false;
  if (isSummaryOrReportTab_(normTitle)) return true;
  if (tabKey === 'dealerSub') {
    return normTitle.indexOf('fooddealer') >= 0 || normTitle.indexOf('omsdealer') >= 0 ||
      normTitle.indexOf('খাদ্যবান্ধবডিলার') >= 0 || normTitle.indexOf('ওএমএসডিলার') >= 0 ||
      normTitle.indexOf('uno') >= 0 || normTitle.indexOf('upazila') >= 0 || normTitle.indexOf('relief') >= 0 ||
      normTitle.indexOf('ইউএনও') >= 0 || normTitle.indexOf('উপজেলা') >= 0 || normTitle.indexOf('ত্রাণ') >= 0 ||
      normTitle.indexOf('ministry') >= 0 || normTitle.indexOf('dc') >= 0 ||
      normTitle.indexOf('মন্ত্রণাল') >= 0 || normTitle.indexOf('ডিসি') >= 0 ||
      normTitle.indexOf('dobook') >= 0 || normTitle.indexOf('doissue') >= 0 ||
      normTitle.indexOf('challan') >= 0 || normTitle.indexOf('চালান') >= 0 ||
      normTitle.indexOf('project') >= 0 || normTitle.indexOf('প্রকল্প') >= 0;
  }
  if (tabKey === 'foodDealers') return normTitle.indexOf('oms') >= 0 || normTitle.indexOf('sub') >= 0 || normTitle.indexOf('allocation') >= 0;
  if (tabKey === 'omsDealers') return normTitle.indexOf('food') >= 0 || normTitle.indexOf('খাদ্য') >= 0 || normTitle.indexOf('sub') >= 0 || normTitle.indexOf('allocation') >= 0;
  if (tabKey === 'unoSub') {
    return normTitle.indexOf('dealer') >= 0 || normTitle.indexOf('ucf') >= 0 ||
      normTitle.indexOf('ডিলার') >= 0 || normTitle.indexOf('ইউসিএফ') >= 0 ||
      normTitle.indexOf('dobook') >= 0 || normTitle.indexOf('doissue') >= 0 ||
      normTitle.indexOf('ministry') >= 0 || normTitle.indexOf('dc') >= 0 ||
      normTitle.indexOf('মন্ত্রণাল') >= 0 || normTitle.indexOf('ডিসি') >= 0 ||
      normTitle.indexOf('challan') >= 0 || normTitle.indexOf('চালান') >= 0 ||
      normTitle.indexOf('project') >= 0 || normTitle.indexOf('প্রকল্প') >= 0;
  }
  if (tabKey === 'doBooks') return normTitle.indexOf('issue') >= 0 || normTitle.indexOf('ইস্যু') >= 0 || normTitle.indexOf('ministry') >= 0 || normTitle.indexOf('dc') >= 0 || normTitle.indexOf('uno') >= 0 || normTitle.indexOf('dealer') >= 0 || normTitle.indexOf('challan') >= 0;
  if (tabKey === 'doIssue') return normTitle.indexOf('book') >= 0 || normTitle.indexOf('বই') >= 0 || normTitle.indexOf('ministry') >= 0 || normTitle.indexOf('dc') >= 0 || normTitle.indexOf('uno') >= 0 || normTitle.indexOf('dealer') >= 0 || normTitle.indexOf('challan') >= 0;
  if (tabKey === 'challans') return normTitle.indexOf('ministry') >= 0 || normTitle.indexOf('dc') >= 0 || normTitle.indexOf('uno') >= 0 || normTitle.indexOf('dealer') >= 0 || normTitle.indexOf('dobook') >= 0 || normTitle.indexOf('doissue') >= 0 || normTitle.indexOf('project') >= 0;
  if (tabKey === 'ministry') {
    return normTitle.indexOf('dc') >= 0 || normTitle.indexOf('uno') >= 0 || normTitle.indexOf('dealer') >= 0 ||
      normTitle.indexOf('challan') >= 0 || normTitle.indexOf('dobook') >= 0 || normTitle.indexOf('doissue') >= 0 ||
      normTitle.indexOf('project') >= 0 || normTitle.indexOf('প্রকল্প') >= 0 ||
      normTitle.indexOf('ডিসি') >= 0 || normTitle.indexOf('ইউএনও') >= 0 || normTitle.indexOf('উপবরাদ্দ') >= 0;
  }
  if (tabKey === 'dc') {
    return normTitle.indexOf('ministry') >= 0 || normTitle.indexOf('uno') >= 0 || normTitle.indexOf('dealer') >= 0 ||
      normTitle.indexOf('challan') >= 0 || normTitle.indexOf('dobook') >= 0 || normTitle.indexOf('doissue') >= 0 ||
      normTitle.indexOf('project') >= 0 || normTitle.indexOf('প্রকল্প') >= 0 ||
      normTitle.indexOf('মন্ত্রণাল') >= 0 || normTitle.indexOf('ইউএনও') >= 0 || normTitle.indexOf('উপবরাদ্দ') >= 0;
  }
  if (tabKey === 'chairmen' || tabKey === 'fireServices' || tabKey === 'rates' || tabKey === 'settings') {
    return normTitle.indexOf('allocation') >= 0 || normTitle.indexOf('uno') >= 0 || normTitle.indexOf('dealer') >= 0 || normTitle.indexOf('dobook') >= 0 || normTitle.indexOf('doissue') >= 0 || normTitle.indexOf('challan') >= 0;
  }
  return false;
}

function findMatchingSheets_(ss, t, readOnly) {
  var sheets = ss.getSheets();
  var targets = [];
  var rawTargets = [t.name].concat(Array.isArray(t.aliases) ? t.aliases : []);
  for (var i = 0; i < rawTargets.length; i++) {
    var na = normTabName_(rawTargets[i]);
    if (na && targets.indexOf(na) === -1) targets.push(na);
  }
  var exact = [];
  for (var k = 0; k < targets.length; k++) {
    var tg = targets[k];
    for (var s = 0; s < sheets.length; s++) {
      if (normTabName_(sheets[s].getName()) === tg && exact.indexOf(sheets[s]) === -1) {
        exact.push(sheets[s]);
      }
    }
  }
  if (exact.length > 0) return exact;
  if (!t.exactOnly) {
    var partial = [];
    for (var s2 = 0; s2 < sheets.length; s2++) {
      var nt = normTabName_(sheets[s2].getName());
      if (!nt || isIncompatiblePartialTab_(t.key, nt)) continue;
      for (var k2 = 0; k2 < targets.length; k2++) {
        var tg2 = targets[k2];
        if (tg2.length >= 3 && (nt.indexOf(tg2) >= 0 || tg2.indexOf(nt) >= 0)) {
          partial.push(sheets[s2]);
          break;
        }
      }
    }
    if (partial.length > 0) return partial;
  }
  if (readOnly) return [];
  return [ss.insertSheet(t.name)];
}

function isRedundantPluralDup_(primaryName, candName) {
  var p = normTabName_(primaryName);
  var c = normTabName_(candName);
  if (!p || !c || p === c) return false;
  if (p === 'chairman' && c === 'chairmen') return true;
  if (p.slice(-1) !== 's' && c === p + 's') return true;
  return false;
}

function cellToVal_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : v;
}

function computeFiscalYear_(dateVal, sampleStyle) {
  var s = String(dateVal || '').trim();
  var m = s.match(/^(\\d{4})-(\\d{2})/);
  var now = new Date();
  var yr = m ? Number(m[1]) : now.getFullYear();
  var mo0 = m ? Number(m[2]) - 1 : now.getMonth();
  var sy = mo0 >= 6 ? yr : yr - 1;
  var ey = sy + 1;
  var useShort = /^\\d{4}-\\d{2}$/.test(String(sampleStyle || '').replace(/[০-৯]/g, function(d) { return String('০১২৩৪৫৬৭৮৯'.indexOf(d)); }).trim());
  var res = useShort ? (sy + '-' + String(ey).slice(2)) : (sy + '-' + ey);
  if (/[০-৯]/.test(String(sampleStyle || ''))) {
    return res.replace(/[0-9]/g, function(d) { return '০১২৩৪৫৬৭৮৯'[Number(d)]; });
  }
  return res;
}

var OLD_ID_FIRST_MAP_ = {
  ministry: ['id','category','memono','memodate','allocationqty','grain','phase','comment','createdat'],
  dc: ['id','category','ministrymemo','dcmemo','dcdate','allocationqty','createdat'],
  dealerSub: ['id','category','sourcememo','unomemo','date','dealername','unionname','salescenter','cardcount','unusedcards','allocationqty','totalamount','taxamount','createdat'],
  grProjects: ['id','category','dcmemo','projectno','projectname','presidentname','unionname','allocationqty','createdat'],
  kabikhaProjects: ['id','category','dcmemo','projectno','projectname','presidentname','unionname','allocationqty','createdat'],
  unoSub: ['id','category','ministrymemo','dcmemo','unomemo','unodate','unionname','recipientname','projectno','projectname','allocationqty','createdat'],
  doBooks: ['id','bookno','startpage','endpage','totalcount','createdat'],
  challans: ['id','challanno','challandate','category','dealername','challantype','bankname','dono','recordedamount','depositedamount','note','createdat','dealerallocationid','dodate','allocationqty']
};

function projectRowsForSheet_(sheet, t) {
  var rows = t.rows;
  if (t.key === 'settings' || t.name === 'Alloc_Settings') {
    return { startRow: 1, rows: rows };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { startRow: 1, rows: rows };
  }
  var existingGrid = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var incomingHeaders = rows[0].map(function(h) { return String(h); });
  var aliasMap = {};
  var ca = t.columnAliases || {};
  for (var j = 0; j < incomingHeaders.length; j++) {
    var hName = incomingHeaders[j];
    var nh = normColKey_(hName);
    if (nh && aliasMap[nh] === undefined) aliasMap[nh] = j;
    var list = ca[hName] || ca[nh] || [];
    for (var a = 0; a < list.length; a++) {
      var na = normColKey_(list[a]);
      if (na && aliasMap[na] === undefined) aliasMap[na] = j;
    }
  }
  var hintTokens = ['id','category','khath','khat','খাত','memo','স্মারক','date','তারিখ','qty','quantity','পরিমাণ','dealer','ডিলার','union','ইউনিয়ন','dono','ডিও','project','প্রকল্প','fy','fiscalyear','অর্থবছর','sl','ক্রমিক','status','অবস্থা'];
  var scanLimit = Math.min(existingGrid.length, 8);
  var headerIdx = 0;
  var bestScore = -1;
  for (var r = 0; r < scanLimit; r++) {
    var row = existingGrid[r] || [];
    var nonEmpty = [];
    for (var c = 0; c < row.length; c++) {
      var sv = String(row[c] === null || row[c] === undefined ? '' : row[c]).trim();
      if (sv !== '') nonEmpty.push(sv);
    }
    if (nonEmpty.length < 2) continue;
    var score = nonEmpty.length;
    for (var n = 0; n < nonEmpty.length; n++) {
      var nk = normColKey_(nonEmpty[n]);
      if (!nk) continue;
      if (aliasMap[nk] !== undefined) score += 6;
      else {
        for (var ht = 0; ht < hintTokens.length; ht++) {
          if (nk === hintTokens[ht] || nk.indexOf(hintTokens[ht]) >= 0) { score += 5; break; }
        }
      }
    }
    if (score > bestScore) { bestScore = score; headerIdx = r; }
  }
  if (bestScore < 6) {
    return { startRow: 1, rows: rows };
  }
  var rawHdr = existingGrid[headerIdx].slice();
  while (rawHdr.length > 0 && String(rawHdr[rawHdr.length - 1] === null || rawHdr[rawHdr.length - 1] === undefined ? '' : rawHdr[rawHdr.length - 1]).trim() === '') {
    rawHdr.pop();
  }
  if (rawHdr.length < 2) {
    return { startRow: 1, rows: rows };
  }
  var existingHeaders = rawHdr.map(function(h) { return String(h === null || h === undefined ? '' : h).trim(); });
  var oldDef = OLD_ID_FIRST_MAP_[t.key];
  if (oldDef && existingHeaders.length === oldDef.length) {
    var allOld = true;
    for (var oi = 0; oi < oldDef.length; oi++) {
      if (normColKey_(existingHeaders[oi]) !== oldDef[oi]) { allOld = false; break; }
    }
    if (allOld) return { startRow: headerIdx + 1, rows: rows };
  }
  var idInIdx = -1;
  var dateInIdx = -1;
  var dkNorm = t.dateFieldKey ? normColKey_(t.dateFieldKey) : '';
  for (var ii = 0; ii < incomingHeaders.length; ii++) {
    var ink = normColKey_(incomingHeaders[ii]);
    if (ink === 'id') idInIdx = ii;
    if (dkNorm && ink === dkNorm) dateInIdx = ii;
  }
  var hasId = false;
  var existingIdCol = -1;
  for (var eh = 0; eh < existingHeaders.length; eh++) {
    if (aliasMap[normColKey_(existingHeaders[eh])] === idInIdx && idInIdx >= 0) {
      hasId = true;
      existingIdCol = eh;
      break;
    }
  }
  var targetHeaders = (hasId || idInIdx < 0) ? existingHeaders.slice() : existingHeaders.concat(['id']);
  var oldRowsById = {};
  var oldDataRows = existingGrid.slice(headerIdx + 1);
  if (existingIdCol >= 0) {
    for (var od = 0; od < oldDataRows.length; od++) {
      var idKey = String(oldDataRows[od][existingIdCol] || '').trim();
      if (idKey) oldRowsById[idKey] = oldDataRows[od];
    }
  }
  var fyTokens = ['fiscalyear','fy','session','অর্থবছর'];
  var slTokens = ['sl','slno','serial','serialno','ক্রমিক','ক্রম'];
  var grainTokens = ['grain','commodity','item','foodgrain','শস্য','পণ্য','খাদ্যশস্য'];
  var modeTokens = ['status','state','অবস্থা','upazila','উপজেলা','district','জেলা','office','division','fiscalyear','fy','অর্থবছর'];
  var fySample = '';
  var colModes = {};
  for (var ci = 0; ci < existingHeaders.length; ci++) {
    var ch = normColKey_(existingHeaders[ci]);
    if (fyTokens.indexOf(ch) >= 0 && !fySample) {
      for (var fr = 0; fr < oldDataRows.length; fr++) {
        var fv = String(oldDataRows[fr][ci] || '').trim();
        if (fv) { fySample = fv; break; }
      }
    }
    var allowMode = false;
    for (var mt = 0; mt < modeTokens.length; mt++) {
      if (ch === modeTokens[mt] || ch.indexOf(modeTokens[mt]) >= 0) { allowMode = true; break; }
    }
    if (allowMode) {
      var counts = {};
      var bestV = '';
      var bestC = 0;
      for (var mr = 0; mr < oldDataRows.length; mr++) {
        var mv = String(cellToVal_(oldDataRows[mr][ci])).trim();
        if (!mv) continue;
        counts[mv] = (counts[mv] || 0) + 1;
        if (counts[mv] > bestC) { bestC = counts[mv]; bestV = mv; }
      }
      if (bestV) colModes[ci] = bestV;
    }
  }
  var out = [targetHeaders];
  var gVals = Array.isArray(t.grainValues) ? t.grainValues : [];
  for (var ri = 1; ri < rows.length; ri++) {
    var srcRow = rows[ri];
    var rid = idInIdx >= 0 ? String(srcRow[idInIdx] || '').trim() : '';
    var oldRow = rid ? oldRowsById[rid] : null;
    var rDate = dateInIdx >= 0 ? String(srcRow[dateInIdx] || '') : '';
    var line = [];
    for (var tj = 0; tj < targetHeaders.length; tj++) {
      var thn = normColKey_(targetHeaders[tj]);
      var sIdx = aliasMap[thn];
      if (sIdx !== undefined) {
        line.push(srcRow[sIdx] === undefined || srcRow[sIdx] === null ? '' : srcRow[sIdx]);
      } else if (slTokens.indexOf(thn) >= 0) {
        line.push(ri);
      } else if (fyTokens.indexOf(thn) >= 0) {
        var oldFy = oldRow && tj < existingHeaders.length ? String(cellToVal_(oldRow[tj])).trim() : '';
        line.push(oldFy || computeFiscalYear_(rDate, fySample));
      } else if (grainTokens.indexOf(thn) >= 0 && gVals[ri - 1]) {
        line.push(gVals[ri - 1]);
      } else if (oldRow && tj < existingHeaders.length && String(cellToVal_(oldRow[tj])).trim() !== '') {
        line.push(cellToVal_(oldRow[tj]));
      } else {
        line.push(colModes[tj] !== undefined ? colModes[tj] : '');
      }
    }
    out.push(line);
  }
  return { startRow: headerIdx + 1, rows: out };
}

function openTargetSpreadsheet_(payload) {
  var ss = null;
  var targetSheetId = String((payload && payload.spreadsheetId) || ALLOC_DEFAULT_SPREADSHEET_ID_ || '').trim();
  var openErr = '';
  if (targetSheetId) {
    try {
      ss = SpreadsheetApp.openById(targetSheetId);
    } catch (eOpen) {
      openErr = eOpen && eOpen.message ? eOpen.message : String(eOpen);
    }
  }
  if (!ss) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (_) {}
  }
  return { ss: ss, openErr: openErr };
}

function readAllTabsFromApp_(payload) {
  var opened = openTargetSpreadsheet_(payload);
  if (!opened.ss) {
    return {
      ok: false,
      error: opened.openErr
        ? ('Google Sheet খোলা যায়নি (' + opened.openErr + ')')
        : 'Spreadsheet পাওয়া যায়নি। অ্যাপের ১ নম্বর ঘরে Google Sheet ID দিন।'
    };
  }
  var ss = opened.ss;
  var sheets = ss.getSheets();
  var tables = {};
  var sheetTitles = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var name = sh.getName();
    sheetTitles.push(name);
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) {
      tables[name] = [];
      continue;
    }
    var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
    tables[name] = vals.map(function(row) {
      return row.map(function(cell) { return cellToVal_(cell); });
    });
  }
  return {
    ok: true,
    title: ss.getName(),
    spreadsheetId: ss.getId(),
    sheetTitles: sheetTitles,
    tables: tables
  };
}

function syncAllTabsFromApp_(payload) {
  var opened = openTargetSpreadsheet_(payload);
  var ss = opened.ss;
  if (!ss) {
    return {
      ok: false,
      error: opened.openErr
        ? ('Google Sheet খোলা যায়নি (' + opened.openErr + ')')
        : 'Spreadsheet পাওয়া যায়নি। অ্যাপের ১ নম্বর ঘরে Google Sheet ID দিন।'
    };
  }
  var tabs = Array.isArray(payload && payload.tabs) ? payload.tabs : [];
  var incomingCoreRows = 0;
  for (var cIdx = 0; cIdx < tabs.length; cIdx++) {
    var ct = tabs[cIdx];
    if (!ct || !ct.name || !Array.isArray(ct.rows)) continue;
    if (ct.key === 'settings' || ct.name === 'Alloc_Settings') continue;
    if (ct.rows.length > 1) incomingCoreRows += (ct.rows.length - 1);
  }
  var allowEmpty = Boolean(payload && payload.allowEmptyOverwrite);
  if (incomingCoreRows === 0 && !allowEmpty) {
    for (var chkIdx = 0; chkIdx < tabs.length; chkIdx++) {
      var chkTab = tabs[chkIdx];
      if (!chkTab || !chkTab.name || chkTab.key === 'settings' || chkTab.name === 'Alloc_Settings') continue;
      var existingMatches = findMatchingSheets_(ss, chkTab, true);
      for (var em = 0; em < existingMatches.length; em++) {
        if (existingMatches[em].getLastRow() > 1) {
          return {
            ok: false,
            error: 'অ্যাপে কোনো ডেটা না থাকায় Google Sheet-এর বিদ্যমান ডেটা সুরক্ষার জন্য ওভাররাইট বাতিল করা হয়েছে। আগে «লিংক সেভ ও শিট থেকে ডেটা আনুন» চাপুন।'
          };
        }
      }
    }
  }
  var updated = 0;
  var duplicatesToDelete = [];
  for (var i = 0; i < tabs.length; i++) {
    var t = tabs[i];
    if (!t || !t.name || !Array.isArray(t.rows) || t.rows.length === 0) continue;
    var matchedSheets = findMatchingSheets_(ss, t, false);
    var primarySheet = matchedSheets[0];
    var writeSheets = [primarySheet];
    var tHasDataRows = t.rows.length > 1;
    for (var ms = 1; ms < matchedSheets.length; ms++) {
      var candSheet = matchedSheets[ms];
      if (isRedundantPluralDup_(primarySheet.getName(), candSheet.getName())) {
        if (tHasDataRows || candSheet.getLastRow() <= 1) {
          if (duplicatesToDelete.indexOf(candSheet) === -1) {
            duplicatesToDelete.push(candSheet);
          }
        }
      } else {
        writeSheets.push(candSheet);
      }
    }
    for (var s = 0; s < writeSheets.length; s++) {
      var sheet = writeSheets[s];
      var projected = projectRowsForSheet_(sheet, t);
      var rows = projected.rows;
      var maxCols = 1;
      for (var r = 0; r < rows.length; r++) {
        if (Array.isArray(rows[r]) && rows[r].length > maxCols) {
          maxCols = rows[r].length;
        }
      }
      var padded = rows.map(function(row) {
        var arr = Array.isArray(row) ? row.slice() : [];
        while (arr.length < maxCols) arr.push('');
        return arr;
      });
      var startRow = projected.startRow || 1;
      var prevLastRow = sheet.getLastRow();
      var prevLastCol = Math.max(sheet.getLastColumn(), maxCols);
      if (startRow === 1) {
        sheet.clearContents();
      } else if (prevLastRow >= startRow) {
        sheet.getRange(startRow, 1, prevLastRow - startRow + 1, prevLastCol).clearContent();
      }
      sheet.getRange(startRow, 1, padded.length, maxCols).setValues(padded);
      updated++;
    }
  }
  var duplicatesRemoved = 0;
  for (var d = 0; d < duplicatesToDelete.length; d++) {
    try {
      if (ss.getSheets().length > 1) {
        ss.deleteSheet(duplicatesToDelete[d]);
        duplicatesRemoved++;
      }
    } catch (_) {}
  }
  SpreadsheetApp.flush();
  return { ok: true, tabsUpdated: updated, duplicatesRemoved: duplicatesRemoved };
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = body.action || '';
    var payload = body.payload || body;
    if (action === 'readAllTabs') {
      return allocJsonOut_(readAllTabsFromApp_(payload));
    }
    if (action !== 'syncAllTabs') {
      return allocJsonOut_({ ok: false, error: 'Unknown action: ' + action });
    }
    return allocJsonOut_(syncAllTabsFromApp_(payload));
  } catch (err) {
    return allocJsonOut_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function allocJsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}`;

export function buildMainAppsScriptCode(
  rawDefaultSpreadsheetId = '',
  includeDoGet = true
): string {
  const cleanId = extractSpreadsheetId(rawDefaultSpreadsheetId).replace(/['"\\]/g, '');
  const header = `var ALLOC_DEFAULT_SPREADSHEET_ID_ = '${cleanId}';\n\n`;
  const doGetBlock = includeDoGet
    ? `\n\nfunction doGet() {\n  return allocJsonOut_({\n    ok: true,\n    status: 'ready',\n    spreadsheetId: ALLOC_DEFAULT_SPREADSHEET_ID_ || '(from payload)',\n    message: 'Khaddo Boraddo Sync Web App Active'\n  });\n}`
    : '';
  return `${header}${ALLOC_MAIN_APPS_SCRIPT_BODY}${doGetBlock}`;
}

export const ALLOC_MAIN_APPS_SCRIPT_CODE = buildMainAppsScriptCode('', false);

function chunkSettingEntries(key: string, jsonStr: string): Array<[string, string]> {
  const CHUNK_SIZE = 40000;
  if (jsonStr.length <= CHUNK_SIZE) {
    return [[key, jsonStr]];
  }
  const chunks: string[] = [];
  for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
    chunks.push(jsonStr.slice(i, i + CHUNK_SIZE));
  }
  return [
    [key, `__CHUNKED__:${chunks.length}`],
    ...chunks.map((chunk, idx): [string, string] => [`${key}__chunk_${idx}`, chunk]),
  ];
}

export function buildAllocationSheetPayloads(
  db: AllocationDatabase,
  opts?: { skipChallansTab?: boolean }
): BuiltSheetTabPayload[] {
  const grProjectsList = db.projects.filter((r) => r.category === 'জি আর');
  const kabikhaProjectsList = db.projects.filter((r) => r.category !== 'জি আর');

  const allTabs: BuiltSheetTabPayload[] = [
    {
      key: 'ministry',
      name: ALLOC_SHEET_TABS.ministry.name,
      aliases: ALLOC_SHEET_TABS.ministry.aliases,
      exactOnly: false,
      dateFieldKey: 'memoDate',
      columnAliases: TAB_COLUMN_ALIASES.ministry,
      grainValues: db.ministryAllocations.map((r) => r.grain || resolveGrainForCategory(r.category, db.ministryAllocations)),
      rows: [
        [...ALLOC_SHEET_TABS.ministry.headers],
        ...db.ministryAllocations.map((r) => [
          r.category,
          r.memoNo,
          toCleanIsoDate(r.memoDate),
          r.allocationQty,
          r.grain,
          r.phase ?? '',
          r.comment ?? '',
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'dc',
      name: ALLOC_SHEET_TABS.dc.name,
      aliases: ALLOC_SHEET_TABS.dc.aliases,
      exactOnly: false,
      dateFieldKey: 'dcDate',
      columnAliases: TAB_COLUMN_ALIASES.dc,
      grainValues: db.dcAllocations.map((r) => resolveGrainForCategory(r.category, db.ministryAllocations, r.ministryMemo)),
      rows: [
        [...ALLOC_SHEET_TABS.dc.headers],
        ...db.dcAllocations.map((r) => [
          r.category,
          r.ministryMemo,
          r.dcMemo,
          toCleanIsoDate(r.dcMemoDate),
          r.allocationQty,
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'dealerSub',
      name: ALLOC_SHEET_TABS.dealerSub.name,
      aliases: ALLOC_SHEET_TABS.dealerSub.aliases,
      exactOnly: false,
      dateFieldKey: 'date',
      columnAliases: TAB_COLUMN_ALIASES.dealerSub,
      grainValues: db.dealerSuballocations.map((r) => resolveGrainForCategory(r.category, db.ministryAllocations, r.sourceMemo)),
      rows: [
        [...ALLOC_SHEET_TABS.dealerSub.headers],
        ...db.dealerSuballocations.map((r) => [
          r.category,
          r.sourceMemo,
          r.unoMemo,
          toCleanIsoDate(r.unoDate),
          r.dealerName,
          r.unionName,
          r.salesCenter,
          r.cardCount,
          r.unusedCards,
          r.allocationQty,
          r.totalAmount,
          r.taxAmount,
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'grProjects',
      name: ALLOC_SHEET_TABS.grProjects.name,
      aliases: ALLOC_SHEET_TABS.grProjects.aliases,
      exactOnly: true,
      columnAliases: TAB_COLUMN_ALIASES.grProjects,
      grainValues: grProjectsList.map((r) => resolveGrainForCategory(r.category, db.ministryAllocations)),
      rows: [
        [...ALLOC_SHEET_TABS.grProjects.headers],
        ...grProjectsList.map((r) => [
          r.category,
          r.dcMemo,
          r.projectNo,
          r.projectName,
          r.presidentName,
          r.unionName,
          r.allocationQty,
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'kabikhaProjects',
      name: ALLOC_SHEET_TABS.kabikhaProjects.name,
      aliases: ALLOC_SHEET_TABS.kabikhaProjects.aliases,
      exactOnly: true,
      columnAliases: TAB_COLUMN_ALIASES.kabikhaProjects,
      grainValues: kabikhaProjectsList.map((r) => resolveGrainForCategory(r.category, db.ministryAllocations)),
      rows: [
        [...ALLOC_SHEET_TABS.kabikhaProjects.headers],
        ...kabikhaProjectsList.map((r) => [
          r.category,
          r.dcMemo,
          r.projectNo,
          r.projectName,
          r.presidentName,
          r.unionName,
          r.allocationQty,
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'unoSub',
      name: ALLOC_SHEET_TABS.unoSub.name,
      aliases: ALLOC_SHEET_TABS.unoSub.aliases,
      exactOnly: false,
      dateFieldKey: 'unoDate',
      columnAliases: TAB_COLUMN_ALIASES.unoSub,
      grainValues: db.unoSuballocations.map((r) => resolveGrainForCategory(r.category, db.ministryAllocations, r.ministryMemo)),
      rows: [
        [...ALLOC_SHEET_TABS.unoSub.headers],
        ...db.unoSuballocations.map((r) => [
          r.category,
          r.ministryMemo,
          r.dcMemo,
          r.unoMemo,
          toCleanIsoDate(r.unoDate),
          r.unionName,
          r.recipientName,
          r.projectNo ?? '',
          r.projectName ?? '',
          r.allocationQty,
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'doBooks',
      name: ALLOC_SHEET_TABS.doBooks.name,
      aliases: ALLOC_SHEET_TABS.doBooks.aliases,
      exactOnly: false,
      columnAliases: TAB_COLUMN_ALIASES.doBooks,
      rows: [
        [...ALLOC_SHEET_TABS.doBooks.headers],
        ...db.doBooks.map((r) => [
          r.bookNo,
          r.startPage,
          r.endPage,
          r.totalCount,
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'doIssue',
      name: ALLOC_SHEET_TABS.doIssue.name,
      aliases: ALLOC_SHEET_TABS.doIssue.aliases,
      exactOnly: false,
      dateFieldKey: 'issueDate',
      columnAliases: TAB_COLUMN_ALIASES.doIssue,
      grainValues: db.doIssues.map((r) => resolveGrainForCategory(r.category, db.ministryAllocations, r.ministryMemo)),
      rows: [
        [...ALLOC_SHEET_TABS.doIssue.headers],
        ...db.doIssues.map((r) => [
          r.doNo,
          toCleanIsoDate(r.issueDate),
          r.category,
          r.unionName,
          r.recipientName,
          r.salesCenter ?? '',
          r.allocationQty,
          r.ministryMemo,
          r.dcMemo,
          r.unoMemo,
          r.projectNo ?? '',
          r.dealerAllocationId ?? '',
          r.challanNo ?? '',
          toCleanIsoDate(r.challanDate ?? ''),
          r.challanAmount ?? '',
          r.taxChallanNo ?? '',
          toCleanIsoDate(r.taxChallanDate ?? ''),
          r.taxChallanAmount ?? '',
          r.requiredAmount ?? '',
          r.requiredTaxAmount ?? '',
          r.createdAt,
          r.id,
        ]),
      ],
    },
    {
      key: 'challans',
      name: ALLOC_SHEET_TABS.challans.name,
      aliases: ALLOC_SHEET_TABS.challans.aliases,
      exactOnly: false,
      dateFieldKey: 'challanDate',
      columnAliases: TAB_COLUMN_ALIASES.challans,
      rows: [
        [...ALLOC_SHEET_TABS.challans.headers],
        ...db.challans.map((r) => [
          r.challanNo,
          toCleanIsoDate(r.challanDate),
          r.category,
          r.dealerName,
          r.challanType,
          r.bankName,
          r.doNo ?? '',
          r.recordedAmount,
          r.depositedAmount,
          r.note ?? '',
          r.createdAt,
          r.dealerAllocationId ?? '',
          toCleanIsoDate(r.doDate ?? ''),
          r.allocationQty ?? '',
          r.id,
        ]),
      ],
    },
    {
      key: 'foodDealers',
      name: ALLOC_SHEET_TABS.foodDealers.name,
      aliases: ALLOC_SHEET_TABS.foodDealers.aliases,
      exactOnly: false,
      dateFieldKey: 'poDate',
      columnAliases: TAB_COLUMN_ALIASES.foodDealers,
      rows: [
        [...ALLOC_SHEET_TABS.foodDealers.headers],
        ...db.foodDealers.map((r) => [
          r.id,
          r.dealerName,
          r.fatherName,
          r.nid,
          r.mobile,
          r.unionName,
          r.salesCenter,
          r.cardCount,
          r.totalQty,
          r.poNo,
          toCleanIsoDate(r.poDate),
          r.bankName,
          r.amount,
          r.supervisorName,
          r.supervisorDesignation,
          r.supervisorMobile,
        ]),
      ],
    },
    {
      key: 'omsDealers',
      name: ALLOC_SHEET_TABS.omsDealers.name,
      aliases: ALLOC_SHEET_TABS.omsDealers.aliases,
      exactOnly: false,
      dateFieldKey: 'poDate',
      columnAliases: TAB_COLUMN_ALIASES.omsDealers,
      rows: [
        [...ALLOC_SHEET_TABS.omsDealers.headers],
        ...db.omsDealers.map((r) => [
          r.id,
          r.dealerName,
          r.fatherName,
          r.nid,
          r.mobile,
          r.areaType,
          r.unionName,
          r.municipality,
          r.salesCenter,
          r.poNo,
          toCleanIsoDate(r.poDate),
          r.bankName,
          r.amount,
          r.supervisorName,
          r.supervisorDesignation,
          r.supervisorMobile,
        ]),
      ],
    },
    {
      key: 'chairmen',
      name: ALLOC_SHEET_TABS.chairmen.name,
      aliases: ALLOC_SHEET_TABS.chairmen.aliases,
      exactOnly: false,
      columnAliases: TAB_COLUMN_ALIASES.chairmen,
      rows: [
        [...ALLOC_SHEET_TABS.chairmen.headers],
        ...db.chairmen.map((r) => [
          r.id,
          r.unionName,
          r.chairmanName,
          r.mobile,
          r.supervisorName,
          r.supervisorMobile,
        ]),
      ],
    },
    {
      key: 'fireServices',
      name: ALLOC_SHEET_TABS.fireServices.name,
      aliases: ALLOC_SHEET_TABS.fireServices.aliases,
      exactOnly: false,
      columnAliases: TAB_COLUMN_ALIASES.fireServices,
      rows: [
        [...ALLOC_SHEET_TABS.fireServices.headers],
        ...db.fireServices.map((r) => [
          r.id,
          r.officeName,
          r.divisionName,
          r.riceRate,
          r.wheatRate,
        ]),
      ],
    },
    {
      key: 'rates',
      name: ALLOC_SHEET_TABS.rates.name,
      aliases: ALLOC_SHEET_TABS.rates.aliases,
      exactOnly: false,
      columnAliases: TAB_COLUMN_ALIASES.rates,
      rows: [
        [...ALLOC_SHEET_TABS.rates.headers],
        ...db.rates.map((r) => [r.id, r.category, r.unitPrice, r.taxRate]),
      ],
    },
    {
      key: 'settings',
      name: ALLOC_SHEET_TABS.settings.name,
      aliases: ALLOC_SHEET_TABS.settings.aliases,
      exactOnly: false,
      columnAliases: TAB_COLUMN_ALIASES.settings,
      rows: [
        [...ALLOC_SHEET_TABS.settings.headers],
        ...chunkSettingEntries('officer', JSON.stringify(db.officer)),
        ...chunkSettingEntries('templates', JSON.stringify(db.templates)),
        ...chunkSettingEntries('reportMemos', JSON.stringify(db.reportMemos || {})),
      ],
    },
  ];
  return opts?.skipChallansTab
    ? allTabs.filter((t) => t.name !== ALLOC_SHEET_TABS.challans.name)
    : allTabs;
}

export async function writeAllocationDatabaseToOAuthSheet(
  accessToken: string,
  rawSpreadsheetId: string,
  db: AllocationDatabase,
  opts?: { skipChallansTab?: boolean; allowEmptyOverwrite?: boolean }
): Promise<{ ok: boolean; tabsUpdated: number }> {
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  if (!spreadsheetId) throw new Error('Google Sheet ID নির্বাচন করুন।');

  const meta = await getSpreadsheetMetadata(accessToken, spreadsheetId);
  const payloads = buildAllocationSheetPayloads(db, opts);

  const missingTabs: string[] = [];
  const duplicateTabsToDelete = new Set<string>();
  const writeTargets: Array<{
    payload: BuiltSheetTabPayload;
    targetName: string;
    matchedExisting: boolean;
  }> = [];

  for (const p of payloads) {
    const plan = resolveSheetWritePlan(
      meta.sheetTitles,
      p.name,
      p.aliases,
      !p.exactOnly,
      p.key
    );
    if (plan.writeTitles.length === 0) {
      missingTabs.push(p.name);
      writeTargets.push({ payload: p, targetName: p.name, matchedExisting: false });
    } else {
      for (const mName of plan.writeTitles) {
        writeTargets.push({ payload: p, targetName: mName, matchedExisting: true });
      }
      if (p.rows.length > 1) {
        for (const dupName of plan.duplicatePluralTitles) {
          duplicateTabsToDelete.add(dupName);
        }
      }
    }
  }

  // Read existing sheet headers/rows in a single batchGet before clearing so existing schemas & extra columns are preserved
  const existingTargetNames = writeTargets.filter((t) => t.matchedExisting).map((t) => t.targetName);
  const existingGridMap = new Map<string, unknown[][]>();
  if (existingTargetNames.length > 0) {
    const params = new URLSearchParams();
    existingTargetNames.forEach((name) => {
      params.append('ranges', `'${name.replace(/'/g, "''")}'!A1:ZZ500`);
    });
    params.set('majorDimension', 'ROWS');
    params.set('valueRenderOption', 'UNFORMATTED_VALUE');
    try {
      const existingResp = await googleFetch<{ valueRanges?: Array<{ values?: unknown[][] }> }>(
        `${SHEETS_API_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchGet?${params.toString()}`,
        accessToken
      );
      (existingResp.valueRanges || []).forEach((vr, idx) => {
        const name = existingTargetNames[idx];
        if (name) existingGridMap.set(name, vr.values || []);
      });
    } catch {
      // If inspecting existing grids fails unexpectedly, proceed safely with default payloads
    }
  }

  const incomingCoreRows = payloads.reduce((sum, p) => {
    if (p.key === 'settings' || p.name === ALLOC_SHEET_TABS.settings.name) return sum;
    return sum + Math.max(0, p.rows.length - 1);
  }, 0);
  if (incomingCoreRows === 0 && !opts?.allowEmptyOverwrite) {
    const sheetHasExistingData = writeTargets.some((t) => {
      if (!t.matchedExisting || t.payload.key === 'settings') return false;
      const grid = existingGridMap.get(t.targetName);
      return Boolean(grid && grid.length > 1);
    });
    if (sheetHasExistingData) {
      throw new Error(
        'অ্যাপে কোনো ডেটা না থাকায় Google Sheet-এর বিদ্যমান ডেটা সুরক্ষার জন্য ওভাররাইট বাতিল করা হয়েছে। আগে «লিংক সেভ ও শিট থেকে ডেটা আনুন» চাপুন।'
      );
    }
  }

  if (missingTabs.length > 0) {
    await googleFetch(`${SHEETS_API_BASE}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, accessToken, {
      method: 'POST',
      body: JSON.stringify({
        requests: missingTabs.map((title) => ({ addSheet: { properties: { title } } })),
      }),
    });
  }

  const projectedWrites = writeTargets.map((t) => {
    const existingGrid = t.matchedExisting ? existingGridMap.get(t.targetName) || [] : [];
    const projected = projectPayloadRowsForExistingSheet(t.payload, existingGrid);
    return {
      targetName: t.targetName,
      startRow: projected.startRow,
      rows: projected.rows,
    };
  });

  await googleFetch(
    `${SHEETS_API_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchClear`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        ranges: projectedWrites.map((p) => {
          const esc = p.targetName.replace(/'/g, "''");
          return p.startRow > 1 ? `'${esc}'!A${p.startRow}:ZZ` : `'${esc}'!A:ZZ`;
        }),
      }),
    }
  );

  await googleFetch(
    `${SHEETS_API_BASE}/${encodeURIComponent(spreadsheetId)}/values:batchUpdate`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: projectedWrites.map((p) => ({
          range: `'${p.targetName.replace(/'/g, "''")}'!A${p.startRow}`,
          values: p.rows,
        })),
      }),
    }
  );

  // Clean up redundant plural (...s) duplicate tabs after primary singular tabs are safely written
  const totalSheetsAfterAdd = meta.sheetTitles.length + missingTabs.length;
  const deleteIds: number[] = [];
  for (const dupTitle of duplicateTabsToDelete) {
    const sid = meta.sheetIdMap[dupTitle];
    if (typeof sid === 'number' && totalSheetsAfterAdd - deleteIds.length > 1) {
      deleteIds.push(sid);
    }
  }
  if (deleteIds.length > 0) {
    try {
      await googleFetch(`${SHEETS_API_BASE}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, accessToken, {
        method: 'POST',
        body: JSON.stringify({
          requests: deleteIds.map((sheetId) => ({ deleteSheet: { sheetId } })),
        }),
      });
    } catch {
      // Non-fatal if a protected duplicate sheet cannot be deleted
    }
  }

  return { ok: true, tabsUpdated: projectedWrites.length };
}

export async function writeAllocationDatabaseToAppsScript(
  rawAppsScriptUrl: string,
  rawSpreadsheetId: string,
  db: AllocationDatabase,
  opts?: { skipChallansTab?: boolean; allowEmptyOverwrite?: boolean }
): Promise<{ ok: boolean; tabsUpdated: number }> {
  const cleanUrl = rawAppsScriptUrl.trim();
  if (!cleanUrl || (!cleanUrl.startsWith('https://') && !cleanUrl.startsWith('http://'))) {
    throw new Error('সঠিক Apps Script Web App URL (https://script.google.com/...) দিন।');
  }
  const spreadsheetId = extractSpreadsheetId(rawSpreadsheetId);
  const payloads = buildAllocationSheetPayloads(db, opts);
  const res = await callGasAction<{ tabsUpdated?: number }>(cleanUrl, 'syncAllTabs', {
    spreadsheetId,
    tabs: payloads,
    allowEmptyOverwrite: Boolean(opts?.allowEmptyOverwrite),
  });
  if (!res.ok) {
    throw new Error(res.error || 'Apps Script-এর মাধ্যমে Google Sheet-এ সেভ করা যায়নি।');
  }
  return {
    ok: true,
    tabsUpdated: Number(res.data?.tabsUpdated || payloads.length),
  };
}

export { uid };

