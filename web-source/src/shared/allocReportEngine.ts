import {
  CATEGORIES,
  CATEGORY_BY_NAME,
  CHALLAN_CATEGORIES,
  type AllocationCategory,
  type AllocationDatabase,
  type ChallanRecord,
  type DOIssue,
  type GrainType,
  type RegisterId,
  type ReportTemplateConfig,
} from '../types/allocation';
import {
  BENGALI_MONTHS,
  bnToEnDigits,
  fiscalYearLabel,
  formatBnDate,
  formatBnInt,
  formatBnMoney,
  formatBnQty,
  isChallanOnOrBeforeDo,
  isMonthInFiscalUpTo,
  isSamePartyName,
  memosLikelyMatch,
  namesFuzzyMatch,
  namesLikelyMatch,
  normChallanNo,
  normDoNo,
  normMemoNo,
  parseMonthYear,
  round2,
  round3,
  toBnDigits,
  toIsoDateStr,
  unionsLikelyMatch,
} from './bnUtils';

export interface RegisterColumnLeaf {
  key: string;
  label: string;
  align?: 'left' | 'center' | 'right';
  isSignature?: boolean;
  widthPct?: number;
}

export interface RegisterColumnGroup {
  label: string;
  children?: RegisterColumnLeaf[];
  leaf?: RegisterColumnLeaf;
}

export interface RegisterRowData {
  id: string;
  groupValues: string[];
  cells: Record<string, string>;
  hasSignature: boolean;
  highlight?: 'excess' | 'deficit' | 'none';
}

export interface BuiltRegisterData {
  id: RegisterId;
  title: string;
  category?: AllocationCategory;
  columns: RegisterColumnGroup[];
  leaves: RegisterColumnLeaf[];
  mergeDepth: number;
  serialPerRow?: boolean;
  colWidths?: string[];
  tableLayoutFixed?: boolean;
  rows: RegisterRowData[];
  summaryText?: string;
}

const LEAF_COL_WEIGHT: Record<string, number> = {
  ministry: 12,
  dc: 11,
  uno: 11.5,
  project: 20,
  union: 7.5,
  dealer: 12,
  recipient: 13,
  cards: 6.5,
  grain: 5,
  qty: 6,
  money: 15.5,
  challan: 15.5,
  doInfo: 9,
  signature: 7,
  date: 8,
  category: 10,
  doNo: 10,
  type: 8,
  challanNo: 15,
  recorded: 9,
  deposited: 9,
  mismatch: 11,
};

function computeRegisterColWidths(
  leaves: RegisterColumnLeaf[],
  showSerial = true,
  flowType?: string
): string[] {
  const serialPct = showSerial ? 3.5 : 0;
  const remainingPct = 100 - serialPct;
  const isDealerOmsOrFood = flowType === 'dealer_oms' || flowType === 'dealer_food';
  const weights = leaves.map((l) => {
    if (isDealerOmsOrFood) {
      if (l.key === 'money' || l.key === 'challan') return 16.5;
      if (l.key === 'ministry') return 11.5;
      if (l.key === 'dc') return 10.5;
      if (l.key === 'uno') return 11;
      if (l.key === 'union') return 7;
      if (l.key === 'dealer') return 11.5;
      if (l.key === 'cards') return 6.5;
      if (l.key === 'qty') return 5.5;
      if (l.key === 'doInfo') return 8.5;
      if (l.key === 'signature') return 6.5;
    }
    return LEAF_COL_WEIGHT[l.key] || 10;
  });
  const totalWeight = weights.reduce((s, w) => s + w, 0) || 1;
  const widths = weights.map((w) =>
    Number(((w / totalWeight) * remainingPct).toFixed(2))
  );
  if (widths.length > 0) {
    const sum = widths.reduce((s, w) => s + w, 0);
    widths[widths.length - 1] = Number(
      Math.max(4, widths[widths.length - 1] + (remainingPct - sum)).toFixed(2)
    );
  }
  return [
    ...(showSerial ? [`${serialPct}%`] : []),
    ...widths.map((w) => `${w}%`),
  ];
}

function dateDaysDiff(a?: string, b?: string): number {
  const tA = a ? Date.parse(bnToEnDigits(a).trim()) : NaN;
  const tB = b ? Date.parse(bnToEnDigits(b).trim()) : NaN;
  if (Number.isNaN(tA) || Number.isNaN(tB)) return 9999;
  return Math.abs(tA - tB) / 86400000;
}

function signedSubToDoDays(subDate?: string, doDate?: string): number {
  const tSub = subDate ? Date.parse(bnToEnDigits(subDate).trim()) : NaN;
  const tDo = doDate ? Date.parse(bnToEnDigits(doDate).trim()) : NaN;
  if (Number.isNaN(tSub) || Number.isNaN(tDo)) return 9999;
  return (tDo - tSub) / 86400000;
}

function isoMonthKey(d?: string): string {
  const s = bnToEnDigits(String(d || '')).trim();
  return /^\d{4}-\d{2}/.test(s) ? s.slice(0, 7) : '';
}

interface SubTarget {
  key: string;
  kind: 'ds' | 'uno';
  id: string;
  category: AllocationCategory;
  sourceMemo: string;
  dcMemo: string;
  unoMemo: string;
  unoDate: string;
  unionName: string;
  recipientName: string;
  projectNo: string;
  salesCenter: string;
  allocationQty: number;
}

export function matchDOsToSuballocations(db: AllocationDatabase): {
  dosBySubKey: Map<string, DOIssue[]>;
  remainingBySubKey: Map<string, number>;
} {
  const targets: SubTarget[] = [
    ...db.dealerSuballocations.map((s) => {
      const catMeta = CATEGORY_BY_NAME[s.category];
      const dcRec = catMeta?.requiresDC
        ? db.dcAllocations.find(
            (d) => d.category === s.category && memosLikelyMatch(d.dcMemo, s.sourceMemo)
          )
        : undefined;
      return {
        key: `ds:${s.id}`,
        kind: 'ds' as const,
        id: s.id,
        category: s.category,
        sourceMemo: dcRec ? dcRec.ministryMemo : s.sourceMemo,
        dcMemo: dcRec ? dcRec.dcMemo : '',
        unoMemo: s.unoMemo,
        unoDate: s.unoDate,
        unionName: s.unionName,
        recipientName: s.dealerName,
        projectNo: '',
        salesCenter: s.salesCenter,
        allocationQty: s.allocationQty,
      };
    }),
    ...db.unoSuballocations.map((u) => ({
      key: `uno:${u.id}`,
      kind: 'uno' as const,
      id: u.id,
      category: u.category,
      sourceMemo: u.ministryMemo,
      dcMemo: u.dcMemo,
      unoMemo: u.unoMemo,
      unoDate: u.unoDate,
      unionName: u.unionName,
      recipientName: u.recipientName,
      projectNo: u.projectNo || '',
      salesCenter: u.projectName || '',
      allocationQty: u.allocationQty,
    })),
  ].sort((a, b) => a.unoDate.localeCompare(b.unoDate));

  const challansBySubId = new Map<string, ChallanRecord[]>();
  for (const c of db.challans) {
    if (!c.dealerAllocationId) continue;
    const k = normMemoNo(c.dealerAllocationId);
    const arr = challansBySubId.get(k) || [];
    arr.push(c);
    challansBySubId.set(k, arr);
  }

  const dosBySubKey = new Map<string, DOIssue[]>();
  const remainingBySubKey = new Map<string, number>();
  for (const t of targets) {
    dosBySubKey.set(t.key, []);
    remainingBySubKey.set(t.key, round3(t.allocationQty));
  }
  const usedDoIds = new Set<string>();

  const assignDo = (t: SubTarget, d: DOIssue) => {
    usedDoIds.add(d.id);
    dosBySubKey.get(t.key)!.push(d);
    const curRem = remainingBySubKey.get(t.key) || 0;
    remainingBySubKey.set(t.key, round3(Math.max(0, curRem - (Number(d.allocationQty) || 0))));
  };

  const scoreCandidate = (t: SubTarget, d: DOIssue): number => {
    const qtyDiff = Math.abs((d.allocationQty || 0) - (t.allocationQty || 0));
    const days = dateDaysDiff(t.unoDate, d.issueDate);
    const exactUnionBonus =
      t.unionName && d.unionName && unionsLikelyMatch(t.unionName, d.unionName) ? 0 : 50;
    return qtyDiff * 1000 + Math.min(days, 365) + exactUnionBonus;
  };

  const runPass = (
    predicate: (t: SubTarget, d: DOIssue) => boolean,
    allowMultiPerSub = true
  ) => {
    for (const t of targets) {
      let guard = 0;
      while ((remainingBySubKey.get(t.key) || 0) > 0.0005 && guard < 20) {
        guard++;
        if (!allowMultiPerSub && (dosBySubKey.get(t.key)?.length || 0) > 0) break;
        const candidates = db.doIssues.filter(
          (d) => !usedDoIds.has(d.id) && predicate(t, d)
        );
        if (candidates.length === 0) break;
        candidates.sort((a, b) => scoreCandidate(t, a) - scoreCandidate(t, b));
        assignDo(t, candidates[0]);
      }
    }
  };

  // Pass 1: Direct ID / Challan link / ProjectNo / exact UNO memo + recipient (strictly same category)
  runPass((t, d) => {
    if (d.category !== t.category) return false;
    if (d.dealerAllocationId && normMemoNo(d.dealerAllocationId) === normMemoNo(t.id)) {
      return true;
    }
    const linkedChallans = challansBySubId.get(normMemoNo(t.id)) || [];
    if (
      linkedChallans.some(
        (c) =>
          (!c.category || c.category === t.category) &&
          isSamePartyName(c.dealerName || t.recipientName, d.recipientName) &&
          isChallanOnOrBeforeDo(c.challanDate, d.issueDate, false) &&
          ((c.doNo && normDoNo(c.doNo) === normDoNo(d.doNo)) ||
            (d.challanNo && normChallanNo(c.challanNo) === normChallanNo(d.challanNo)) ||
            (d.taxChallanNo && normChallanNo(c.challanNo) === normChallanNo(d.taxChallanNo)))
      )
    ) {
      return true;
    }
    if (
      t.projectNo &&
      d.projectNo &&
      memosLikelyMatch(t.projectNo, d.projectNo) &&
      (!t.dcMemo || !d.dcMemo || memosLikelyMatch(t.dcMemo, d.dcMemo))
    ) {
      return true;
    }
    if (
      t.unoMemo &&
      d.unoMemo &&
      memosLikelyMatch(t.unoMemo, d.unoMemo) &&
      namesFuzzyMatch(d.recipientName, t.recipientName) &&
      (!t.unionName || !d.unionName || unionsLikelyMatch(t.unionName, d.unionName))
    ) {
      return true;
    }
    return false;
  });

  // Pass 2: Source / Ministry / DC memo match + recipient
  runPass((t, d) => {
    if (d.category !== t.category) return false;
    if (!namesFuzzyMatch(d.recipientName, t.recipientName)) return false;
    if (t.unionName && d.unionName && !unionsLikelyMatch(t.unionName, d.unionName)) return false;
    return Boolean(
      (t.sourceMemo &&
        (memosLikelyMatch(d.ministryMemo, t.sourceMemo) ||
          memosLikelyMatch(d.dcMemo, t.sourceMemo) ||
          memosLikelyMatch(d.unoMemo, t.sourceMemo))) ||
        (t.unoMemo && memosLikelyMatch(d.ministryMemo, t.unoMemo)) ||
        (t.dcMemo && memosLikelyMatch(d.dcMemo, t.dcMemo))
    );
  });

  // Pass 3: Exact quantity (±0.02 MT) + month/date window (-15 to +60 days) + recipient
  runPass((t, d) => {
    if (d.category !== t.category) return false;
    if (!namesFuzzyMatch(d.recipientName, t.recipientName)) return false;
    if (t.unionName && d.unionName && !unionsLikelyMatch(t.unionName, d.unionName)) return false;
    if (Math.abs((d.allocationQty || 0) - (t.allocationQty || 0)) > 0.02) return false;
    const sDays = signedSubToDoDays(t.unoDate, d.issueDate);
    return (
      isoMonthKey(d.issueDate) === isoMonthKey(t.unoDate) ||
      (sDays >= -15 && sDays <= 60)
    );
  });

  // Pass 4: Recipient + memo (ignoring union label mismatch) OR recipient + date window (-15..+60d)
  runPass((t, d) => {
    if (d.category !== t.category) return false;
    if (!namesFuzzyMatch(d.recipientName, t.recipientName)) return false;
    const memoHit =
      (t.unoMemo && memosLikelyMatch(d.unoMemo, t.unoMemo)) ||
      (t.sourceMemo &&
        (memosLikelyMatch(d.ministryMemo, t.sourceMemo) ||
          memosLikelyMatch(d.dcMemo, t.sourceMemo) ||
          memosLikelyMatch(d.unoMemo, t.sourceMemo)));
    if (memoHit) return true;
    const sDays = signedSubToDoDays(t.unoDate, d.issueDate);
    return isoMonthKey(d.issueDate) === isoMonthKey(t.unoDate) || (sDays >= -15 && sDays <= 60);
  });

  // Pass 5: Exact union + exact qty (±0.01 MT) + tight date window (-15..+45d) when name spelled differently
  runPass((t, d) => {
    if (d.category !== t.category) return false;
    if (!t.unionName || !d.unionName || !unionsLikelyMatch(t.unionName, d.unionName)) return false;
    if (Math.abs((d.allocationQty || 0) - (t.allocationQty || 0)) > 0.01) return false;
    const sDays = signedSubToDoDays(t.unoDate, d.issueDate);
    return isoMonthKey(d.issueDate) === isoMonthKey(t.unoDate) || (sDays >= -15 && sDays <= 45);
  });

  // Pass 6: EP ('ইপি চাল' / 'ইপি গম') historical catch-all so previously issued EP DOs always resolve
  const isEp = (c: AllocationCategory) => c === 'ইপি চাল' || c === 'ইপি গম';
  runPass((t, d) => isEp(t.category) && d.category === t.category);
  runPass(
    (t, d) =>
      isEp(t.category) &&
      isEp(d.category) &&
      (Math.abs((d.allocationQty || 0) - (t.allocationQty || 0)) <= 0.02 ||
        (t.sourceMemo &&
          (memosLikelyMatch(d.ministryMemo, t.sourceMemo) || memosLikelyMatch(d.unoMemo, t.sourceMemo))) ||
        dateDaysDiff(t.unoDate, d.issueDate) <= 45)
  );

  return { dosBySubKey, remainingBySubKey };
}

function inDateRange(dateStr: string, start: string, end: string): boolean {
  if (!start && !end) return true;
  if (!dateStr) return false;
  if (start && dateStr < start) return false;
  if (end && dateStr > end) return false;
  return true;
}

export function getUnusedDoNumbers(db: AllocationDatabase): Array<{
  bookNo: number;
  doNo: number;
  startPage: number;
  endPage: number;
}> {
  const usedSet = new Set(db.doIssues.map((d) => normDoNo(d.doNo)).filter(Boolean));
  const out: Array<{ bookNo: number; doNo: number; startPage: number; endPage: number }> = [];
  const books = [...db.doBooks].sort((a, b) => a.bookNo - b.bookNo);
  for (const b of books) {
    for (let p = b.startPage; p <= b.endPage; p++) {
      if (!usedSet.has(normDoNo(p))) {
        out.push({ bookNo: b.bookNo, doNo: p, startPage: b.startPage, endPage: b.endPage });
      }
    }
  }
  return out;
}

export function getDoBookRegisterStats(db: AllocationDatabase) {
  const usedSet = new Set(db.doIssues.map((d) => normDoNo(d.doNo)).filter(Boolean));
  const books = [...db.doBooks].sort((a, b) => a.bookNo - b.bookNo);
  let totalDO = 0;
  let totalUsed = 0;

  const usedRows = books.map((b) => {
    let usedInBook = 0;
    for (let p = b.startPage; p <= b.endPage; p++) {
      if (usedSet.has(normDoNo(p))) usedInBook++;
    }
    const count = Math.max(0, b.endPage - b.startPage + 1);
    totalDO += count;
    totalUsed += usedInBook;
    return {
      bookNo: b.bookNo,
      startPage: b.startPage,
      endPage: b.endPage,
      totalCount: count,
      usedCount: usedInBook,
    };
  });

  const unusedRows = books
    .map((b) => {
      const ranges: string[] = [];
      let runStart: number | null = null;
      let runPrev = 0;
      let unusedCount = 0;
      for (let p = b.startPage; p <= b.endPage; p++) {
        if (!usedSet.has(normDoNo(p))) {
          unusedCount++;
          if (runStart === null) {
            runStart = p;
            runPrev = p;
          } else if (p === runPrev + 1) {
            runPrev = p;
          } else {
            ranges.push(runStart === runPrev ? toBnDigits(runStart) : `${toBnDigits(runStart)}-${toBnDigits(runPrev)}`);
            runStart = p;
            runPrev = p;
          }
        }
      }
      if (runStart !== null) {
        ranges.push(runStart === runPrev ? toBnDigits(runStart) : `${toBnDigits(runStart)}-${toBnDigits(runPrev)}`);
      }
      return {
        bookNo: b.bookNo,
        unusedRange: ranges.join(', ') || '—',
        unusedCount,
      };
    })
    .filter((r) => r.unusedCount > 0);

  return {
    usedRows,
    unusedRows,
    summary: {
      totalDO,
      totalUsed,
      totalUnused: Math.max(0, totalDO - totalUsed),
    },
  };
}

export function buildRegisterData(
  regId: RegisterId,
  db: AllocationDatabase,
  startDate: string,
  endDate: string,
  template?: ReportTemplateConfig
): BuiltRegisterData {
  if (regId === 'reg-challan') {
    return buildChallanRegister(db, startDate, endDate, template);
  }
  if (regId === 'reg-kernel') {
    return buildKernelRegister(db, startDate, endDate, template);
  }

  const catMeta = CATEGORIES.find((c) => c.regId === regId) || CATEGORIES[0];
  const category = catMeta.category;
  const colOverride = template?.columnLabels || {};
  const label = (key: string, def: string) => colOverride[key]?.trim() || def;

  const { dosBySubKey } = matchDOsToSuballocations(db);
  const rateMap = new Map(db.rates.map((r) => [r.category, r]));
  const subById = new Map(db.dealerSuballocations.map((s) => [s.id.trim(), s]));

  const getChallanPairBaseId = (c?: ChallanRecord): string => {
    if (!c) return '';
    const raw = (
      c.sourceRowId ||
      (!c.id.startsWith('ch-ext-') && !c.id.startsWith('ch-rec-') && !c.id.startsWith('ch-paste-')
        ? c.id
        : '')
    ).trim();
    return raw ? raw.replace(/-(TAX|MAIN)$/i, '').trim() : '';
  };

  const resolveChallanDealer = (c: ChallanRecord): string => {
    if (c.dealerName && c.dealerName !== '—') return c.dealerName;
    if (c.dealerAllocationId) {
      const sub = subById.get(c.dealerAllocationId.trim());
      if (sub && (!c.category || sub.category === c.category)) return sub.dealerName;
    }
    return '';
  };

  const reservedChallanKeysByDo = new Set<string>();
  for (const d of db.doIssues) {
    if (d.challanNo) reservedChallanKeysByDo.add(`${normChallanNo(d.challanNo)}::মূল চালান`);
    if (d.taxChallanNo) reservedChallanKeysByDo.add(`${normChallanNo(d.taxChallanNo)}::আয়কর চালান`);
  }
  const usedChallanIdsInReg = new Set<string>();
  const usedChallanKeysInReg = new Set<string>();

  const findMatchingChallanForSub = (
    subId: string,
    subDate: string,
    subQty: number,
    dealerName: string,
    doRec: DOIssue | undefined,
    challanType: 'মূল চালান' | 'আয়কর চালান',
    pairedMainBaseId?: string
  ): ChallanRecord | undefined => {
    const targetSubNorm = normMemoNo(subId);
    const isUnusedInReg = (c: ChallanRecord) =>
      !usedChallanIdsInReg.has(c.id) &&
      (!c.challanNo || !usedChallanKeysInReg.has(`${normChallanNo(c.challanNo)}::${challanType}`));

    const markUsed = (c: ChallanRecord): ChallanRecord => {
      usedChallanIdsInReg.add(c.id);
      if (c.challanNo) usedChallanKeysInReg.add(`${normChallanNo(c.challanNo)}::${challanType}`);
      return c;
    };

    if (doRec) {
      const targetChallanNo = challanType === 'মূল চালান' ? doRec.challanNo : doRec.taxChallanNo;
      const targetNorm = normChallanNo(targetChallanNo);
      const isCompatibleForDo = (c: ChallanRecord) => {
        if (c.challanType !== challanType || c.category !== category) return false;
        const effDealer = resolveChallanDealer(c);
        const sameParty = effDealer
          ? isSamePartyName(effDealer, dealerName)
          : Boolean(c.dealerAllocationId && normMemoNo(c.dealerAllocationId) === targetSubNorm);
        return (
          sameParty &&
          isChallanOnOrBeforeDo(c.challanDate, doRec.issueDate, false) &&
          isUnusedInReg(c)
        );
      };
      // Prefer the DO's explicitly pinned challanNo first so stale c.doNo on another challan never overrides it
      if (targetNorm) {
        const byPinnedNo = db.challans.find(
          (c) => normChallanNo(c.challanNo) === targetNorm && isCompatibleForDo(c)
        );
        if (byPinnedNo) return markUsed(byPinnedNo);
        return undefined;
      }
      const doNorm = normDoNo(doRec.doNo);
      const byDoNo = doNorm
        ? db.challans.find(
            (c) =>
              c.doNo &&
              normDoNo(c.doNo) === doNorm &&
              (!c.dealerAllocationId || normMemoNo(c.dealerAllocationId) === targetSubNorm) &&
              (!c.challanNo || !reservedChallanKeysByDo.has(`${normChallanNo(c.challanNo)}::${challanType}`)) &&
              isCompatibleForDo(c)
          )
        : undefined;
      if (byDoNo) return markUsed(byDoNo);
      return undefined;
    }

    // No DO issued yet: only match an UNUSED challan of the exact same category & dealer
    const bySubId = db.challans.find((c) => {
      if (c.challanType !== challanType || c.category !== category) return false;
      if (!c.dealerAllocationId || normMemoNo(c.dealerAllocationId) !== targetSubNorm) return false;
      const effDealer = resolveChallanDealer(c) || dealerName;
      return (
        isSamePartyName(effDealer, dealerName) &&
        !c.doNo &&
        (!c.challanNo || !reservedChallanKeysByDo.has(`${normChallanNo(c.challanNo)}::${challanType}`)) &&
        isUnusedInReg(c)
      );
    });
    if (bySubId) return markUsed(bySubId);

    const pool = db.challans.filter((c) => {
      if (c.challanType !== challanType || c.category !== category) return false;
      const effDealer = resolveChallanDealer(c);
      return (
        Boolean(effDealer && isSamePartyName(effDealer, dealerName)) &&
        !c.doNo &&
        (!c.challanNo || !reservedChallanKeysByDo.has(`${normChallanNo(c.challanNo)}::${challanType}`)) &&
        isUnusedInReg(c) &&
        (!c.dealerAllocationId || normMemoNo(c.dealerAllocationId) === targetSubNorm)
      );
    });
    if (pool.length === 0) return undefined;

    const withinWindow = pool.filter((c) => {
      const cDate = c.challanDate;
      if (!cDate || !subDate) return true;
      return dateDaysDiff(cDate, subDate) <= 60;
    });
    const searchPool = withinWindow.length > 0 ? withinWindow : pool.length === 1 && !subDate ? pool : [];
    if (searchPool.length === 0) return undefined;

    const picked = [...searchPool].sort((a, b) => {
      if (pairedMainBaseId) {
        const pairA = getChallanPairBaseId(a) === pairedMainBaseId ? 0 : 1;
        const pairB = getChallanPairBaseId(b) === pairedMainBaseId ? 0 : 1;
        if (pairA !== pairB) return pairA - pairB;
      }
      const aIso = toIsoDateStr(a.challanDate) || '9999-99-99';
      const bIso = toIsoDateStr(b.challanDate) || '9999-99-99';
      if (aIso !== bIso) return aIso.localeCompare(bIso);
      const qa = a.allocationQty ? Math.abs(a.allocationQty - subQty) : 999;
      const qb = b.allocationQty ? Math.abs(b.allocationQty - subQty) : 999;
      if (Math.abs(qa - qb) > 0.01) return qa - qb;
      return bnToEnDigits(a.challanNo || '').localeCompare(bnToEnDigits(b.challanNo || ''), undefined, { numeric: true });
    })[0];
    return picked ? markUsed(picked) : undefined;
  };

  const formatDoCell = (doList: DOIssue[]): string => {
    if (doList.length === 0) return '—';
    return doList
      .map(
        (d) =>
          `ডিও নং- ${toBnDigits(d.doNo)}${d.issueDate ? `\nতাং- ${formatBnDate(d.issueDate)}` : ''}`
      )
      .join('\n');
  };

  const catMinAllocs = db.ministryAllocations.filter((m) => m.category === category);
  const catDcAllocs = db.dcAllocations.filter((d) => d.category === category);

  const findMinRec = (rawMemo?: string) => {
    const t = (rawMemo || '').trim();
    if (!t) return undefined;
    const nt = normMemoNo(t);
    return (
      catMinAllocs.find((m) => m.memoNo === t) ||
      (nt ? catMinAllocs.find((m) => normMemoNo(m.memoNo) === nt) : undefined) ||
      catMinAllocs.find((m) => memosLikelyMatch(m.memoNo, t))
    );
  };

  const findDcRec = (rawMemo?: string) => {
    const t = (rawMemo || '').trim();
    if (!t) return undefined;
    const nt = normMemoNo(t);
    return (
      catDcAllocs.find((d) => d.dcMemo === t) ||
      (nt ? catDcAllocs.find((d) => normMemoNo(d.dcMemo) === nt) : undefined) ||
      catDcAllocs.find((d) => memosLikelyMatch(d.dcMemo, t))
    );
  };

  const fyKeyOf = (dateStr?: string): string => {
    const parsed = dateStr ? parseMonthYear(dateStr) : null;
    return parsed ? fiscalYearLabel(parsed.year, parsed.month) : '';
  };

  const computeTierRemainingMap = <T extends { id: string; allocationQty: number }>(
    upstreams: Array<{ rec: T; date: string; memo: string }>,
    downstreams: Array<{ memoKey: string; qty: number; date: string }>
  ): Map<string, number> => {
    const directUsed = new Map<string, number>();
    const unlinked: Array<{ qty: number; date: string }> = [];

    for (const item of downstreams) {
      if (item.qty <= 0) continue;
      const key = (item.memoKey || '').trim();
      let matched: T | undefined;
      if (key) {
        const nk = normMemoNo(key);
        matched =
          upstreams.find((u) => normMemoNo(u.memo) === nk)?.rec ||
          upstreams.find((u) => memosLikelyMatch(u.memo, key))?.rec;
      }
      if (matched) {
        directUsed.set(matched.id, round3((directUsed.get(matched.id) || 0) + item.qty));
      } else {
        unlinked.push({ qty: item.qty, date: item.date });
      }
    }

    const sortedUpstreams = [...upstreams].sort((a, b) =>
      (a.date || '').localeCompare(b.date || '')
    );
    const remCap = new Map<string, number>();
    const fifoUsed = new Map<string, number>();
    for (const u of sortedUpstreams) {
      const total = Number(u.rec.allocationQty) || 0;
      const used = directUsed.get(u.rec.id) || 0;
      remCap.set(u.rec.id, Math.max(0, round3(total - used)));
      fifoUsed.set(u.rec.id, 0);
    }

    for (const item of unlinked) {
      let left = item.qty;
      const itemFy = fyKeyOf(item.date);
      const passList = itemFy
        ? [
            sortedUpstreams.filter((u) => {
              const uFy = fyKeyOf(u.date);
              return !uFy || uFy === itemFy;
            }),
            sortedUpstreams,
          ]
        : [sortedUpstreams];

      for (const candidates of passList) {
        if (left <= 0.0005) break;
        for (const u of candidates) {
          if (left <= 0.0005) break;
          const cap = remCap.get(u.rec.id) || 0;
          if (cap <= 0.0005) continue;
          const take = Math.min(cap, left);
          remCap.set(u.rec.id, round3(cap - take));
          fifoUsed.set(u.rec.id, round3((fifoUsed.get(u.rec.id) || 0) + take));
          left = round3(left - take);
        }
      }
    }

    const out = new Map<string, number>();
    for (const u of upstreams) {
      const total = round3(Number(u.rec.allocationQty) || 0);
      const used = round3((directUsed.get(u.rec.id) || 0) + (fifoUsed.get(u.rec.id) || 0));
      out.set(u.rec.id, round3(Math.max(0, total - used)));
    }
    return out;
  };

  const minDownstreams = catMeta.requiresDC
    ? catDcAllocs.map((d) => ({
        memoKey: d.ministryMemo,
        qty: Number(d.allocationQty) || 0,
        date: d.dcMemoDate || d.createdAt,
      }))
    : catMeta.flowType === 'vwb'
      ? db.unoSuballocations
          .filter((u) => u.category === category)
          .map((u) => ({
            memoKey: u.ministryMemo || u.dcMemo,
            qty: Number(u.allocationQty) || 0,
            date: u.unoDate || u.createdAt,
          }))
      : db.dealerSuballocations
          .filter((s) => s.category === category)
          .map((s) => ({
            memoKey: s.sourceMemo,
            qty: Number(s.allocationQty) || 0,
            date: s.unoDate || s.createdAt,
          }));

  const minRemainingById = computeTierRemainingMap(
    catMinAllocs.map((m) => ({ rec: m, date: m.memoDate || m.createdAt, memo: m.memoNo })),
    minDownstreams
  );

  const dcDownstreams =
    catMeta.flowType === 'dealer_oms'
      ? db.dealerSuballocations
          .filter((s) => s.category === category)
          .map((s) => ({
            memoKey: s.sourceMemo,
            qty: Number(s.allocationQty) || 0,
            date: s.unoDate || s.createdAt,
          }))
      : db.unoSuballocations
          .filter((u) => u.category === category)
          .map((u) => ({
            memoKey: u.dcMemo,
            qty: Number(u.allocationQty) || 0,
            date: u.unoDate || u.createdAt,
          }));

  const dcRemainingById = computeTierRemainingMap(
    catDcAllocs.map((d) => ({ rec: d, date: d.dcMemoDate || d.createdAt, memo: d.dcMemo })),
    dcDownstreams
  );

  const catTargets =
    catMeta.flowType === 'dealer_food' ||
    catMeta.flowType === 'dealer_oms' ||
    catMeta.flowType === 'ep'
      ? db.dealerSuballocations
          .filter((s) => s.category === category)
          .map((s) => {
            const dcRec = catMeta.requiresDC ? findDcRec(s.sourceMemo) : undefined;
            const rawMin = dcRec ? dcRec.ministryMemo : s.sourceMemo;
            const minRec = findMinRec(rawMin);
            return {
              subKey: `ds:${s.id}`,
              minMemo: minRec?.memoNo || rawMin,
              dcMemo: dcRec ? dcRec.dcMemo : catMeta.requiresDC ? s.sourceMemo : '',
              unoMemo: s.unoMemo || s.sourceMemo || 'ইউসিএফ',
              unoDate: s.unoDate,
              qty: Number(s.allocationQty) || 0,
            };
          })
      : db.unoSuballocations
          .filter((u) => u.category === category)
          .map((u) => {
            const dcRec = catMeta.requiresDC ? findDcRec(u.dcMemo) : undefined;
            const rawMin = dcRec?.ministryMemo || u.ministryMemo;
            const minRec = findMinRec(rawMin);
            return {
              subKey: `uno:${u.id}`,
              minMemo: minRec?.memoNo || rawMin,
              dcMemo: dcRec?.dcMemo || u.dcMemo || '',
              unoMemo: u.unoMemo || 'ইউএনও',
              unoDate: u.unoDate,
              qty: Number(u.allocationQty) || 0,
            };
          });

  const unoGroupKey = (minMemo: string, dcMemo: string, unoMemo: string, unoDate: string) =>
    `${normMemoNo(minMemo)}::${normMemoNo(dcMemo)}::${normMemoNo(unoMemo)}::${unoDate}`;
  const unoTotalsMap = new Map<string, { alloc: number; issued: number }>();

  for (const t of catTargets) {
    const matchedDos = dosBySubKey.get(t.subKey) || [];
    const subDoQty = round3(matchedDos.reduce((s, d) => s + (Number(d.allocationQty) || 0), 0));
    const uKey = unoGroupKey(t.minMemo, t.dcMemo, t.unoMemo, t.unoDate);
    const cur = unoTotalsMap.get(uKey) || { alloc: 0, issued: 0 };
    cur.alloc = round3(cur.alloc + t.qty);
    cur.issued = round3(cur.issued + subDoQty);
    unoTotalsMap.set(uKey, cur);
  }

  const buildMinistryCellText = (
    minMemoFallback: string,
    grainFallback?: string
  ): string => {
    const minRec = findMinRec(minMemoFallback);
    if (minRec) {
      const remaining = minRemainingById.get(minRec.id) ?? 0;
      return [
        toBnDigits(minRec.memoNo),
        `তাং- ${formatBnDate(minRec.memoDate)}`,
        minRec.phase ? `পর্যায়-${toBnDigits(minRec.phase)}` : '',
        `মোট বরাদ্দ: ${formatBnQty(minRec.allocationQty)} মে.টন (${minRec.grain || grainFallback || catMeta.defaultGrain || 'চাল'})`,
        `অবশিষ্ট: ${formatBnQty(remaining)} মে.টন`,
      ]
        .filter(Boolean)
        .join('\n');
    }
    if (!minMemoFallback) return '—';
    const subRowsForMemo = catTargets.filter((t) => memosLikelyMatch(t.minMemo, minMemoFallback));
    if (subRowsForMemo.length > 0) {
      const totAlloc = round3(subRowsForMemo.reduce((s, t) => s + t.qty, 0));
      return [
        toBnDigits(minMemoFallback),
        `মোট বরাদ্দ: ${formatBnQty(totAlloc)} মে.টন`,
        `অবশিষ্ট: ০.০০০ মে.টন`,
      ].join('\n');
    }
    return toBnDigits(minMemoFallback);
  };

  const buildDcCellText = (dcMemoFallback: string): string => {
    const dcRec = findDcRec(dcMemoFallback);
    if (dcRec) {
      const remaining = dcRemainingById.get(dcRec.id) ?? 0;
      return [
        toBnDigits(dcRec.dcMemo),
        `তাং- ${formatBnDate(dcRec.dcMemoDate)}`,
        `মোট বরাদ্দ: ${formatBnQty(dcRec.allocationQty)} মে.টন`,
        `অবশিষ্ট: ${formatBnQty(remaining)} মে.টন`,
      ].join('\n');
    }
    if (!dcMemoFallback) return '—';
    const subRowsForDc = catTargets.filter((t) => memosLikelyMatch(t.dcMemo, dcMemoFallback));
    if (subRowsForDc.length > 0) {
      const totAlloc = round3(subRowsForDc.reduce((s, t) => s + t.qty, 0));
      return [
        toBnDigits(dcMemoFallback),
        `মোট বরাদ্দ: ${formatBnQty(totAlloc)} মে.টন`,
        `অবশিষ্ট: ০.০০০ মে.টন`,
      ].join('\n');
    }
    return toBnDigits(dcMemoFallback);
  };

  const buildUnoCellText = (
    minMemo: string,
    dcMemo: string,
    unoMemoLabel: string,
    unoDate: string
  ): string => {
    const uKey = unoGroupKey(minMemo, dcMemo, unoMemoLabel, unoDate);
    const stat = unoTotalsMap.get(uKey) || { alloc: 0, issued: 0 };
    const remaining = round3(Math.max(0, stat.alloc - stat.issued));
    return [
      toBnDigits(unoMemoLabel),
      `তাং- ${formatBnDate(unoDate)}`,
      `মোট বরাদ্দ: ${formatBnQty(stat.alloc)} মে.টন`,
      `অবশিষ্ট: ${formatBnQty(remaining)} মে.টন`,
    ].join('\n');
  };

  const compareNumericText = (a: string, b: string) =>
    bnToEnDigits(a || '').localeCompare(bnToEnDigits(b || ''), undefined, { numeric: true });

  if (catMeta.flowType === 'dealer_food' || catMeta.flowType === 'dealer_oms') {
    const isFood = catMeta.flowType === 'dealer_food';
    const isTCB = catMeta.requiresDC;
    const unionOrder = db.officer.unions || [];

    const columns: RegisterColumnGroup[] = [
      { label: label('ministry', 'মন্ত্রণালয় বরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'ministry', label: label('ministry', 'মন্ত্রণালয় বরাদ্দ'), align: 'left' } },
      ...(isTCB
        ? [{ label: label('dc', 'ডিসি বরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'dc', label: label('dc', 'ডিসি বরাদ্দ'), align: 'left' as const } }]
        : []),
      { label: label('uno', 'ইউসিএফ উপবরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'uno', label: label('uno', 'ইউসিএফ উপবরাদ্দ'), align: 'left' } },
      { label: label('union', 'ইউনিয়ন / এলাকা'), leaf: { key: 'union', label: label('union', 'ইউনিয়ন / এলাকা'), align: 'left' } },
      { label: label('dealer', 'ডিলার ও বিক্রয় কেন্দ্র'), leaf: { key: 'dealer', label: label('dealer', 'ডিলার ও বিক্রয় কেন্দ্র'), align: 'left' } },
      ...(isFood
        ? [{ label: label('cards', 'কার্ড (মোট / অবিক্রিত)'), leaf: { key: 'cards', label: label('cards', 'কার্ড'), align: 'center' as const } }]
        : []),
      { label: label('qty', 'পরিমাণ (মে.টন)'), leaf: { key: 'qty', label: label('qty', 'পরিমাণ (মে.টন)'), align: 'right' } },
      { label: label('money', 'মূল চালান ও টাকা'), leaf: { key: 'money', label: label('money', 'মূল চালান ও টাকা'), align: 'left' } },
      { label: label('challan', 'আয়কর চালান ও টাকা'), leaf: { key: 'challan', label: label('challan', 'আয়কর চালান ও টাকা'), align: 'left' } },
      { label: label('doInfo', 'ডিও নং ও তারিখ'), leaf: { key: 'doInfo', label: label('doInfo', 'ডিও নং ও তারিখ'), align: 'center' } },
      { label: label('signature', 'স্বাক্ষর'), leaf: { key: 'signature', label: label('signature', 'স্বাক্ষর'), align: 'center', isSignature: true } },
    ];

    const subs = db.dealerSuballocations
      .filter((s) => s.category === category && inDateRange(s.unoDate, startDate, endDate))
      .map((s) => {
        const dcRec = isTCB ? findDcRec(s.sourceMemo) : undefined;
        const rawMin = dcRec ? dcRec.ministryMemo : s.sourceMemo;
        const minRec = findMinRec(rawMin);
        return {
          s,
          dcRec,
          minRec,
          minMemoResolved: minRec?.memoNo || rawMin || '',
          dcMemoResolved: dcRec?.dcMemo || (isTCB ? s.sourceMemo : ''),
        };
      })
      .sort((a, b) => {
        const minDateCmp = (a.minRec?.memoDate || '9999-99-99').localeCompare(
          b.minRec?.memoDate || '9999-99-99'
        );
        if (minDateCmp !== 0) return minDateCmp;
        const minMemoCmp = compareNumericText(a.minMemoResolved, b.minMemoResolved);
        if (minMemoCmp !== 0) return minMemoCmp;
        if (isTCB) {
          const dcDateCmp = (a.dcRec?.dcMemoDate || '9999-99-99').localeCompare(
            b.dcRec?.dcMemoDate || '9999-99-99'
          );
          if (dcDateCmp !== 0) return dcDateCmp;
          const dcMemoCmp = compareNumericText(a.dcMemoResolved, b.dcMemoResolved);
          if (dcMemoCmp !== 0) return dcMemoCmp;
        }
        const dateCmp = a.s.unoDate.localeCompare(b.s.unoDate);
        if (dateCmp !== 0) return dateCmp;
        const memoCmp = compareNumericText(a.s.unoMemo || a.s.sourceMemo, b.s.unoMemo || b.s.sourceMemo);
        if (memoCmp !== 0) return memoCmp;
        const ia = unionOrder.findIndex((u) => unionsLikelyMatch(u, a.s.unionName));
        const ib = unionOrder.findIndex((u) => unionsLikelyMatch(u, b.s.unionName));
        if (ia !== -1 && ib !== -1 && ia !== ib) return ia - ib;
        if (ia !== -1 && ib === -1) return -1;
        if (ib !== -1 && ia === -1) return 1;
        return a.s.dealerName.localeCompare(b.s.dealerName);
      });

    const catRate = rateMap.get(category);
    const effUnitPrice =
      catRate && catRate.unitPrice > 0 ? catRate.unitPrice : category === 'খাদ্য বান্ধব' ? 13000 : 0;
    const effTaxRate =
      catRate && catRate.taxRate > 0 ? catRate.taxRate : category === 'খাদ্য বান্ধব' ? 200 : 0;

    const rows: RegisterRowData[] = subs.map(({ s, minRec, minMemoResolved, dcMemoResolved }) => {
      const unoMemoLabel = s.unoMemo || s.sourceMemo || 'ইউসিএফ';

      const minText = buildMinistryCellText(minMemoResolved);
      const dcText = isTCB ? buildDcCellText(dcMemoResolved || s.sourceMemo) : '—';
      const unoText = buildUnoCellText(minMemoResolved, dcMemoResolved, unoMemoLabel, s.unoDate);

      const matchedDos = dosBySubKey.get(`ds:${s.id}`) || [];
      const doRec = matchedDos[0];
      const hasDo = matchedDos.some((d) => Boolean(d.doNo && String(d.doNo).trim()));
      const isKernelRow = minRec?.grain === 'কার্নেল';

      const matchedMainChallanObj = isKernelRow
        ? undefined
        : findMatchingChallanForSub(
            s.id,
            s.unoDate,
            s.allocationQty,
            s.dealerName,
            doRec,
            'মূল চালান'
          );
      const matchedTaxChallanObj = isKernelRow
        ? undefined
        : findMatchingChallanForSub(
            s.id,
            s.unoDate,
            s.allocationQty,
            s.dealerName,
            doRec,
            'আয়কর চালান',
            getChallanPairBaseId(matchedMainChallanObj) || undefined
          );
      const matchedMainChallan = isKernelRow ? '' : doRec?.challanNo || matchedMainChallanObj?.challanNo || '';
      const matchedTaxChallan = isKernelRow ? '' : doRec?.taxChallanNo || matchedTaxChallanObj?.challanNo || '';
      const mainDate = isKernelRow ? '' : doRec?.challanDate || matchedMainChallanObj?.challanDate || '';
      const taxDate = isKernelRow ? '' : doRec?.taxChallanDate || matchedTaxChallanObj?.challanDate || '';
      const mainDep = isKernelRow ? 0 : matchedMainChallanObj?.depositedAmount || doRec?.challanAmount || 0;
      const taxDep = isKernelRow ? 0 : matchedTaxChallanObj?.depositedAmount || doRec?.taxChallanAmount || 0;

      const expectedMain = isKernelRow
        ? 0
        : s.totalAmount > 0
          ? s.totalAmount
          : round2(s.allocationQty * effUnitPrice);
      const expectedTax = isKernelRow
        ? 0
        : s.taxAmount > 0
          ? s.taxAmount
          : round2(s.allocationQty * effTaxRate);

      const mainChallanLines = isKernelRow
        ? ['প্রযোজ্য নয়', '(কার্নেল)']
        : [
            matchedMainChallan ? `নং- ${toBnDigits(matchedMainChallan)}` : '—',
            mainDate ? `(${formatBnDate(mainDate)})` : '',
            mainDep > 0
              ? `টাকা: ${formatBnMoney(mainDep)} ৳`
              : expectedMain > 0
                ? `প্রাপ্য: ${formatBnMoney(expectedMain)} ৳`
                : '',
          ].filter(Boolean);

      const taxChallanLines = isKernelRow
        ? ['প্রযোজ্য নয়', '(কার্নেল)']
        : [
            matchedTaxChallan ? `নং- ${toBnDigits(matchedTaxChallan)}` : '—',
            taxDate ? `(${formatBnDate(taxDate)})` : '',
            taxDep > 0
              ? `টাকা: ${formatBnMoney(taxDep)} ৳`
              : expectedTax > 0
                ? `প্রাপ্য: ${formatBnMoney(expectedTax)} ৳`
                : '',
          ].filter(Boolean);

      const groupValues = isTCB ? [minText, dcText, unoText] : [minText, unoText];

      return {
        id: s.id,
        groupValues,
        cells: {
          ministry: minText,
          ...(isTCB ? { dc: dcText } : {}),
          uno: unoText,
          union: toBnDigits(s.unionName || '—'),
          dealer: `${s.dealerName}${s.salesCenter ? `\n(${s.salesCenter})` : ''}`,
          ...(isFood
            ? {
                cards: `${formatBnInt(s.cardCount)}টি${s.unusedCards > 0 ? `\nঅবিক্রিত: ${formatBnInt(s.unusedCards)}` : ''}`,
              }
            : {}),
          qty: isKernelRow ? `${formatBnQty(s.allocationQty)}\n(কার্নেল)` : formatBnQty(s.allocationQty),
          money: mainChallanLines.join('\n') || '—',
          challan: taxChallanLines.join('\n') || '—',
          doInfo: formatDoCell(matchedDos),
          signature: '',
        },
        hasSignature: hasDo,
      };
    });

    const leaves = columns.map((c) => c.leaf!);
    return {
      id: regId,
      title: `${category} রেজিস্টার`,
      category,
      columns,
      leaves,
      mergeDepth: isTCB ? 3 : 2,
      serialPerRow: true,
      colWidths: computeRegisterColWidths(leaves, true, catMeta.flowType),
      tableLayoutFixed: true,
      rows,
    };
  }

  if (catMeta.flowType === 'ep') {
    const columns: RegisterColumnGroup[] = [
      { label: label('ministry', 'বিভাগীয় বরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'ministry', label: label('ministry', 'বিভাগীয় বরাদ্দ'), align: 'left' } },
      { label: label('uno', 'ইউসিএফ উপবরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'uno', label: label('uno', 'ইউসিএফ উপবরাদ্দ'), align: 'left' } },
      { label: label('dealer', 'প্রাপক (ফায়ার সার্ভিস)'), leaf: { key: 'dealer', label: label('dealer', 'প্রাপক'), align: 'left' } },
      { label: label('grain', 'শস্য'), leaf: { key: 'grain', label: label('grain', 'শস্য'), align: 'center' } },
      { label: label('qty', 'পরিমাণ (মে.টন)'), leaf: { key: 'qty', label: label('qty', 'পরিমাণ (মে.টন)'), align: 'right' } },
      { label: label('money', 'মোট টাকা (৳)'), leaf: { key: 'money', label: label('money', 'মোট টাকা (৳)'), align: 'right' } },
      { label: label('challan', 'চালান নং ও তারিখ'), leaf: { key: 'challan', label: label('challan', 'চালান'), align: 'left' } },
      { label: label('doInfo', 'ডিও নং ও তারিখ'), leaf: { key: 'doInfo', label: label('doInfo', 'ডিও নং ও তারিখ'), align: 'center' } },
      { label: label('signature', 'স্বাক্ষর'), leaf: { key: 'signature', label: label('signature', 'স্বাক্ষর'), align: 'center', isSignature: true } },
    ];

    const subs = db.dealerSuballocations
      .filter((s) => s.category === category && inDateRange(s.unoDate, startDate, endDate))
      .map((s) => {
        const minRec = findMinRec(s.sourceMemo);
        return { s, minRec, minMemoResolved: minRec?.memoNo || s.sourceMemo || '' };
      })
      .sort((a, b) => {
        const minDateCmp = (a.minRec?.memoDate || '9999-99-99').localeCompare(
          b.minRec?.memoDate || '9999-99-99'
        );
        if (minDateCmp !== 0) return minDateCmp;
        const minMemoCmp = compareNumericText(a.minMemoResolved, b.minMemoResolved);
        if (minMemoCmp !== 0) return minMemoCmp;
        const dc = a.s.unoDate.localeCompare(b.s.unoDate);
        if (dc !== 0) return dc;
        return compareNumericText(a.s.unoMemo || a.s.sourceMemo, b.s.unoMemo || b.s.sourceMemo);
      });

    const rows: RegisterRowData[] = subs.map(({ s, minRec, minMemoResolved }) => {
      const unoMemoLabel = s.unoMemo || s.sourceMemo || 'ইউসিএফ';
      const minText = buildMinistryCellText(minMemoResolved, catMeta.defaultGrain);
      const unoText = buildUnoCellText(minMemoResolved, '', unoMemoLabel, s.unoDate);

      const matchedDos = dosBySubKey.get(`ds:${s.id}`) || [];
      const doRec = matchedDos[0];
      const hasDo = matchedDos.some((d) => Boolean(d.doNo && String(d.doNo).trim()));
      const isKernelRow = minRec?.grain === 'কার্নেল';

      const epChallanObj = isKernelRow
        ? undefined
        : findMatchingChallanForSub(
            s.id,
            s.unoDate,
            s.allocationQty,
            s.dealerName,
            doRec,
            'মূল চালান'
          );
      const epChallanNo = isKernelRow ? '' : doRec?.challanNo || epChallanObj?.challanNo || '';
      const epChallanDate = isKernelRow ? '' : doRec?.challanDate || epChallanObj?.challanDate || '';
      const epDep = isKernelRow ? 0 : epChallanObj?.depositedAmount ?? doRec?.challanAmount ?? 0;

      return {
        id: s.id,
        groupValues: [minText, unoText],
        cells: {
          ministry: minText,
          uno: unoText,
          dealer: s.dealerName,
          grain: minRec?.grain || catMeta.defaultGrain || 'চাল',
          qty: formatBnQty(s.allocationQty),
          money: isKernelRow ? 'প্রযোজ্য নয়' : formatBnMoney(s.totalAmount || epDep),
          challan: isKernelRow
            ? 'প্রযোজ্য নয়\n(কার্নেল)'
            : epChallanNo || epDep > 0
              ? `${epChallanNo ? toBnDigits(epChallanNo) : '—'}${epChallanDate ? `\n${formatBnDate(epChallanDate)}` : ''}${epDep > 0 ? `\nজমা: ${formatBnMoney(epDep)} ৳` : ''}`
              : '—',
          doInfo: formatDoCell(matchedDos),
          signature: '',
        },
        hasSignature: hasDo,
      };
    });

    const leaves = columns.map((c) => c.leaf!);
    return {
      id: regId,
      title: `${category} রেজিস্টার`,
      category,
      columns,
      leaves,
      mergeDepth: 2,
      serialPerRow: true,
      colWidths: computeRegisterColWidths(leaves, true),
      tableLayoutFixed: true,
      rows,
    };
  }

  const showDC = catMeta.requiresDC;
  const showProject = catMeta.supportsProjects;
  const unionOrder = db.officer.unions || [];
  const projectColTitle = label(
    'project',
    'প্রকল্প নং, প্রকল্পের নাম,\nসভাপতি ও সেক্রেটারি'
  );
  const columns: RegisterColumnGroup[] = [
    { label: label('ministry', 'মন্ত্রণালয় বরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'ministry', label: label('ministry', 'মন্ত্রণালয় বরাদ্দ'), align: 'left' } },
    ...(showDC
      ? [{ label: label('dc', 'ডিসি বরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'dc', label: label('dc', 'ডিসি বরাদ্দ'), align: 'left' as const } }]
      : []),
    { label: label('uno', 'ইউএনও উপবরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'), leaf: { key: 'uno', label: label('uno', 'ইউএনও উপবরাদ্দ'), align: 'left' } },
    ...(showProject
      ? [
          {
            label: projectColTitle,
            leaf: {
              key: 'project',
              label: projectColTitle,
              align: 'left' as const,
            },
          },
        ]
      : []),
    { label: label('union', 'ইউনিয়ন'), leaf: { key: 'union', label: label('union', 'ইউনিয়ন'), align: 'left' } },
    ...(!showProject
      ? [
          {
            label: label('recipient', 'চেয়ারম্যান / প্রাপক'),
            leaf: { key: 'recipient', label: label('recipient', 'প্রাপক'), align: 'left' as const },
          },
        ]
      : []),
    { label: label('grain', 'শস্য'), leaf: { key: 'grain', label: label('grain', 'শস্য'), align: 'center' } },
    { label: label('qty', 'পরিমাণ (মে.টন)'), leaf: { key: 'qty', label: label('qty', 'পরিমাণ (মে.টন)'), align: 'right' } },
    { label: label('doInfo', 'ডিও নং ও তারিখ'), leaf: { key: 'doInfo', label: label('doInfo', 'ডিও নং ও তারিখ'), align: 'center' } },
    { label: label('signature', 'স্বাক্ষর'), leaf: { key: 'signature', label: label('signature', 'স্বাক্ষর'), align: 'center', isSignature: true } },
  ];

  const unoSubs = db.unoSuballocations
    .filter((u) => u.category === category && inDateRange(u.unoDate, startDate, endDate))
    .map((u) => {
      const dcRec = showDC ? findDcRec(u.dcMemo) : undefined;
      const rawMin = dcRec?.ministryMemo || u.ministryMemo;
      const minRec = findMinRec(rawMin);
      return {
        u,
        dcRec,
        minRec,
        minMemoResolved: minRec?.memoNo || rawMin || '',
        dcMemoResolved: dcRec?.dcMemo || u.dcMemo || '',
      };
    })
    .sort((a, b) => {
      const minDateCmp = (a.minRec?.memoDate || '9999-99-99').localeCompare(
        b.minRec?.memoDate || '9999-99-99'
      );
      if (minDateCmp !== 0) return minDateCmp;
      const minMemoCmp = compareNumericText(a.minMemoResolved, b.minMemoResolved);
      if (minMemoCmp !== 0) return minMemoCmp;
      if (showDC) {
        const dcDateCmp = (a.dcRec?.dcMemoDate || '9999-99-99').localeCompare(
          b.dcRec?.dcMemoDate || '9999-99-99'
        );
        if (dcDateCmp !== 0) return dcDateCmp;
        const dcMemoCmp = compareNumericText(a.dcMemoResolved, b.dcMemoResolved);
        if (dcMemoCmp !== 0) return dcMemoCmp;
      }
      const dateCmp = a.u.unoDate.localeCompare(b.u.unoDate);
      if (dateCmp !== 0) return dateCmp;
      const memoCmp = compareNumericText(a.u.unoMemo, b.u.unoMemo);
      if (memoCmp !== 0) return memoCmp;
      const ia = unionOrder.findIndex((un) => unionsLikelyMatch(un, a.u.unionName));
      const ib = unionOrder.findIndex((un) => unionsLikelyMatch(un, b.u.unionName));
      if (ia !== -1 && ib !== -1 && ia !== ib) return ia - ib;
      if (ia !== -1 && ib === -1) return -1;
      if (ib !== -1 && ia === -1) return 1;
      return compareNumericText(
        a.u.projectNo || a.u.recipientName || '',
        b.u.projectNo || b.u.recipientName || ''
      );
    });

  const rows: RegisterRowData[] = unoSubs.map(({ u, minRec, minMemoResolved, dcMemoResolved }) => {
    const minText = buildMinistryCellText(minMemoResolved, catMeta.defaultGrain);
    const dcText = showDC ? buildDcCellText(dcMemoResolved) : '—';
    const unoText = buildUnoCellText(
      minMemoResolved,
      dcMemoResolved,
      u.unoMemo || 'ইউএনও',
      u.unoDate
    );

    const matchedDos = dosBySubKey.get(`uno:${u.id}`) || [];
    const hasDo = matchedDos.some((d) => Boolean(d.doNo && String(d.doNo).trim()));
    const groupValues = showDC ? [minText, dcText, unoText] : [minText, unoText];

    let projectCell = '';
    if (showProject) {
      const matchedProject = db.projects.find((p) => {
        if (p.category !== category) return false;
        if (u.projectNo && p.projectNo && normMemoNo(p.projectNo) === normMemoNo(u.projectNo)) {
          return !u.dcMemo || !p.dcMemo || normMemoNo(p.dcMemo) === normMemoNo(u.dcMemo);
        }
        if (
          u.dcMemo &&
          p.dcMemo &&
          normMemoNo(p.dcMemo) === normMemoNo(u.dcMemo) &&
          (!u.unionName || !p.unionName || unionsLikelyMatch(u.unionName, p.unionName)) &&
          (namesLikelyMatch(p.presidentName, u.recipientName) ||
            namesLikelyMatch(p.projectName, u.projectName || '') ||
            namesLikelyMatch(p.projectName, u.recipientName))
        ) {
          return true;
        }
        return false;
      });

      const pNo = (u.projectNo || matchedProject?.projectNo || '').trim();
      const pName = (u.projectName || matchedProject?.projectName || '').trim();
      let pPresident = (u.recipientName || matchedProject?.presidentName || '').trim();
      if (pName && pPresident === pName && matchedProject?.presidentName?.trim()) {
        pPresident = matchedProject.presidentName.trim();
      }

      const lines: string[] = [];
      if (pNo) {
        const bnNo = toBnDigits(pNo);
        lines.push(/^(প্রকল্প|প্র[:.\s-])/i.test(pNo) ? bnNo : `প্র: নং- ${bnNo}`);
      }
      if (pName) {
        lines.push(pName);
      }
      if (pPresident && pPresident !== pName) {
        lines.push(lines.length > 0 ? `সভাপতি/সেক্রেটারি: ${pPresident}` : pPresident);
      }
      projectCell = lines.length > 0 ? lines.join('\n') : '—';
    }

    return {
      id: u.id,
      groupValues,
      cells: {
        ministry: minText,
        ...(showDC ? { dc: dcText } : {}),
        uno: unoText,
        ...(showProject ? { project: projectCell } : {}),
        union: u.unionName,
        ...(!showProject ? { recipient: u.recipientName } : {}),
        grain: minRec?.grain || catMeta.defaultGrain || 'চাল',
        qty: formatBnQty(u.allocationQty),
        doInfo: formatDoCell(matchedDos),
        signature: '',
      },
      hasSignature: hasDo,
    };
  });

  const leaves = columns.map((c) => c.leaf!);
  return {
    id: regId,
    title: `${category} রেজিস্টার`,
    category,
    columns,
    leaves,
    mergeDepth: showDC ? 3 : 2,
    serialPerRow: true,
    colWidths: computeRegisterColWidths(leaves, true),
    tableLayoutFixed: true,
    rows,
  };
}

function buildChallanRegister(
  db: AllocationDatabase,
  startDate: string,
  endDate: string,
  template?: ReportTemplateConfig
): BuiltRegisterData {
  const colOverride = template?.columnLabels || {};
  const label = (key: string, def: string) => colOverride[key]?.trim() || def;
  const signRule = template?.challanSignRule || 'excess_or_equal';

  const columns: RegisterColumnGroup[] = [
    { label: label('doNo', 'ডিও নং ও তারিখ'), leaf: { key: 'doNo', label: label('doNo', 'ডিও নং ও তারিখ'), align: 'center' } },
    { label: label('dealer', 'ডিলার / প্রাপকের নাম'), leaf: { key: 'dealer', label: label('dealer', 'ডিলার / প্রাপক'), align: 'left' } },
    { label: label('category', 'খাত'), leaf: { key: 'category', label: label('category', 'খাত'), align: 'left' } },
    { label: label('type', 'চালানের ধরন'), leaf: { key: 'type', label: label('type', 'ধরন'), align: 'center' } },
    { label: label('date', 'চালানের তারিখ'), leaf: { key: 'date', label: label('date', 'তারিখ'), align: 'center' } },
    { label: label('challanNo', 'চালান নং ও ব্যাংক'), leaf: { key: 'challanNo', label: label('challanNo', 'চালান নং ও ব্যাংক'), align: 'left' } },
    { label: label('recorded', 'রেকর্ড অনুযায়ী টাকা (৳)'), leaf: { key: 'recorded', label: label('recorded', 'রেকর্ড অনুযায়ী টাকা'), align: 'right' } },
    { label: label('deposited', 'জমাকৃত টাকা (৳)'), leaf: { key: 'deposited', label: label('deposited', 'জমাকৃত টাকা'), align: 'right' } },
    { label: label('mismatch', 'গরমিল / পার্থক্য (৳)'), leaf: { key: 'mismatch', label: label('mismatch', 'গরমিল (৳)'), align: 'right' } },
    { label: label('signature', 'স্বাক্ষর'), leaf: { key: 'signature', label: label('signature', 'স্বাক্ষর'), align: 'center', isSignature: true } },
  ];

  const getChallanPairId = (c?: ChallanRecord): string => {
    if (!c) return '';
    const raw = (
      c.sourceRowId ||
      (!c.id.startsWith('ch-ext-') && !c.id.startsWith('ch-rec-') && !c.id.startsWith('ch-paste-')
        ? c.id
        : '')
    ).trim();
    return raw ? raw.replace(/-(TAX|MAIN)$/i, '').trim() : '';
  };

  // Deduplicate and propagate base-ID pair metadata (dealerName, dealerAllocationId, allocationQty) onto split -TAX rows
  const dedupMap = new Map<string, ChallanRecord>();
  for (const c of db.challans) {
    const chKey = normChallanNo(c.challanNo);
    const dKey = chKey ? `${chKey}::${c.challanType}` : `id::${c.id}`;
    if (!dedupMap.has(dKey)) dedupMap.set(dKey, c);
  }
  const rawList: ChallanRecord[] = Array.from(dedupMap.values());
  const pairMetaByBase = new Map<
    string,
    { dealerName?: string; dealerAllocationId?: string; allocationQty?: number; challanDate?: string }
  >();
  for (const c of rawList) {
    const base = getChallanPairId(c);
    if (!base) continue;
    const cur = pairMetaByBase.get(base) || {};
    if ((!cur.dealerName || cur.dealerName === '—') && c.dealerName && c.dealerName !== '—') {
      cur.dealerName = c.dealerName;
    }
    if (!cur.dealerAllocationId && c.dealerAllocationId) {
      cur.dealerAllocationId = c.dealerAllocationId;
    }
    if ((!cur.allocationQty || cur.allocationQty <= 0) && c.allocationQty && c.allocationQty > 0) {
      cur.allocationQty = c.allocationQty;
    }
    if (!cur.challanDate && c.challanDate) {
      cur.challanDate = c.challanDate;
    }
    pairMetaByBase.set(base, cur);
  }
  const list: ChallanRecord[] = rawList.map((c) => {
    const base = getChallanPairId(c);
    const info = base ? pairMetaByBase.get(base) : undefined;
    if (!info) return c;
    return {
      ...c,
      sourceRowId: base,
      dealerName: c.dealerName && c.dealerName !== '—' ? c.dealerName : info.dealerName || c.dealerName,
      dealerAllocationId: c.dealerAllocationId || info.dealerAllocationId,
      allocationQty: c.allocationQty && c.allocationQty > 0 ? c.allocationQty : info.allocationQty,
      challanDate: c.challanDate || info.challanDate || '',
    };
  });
  const rateMap = new Map(db.rates.map((r) => [r.category, r]));
  const subById = new Map(db.dealerSuballocations.map((s) => [s.id.trim(), s]));

  // Collect canonical Bengali/app party names for English or bilingual sheet names
  const masterPartyNamesByCat = new Map<AllocationCategory, string[]>();
  const allMasterPartyNames: string[] = [];
  const pushParty = (cat: AllocationCategory | undefined, name: string | undefined) => {
    const clean = (name || '').trim();
    if (!clean || clean === '—') return;
    if (cat) {
      const arr = masterPartyNamesByCat.get(cat) || [];
      if (!arr.includes(clean)) {
        arr.push(clean);
        masterPartyNamesByCat.set(cat, arr);
      }
    }
    if (!allMasterPartyNames.includes(clean)) allMasterPartyNames.push(clean);
  };
  for (const s of db.dealerSuballocations) pushParty(s.category, s.dealerName);
  for (const d of db.foodDealers) pushParty('খাদ্য বান্ধব', d.dealerName);
  for (const d of db.omsDealers) {
    pushParty('ওএমএস সাধারণ', d.dealerName);
    pushParty('ওএমএস বিশেষ', d.dealerName);
    pushParty('ওএমএস টিসিবি', d.dealerName);
  }
  for (const f of db.fireServices) {
    pushParty('ইপি চাল', f.officeName);
    pushParty('ইপি গম', f.officeName);
  }
  const resolveCanonicalDealerName = (cat: AllocationCategory, sheetName: string): string => {
    const clean = (sheetName || '').trim();
    if (!clean || clean === '—') return '';
    const catPool = masterPartyNamesByCat.get(cat) || [];
    const hitCat = catPool.find((m) => isSamePartyName(m, clean));
    if (hitCat) return hitCat;
    const hitAny = allMasterPartyNames.find((m) => isSamePartyName(m, clean));
    return hitAny || clean;
  };

  // Strict chronological 1-to-1 matching across ALL challans & DOs:
  // Same category, same dealer/party, challanDate <= doIssueDate, earliest challan -> earliest DO
  const sortedAllDos = [...db.doIssues].sort((a, b) => {
    const da = toIsoDateStr(a.issueDate) || '9999-99-99';
    const dbDate = toIsoDateStr(b.issueDate) || '9999-99-99';
    if (da !== dbDate) return da.localeCompare(dbDate);
    const doCmp = bnToEnDigits(a.doNo || '').localeCompare(bnToEnDigits(b.doNo || ''), undefined, {
      numeric: true,
    });
    if (doCmp !== 0) return doCmp;
    return (a.createdAt || '').localeCompare(b.createdAt || '');
  });

  const usedChallanIds = new Set<string>();
  const usedChallanNoTypes = new Set<string>();
  const matchedDoByChallanId = new Map<string, DOIssue>();
  const matchedMainPairByDoId = new Map<string, string>();
  const pinnedMainDoIds = new Set<string>();
  const pinnedTaxDoIds = new Set<string>();
  const doNoExistsSet = new Set(db.doIssues.map((d) => normDoNo(d.doNo)).filter(Boolean));

  const getDealerForChallan = (c: ChallanRecord): string => {
    const rawSub = c.dealerAllocationId ? subById.get(c.dealerAllocationId.trim()) : undefined;
    const validSub =
      rawSub &&
      rawSub.category === c.category &&
      (!c.dealerName || c.dealerName === '—' || isSamePartyName(c.dealerName, rawSub.dealerName))
        ? rawSub
        : undefined;
    return (c.dealerName && c.dealerName !== '—' ? c.dealerName : validSub?.dealerName || '').trim();
  };

  const lockRegMatch = (type: 'মূল চালান' | 'আয়কর চালান', d: DOIssue, c: ChallanRecord) => {
    usedChallanIds.add(c.id);
    const bKey = normChallanNo(c.challanNo);
    if (bKey) usedChallanNoTypes.add(`${bKey}::${type}`);
    matchedDoByChallanId.set(c.id, d);
    if (type === 'মূল চালান') {
      pinnedMainDoIds.add(d.id);
      const pairId = getChallanPairId(c);
      if (pairId) matchedMainPairByDoId.set(d.id, pairId);
    } else {
      pinnedTaxDoIds.add(d.id);
    }
  };

  // Pass 0A: LOCK saved DO -> Challan pairings FIRST
  for (const type of ['মূল চালান', 'আয়কর চালান'] as const) {
    for (const d of sortedAllDos) {
      const targetDoChallanNo = type === 'মূল চালান' ? d.challanNo : d.taxChallanNo;
      const targetNorm = normChallanNo(targetDoChallanNo);
      if (!targetNorm) continue;
      const dIso = toIsoDateStr(d.issueDate);
      const pinned = list.find((c) => {
        if (c.challanType !== type || c.category !== d.category) return false;
        if (usedChallanIds.has(c.id)) return false;
        const chKey = normChallanNo(c.challanNo);
        if (!chKey || chKey !== targetNorm) return false;
        if (usedChallanNoTypes.has(`${chKey}::${type}`)) return false;
        const dealerName = getDealerForChallan(c);
        if (dealerName && !isSamePartyName(dealerName, d.recipientName)) return false;
        const cIso = toIsoDateStr(c.challanDate);
        if (cIso && dIso && cIso > dIso) return false;
        return true;
      });
      if (pinned) {
        lockRegMatch(type, d, pinned);
      } else {
        const savedDate = type === 'মূল চালান' ? d.challanDate : d.taxChallanDate;
        if (isChallanOnOrBeforeDo(savedDate, d.issueDate, false)) {
          usedChallanNoTypes.add(`${targetNorm}::${type}`);
          if (type === 'মূল চালান') pinnedMainDoIds.add(d.id);
          else pinnedTaxDoIds.add(d.id);
        }
      }
    }
  }

  // Pass 0B: Explicit c.doNo on unpinned DOs
  for (const type of ['মূল চালান', 'আয়কর চালান'] as const) {
    for (const d of sortedAllDos) {
      if (type === 'মূল চালান' && pinnedMainDoIds.has(d.id)) continue;
      if (type === 'আয়কর চালান' && pinnedTaxDoIds.has(d.id)) continue;
      const dDoNorm = normDoNo(d.doNo);
      if (!dDoNorm) continue;
      const dIso = toIsoDateStr(d.issueDate);
      const byExplicitDo = list.find((c) => {
        if (c.challanType !== type || c.category !== d.category) return false;
        if (usedChallanIds.has(c.id)) return false;
        const chKey = normChallanNo(c.challanNo);
        if (chKey && usedChallanNoTypes.has(`${chKey}::${type}`)) return false;
        if (!c.doNo || normDoNo(c.doNo) !== dDoNorm) return false;
        const dealerName = getDealerForChallan(c);
        if (!dealerName || !isSamePartyName(dealerName, d.recipientName)) return false;
        const cIso = toIsoDateStr(c.challanDate);
        if (cIso && dIso && cIso > dIso) return false;
        return true;
      });
      if (byExplicitDo) lockRegMatch(type, d, byExplicitDo);
    }
  }

  // Pass 1: Remaining unlinked DOs only
  for (const type of ['মূল চালান', 'আয়কর চালান'] as const) {
    for (const d of sortedAllDos) {
      if (type === 'মূল চালান' && pinnedMainDoIds.has(d.id)) continue;
      if (type === 'আয়কর চালান' && pinnedTaxDoIds.has(d.id)) continue;
      const dIso = toIsoDateStr(d.issueDate);
      const dDoNorm = normDoNo(d.doNo);
      const dSubNorm = normMemoNo(d.dealerAllocationId);
      const pairedMainPairId = type === 'আয়কর চালান' ? matchedMainPairByDoId.get(d.id) : undefined;

      const candidates = list.filter((c) => {
        if (c.challanType !== type) return false;
        if (usedChallanIds.has(c.id)) return false;
        const chKey = normChallanNo(c.challanNo);
        if (chKey && usedChallanNoTypes.has(`${chKey}::${type}`)) return false;
        if (c.category !== d.category) return false;

        const cDoNorm = normDoNo(c.doNo);
        if (cDoNorm && (!dDoNorm || cDoNorm !== dDoNorm)) {
          return false;
        }
        const cSubNorm = normMemoNo(c.dealerAllocationId);
        if (cSubNorm) {
          if (!dSubNorm || cSubNorm !== dSubNorm) return false;
        }

        const dealerName = getDealerForChallan(c);
        if (!dealerName || !isSamePartyName(dealerName, d.recipientName)) return false;

        const cIso = toIsoDateStr(c.challanDate);
        if (cIso && dIso) {
          if (cIso > dIso) return false;
        } else if (!cIso && dIso) {
          const explicit =
            (cDoNorm && dDoNorm && cDoNorm === dDoNorm) ||
            (d.dealerAllocationId &&
              c.dealerAllocationId &&
              normMemoNo(d.dealerAllocationId) === normMemoNo(c.dealerAllocationId));
          if (!explicit) return false;
        }
        return true;
      });

      if (candidates.length === 0) continue;

      candidates.sort((a, b) => {
        if (pairedMainPairId) {
          const aSamePair = getChallanPairId(a) === pairedMainPairId ? 0 : 1;
          const bSamePair = getChallanPairId(b) === pairedMainPairId ? 0 : 1;
          if (aSamePair !== bSamePair) return aSamePair - bSamePair;
        }
        const aExplicit =
          d.dealerAllocationId &&
          a.dealerAllocationId &&
          normMemoNo(a.dealerAllocationId) === normMemoNo(d.dealerAllocationId)
            ? 0
            : 1;
        const bExplicit =
          d.dealerAllocationId &&
          b.dealerAllocationId &&
          normMemoNo(b.dealerAllocationId) === normMemoNo(d.dealerAllocationId)
            ? 0
            : 1;
        if (aExplicit !== bExplicit) return aExplicit - bExplicit;

        const aIso = toIsoDateStr(a.challanDate);
        const bIso = toIsoDateStr(b.challanDate);
        if (Boolean(aIso) !== Boolean(bIso)) return aIso ? -1 : 1;
        if (aIso && bIso && aIso !== bIso) return aIso.localeCompare(bIso);
        if (d.allocationQty > 0) {
          const qa = a.allocationQty ? Math.abs(a.allocationQty - d.allocationQty) : 999;
          const qb = b.allocationQty ? Math.abs(b.allocationQty - d.allocationQty) : 999;
          if (Math.abs(qa - qb) > 0.01) return qa - qb;
        }
        return bnToEnDigits(a.challanNo || '').localeCompare(
          bnToEnDigits(b.challanNo || ''),
          undefined,
          { numeric: true }
        );
      });

      lockRegMatch(type, d, candidates[0]);
    }
  }

  const filtered = list.filter((c) => inDateRange(c.challanDate, startDate, endDate));
  const monthKey = (iso?: string) => {
    const s = toIsoDateStr(iso);
    return s ? s.slice(0, 7) : '';
  };

  interface ResolvedChallanItem {
    c: ChallanRecord;
    directSub?: (typeof db.dealerSuballocations)[number];
    dealerName: string;
    refMonth: string;
    matchedDo?: DOIssue;
    resolvedDoNo: string;
    resolvedDoDate: string;
  }

  const resolvedItems: ResolvedChallanItem[] = filtered.map((c) => {
    const rawSub = c.dealerAllocationId ? subById.get(c.dealerAllocationId.trim()) : undefined;
    const directSub =
      rawSub &&
      rawSub.category === c.category &&
      (!c.dealerName || c.dealerName === '—' || isSamePartyName(c.dealerName, rawSub.dealerName))
        ? rawSub
        : undefined;
    const matchedDo = matchedDoByChallanId.get(c.id);
    const rawChallanDealer = c.dealerName && c.dealerName !== '—' ? c.dealerName : '';
    const dealerName = (
      matchedDo?.recipientName ||
      directSub?.dealerName ||
      resolveCanonicalDealerName(c.category, rawChallanDealer) ||
      rawChallanDealer ||
      '—'
    ).trim();
    const refDate = c.challanDate || directSub?.unoDate || '';
    const refMonth = monthKey(refDate);

    let resolvedDoNo = '';
    let resolvedDoDate = '';
    if (matchedDo) {
      resolvedDoNo = (matchedDo.doNo || '').trim();
      resolvedDoDate = (matchedDo.issueDate || '').trim();
    } else if (c.doNo && !doNoExistsSet.has(normDoNo(c.doNo))) {
      // Unsynced external DO number written directly on Challan sheet (only keep if date valid)
      if (isChallanOnOrBeforeDo(c.challanDate, c.doDate, false)) {
        resolvedDoNo = c.doNo.trim();
        resolvedDoDate = (c.doDate || '').trim();
      }
    }

    return {
      c,
      directSub,
      dealerName,
      refMonth,
      matchedDo,
      resolvedDoNo,
      resolvedDoDate,
    };
  });

  // Normalize canonical DO display / date & group metadata so Main & Tax rows of the same DO share identical groupValues
  const groupMeta = new Map<
    string,
    { sortDate: string; doNo: string; doDate: string; dealerDisplay: string }
  >();
  const getGroupKey = (item: ResolvedChallanItem) =>
    item.resolvedDoNo
      ? `do::${normDoNo(item.resolvedDoNo)}::${item.c.category}`
      : `nodo::${item.c.id}`;

  for (const item of resolvedItems) {
    const gk = getGroupKey(item);
    const cur = groupMeta.get(gk);
    const candDate = toIsoDateStr(item.resolvedDoDate || item.c.challanDate) || '9999-99-99';
    if (!cur) {
      groupMeta.set(gk, {
        sortDate: candDate,
        doNo: item.resolvedDoNo,
        doDate: item.resolvedDoDate,
        dealerDisplay: item.dealerName,
      });
    } else {
      if (candDate < cur.sortDate) cur.sortDate = candDate;
      if (!cur.doDate && item.resolvedDoDate) cur.doDate = item.resolvedDoDate;
      if (cur.dealerDisplay === '—' && item.dealerName !== '—') {
        cur.dealerDisplay = item.dealerName;
      }
    }
  }

  resolvedItems.sort((a, b) => {
    const gkA = getGroupKey(a);
    const gkB = getGroupKey(b);
    if (gkA !== gkB) {
      const metaA = groupMeta.get(gkA)!;
      const metaB = groupMeta.get(gkB)!;
      const dateCmp = metaA.sortDate.localeCompare(metaB.sortDate);
      if (dateCmp !== 0) return dateCmp;
      const doCmp = bnToEnDigits(metaA.doNo || '').localeCompare(
        bnToEnDigits(metaB.doNo || ''),
        undefined,
        { numeric: true }
      );
      if (doCmp !== 0) return doCmp;
      return gkA.localeCompare(gkB);
    }
    const typeRankA = a.c.challanType === 'মূল চালান' ? 0 : 1;
    const typeRankB = b.c.challanType === 'মূল চালান' ? 0 : 1;
    if (typeRankA !== typeRankB) return typeRankA - typeRankB;
    const chDateCmp = (toIsoDateStr(a.c.challanDate) || '').localeCompare(
      toIsoDateStr(b.c.challanDate) || ''
    );
    if (chDateCmp !== 0) return chDateCmp;
    return bnToEnDigits(a.c.challanNo || '').localeCompare(
      bnToEnDigits(b.c.challanNo || ''),
      undefined,
      { numeric: true }
    );
  });

  const rows: RegisterRowData[] = resolvedItems.map((item) => {
    const { c, directSub, dealerName, refMonth, matchedDo } = item;
    const gk = getGroupKey(item);
    const gMeta = groupMeta.get(gk)!;
    const effDealer = item.resolvedDoNo ? gMeta.dealerDisplay : dealerName;
    const effDoNo = item.resolvedDoNo ? gMeta.doNo : '';
    const effDoDate = item.resolvedDoNo ? gMeta.doDate : item.resolvedDoDate;

    let recVal = c.recordedAmount;
    if (recVal <= 0) {
      const sameMonthSub =
        !directSub && effDealer !== '—'
          ? db.dealerSuballocations.find(
              (s) =>
                s.category === c.category &&
                isSamePartyName(s.dealerName, effDealer) &&
                ((matchedDo?.unoMemo && normMemoNo(s.unoMemo) === normMemoNo(matchedDo.unoMemo)) ||
                  (refMonth && monthKey(s.unoDate) === refMonth))
            )
          : undefined;
      const matchedSub = directSub || sameMonthSub;
      const effCat = c.category;
      const rate = rateMap.get(effCat);
      const effUnitPrice = rate && rate.unitPrice > 0 ? rate.unitPrice : effCat === 'খাদ্য বান্ধব' ? 13000 : 0;
      const effTaxRate = rate && rate.taxRate > 0 ? rate.taxRate : effCat === 'খাদ্য বান্ধব' ? 200 : 0;
      const effQty =
        matchedSub?.allocationQty || c.allocationQty || matchedDo?.allocationQty || 0;

      if (c.challanType === 'আয়কর চালান') {
        recVal =
          matchedDo?.requiredTaxAmount ||
          matchedSub?.taxAmount ||
          round2(effQty * effTaxRate);
      } else {
        recVal =
          matchedDo?.requiredAmount ||
          matchedSub?.totalAmount ||
          round2(effQty * effUnitPrice);
      }
    }
    const rec = round2(recVal);
    const dep = round2(c.depositedAmount);
    const diff = round2(dep - rec);
    const isExcess = dep > rec;
    const isDeficit = dep < rec;

    const hasSignature =
      signRule === 'excess_or_equal' ? dep > 0 && dep >= rec : dep > rec;

    let mismatchStr = '০.০০ (সঠিক)';
    if (isExcess) {
      mismatchStr = `+${formatBnMoney(diff)} (বেশি জমা)`;
    } else if (isDeficit) {
      mismatchStr = `-${formatBnMoney(Math.abs(diff))} (কম জমা)`;
    }

    const doDisplay = effDoNo
      ? effDoDate
        ? `ডিও- ${toBnDigits(effDoNo)}\nতাং- ${formatBnDate(effDoDate)}`
        : `ডিও- ${toBnDigits(effDoNo)}`
      : '—';

    return {
      id: c.id,
      groupValues: effDoNo ? [doDisplay, effDealer, c.category] : ['', '', ''],
      cells: {
        doNo: doDisplay,
        dealer: effDealer,
        category: c.category,
        type: c.challanType,
        date: formatBnDate(c.challanDate),
        challanNo: `${toBnDigits(c.challanNo)}${c.bankName ? `\n(${c.bankName})` : ''}`,
        recorded: formatBnMoney(rec),
        deposited: formatBnMoney(dep),
        mismatch: mismatchStr,
        signature: '',
      },
      hasSignature,
      highlight: isExcess ? 'excess' : isDeficit ? 'deficit' : 'none',
    };
  });

  const leaves = columns.map((c) => c.leaf!);
  return {
    id: 'reg-challan',
    title: 'চালান রেজিস্টার',
    columns,
    leaves,
    mergeDepth: 3,
    colWidths: computeRegisterColWidths(leaves, true),
    tableLayoutFixed: true,
    rows,
  };
}

function buildKernelRegister(
  db: AllocationDatabase,
  startDate: string,
  endDate: string,
  template?: ReportTemplateConfig
): BuiltRegisterData {
  const colOverride = template?.columnLabels || {};
  const label = (key: string, def: string) => colOverride[key]?.trim() || def;

  const columns: RegisterColumnGroup[] = [
    { label: label('category', 'খাত'), leaf: { key: 'category', label: label('category', 'খাত'), align: 'left' } },
    {
      label: label('ministry', 'মন্ত্রণালয় বরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'),
      leaf: { key: 'ministry', label: label('ministry', 'মন্ত্রণালয় বরাদ্দ'), align: 'left' },
    },
    {
      label: label('uno', 'উপবরাদ্দ (স্মারক, তারিখ, বরাদ্দ ও অবশিষ্ট)'),
      leaf: { key: 'uno', label: label('uno', 'উপবরাদ্দ'), align: 'left' },
    },
    { label: label('union', 'ইউনিয়ন / এলাকা'), leaf: { key: 'union', label: label('union', 'ইউনিয়ন / এলাকা'), align: 'left' } },
    {
      label: label('recipient', 'প্রাপক / ডিলার / প্রকল্প'),
      leaf: { key: 'recipient', label: label('recipient', 'প্রাপক / ডিলার'), align: 'left' },
    },
    { label: label('grain', 'পণ্য'), leaf: { key: 'grain', label: label('grain', 'পণ্য'), align: 'center' } },
    { label: label('qty', 'পরিমাণ (মে.টন)'), leaf: { key: 'qty', label: label('qty', 'পরিমাণ (মে.টন)'), align: 'right' } },
    { label: label('doInfo', 'ডিও নং ও তারিখ'), leaf: { key: 'doInfo', label: label('doInfo', 'ডিও নং ও তারিখ'), align: 'center' } },
    { label: label('signature', 'স্বাক্ষর'), leaf: { key: 'signature', label: label('signature', 'স্বাক্ষর'), align: 'center', isSignature: true } },
  ];

  const kernelMinAllocs = db.ministryAllocations.filter(
    (m) => m.category === 'কার্নেল' || m.grain === 'কার্নেল'
  );
  const { dosBySubKey } = matchDOsToSuballocations(db);

  const findKernelMin = (cat: AllocationCategory, memoRaw?: string) => {
    const t = (memoRaw || '').trim();
    if (!t) return undefined;
    const catMeta = CATEGORY_BY_NAME[cat];
    const dcRec = catMeta?.requiresDC
      ? db.dcAllocations.find(
          (d) =>
            d.category === cat &&
            (d.dcMemo === t || normMemoNo(d.dcMemo) === normMemoNo(t) || memosLikelyMatch(d.dcMemo, t))
        )
      : undefined;
    const resolved = dcRec ? dcRec.ministryMemo : t;
    const nt = normMemoNo(resolved);
    return kernelMinAllocs.find(
      (m) =>
        m.category === cat &&
        (m.memoNo === resolved ||
          (nt && normMemoNo(m.memoNo) === nt) ||
          memosLikelyMatch(m.memoNo, resolved))
    );
  };

  const minUsedById = new Map<string, number>();
  const subCountByMinId = new Map<string, number>();
  const unoTotalsMap = new Map<string, { alloc: number; issued: number }>();

  interface KernelRowItem {
    id: string;
    category: AllocationCategory;
    minRec: (typeof kernelMinAllocs)[number];
    unoMemo: string;
    unoDate: string;
    unionName: string;
    recipientText: string;
    qty: number;
    matchedDos: DOIssue[];
    isUnallocatedMin?: boolean;
  }

  const allItems: KernelRowItem[] = [];

  for (const s of db.dealerSuballocations) {
    const minRec = findKernelMin(s.category, s.sourceMemo);
    if (!minRec) continue;
    minUsedById.set(minRec.id, round3((minUsedById.get(minRec.id) || 0) + s.allocationQty));
    subCountByMinId.set(minRec.id, (subCountByMinId.get(minRec.id) || 0) + 1);

    const matchedDos = dosBySubKey.get(`ds:${s.id}`) || [];
    const doSum = round3(matchedDos.reduce((acc, d) => acc + (Number(d.allocationQty) || 0), 0));
    const uMemo = s.unoMemo || s.sourceMemo || 'ইউসিএফ';
    const uKey = `${s.category}::${minRec.id}::${normMemoNo(uMemo)}::${s.unoDate}`;
    const uStat = unoTotalsMap.get(uKey) || { alloc: 0, issued: 0 };
    uStat.alloc = round3(uStat.alloc + s.allocationQty);
    uStat.issued = round3(uStat.issued + doSum);
    unoTotalsMap.set(uKey, uStat);

    if (inDateRange(s.unoDate, startDate, endDate)) {
      allItems.push({
        id: `ds-${s.id}`,
        category: s.category,
        minRec,
        unoMemo: uMemo,
        unoDate: s.unoDate,
        unionName: s.unionName || '—',
        recipientText: `${s.dealerName}${s.salesCenter ? `\n(${s.salesCenter})` : ''}`,
        qty: s.allocationQty,
        matchedDos,
      });
    }
  }

  for (const u of db.unoSuballocations) {
    const minRec = findKernelMin(u.category, u.ministryMemo || u.dcMemo);
    if (!minRec) continue;
    minUsedById.set(minRec.id, round3((minUsedById.get(minRec.id) || 0) + u.allocationQty));
    subCountByMinId.set(minRec.id, (subCountByMinId.get(minRec.id) || 0) + 1);

    const matchedDos = dosBySubKey.get(`uno:${u.id}`) || [];
    const doSum = round3(matchedDos.reduce((acc, d) => acc + (Number(d.allocationQty) || 0), 0));
    const uMemo = u.unoMemo || 'ইউএনও';
    const uKey = `${u.category}::${minRec.id}::${normMemoNo(uMemo)}::${u.unoDate}`;
    const uStat = unoTotalsMap.get(uKey) || { alloc: 0, issued: 0 };
    uStat.alloc = round3(uStat.alloc + u.allocationQty);
    uStat.issued = round3(uStat.issued + doSum);
    unoTotalsMap.set(uKey, uStat);

    if (inDateRange(u.unoDate, startDate, endDate)) {
      const recText = u.projectNo
        ? `প্র: নং- ${toBnDigits(u.projectNo)}${u.projectName ? `\n${u.projectName}` : ''}\n${u.recipientName}`
        : u.recipientName;
      allItems.push({
        id: `uno-${u.id}`,
        category: u.category,
        minRec,
        unoMemo: uMemo,
        unoDate: u.unoDate,
        unionName: u.unionName || '—',
        recipientText: recText,
        qty: u.allocationQty,
        matchedDos,
      });
    }
  }

  // Also display newly entered kernel ministry allocations that have not been suballocated yet
  for (const m of kernelMinAllocs) {
    if ((subCountByMinId.get(m.id) || 0) > 0) continue;
    if (!inDateRange(m.memoDate, startDate, endDate)) continue;
    allItems.push({
      id: `min-${m.id}`,
      category: m.category,
      minRec: m,
      unoMemo: '',
      unoDate: '',
      unionName: '—',
      recipientText: '—',
      qty: m.allocationQty,
      matchedDos: [],
      isUnallocatedMin: true,
    });
  }

  const unionOrder = db.officer.unions || [];
  allItems.sort((a, b) => {
    const catCmp = a.category.localeCompare(b.category);
    if (catCmp !== 0) return catCmp;
    const minDateCmp = (a.minRec.memoDate || '9999-99-99').localeCompare(
      b.minRec.memoDate || '9999-99-99'
    );
    if (minDateCmp !== 0) return minDateCmp;
    const memoCmp = bnToEnDigits(a.minRec.memoNo).localeCompare(bnToEnDigits(b.minRec.memoNo), undefined, {
      numeric: true,
    });
    if (memoCmp !== 0) return memoCmp;
    const uDateCmp = (a.unoDate || '9999-99-99').localeCompare(b.unoDate || '9999-99-99');
    if (uDateCmp !== 0) return uDateCmp;
    const ia = unionOrder.findIndex((u) => unionsLikelyMatch(u, a.unionName));
    const ib = unionOrder.findIndex((u) => unionsLikelyMatch(u, b.unionName));
    if (ia !== -1 && ib !== -1 && ia !== ib) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.recipientText.localeCompare(b.recipientText);
  });

  const rows: RegisterRowData[] = allItems.map((it) => {
    const remMin = round3(Math.max(0, it.minRec.allocationQty - (minUsedById.get(it.minRec.id) || 0)));
    const minText = [
      toBnDigits(it.minRec.memoNo),
      `তাং- ${formatBnDate(it.minRec.memoDate)}`,
      it.minRec.phase ? `পর্যায়-${toBnDigits(it.minRec.phase)}` : '',
      `মোট বরাদ্দ: ${formatBnQty(it.minRec.allocationQty)} মে.টন (কার্নেল)`,
      `অবশিষ্ট: ${formatBnQty(remMin)} মে.টন`,
    ]
      .filter(Boolean)
      .join('\n');

    let unoText = 'এখনো উপবরাদ্দ হয়নি';
    if (!it.isUnallocatedMin) {
      const uKey = `${it.category}::${it.minRec.id}::${normMemoNo(it.unoMemo)}::${it.unoDate}`;
      const uStat = unoTotalsMap.get(uKey) || { alloc: it.qty, issued: 0 };
      const remUno = round3(Math.max(0, uStat.alloc - uStat.issued));
      unoText = [
        toBnDigits(it.unoMemo),
        `তাং- ${formatBnDate(it.unoDate)}`,
        `মোট বরাদ্দ: ${formatBnQty(uStat.alloc)} মে.টন`,
        `অবশিষ্ট: ${formatBnQty(remUno)} মে.টন`,
      ].join('\n');
    }

    const hasDo = it.matchedDos.some((d) => Boolean(d.doNo && String(d.doNo).trim()));
    const doInfoText =
      it.matchedDos.length === 0
        ? '—'
        : it.matchedDos
            .map(
              (d) =>
                `ডিও নং- ${toBnDigits(d.doNo)}${d.issueDate ? `\nতাং- ${formatBnDate(d.issueDate)}` : ''}`
            )
            .join('\n');

    return {
      id: it.id,
      groupValues: [it.category, minText, unoText],
      cells: {
        category: it.category,
        ministry: minText,
        uno: unoText,
        union: toBnDigits(it.unionName || '—'),
        recipient: it.recipientText,
        grain: 'কার্নেল',
        qty: formatBnQty(it.qty),
        doInfo: doInfoText,
        signature: '',
      },
      hasSignature: hasDo,
    };
  });

  const leaves = columns.map((c) => c.leaf!);
  return {
    id: 'reg-kernel',
    title: 'কার্নেল রেজিস্টার',
    columns,
    leaves,
    mergeDepth: 3,
    serialPerRow: true,
    colWidths: computeRegisterColWidths(leaves, true),
    tableLayoutFixed: true,
    rows,
  };
}

export interface MonthlyDistRow {
  category: AllocationCategory;
  categoryBn: string;
  previousRice: number;
  previousWheat: number;
  previousAtta: number;
  previousKernel: number;
  currentRice: number;
  currentWheat: number;
  currentAtta: number;
  currentKernel: number;
  cumulativeRice: number;
  cumulativeWheat: number;
  cumulativeAtta: number;
  cumulativeKernel: number;
}

export function buildMonthlyDistributionReport(
  db: AllocationDatabase,
  month0: number,
  year: number,
  categoryFilter: string
) {
  const fy = fiscalYearLabel(year, month0);
  const grainMap = new Map<string, string>();
  for (const m of db.ministryAllocations) {
    grainMap.set(`${m.category}::${m.memoNo}`, m.grain);
    const nm = normMemoNo(m.memoNo);
    if (nm) grainMap.set(`${m.category}::${nm}`, m.grain);
  }

  const targetCats =
    categoryFilter && categoryFilter !== 'ALL'
      ? CATEGORIES.filter((c) => c.category === categoryFilter)
      : CATEGORIES;

  const rows: MonthlyDistRow[] = targetCats.map((meta) => {
    const r: MonthlyDistRow = {
      category: meta.category,
      categoryBn: meta.category,
      previousRice: 0,
      previousWheat: 0,
      previousAtta: 0,
      previousKernel: 0,
      currentRice: 0,
      currentWheat: 0,
      currentAtta: 0,
      currentKernel: 0,
      cumulativeRice: 0,
      cumulativeWheat: 0,
      cumulativeAtta: 0,
      cumulativeKernel: 0,
    };

    for (const d of db.doIssues) {
      if (d.category !== meta.category) continue;
      const my = parseMonthYear(d.issueDate);
      if (!my) continue;
      const rel = isMonthInFiscalUpTo(my.year, my.month, year, month0);
      if (rel === 'after' || rel === 'outside') continue;
      const grain =
        grainMap.get(`${d.category}::${d.ministryMemo}`) ||
        grainMap.get(`${d.category}::${normMemoNo(d.ministryMemo)}`) ||
        meta.defaultGrain ||
        'চাল';
      const qty = Number(d.allocationQty) || 0;
      if (rel === 'before') {
        if (grain === 'গম') r.previousWheat = round3(r.previousWheat + qty);
        else if (grain === 'আটা') r.previousAtta = round3(r.previousAtta + qty);
        else if (grain === 'কার্নেল') r.previousKernel = round3(r.previousKernel + qty);
        else r.previousRice = round3(r.previousRice + qty);
      } else if (rel === 'current') {
        if (grain === 'গম') r.currentWheat = round3(r.currentWheat + qty);
        else if (grain === 'আটা') r.currentAtta = round3(r.currentAtta + qty);
        else if (grain === 'কার্নেল') r.currentKernel = round3(r.currentKernel + qty);
        else r.currentRice = round3(r.currentRice + qty);
      }
    }

    r.cumulativeRice = round3(r.previousRice + r.currentRice);
    r.cumulativeWheat = round3(r.previousWheat + r.currentWheat);
    r.cumulativeAtta = round3(r.previousAtta + r.currentAtta);
    r.cumulativeKernel = round3(r.previousKernel + r.currentKernel);
    return r;
  });

  const totals = rows.reduce(
    (acc, r) => ({
      previousRice: round3(acc.previousRice + r.previousRice),
      previousWheat: round3(acc.previousWheat + r.previousWheat),
      previousAtta: round3(acc.previousAtta + r.previousAtta),
      previousKernel: round3(acc.previousKernel + r.previousKernel),
      currentRice: round3(acc.currentRice + r.currentRice),
      currentWheat: round3(acc.currentWheat + r.currentWheat),
      currentAtta: round3(acc.currentAtta + r.currentAtta),
      currentKernel: round3(acc.currentKernel + r.currentKernel),
      cumulativeRice: round3(acc.cumulativeRice + r.cumulativeRice),
      cumulativeWheat: round3(acc.cumulativeWheat + r.cumulativeWheat),
      cumulativeAtta: round3(acc.cumulativeAtta + r.cumulativeAtta),
      cumulativeKernel: round3(acc.cumulativeKernel + r.cumulativeKernel),
    }),
    {
      previousRice: 0,
      previousWheat: 0,
      previousAtta: 0,
      previousKernel: 0,
      currentRice: 0,
      currentWheat: 0,
      currentAtta: 0,
      currentKernel: 0,
      cumulativeRice: 0,
      cumulativeWheat: 0,
      cumulativeAtta: 0,
      cumulativeKernel: 0,
    }
  );

  return { fiscalYear: fy, monthName: BENGALI_MONTHS[month0] || '', rows, totals };
}

export interface AllocationSummaryRow {
  category: AllocationCategory;
  categoryBn: string;
  grain: string;
  requiresDC: boolean;
  ministryQty: number;
  dcQty: number | null;
  subQty: number;
  projectQty: number;
  doQty: number;
  unissuedQty: number;
}

export function buildAllocationSummaryReport(
  db: AllocationDatabase,
  categoryFilter: string
) {
  const targetCats =
    categoryFilter && categoryFilter !== 'ALL'
      ? CATEGORIES.filter((c) => c.category === categoryFilter)
      : CATEGORIES;

  const rows: AllocationSummaryRow[] = targetCats.map((meta) => {
    const minItems = db.ministryAllocations.filter((m) => m.category === meta.category);
    const ministryQty = round3(
      minItems.reduce((s, m) => s + (Number(m.allocationQty) || 0), 0)
    );
    const dcRaw = db.dcAllocations
      .filter((d) => d.category === meta.category)
      .reduce((s, d) => s + (Number(d.allocationQty) || 0), 0);
    const dcQty = meta.requiresDC ? round3(dcRaw) : null;
    const subQty = round3(
      db.dealerSuballocations
        .filter((s) => s.category === meta.category)
        .reduce((acc, s) => acc + (Number(s.allocationQty) || 0), 0) +
        db.unoSuballocations
          .filter((u) => u.category === meta.category)
          .reduce((acc, u) => acc + (Number(u.allocationQty) || 0), 0)
    );
    const projectQty = round3(
      db.projects
        .filter((p) => p.category === meta.category)
        .reduce((acc, p) => acc + (Number(p.allocationQty) || 0), 0)
    );
    const doQty = round3(
      db.doIssues
        .filter((d) => d.category === meta.category)
        .reduce((acc, d) => acc + (Number(d.allocationQty) || 0), 0)
    );
    const baseAlloc = Math.max(ministryQty, dcQty ?? 0, subQty, projectQty);
    const unissuedQty = Math.max(0, round3(baseAlloc - doQty));
    const distinctGrains = Array.from(new Set(minItems.map((m) => m.grain).filter(Boolean)));
    const grain =
      distinctGrains.length > 0 ? distinctGrains.join(', ') : meta.defaultGrain || 'চাল';

    return {
      category: meta.category,
      categoryBn: meta.category,
      grain,
      requiresDC: meta.requiresDC,
      ministryQty,
      dcQty,
      subQty,
      projectQty,
      doQty,
      unissuedQty,
    };
  });

  const totals = rows.reduce(
    (acc, r) => ({
      ministryQty: round3(acc.ministryQty + r.ministryQty),
      dcQty: round3(acc.dcQty + (r.dcQty ?? 0)),
      subQty: round3(acc.subQty + r.subQty),
      projectQty: round3(acc.projectQty + r.projectQty),
      doQty: round3(acc.doQty + r.doQty),
      unissuedQty: round3(acc.unissuedQty + r.unissuedQty),
    }),
    { ministryQty: 0, dcQty: 0, subQty: 0, projectQty: 0, doQty: 0, unissuedQty: 0 }
  );

  return { rows, totals };
}

export function buildCashSaleReport(db: AllocationDatabase, month0: number, year: number) {
  const rateMap = new Map(db.rates.map((r) => [r.category, r]));
  const grainMap = new Map(
    db.ministryAllocations.map((m) => [`${m.category}::${normMemoNo(m.memoNo)}`, m.grain])
  );
  const subById = new Map(db.dealerSuballocations.map((s) => [s.id.trim(), s]));
  const isKernelSub = (s?: (typeof db.dealerSuballocations)[number]): boolean => {
    if (!s) return false;
    if (s.category === 'কার্নেল') return true;
    return grainMap.get(`${s.category}::${normMemoNo(s.sourceMemo)}`) === 'কার্নেল';
  };
  const isKernelDo = (d: DOIssue): boolean => {
    if (d.category === 'কার্নেল') return true;
    if (grainMap.get(`${d.category}::${normMemoNo(d.ministryMemo)}`) === 'কার্নেল') return true;
    const linked = d.dealerAllocationId ? subById.get(d.dealerAllocationId.trim()) : undefined;
    return isKernelSub(linked);
  };

  interface CashChallanItem {
    challanNo: string;
    challanDate: string;
    qty: number;
    amount: number;
    grain: string;
  }

  const rows: Array<{
    channel: AllocationCategory;
    product: string;
    suppliedQuantity: number;
    rate: number;
    challan: string;
    amount: number;
  }> = [];

  const allMainChallanNos = new Set(
    db.challans
      .filter((c) => c.challanType === 'মূল চালান' && c.challanNo)
      .map((c) => normChallanNo(c.challanNo))
      .filter(Boolean)
  );

  for (const channel of CHALLAN_CATEGORIES) {
    const rateRec = rateMap.get(channel);
    const configuredRate =
      rateRec && rateRec.unitPrice > 0
        ? rateRec.unitPrice
        : channel === 'খাদ্য বান্ধব'
          ? 13000
          : 0;

    // Supplied quantity strictly equals total non-kernel DO issues for this category in the selected month
    const channelMonthlyDos = db.doIssues.filter((d) => {
      if (d.category !== channel) return false;
      if (isKernelDo(d)) return false;
      const my = parseMonthYear(d.issueDate);
      return Boolean(my && my.year === year && my.month === month0);
    });
    const doTotalQty = round3(
      channelMonthlyDos.reduce((s, d) => s + (Number(d.allocationQty) || 0), 0)
    );

    const items: CashChallanItem[] = [];
    const seenChallanNos = new Set<string>();
    const coveredDoNos = new Set<string>();

    // Only main challans ('মূল চালান'), never 'আয়কর চালান'
    for (const c of db.challans) {
      if (c.challanType !== 'মূল চালান') continue;
      if (c.category !== channel) continue;
      const my = parseMonthYear(c.challanDate || c.doDate || '');
      if (!my || my.year !== year || my.month !== month0) continue;

      const chKey = normChallanNo(c.challanNo);
      if (chKey) {
        if (seenChallanNos.has(chKey)) continue;
        seenChallanNos.add(chKey);
      }

      const matchedDo = db.doIssues.find(
        (d) =>
          d.category === channel &&
          !isKernelDo(d) &&
          (!c.dealerName || isSamePartyName(c.dealerName, d.recipientName)) &&
          isChallanOnOrBeforeDo(c.challanDate, d.issueDate, false) &&
          ((c.doNo && d.doNo && normMemoNo(d.doNo) === normMemoNo(c.doNo)) ||
            (chKey && d.challanNo && normChallanNo(d.challanNo) === chKey) ||
            (c.dealerAllocationId &&
              d.dealerAllocationId &&
              normMemoNo(d.dealerAllocationId) === normMemoNo(c.dealerAllocationId)))
      );
      if (c.doNo) coveredDoNos.add(normMemoNo(c.doNo));
      if (matchedDo?.doNo) coveredDoNos.add(normMemoNo(matchedDo.doNo));

      const rawMatchedSub = c.dealerAllocationId ? subById.get(c.dealerAllocationId.trim()) : undefined;
      const matchedSub = rawMatchedSub && !isKernelSub(rawMatchedSub) ? rawMatchedSub : undefined;
      const amount = round2(c.depositedAmount > 0 ? c.depositedAmount : matchedDo?.challanAmount || 0);
      let qty =
        matchedDo?.allocationQty ||
        c.allocationQty ||
        matchedSub?.allocationQty ||
        0;
      if (qty <= 0 && configuredRate > 0 && amount > 0) {
        qty = round3(amount / configuredRate);
      }
      const grain =
        (matchedDo
          ? grainMap.get(`${channel}::${normMemoNo(matchedDo.ministryMemo)}`)
          : undefined) ||
        CATEGORY_BY_NAME[channel]?.defaultGrain ||
        'চাল';

      items.push({
        challanNo: (c.challanNo || '').trim(),
        challanDate: c.challanDate || c.doDate || matchedDo?.issueDate || '',
        qty: round3(qty),
        amount,
        grain,
      });
    }

    // Include main challans recorded directly on DOIssue only if not present anywhere in db.challans
    for (const d of db.doIssues) {
      if (d.category !== channel) continue;
      if (isKernelDo(d)) continue;
      if (!d.challanNo && !(d.challanAmount && d.challanAmount > 0)) continue;
      const effDate = d.challanDate || d.issueDate;
      const my = parseMonthYear(effDate);
      if (!my || my.year !== year || my.month !== month0) continue;

      const chKey = normChallanNo(d.challanNo);
      if (chKey && (seenChallanNos.has(chKey) || allMainChallanNos.has(chKey))) continue;
      if (!chKey && d.doNo && coveredDoNos.has(normMemoNo(d.doNo))) continue;
      if (chKey) seenChallanNos.add(chKey);
      if (d.doNo) coveredDoNos.add(normMemoNo(d.doNo));

      const amount = round2(
        d.challanAmount && d.challanAmount > 0
          ? d.challanAmount
          : configuredRate > 0
            ? d.allocationQty * configuredRate
            : 0
      );
      const grain =
        grainMap.get(`${channel}::${normMemoNo(d.ministryMemo)}`) ||
        CATEGORY_BY_NAME[channel]?.defaultGrain ||
        'চাল';

      items.push({
        challanNo: (d.challanNo || '').trim(),
        challanDate: effDate,
        qty: round3(d.allocationQty || 0),
        amount,
        grain,
      });
    }

    if (items.length === 0 && doTotalQty <= 0) continue;

    items.sort((a, b) => {
      const dc = (a.challanDate || '').localeCompare(b.challanDate || '');
      if (dc !== 0) return dc;
      return bnToEnDigits(a.challanNo).localeCompare(bnToEnDigits(b.challanNo), undefined, {
        numeric: true,
      });
    });

    const totalChannelAmt = round2(items.reduce((s, it) => s + it.amount, 0));
    const effRate =
      configuredRate > 0
        ? configuredRate
        : doTotalQty > 0 && totalChannelAmt > 0
          ? round2(totalChannelAmt / doTotalQty)
          : 0;

    const challanEntries = items.map((it) => {
      const noStr = it.challanNo ? toBnDigits(it.challanNo) : '—';
      const dtStr = it.challanDate ? ` (${formatBnDate(it.challanDate)})` : '';
      return `${noStr}${dtStr}`;
    });
    const challanLines: string[] = [];
    for (let i = 0; i < challanEntries.length; i += 2) {
      challanLines.push(challanEntries.slice(i, i + 2).join(', '));
    }
    const challanListText = challanLines.join('\n');

    const doGrains = Array.from(
      new Set(
        channelMonthlyDos
          .map((d) => grainMap.get(`${channel}::${normMemoNo(d.ministryMemo)}`))
          .filter((g): g is NonNullable<typeof g> => Boolean(g))
      )
    );

    rows.push({
      channel,
      product:
        doGrains.join(', ') ||
        items[0]?.grain ||
        CATEGORY_BY_NAME[channel]?.defaultGrain ||
        'চাল',
      suppliedQuantity: doTotalQty,
      rate: effRate,
      challan: challanListText || '—',
      amount: totalChannelAmt,
    });
  }

  const totalQty = round3(rows.reduce((s, r) => s + r.suppliedQuantity, 0));
  const totalAmount = round2(rows.reduce((s, r) => s + r.amount, 0));

  return { monthName: BENGALI_MONTHS[month0] || '', rows, totalQty, totalAmount };
}

export function buildEPRationReport(db: AllocationDatabase, month0: number, year: number) {
  const fy = fiscalYearLabel(year, month0);
  const byInst = new Map<
    string,
    { institution: string; rice: number; wheat: number; cumulativeRice: number; cumulativeWheat: number }
  >();

  const defaultInst = db.fireServices[0]?.officeName || 'ফায়ার সার্ভিস ও সিভিল ডিফেন্স';

  for (const d of db.doIssues) {
    if (d.category !== 'ইপি চাল' && d.category !== 'ইপি গম') continue;
    const my = parseMonthYear(d.issueDate);
    if (!my) continue;
    const rel = isMonthInFiscalUpTo(my.year, my.month, year, month0);
    if (rel === 'after' || rel === 'outside') continue;
    const inst = d.recipientName || defaultInst;
    const cur = byInst.get(inst) || {
      institution: inst,
      rice: 0,
      wheat: 0,
      cumulativeRice: 0,
      cumulativeWheat: 0,
    };
    if (d.category === 'ইপি গম') {
      if (rel === 'current') cur.wheat = round3(cur.wheat + d.allocationQty);
      cur.cumulativeWheat = round3(cur.cumulativeWheat + d.allocationQty);
    } else {
      if (rel === 'current') cur.rice = round3(cur.rice + d.allocationQty);
      cur.cumulativeRice = round3(cur.cumulativeRice + d.allocationQty);
    }
    byInst.set(inst, cur);
  }

  return {
    fiscalYear: fy,
    monthName: BENGALI_MONTHS[month0] || '',
    rows: Array.from(byInst.values()),
  };
}

function resolveFoodSubGrain(db: AllocationDatabase, sourceMemo: string): string {
  const t = (sourceMemo || '').trim();
  if (!t) return 'চাল';
  const nt = normMemoNo(t);
  const found = db.ministryAllocations.find(
    (m) =>
      m.category === 'খাদ্য বান্ধব' &&
      (m.memoNo === t || (nt && normMemoNo(m.memoNo) === nt) || memosLikelyMatch(m.memoNo, t))
  );
  return found?.grain || 'চাল';
}

export function buildFoodAllocOrderMemos(db: AllocationDatabase, month0: number, year: number) {
  const byMemo = new Map<
    string,
    { memo: string; date: string; count: number; totalQty: number; grain: string }
  >();
  for (const s of db.dealerSuballocations) {
    if (s.category !== 'খাদ্য বান্ধব') continue;
    const my = parseMonthYear(s.unoDate);
    if (!my || my.year !== year || my.month !== month0) continue;
    const grain = resolveFoodSubGrain(db, s.sourceMemo);
    const cur = byMemo.get(s.unoMemo) || {
      memo: s.unoMemo,
      date: s.unoDate,
      count: 0,
      totalQty: 0,
      grain,
    };
    cur.count += 1;
    cur.totalQty = round3(cur.totalQty + s.allocationQty);
    byMemo.set(s.unoMemo, cur);
  }
  return Array.from(byMemo.values()).sort((a, b) => {
    if (a.grain !== b.grain) {
      if (a.grain === 'কার্নেল') return 1;
      if (b.grain === 'কার্নেল') return -1;
    }
    return a.date.localeCompare(b.date);
  });
}

export function buildFoodKernelMemos(db: AllocationDatabase, month0: number, year: number) {
  const byMemo = new Map<
    string,
    { memo: string; date: string; count: number; totalQty: number; sameMonth: boolean }
  >();
  for (const s of db.dealerSuballocations) {
    const isKernelCat =
      s.category === 'কার্নেল' ||
      (s.category === 'খাদ্য বান্ধব' && resolveFoodSubGrain(db, s.sourceMemo) === 'কার্নেল');
    if (!isKernelCat) continue;
    const my = parseMonthYear(s.unoDate);
    const sameMonth = Boolean(my && my.year === year && my.month === month0);
    const key = normMemoNo(s.unoMemo) || s.unoMemo;
    const cur = byMemo.get(key) || {
      memo: s.unoMemo,
      date: s.unoDate,
      count: 0,
      totalQty: 0,
      sameMonth,
    };
    cur.count += 1;
    cur.totalQty = round3(cur.totalQty + s.allocationQty);
    if (sameMonth) cur.sameMonth = true;
    byMemo.set(key, cur);
  }
  return Array.from(byMemo.values()).sort((a, b) => {
    if (a.sameMonth !== b.sameMonth) return a.sameMonth ? -1 : 1;
    return b.date.localeCompare(a.date);
  });
}

export function buildFoodAllocOrderRows(
  db: AllocationDatabase,
  month0: number,
  year: number,
  unoMemo: string,
  kernelOpts?: {
    includeKernel?: boolean;
    kernelMemo?: string;
    kernelPercent?: number;
  }
) {
  const findDealerInfo = (dealerName: string, salesCenter?: string, unionName?: string) =>
    db.foodDealers.find((d) => namesLikelyMatch(d.dealerName, dealerName)) ||
    db.foodDealers.find(
      (d) =>
        Boolean(salesCenter && d.salesCenter) &&
        namesFuzzyMatch(d.salesCenter, salesCenter!) &&
        unionsLikelyMatch(d.unionName, unionName || '')
    );

  const foodRate = db.rates.find((r) => r.category === 'খাদ্য বান্ধব');
  const effUnitPrice = foodRate && foodRate.unitPrice > 0 ? foodRate.unitPrice : 13000;
  const effTaxRate = foodRate && foodRate.taxRate > 0 ? foodRate.taxRate : 200;

  const includeKernel = Boolean(kernelOpts?.includeKernel);
  const trimmedUnoMemo = (unoMemo || '').trim();
  if (!trimmedUnoMemo) {
    return {
      monthName: BENGALI_MONTHS[month0] || '',
      year,
      memo: '',
      memoDate: '',
      sourceMinistryMemo: '',
      sourceMinistryDate: '',
      sourceDfcMemo: '',
      sourceDfcDate: '',
      unitPrice: effUnitPrice,
      taxRate: effTaxRate,
      baseGrain: 'চাল' as GrainType,
      includeKernel,
      rows: [],
    };
  }

  const kernelMemoNorm =
    kernelOpts?.kernelMemo && kernelOpts.kernelMemo !== 'PERCENT'
      ? normMemoNo(kernelOpts.kernelMemo)
      : '';
  const kernelPercent = Math.max(0, Number(kernelOpts?.kernelPercent) || 0);

  const kernelSubs = kernelMemoNorm
    ? db.dealerSuballocations.filter(
        (s) =>
          (s.category === 'কার্নেল' ||
            (s.category === 'খাদ্য বান্ধব' && resolveFoodSubGrain(db, s.sourceMemo) === 'কার্নেল')) &&
          normMemoNo(s.unoMemo) === kernelMemoNorm
      )
    : [];

  const targetMemoNorm = normMemoNo(trimmedUnoMemo);
  const matchedSubIds = db.dealerSuballocations.filter((s) => {
    if (s.category !== 'খাদ্য বান্ধব') return false;
    const my = parseMonthYear(s.unoDate);
    if (!my || my.year !== year || my.month !== month0) return false;
    return normMemoNo(s.unoMemo) === targetMemoNorm || memosLikelyMatch(s.unoMemo, trimmedUnoMemo);
  });

  const rows = matchedSubIds.map((s) => {
    const info = findDealerInfo(s.dealerName, s.salesCenter, s.unionName);
    const baseGrain = resolveFoodSubGrain(db, s.sourceMemo);
    const isBaseKernel = baseGrain === 'কার্নেল';

    const dealerName = info?.dealerName || s.dealerName;
    const union = info?.unionName || s.unionName;
    const salesCenter = info?.salesCenter || s.salesCenter;

    // Master Card Count always comes from Settings (NEVER subtract unused cards)
    const cards = isBaseKernel
      ? 0
      : info && info.cardCount > 0
        ? info.cardCount
        : Math.max(0, s.cardCount);

    // Net rice allocation in UCF ("চলতি মাসের বরাদ্দের পরিমাণ") drives cash & tax
    const qtyWithoutKernel = isBaseKernel ? 0 : round3(s.allocationQty);
    const explicitUnusedRice = s.unusedCards > 0 ? round3(s.unusedCards * 0.03) : 0;
    const explicitOpening = s.openingStock && s.openingStock > 0 ? round3(s.openingStock) : 0;
    const totalRiceRequirement = isBaseKernel
      ? 0
      : cards > 0
        ? round3(cards * 0.03)
        : round3(qtyWithoutKernel + Math.max(explicitUnusedRice, explicitOpening));
    const openingStockQty = isBaseKernel
      ? 0
      : explicitUnusedRice > 0
        ? explicitUnusedRice
        : explicitOpening > 0
          ? explicitOpening
          : totalRiceRequirement > qtyWithoutKernel
            ? round3(totalRiceRequirement - qtyWithoutKernel)
            : 0;
    let kernelQty = 0;

    if (isBaseKernel) {
      kernelQty = round3(s.allocationQty);
    } else if (includeKernel) {
      if (kernelMemoNorm) {
        const kMatch = kernelSubs.find(
          (k) =>
            namesLikelyMatch(k.dealerName, dealerName) ||
            (k.salesCenter &&
              salesCenter &&
              namesFuzzyMatch(k.salesCenter, salesCenter) &&
              unionsLikelyMatch(k.unionName, union))
        );
        kernelQty = kMatch ? round3(kMatch.allocationQty) : 0;
      } else if (kernelPercent > 0) {
        kernelQty = round3(qtyWithoutKernel * (kernelPercent / 100));
      }
    }

    const qtyWithKernel = round3(qtyWithoutKernel + kernelQty);
    const amount = isBaseKernel
      ? 0
      : round2(s.totalAmount > 0 ? s.totalAmount : qtyWithoutKernel * effUnitPrice);
    const tax = isBaseKernel
      ? 0
      : round2(s.taxAmount > 0 ? s.taxAmount : qtyWithoutKernel * effTaxRate);

    return {
      subId: s.id,
      union,
      salesCenter,
      dealerName,
      mobile: info?.mobile || '',
      cards,
      grossCards: cards,
      openingStockQty,
      totalRiceRequirement,
      baseGrain,
      qty: includeKernel ? qtyWithKernel : qtyWithoutKernel,
      qtyWithoutKernel,
      kernelQty,
      qtyWithKernel,
      amount,
      tax,
      unoDate: s.unoDate,
      sourceMemo: s.sourceMemo,
    };
  });

  const unionOrder = db.officer.unions || [];
  rows.sort((a, b) => {
    const ia = unionOrder.indexOf(a.union);
    const ib = unionOrder.indexOf(b.union);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.union.localeCompare(b.union);
  });

  const firstSourceMemo = rows.find((r) => r.sourceMemo)?.sourceMemo || '';
  const dcOrDfcRec = db.dcAllocations.find((d) => {
    if (d.category !== 'খাদ্য বান্ধব') return false;
    if (firstSourceMemo) {
      return (
        d.dcMemo === firstSourceMemo ||
        normMemoNo(d.dcMemo) === normMemoNo(firstSourceMemo) ||
        memosLikelyMatch(d.dcMemo, firstSourceMemo)
      );
    }
    const my = parseMonthYear(d.dcMemoDate);
    return Boolean(my && my.year === year && my.month === month0);
  });

  const ministryLookupMemo = dcOrDfcRec?.ministryMemo || firstSourceMemo;
  const sourceMinistryRec =
    (ministryLookupMemo
      ? db.ministryAllocations.find(
          (m) =>
            m.category === 'খাদ্য বান্ধব' &&
            (m.memoNo === ministryLookupMemo ||
              normMemoNo(m.memoNo) === normMemoNo(ministryLookupMemo) ||
              memosLikelyMatch(m.memoNo, ministryLookupMemo))
        )
      : undefined) ||
    db.ministryAllocations.find((m) => {
      if (m.category !== 'খাদ্য বান্ধব') return false;
      const my = parseMonthYear(m.memoDate);
      return Boolean(my && my.year === year && my.month === month0);
    });

  const resolvedMinistryMemo = sourceMinistryRec?.memoNo || ministryLookupMemo || '';
  const resolvedMinistryDate = sourceMinistryRec?.memoDate || '';
  const dfcFromSub =
    firstSourceMemo &&
    resolvedMinistryMemo &&
    normMemoNo(firstSourceMemo) !== normMemoNo(resolvedMinistryMemo) &&
    !memosLikelyMatch(firstSourceMemo, resolvedMinistryMemo)
      ? firstSourceMemo
      : '';
  const sourceDfcMemo = dcOrDfcRec?.dcMemo || dfcFromSub || '';
  const sourceDfcDate = dcOrDfcRec?.dcMemoDate || (dfcFromSub ? rows[0]?.unoDate || '' : '');

  return {
    monthName: BENGALI_MONTHS[month0] || '',
    year,
    memo: trimmedUnoMemo,
    memoDate: rows[0]?.unoDate || '',
    sourceMinistryMemo: resolvedMinistryMemo,
    sourceMinistryDate: resolvedMinistryDate,
    sourceDfcMemo,
    sourceDfcDate,
    unitPrice: effUnitPrice,
    taxRate: effTaxRate,
    baseGrain: rows[0]?.baseGrain || ('চাল' as GrainType),
    includeKernel,
    rows,
  };
}

export function buildOmsAllocOrderMemos(
  db: AllocationDatabase,
  category: 'ওএমএস সাধারণ' | 'ওএমএস বিশেষ' | 'ওএমএস টিসিবি',
  month0: number,
  year: number
) {
  const byMemo = new Map<
    string,
    { memo: string; date: string; count: number; totalQty: number; sourceMemo: string }
  >();
  for (const s of db.dealerSuballocations) {
    if (s.category !== category) continue;
    const my = parseMonthYear(s.unoDate);
    if (!my || my.year !== year || my.month !== month0) continue;
    const key = normMemoNo(s.unoMemo) || s.unoMemo || s.sourceMemo || 'মাসিক';
    const cur = byMemo.get(key) || {
      memo: s.unoMemo || s.sourceMemo || '',
      date: s.unoDate,
      count: 0,
      totalQty: 0,
      sourceMemo: s.sourceMemo || '',
    };
    cur.count += 1;
    cur.totalQty = round3(cur.totalQty + s.allocationQty);
    byMemo.set(key, cur);
  }
  return Array.from(byMemo.values()).sort((a, b) => a.date.localeCompare(b.date));
}

export function buildOmsAllocOrderRows(
  db: AllocationDatabase,
  category: 'ওএমএস সাধারণ' | 'ওএমএস বিশেষ' | 'ওএমএস টিসিবি',
  month0: number,
  year: number,
  unoMemo: string
) {
  const rateRec = db.rates.find((r) => r.category === category);
  const effUnitPrice = rateRec && rateRec.unitPrice > 0 ? rateRec.unitPrice : 26000;
  const effTaxRate = rateRec && rateRec.taxRate >= 0 ? rateRec.taxRate : 0;

  const trimmedUnoMemo = (unoMemo || '').trim();
  if (!trimmedUnoMemo) {
    return {
      monthName: BENGALI_MONTHS[month0] || '',
      year,
      memo: '',
      memoDate: '',
      sourceDcMemo: '',
      sourceDcDate: '',
      sourceDfcMemo: '',
      sourceDfcDate: '',
      sourceMinistryMemo: '',
      sourceMinistryDate: '',
      unitPrice: effUnitPrice,
      taxRate: effTaxRate,
      rows: [],
    };
  }

  const findDealerMatches = (name: string, salesCenter?: string, unionName?: string) => {
    const omsRec =
      db.omsDealers.find((d) => namesLikelyMatch(d.dealerName, name)) ||
      db.omsDealers.find(
        (d) =>
          Boolean(salesCenter && d.salesCenter) &&
          namesFuzzyMatch(d.salesCenter, salesCenter!) &&
          unionsLikelyMatch(
            d.areaType === 'পৌরসভা' ? d.municipality : d.unionName,
            unionName || ''
          )
      );
    const foodRec =
      category === 'ওএমএস টিসিবি' || !omsRec
        ? db.foodDealers.find((d) => namesLikelyMatch(d.dealerName, name)) ||
          db.foodDealers.find(
            (d) =>
              Boolean(salesCenter && d.salesCenter) &&
              namesFuzzyMatch(d.salesCenter, salesCenter!) &&
              unionsLikelyMatch(d.unionName, unionName || '')
          )
        : undefined;
    return { omsRec, foodRec };
  };

  const targetMemoNorm = normMemoNo(trimmedUnoMemo);
  const matchedSubs = db.dealerSuballocations.filter((s) => {
    if (s.category !== category) return false;
    const my = parseMonthYear(s.unoDate);
    if (!my || my.year !== year || my.month !== month0) return false;
    return normMemoNo(s.unoMemo) === targetMemoNorm || memosLikelyMatch(s.unoMemo, trimmedUnoMemo);
  });

  const rows = matchedSubs.map((s) => {
    const { omsRec, foodRec } = findDealerMatches(s.dealerName, s.salesCenter, s.unionName);
    const dealerName = omsRec?.dealerName || foodRec?.dealerName || s.dealerName;
    const settingsUnion = omsRec
      ? omsRec.areaType === 'পৌরসভা'
        ? omsRec.municipality
        : omsRec.unionName
      : foodRec?.unionName;
    const union = settingsUnion || s.unionName || '—';
    const salesCenter = omsRec?.salesCenter || foodRec?.salesCenter || s.salesCenter || '—';
    const mobile = omsRec?.mobile || foodRec?.mobile || '';

    const allocQty = round3(s.allocationQty);
    let cards = 0;
    let openingStockQty = 0;
    let totalRiceRequirement = allocQty;

    if (category === 'ওএমএস টিসিবি') {
      // STRICT ISOLATION: use tcbCardCount ONLY (never foodDealer.cardCount)
      const settingsTcbCards = omsRec?.tcbCardCount ?? foodRec?.tcbCardCount ?? 0;
      cards = settingsTcbCards > 0 ? settingsTcbCards : Math.max(0, s.cardCount);
      const explicitUnusedRice = s.unusedCards > 0 ? round3(s.unusedCards * 0.005) : 0;
      const explicitOpening = s.openingStock && s.openingStock > 0 ? round3(s.openingStock) : 0;
      totalRiceRequirement =
        cards > 0
          ? round3(cards * 0.005)
          : round3(allocQty + Math.max(explicitUnusedRice, explicitOpening));
      openingStockQty =
        explicitUnusedRice > 0
          ? explicitUnusedRice
          : explicitOpening > 0
            ? explicitOpening
            : totalRiceRequirement > allocQty
              ? round3(totalRiceRequirement - allocQty)
              : 0;
    } else {
      openingStockQty = round3(Math.max(0, s.openingStock || 0));
      totalRiceRequirement = round3(allocQty + openingStockQty);
    }

    const amount = round2(s.totalAmount > 0 ? s.totalAmount : allocQty * effUnitPrice);
    const tax = round2(s.taxAmount > 0 ? s.taxAmount : allocQty * effTaxRate);

    return {
      subId: s.id,
      union,
      salesCenter,
      dealerName,
      mobile,
      cards,
      openingStockQty,
      totalRiceRequirement,
      qty: allocQty,
      amount,
      tax,
      comment: '',
      unoDate: s.unoDate,
      sourceMemo: s.sourceMemo,
    };
  });

  const unionOrder = db.officer.unions || [];
  rows.sort((a, b) => {
    const ia = unionOrder.indexOf(a.union);
    const ib = unionOrder.indexOf(b.union);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.union.localeCompare(b.union);
  });

  const firstSourceMemo = rows.find((r) => r.sourceMemo)?.sourceMemo || '';
  const dcOrDfcRec = db.dcAllocations.find((d) => {
    if (d.category !== category) return false;
    if (firstSourceMemo) {
      return (
        d.dcMemo === firstSourceMemo ||
        normMemoNo(d.dcMemo) === normMemoNo(firstSourceMemo) ||
        memosLikelyMatch(d.dcMemo, firstSourceMemo)
      );
    }
    const my = parseMonthYear(d.dcMemoDate);
    return Boolean(my && my.year === year && my.month === month0);
  });

  const dcFallbackByMonth =
    category === 'ওএমএস টিসিবি' && !dcOrDfcRec
      ? db.dcAllocations.find((d) => {
          if (d.category !== category) return false;
          const my = parseMonthYear(d.dcMemoDate);
          return Boolean(my && my.year === year && my.month === month0);
        })
      : undefined;

  const dcRec = category === 'ওএমএস টিসিবি' ? dcOrDfcRec || dcFallbackByMonth : undefined;
  const nonTcbDcRec = category !== 'ওএমএস টিসিবি' ? dcOrDfcRec : undefined;

  const ministryLookupMemo = dcRec?.ministryMemo || nonTcbDcRec?.ministryMemo || firstSourceMemo;
  const minRec =
    (ministryLookupMemo
      ? db.ministryAllocations.find(
          (m) =>
            m.category === category &&
            (m.memoNo === ministryLookupMemo ||
              normMemoNo(m.memoNo) === normMemoNo(ministryLookupMemo) ||
              memosLikelyMatch(m.memoNo, ministryLookupMemo))
        )
      : undefined) ||
    db.ministryAllocations.find((m) => {
      if (m.category !== category) return false;
      const my = parseMonthYear(m.memoDate);
      return Boolean(my && my.year === year && my.month === month0);
    });

  const resolvedMinistryMemo = minRec?.memoNo || ministryLookupMemo || '';
  const resolvedMinistryDate = minRec?.memoDate || '';

  let sourceDfcMemo = '';
  let sourceDfcDate = '';
  if (category === 'ওএমএস টিসিবি') {
    const tcbSub =
      firstSourceMemo &&
      dcRec?.dcMemo &&
      normMemoNo(firstSourceMemo) !== normMemoNo(dcRec.dcMemo) &&
      !memosLikelyMatch(firstSourceMemo, dcRec.dcMemo)
        ? firstSourceMemo
        : '';
    sourceDfcMemo = tcbSub;
    sourceDfcDate = tcbSub ? rows[0]?.unoDate || '' : '';
  } else {
    const dfcFromSub =
      firstSourceMemo &&
      resolvedMinistryMemo &&
      normMemoNo(firstSourceMemo) !== normMemoNo(resolvedMinistryMemo) &&
      !memosLikelyMatch(firstSourceMemo, resolvedMinistryMemo)
        ? firstSourceMemo
        : '';
    sourceDfcMemo = nonTcbDcRec?.dcMemo || dfcFromSub || '';
    sourceDfcDate = nonTcbDcRec?.dcMemoDate || (dfcFromSub ? rows[0]?.unoDate || '' : '');
  }

  return {
    monthName: BENGALI_MONTHS[month0] || '',
    year,
    memo: trimmedUnoMemo,
    memoDate: rows[0]?.unoDate || '',
    sourceDcMemo: dcRec?.dcMemo || (category === 'ওএমএস টিসিবি' ? firstSourceMemo : ''),
    sourceDcDate: dcRec?.dcMemoDate || '',
    sourceDfcMemo,
    sourceDfcDate,
    sourceMinistryMemo: resolvedMinistryMemo,
    sourceMinistryDate: resolvedMinistryDate,
    unitPrice: effUnitPrice,
    taxRate: effTaxRate,
    rows,
  };
}

export function getPendingDORecipients(db: AllocationDatabase) {
  const { remainingBySubKey } = matchDOsToSuballocations(db);

  const out: Array<{
    key: string;
    subId: string;
    category: AllocationCategory;
    unionName: string;
    recipientName: string;
    salesCenter: string;
    ministryMemo: string;
    dcMemo: string;
    unoMemo: string;
    unoDate: string;
    projectNo: string;
    allocated: number;
    remaining: number;
  }> = [];

  for (const s of db.dealerSuballocations) {
    const remaining = remainingBySubKey.get(`ds:${s.id}`) ?? s.allocationQty;
    if (remaining > 0.0005) {
      const catMeta = CATEGORY_BY_NAME[s.category];
      const dcRec = catMeta?.requiresDC
        ? db.dcAllocations.find(
            (d) => d.category === s.category && memosLikelyMatch(d.dcMemo, s.sourceMemo)
          )
        : undefined;
      out.push({
        key: `ds-${s.id}-${out.length}`,
        subId: s.id,
        category: s.category,
        unionName: s.unionName,
        recipientName: s.dealerName,
        salesCenter: s.salesCenter,
        ministryMemo: dcRec ? dcRec.ministryMemo : s.sourceMemo,
        dcMemo: dcRec ? dcRec.dcMemo : '',
        unoMemo: s.unoMemo,
        unoDate: s.unoDate,
        projectNo: '',
        allocated: s.allocationQty,
        remaining,
      });
    }
  }

  for (const u of db.unoSuballocations) {
    const remaining = remainingBySubKey.get(`uno:${u.id}`) ?? u.allocationQty;
    if (remaining > 0.0005) {
      out.push({
        key: `uno-${u.id}-${out.length}`,
        subId: u.id,
        category: u.category,
        unionName: u.unionName,
        recipientName: u.recipientName,
        salesCenter: u.projectName || '',
        ministryMemo: u.ministryMemo,
        dcMemo: u.dcMemo,
        unoMemo: u.unoMemo,
        unoDate: u.unoDate,
        projectNo: u.projectNo || '',
        allocated: u.allocationQty,
        remaining,
      });
    }
  }

  return out;
}



