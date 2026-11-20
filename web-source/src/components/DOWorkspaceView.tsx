import React, { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  ArrowLeft,
  BookOpen,
  Check,
  CheckCircle2,
  FileCheck2,
  Link2,
  Lock,
  Pencil,
  Search,
  Sparkles,
  Trash2,
  Unlock,
  X,
} from 'lucide-react';
import {
  CATEGORIES,
  type AllocationCategory,
  type AllocationDatabase,
  type ChallanRecord,
} from '../types/allocation';
import {
  getDoBookRegisterStats,
  getPendingDORecipients,
  getUnusedDoNumbers,
} from '../shared/allocReportEngine';
import {
  bnToEnDigits,
  formatBnDate,
  formatBnInt,
  formatBnMoney,
  formatBnQty,
  isChallanOnOrBeforeDo,
  isSamePartyName,
  normChallanNo,
  normDoNo,
  normMemoNo,
  parseNum,
  round3,
  toBnDigits,
  toIsoDateStr,
  todayIsoDate,
} from '../shared/bnUtils';
import type { ActionResult } from '../store/useAllocationStore';

export type DOTabKey = 'do-book' | 'do-issue';

interface DOWorkspaceViewProps {
  db: AllocationDatabase;
  activeTab: DOTabKey;
  onChangeTab: (tab: DOTabKey) => void;
  initialCategory?: AllocationCategory | 'ALL';
  lockedCategory?: AllocationCategory | null;
  embedded?: boolean;
  onBack?: () => void;
  onSaveDOBook: (input: { bookNo: number; startPage: number; endPage: number }) => ActionResult;
  onUpdateDOBook?: (
    id: string,
    input: { bookNo: number; startPage: number; endPage: number }
  ) => ActionResult;
  onDeleteDOBook: (id: string) => void;
  verifyChallan: (
    category: AllocationCategory,
    dealerName: string,
    qty: number,
    opts?: {
      subId?: string;
      unoDate?: string;
      unoMemo?: string;
      ministryMemo?: string;
      issueDate?: string;
      excludeDoId?: string;
      excludeChallanIds?: Set<string>;
    }
  ) => {
    applicable: boolean;
    ok: boolean;
    message?: string;
    requiredAmount?: number;
    requiredTaxAmount?: number;
    mainChallan?: ChallanRecord;
    taxChallan?: ChallanRecord;
    errors?: string[];
  };
  onIssueDOBatch: (
    items: Array<{
      category: AllocationCategory;
      unionName: string;
      recipientName: string;
      salesCenter?: string;
      allocationQty: number;
      ministryMemo: string;
      dcMemo: string;
      unoMemo: string;
      unoDate?: string;
      subId?: string;
      projectNo?: string;
      doNo: string;
      issueDate: string;
    }>
  ) => ActionResult;
  onUpdateDOIssue?: (
    id: string,
    patch: { doNo: string; issueDate: string; allocationQty: number }
  ) => ActionResult;
  onDeleteDOIssue: (id: string) => void;
  onRelinkChallan?: (
    challanId: string,
    target: { dealerAllocationId?: string; doNo?: string; clearLink?: boolean }
  ) => ActionResult;
}

export const DOWorkspaceView: React.FC<DOWorkspaceViewProps> = ({
  db,
  activeTab,
  onChangeTab,
  initialCategory = 'ALL',
  lockedCategory = null,
  embedded = false,
  onBack,
  onSaveDOBook,
  onUpdateDOBook,
  onDeleteDOBook,
  verifyChallan,
  onIssueDOBatch,
  onUpdateDOIssue,
  onDeleteDOIssue,
  onRelinkChallan,
}) => {
  const [feedback, setFeedback] = useState<{ ok: boolean; text: string; warning?: string } | null>(null);

  // DO Book form
  const [editingBookId, setEditingBookId] = useState<string | null>(null);
  const [bookNo, setBookNo] = useState('');
  const [startPage, setStartPage] = useState('');
  const [endPage, setEndPage] = useState('');
  const [selectedBookId, setSelectedBookId] = useState<string>('');

  // DO Issue state
  const [doCatFilter, setDoCatFilter] = useState<string>(lockedCategory || initialCategory || 'ALL');
  const [doBatchDate, setDoBatchDate] = useState(todayIsoDate());
  const [pendingInputs, setPendingInputs] = useState<
    Record<string, { doNo: string; date: string }>
  >({});
  const [verifyTargetKey, setVerifyTargetKey] = useState<string | null>(null);
  const [showChallanLocks, setShowChallanLocks] = useState(false);
  const [challanLockSearch, setChallanLockSearch] = useState('');
  const [challanTargetDrafts, setChallanTargetDrafts] = useState<Record<string, string>>({});
  const [manualRelinkPickByRow, setManualRelinkPickByRow] = useState<Record<string, string>>({});
  const [editingDoChallanId, setEditingDoChallanId] = useState<string>('');
  const [doSearch, setDoSearch] = useState('');
  const [editingDoRow, setEditingDoRow] = useState<{
    id: string;
    doNo: string;
    issueDate: string;
    qty: string;
  } | null>(null);

  useEffect(() => {
    if (lockedCategory) {
      setDoCatFilter(lockedCategory);
    } else if (initialCategory) {
      setDoCatFilter(initialCategory);
    }
  }, [initialCategory, lockedCategory]);

  const unusedList = useMemo(() => getUnusedDoNumbers(db), [db]);
  const bookStats = useMemo(() => getDoBookRegisterStats(db), [db]);
  const pendingRecipients = useMemo(() => getPendingDORecipients(db), [db]);

  const showRes = (res: ActionResult) => {
    setFeedback({
      ok: res.ok,
      text: res.ok ? res.message || 'সম্পন্ন হয়েছে।' : res.error || 'ত্রুটি ঘটেছে।',
      warning: res.warning,
    });
  };

  const resetBookForm = () => {
    setEditingBookId(null);
    setBookNo('');
    setStartPage('');
    setEndPage('');
  };

  const handleBookSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (editingBookId && onUpdateDOBook) {
      const res = onUpdateDOBook(editingBookId, {
        bookNo: parseNum(bookNo),
        startPage: parseNum(startPage),
        endPage: parseNum(endPage),
      });
      showRes(res);
      if (res.ok) {
        resetBookForm();
      }
      return;
    }
    const res = onSaveDOBook({
      bookNo: parseNum(bookNo),
      startPage: parseNum(startPage),
      endPage: parseNum(endPage),
    });
    showRes(res);
    if (res.ok) {
      resetBookForm();
    }
  };

  const filteredPending = useMemo(() => {
    return pendingRecipients.filter(
      (p) => doCatFilter === 'ALL' || p.category === doCatFilter
    );
  }, [pendingRecipients, doCatFilter]);

  const issuedSet = useMemo(
    () => new Set(db.doIssues.map((d) => normDoNo(d.doNo)).filter(Boolean)),
    [db.doIssues]
  );

  const sortedUnusedList = useMemo(
    () => [...unusedList].sort((a, b) => a.doNo - b.doNo),
    [unusedList]
  );

  const ministryGrainByKey = useMemo(() => {
    const map = new Map<string, string>();
    for (const m of db.ministryAllocations) {
      map.set(`${m.category}::${m.memoNo}`, m.grain);
      const nm = normMemoNo(m.memoNo);
      if (nm) map.set(`${m.category}::${nm}`, m.grain);
    }
    return map;
  }, [db.ministryAllocations]);

  const resolveItemGrain = (category: AllocationCategory, ministryMemo?: string): string => {
    if (category === 'কার্নেল') return 'কার্নেল';
    if (!ministryMemo) return '';
    return (
      ministryGrainByKey.get(`${category}::${ministryMemo}`) ||
      ministryGrainByKey.get(`${category}::${normMemoNo(ministryMemo)}`) ||
      ''
    );
  };

  const pendingGroups = useMemo(() => {
    const groups: Array<{
      groupKey: string;
      category: AllocationCategory;
      unoMemo: string;
      unoDate: string;
      grains: string[];
      items: typeof filteredPending;
      totalRemaining: number;
    }> = [];
    const byKey = new Map<string, (typeof groups)[number]>();
    for (const item of filteredPending) {
      const groupKey = `${item.category}::${normMemoNo(item.unoMemo)}::${item.unoDate}`;
      let grp = byKey.get(groupKey);
      if (!grp) {
        grp = {
          groupKey,
          category: item.category,
          unoMemo: item.unoMemo,
          unoDate: item.unoDate,
          grains: [],
          items: [],
          totalRemaining: 0,
        };
        byKey.set(groupKey, grp);
        groups.push(grp);
      }
      const g = resolveItemGrain(item.category, item.ministryMemo);
      if (g && !grp.grains.includes(g)) grp.grains.push(g);
      grp.items.push(item);
      grp.totalRemaining = round3(grp.totalRemaining + item.remaining);
    }
    return groups;
  }, [filteredPending, ministryGrainByKey]);

  const previewChecksByKey = useMemo(() => {
    const map = new Map<string, ReturnType<typeof verifyChallan>>();
    const usedIds = new Set<string>();
    const doNumVal = (no: string): number => {
      const digits = bnToEnDigits(no || '').replace(/\D/g, '');
      return digits ? Number(digits) : Number.MAX_SAFE_INTEGER;
    };
    const subChallanNormSet = new Set(
      db.challans.map((c) => normMemoNo(c.dealerAllocationId)).filter(Boolean)
    );
    const sameMonthUnlinkedChallans = db.challans.filter((c) => !c.doNo);
    const itemHasExactMonthChallan = (p: (typeof filteredPending)[number]): boolean => {
      const uIso = toIsoDateStr(p.unoDate);
      const uMonth = uIso ? uIso.slice(0, 7) : '';
      if (!uMonth) return false;
      return sameMonthUnlinkedChallans.some((c) => {
        if (c.category !== p.category) return false;
        const cIso = toIsoDateStr(c.challanDate);
        if (!cIso || cIso.slice(0, 7) !== uMonth) return false;
        if (c.dealerAllocationId) {
          return Boolean(p.subId && normMemoNo(c.dealerAllocationId) === normMemoNo(p.subId));
        }
        return Boolean(c.dealerName && isSamePartyName(c.dealerName, p.recipientName));
      });
    };

    // Pass 1: Items with explicit draft DO number, direct subId challan, or same-month challan match
    // so an earlier month's unlinked pending row can NEVER preview-steal a later month's challan.
    const pass1Items: typeof filteredPending = [];
    const pass2Items: typeof filteredPending = [];
    for (const p of filteredPending) {
      const hasDraftDo = Boolean(pendingInputs[p.key]?.doNo?.trim());
      const hasDirectSub = Boolean(p.subId && subChallanNormSet.has(normMemoNo(p.subId)));
      if (hasDraftDo || hasDirectSub || itemHasExactMonthChallan(p)) {
        pass1Items.push(p);
      } else {
        pass2Items.push(p);
      }
    }

    const sortByIssueAndUno = (list: typeof filteredPending) =>
      [...list].sort((a, b) => {
        const draftA = pendingInputs[a.key]?.doNo?.trim() ? 0 : 1;
        const draftB = pendingInputs[b.key]?.doNo?.trim() ? 0 : 1;
        if (draftA !== draftB) return draftA - draftB;
        const directA = a.subId && subChallanNormSet.has(normMemoNo(a.subId)) ? 0 : 1;
        const directB = b.subId && subChallanNormSet.has(normMemoNo(b.subId)) ? 0 : 1;
        if (directA !== directB) return directA - directB;
        const dA = toIsoDateStr(pendingInputs[a.key]?.date || doBatchDate) || '9999-12-31';
        const dB = toIsoDateStr(pendingInputs[b.key]?.date || doBatchDate) || '9999-12-31';
        if (dA !== dB) return dA.localeCompare(dB);
        const uA = toIsoDateStr(a.unoDate) || '9999-12-31';
        const uB = toIsoDateStr(b.unoDate) || '9999-12-31';
        if (uA !== uB) return uA.localeCompare(uB);
        return doNumVal(pendingInputs[a.key]?.doNo || '') - doNumVal(pendingInputs[b.key]?.doNo || '');
      });

    for (const p of [...sortByIssueAndUno(pass1Items), ...sortByIssueAndUno(pass2Items)]) {
      const targetDate = pendingInputs[p.key]?.date || doBatchDate;
      const check = verifyChallan(p.category, p.recipientName, p.remaining, {
        subId: p.subId,
        unoDate: p.unoDate,
        unoMemo: p.unoMemo,
        ministryMemo: p.ministryMemo,
        issueDate: targetDate,
        excludeChallanIds: usedIds,
      });
      if (check.mainChallan) usedIds.add(check.mainChallan.id);
      if (check.taxChallan) usedIds.add(check.taxChallan.id);
      map.set(p.key, check);
    }
    return map;
  }, [db.challans, filteredPending, pendingInputs, doBatchDate, verifyChallan]);

  const buildUsedDoDraftSet = (
    draftMap: Record<string, { doNo: string; date: string }>
  ): Set<string> => {
    const used = new Set<string>(issuedSet);
    for (const val of Object.values(draftMap)) {
      const cleaned = normDoNo(val?.doNo);
      if (cleaned) used.add(cleaned);
    }
    return used;
  };

  const draftUsedSet = useMemo(
    () => buildUsedDoDraftSet(pendingInputs),
    [issuedSet, pendingInputs]
  );

  const availableDatalistDos = useMemo(
    () =>
      sortedUnusedList.filter((u) => {
        const key = normDoNo(u.doNo);
        return Boolean(key) && !draftUsedSet.has(key);
      }),
    [sortedUnusedList, draftUsedSet]
  );

  const findNextLowestUnusedDo = (usedSet: Set<string>): string | null => {
    for (const u of sortedUnusedList) {
      const candidateKey = normDoNo(u.doNo);
      if (candidateKey && !usedSet.has(candidateKey)) return String(u.doNo);
    }
    return null;
  };

  const toggleSingleRowIssue = (
    p: (typeof pendingRecipients)[number],
    checked: boolean
  ) => {
    if (!checked) {
      setPendingInputs((prev) => {
        const copy = { ...prev };
        delete copy[p.key];
        return copy;
      });
      return;
    }
    setPendingInputs((prev) => {
      const existing = prev[p.key]?.doNo?.trim();
      if (existing) return prev;
      const usedSet = buildUsedDoDraftSet(prev);
      const nextDo = findNextLowestUnusedDo(usedSet);
      if (!nextDo) {
        setFeedback({
          ok: false,
          text: 'খালি ডিও বইয়ের পাতা অবশিষ্ট নেই। নতুন ডিও বই এন্ট্রি করুন অথবা হাতে ডিও নং লিখুন।',
        });
        return prev;
      }
      return {
        ...prev,
        [p.key]: {
          doNo: nextDo,
          date: prev[p.key]?.date || doBatchDate,
        },
      };
    });
  };

  const toggleGroupSelectAll = (
    items: typeof filteredPending,
    checked: boolean
  ) => {
    if (!checked) {
      setPendingInputs((prev) => {
        const copy = { ...prev };
        for (const item of items) {
          delete copy[item.key];
        }
        return copy;
      });
      return;
    }
    setPendingInputs((prev) => {
      const next = { ...prev };
      const usedSet = buildUsedDoDraftSet(next);
      let missingCount = 0;
      for (const item of items) {
        if (next[item.key]?.doNo?.trim()) continue;
        const candidate = findNextLowestUnusedDo(usedSet);
        if (candidate) {
          usedSet.add(normDoNo(candidate));
          next[item.key] = {
            doNo: candidate,
            date: next[item.key]?.date || doBatchDate,
          };
        } else {
          missingCount++;
        }
      }
      if (missingCount > 0) {
        setFeedback({
          ok: false,
          text: `${formatBnInt(missingCount)} জন প্রাপকের জন্য খালি ডিও পাতা পাওয়া যায়নি।`,
        });
      }
      return next;
    });
  };

  const autoAssignUnusedDos = () => {
    const next: Record<string, { doNo: string; date: string }> = { ...pendingInputs };
    const usedInDraft = buildUsedDoDraftSet(next);
    for (const p of filteredPending) {
      if (next[p.key]?.doNo?.trim()) continue;
      const assigned = findNextLowestUnusedDo(usedInDraft);
      if (!assigned) break;
      usedInDraft.add(normDoNo(assigned));
      next[p.key] = {
        doNo: assigned,
        date: next[p.key]?.date || doBatchDate,
      };
    }
    setPendingInputs(next);
  };

  const handleSaveBatchDos = () => {
    const toIssue = filteredPending
      .filter((p) => Boolean(pendingInputs[p.key]?.doNo?.trim()))
      .map((p) => ({
        category: p.category,
        unionName: p.unionName,
        recipientName: p.recipientName,
        salesCenter: p.salesCenter,
        allocationQty: p.remaining,
        ministryMemo: p.ministryMemo,
        dcMemo: p.dcMemo,
        unoMemo: p.unoMemo,
        unoDate: p.unoDate,
        subId: p.subId,
        projectNo: p.projectNo || undefined,
        doNo: pendingInputs[p.key].doNo.trim(),
        issueDate: pendingInputs[p.key].date || doBatchDate,
      }));
    const res = onIssueDOBatch(toIssue);
    showRes(res);
    if (res.ok) {
      setPendingInputs({});
    }
  };

  const handleIssueSingleRow = (p: (typeof pendingRecipients)[number]) => {
    const val = pendingInputs[p.key];
    const dNo = val?.doNo?.trim() || '';
    const res = onIssueDOBatch([
      {
        category: p.category,
        unionName: p.unionName,
        recipientName: p.recipientName,
        salesCenter: p.salesCenter,
        allocationQty: p.remaining,
        ministryMemo: p.ministryMemo,
        dcMemo: p.dcMemo,
        unoMemo: p.unoMemo,
        unoDate: p.unoDate,
        subId: p.subId,
        projectNo: p.projectNo || undefined,
        doNo: dNo,
        issueDate: val?.date || doBatchDate,
      },
    ]);
    showRes(res);
    if (res.ok) {
      setPendingInputs((prev) => {
        const copy = { ...prev };
        delete copy[p.key];
        return copy;
      });
    }
  };

  const selectedBook =
    db.doBooks.find((b) => b.id === selectedBookId) || db.doBooks[0] || null;

  const filteredIssuedDos = useMemo(() => {
    const q = doSearch.trim().toLowerCase();
    const qDo = normDoNo(doSearch);
    return db.doIssues.filter((d) => {
      if (doCatFilter !== 'ALL' && d.category !== doCatFilter) return false;
      if (!q) return true;
      return (
        d.doNo.toLowerCase().includes(q) ||
        toBnDigits(d.doNo).toLowerCase().includes(q) ||
        (Boolean(qDo) && normDoNo(d.doNo).includes(qDo)) ||
        d.recipientName.toLowerCase().includes(q) ||
        d.category.toLowerCase().includes(q) ||
        d.unionName.toLowerCase().includes(q)
      );
    });
  }, [db.doIssues, doSearch, doCatFilter]);

  return (
    <div className="space-y-4 font-nikosh">
      {!embedded && (
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-2">
            {onBack && (
              <button
                type="button"
                onClick={onBack}
                className="min-h-[44px] inline-flex items-center gap-1.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 px-3 py-2 text-xs font-bold text-slate-700 cursor-pointer"
              >
                <ArrowLeft className="h-4 w-4 text-[#E2136E]" />
                <span>ফিরে যান</span>
              </button>
            )}
            {lockedCategory ? (
              <h2 className="text-sm sm:text-base font-bold text-slate-900">
                {lockedCategory} — ডিও এন্ট্রি ও তালিকা
              </h2>
            ) : (
              <>
                <button
                  type="button"
                  onClick={() => onChangeTab('do-issue')}
                  className={`min-h-[44px] inline-flex items-center gap-1.5 rounded-xl px-4 py-2 text-xs font-bold border cursor-pointer ${
                    activeTab === 'do-issue'
                      ? 'bg-[#E2136E] text-white border-[#E2136E]'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-rose-50/50'
                  }`}
                >
                  <FileCheck2 className="h-4 w-4" />
                  <span>ডিও ইস্যু ({formatBnInt(filteredPending.length)} পেন্ডিং)</span>
                </button>

                <button
                  type="button"
                  onClick={() => onChangeTab('do-book')}
                  className={`min-h-[44px] inline-flex items-center gap-1.5 rounded-xl px-4 py-2 text-xs font-bold border cursor-pointer ${
                    activeTab === 'do-book'
                      ? 'bg-[#E2136E] text-white border-[#E2136E]'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-rose-50/50'
                  }`}
                >
                  <BookOpen className="h-4 w-4" />
                  <span>ডিও বই এন্ট্রি ({formatBnInt(unusedList.length)} খালি)</span>
                </button>
              </>
            )}
          </div>
        </div>
      )}

      {feedback && (
        <div className="space-y-2">
          <div
            className={`rounded-xl border px-4 py-2.5 text-xs font-bold flex items-center gap-2 ${
              feedback.ok
                ? 'border-emerald-200 bg-emerald-50 text-emerald-900'
                : 'border-rose-200 bg-rose-50 text-rose-900'
            }`}
          >
            <CheckCircle2 className="h-4 w-4 shrink-0" />
            <span>{feedback.text}</span>
          </div>
          {feedback.warning && (
            <div className="rounded-xl border border-amber-300 bg-amber-50 px-4 py-2.5 text-xs font-bold text-amber-900 flex items-start gap-2">
              <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5 text-amber-600" />
              <span>চালান সিগন্যাল (ডিও সংরক্ষিত হয়েছে): {feedback.warning}</span>
            </div>
          )}
        </div>
      )}

      {/* TAB 1: DO BOOK ENTRY */}
      {!lockedCategory && activeTab === 'do-book' && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
          <form
            onSubmit={handleBookSubmit}
            className="lg:col-span-4 rounded-2xl border border-slate-200 bg-white p-4 space-y-3"
          >
            <div className="flex items-center justify-between gap-2">
              <div className="text-sm font-bold text-slate-900">
                {editingBookId ? 'ডিও বই এডিট করুন' : 'নতুন ডিও বই যুক্ত করুন'}
              </div>
              {editingBookId && (
                <button
                  type="button"
                  onClick={resetBookForm}
                  className="inline-flex items-center gap-1 rounded-lg border border-slate-300 bg-slate-50 px-2 py-1 text-[11px] font-bold text-slate-700 cursor-pointer"
                >
                  <X className="h-3 w-3" /> বাতিল
                </button>
              )}
            </div>
            <div className="space-y-2.5 text-xs">
              <div>
                <label className="block font-bold text-slate-700 mb-1">বই নম্বর *</label>
                <input
                  type="number"
                  required
                  value={bookNo}
                  onChange={(e) => setBookNo(e.target.value)}
                  placeholder="যেমন: ১০২"
                  className="w-full rounded-lg border border-slate-300 px-3 py-2"
                />
              </div>
              <div className="grid grid-cols-2 gap-2.5">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">শুরু পাতা নং *</label>
                  <input
                    type="number"
                    required
                    value={startPage}
                    onChange={(e) => setStartPage(e.target.value)}
                    placeholder="২৪০৪০১"
                    className="w-full rounded-lg border border-slate-300 px-3 py-2"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">শেষ পাতা নং *</label>
                  <input
                    type="number"
                    required
                    value={endPage}
                    onChange={(e) => setEndPage(e.target.value)}
                    placeholder="২৪০৪৫০"
                    className="w-full rounded-lg border border-slate-300 px-3 py-2"
                  />
                </div>
              </div>
            </div>
            <button
              type="submit"
              className="w-full min-h-[44px] rounded-xl bg-[#E2136E] hover:bg-[#B80D58] text-white text-xs font-bold cursor-pointer"
            >
              {editingBookId ? 'ডিও বই হালনাগাদ করুন' : 'ডিও বই সংরক্ষণ করুন'}
            </button>

            <div className="rounded-xl border border-slate-200 bg-slate-50 p-3 text-xs space-y-1 tabular-nums">
              <div>মোট ডিও পাতা: <strong>{formatBnInt(bookStats.summary.totalDO)}</strong></div>
              <div>ব্যবহৃত পাতা: <strong className="text-[#E2136E]">{formatBnInt(bookStats.summary.totalUsed)}</strong></div>
              <div>অব্যবহৃত পাতা: <strong className="text-amber-700">{formatBnInt(bookStats.summary.totalUnused)}</strong></div>
            </div>
          </form>

          <div className="lg:col-span-8 space-y-4">
            <div className="rounded-2xl border border-slate-200 bg-white overflow-hidden">
              <div className="px-4 py-3 border-b border-slate-200 text-xs font-bold text-slate-900">
                নিবন্ধিত ডিও বই তালিকা (ব্যবহৃত / খালি পাতার হিসাব)
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs border-collapse tabular-nums">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-slate-700 font-bold">
                      <th className="py-2 px-3 text-center">বই নং</th>
                      <th className="py-2 px-3 text-center">সিরিয়াল রেঞ্জ</th>
                      <th className="py-2 px-3 text-center">মোট পাতা</th>
                      <th className="py-2 px-3 text-center">ব্যবহৃত</th>
                      <th className="py-2 px-3 text-center">অবশিষ্ট</th>
                      <th className="py-2 px-3 text-center">কাজ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {db.doBooks.map((b) => {
                      const stat = bookStats.usedRows.find((r) => r.bookNo === b.bookNo);
                      const used = stat?.usedCount || 0;
                      const rem = Math.max(0, b.totalCount - used);
                      return (
                        <tr key={b.id} className={selectedBook?.id === b.id ? 'bg-rose-50/40' : ''}>
                          <td className="py-2 px-3 text-center font-bold">{toBnDigits(b.bookNo)}</td>
                          <td className="py-2 px-3 text-center">
                            {toBnDigits(b.startPage)} — {toBnDigits(b.endPage)}
                          </td>
                          <td className="py-2 px-3 text-center">{formatBnInt(b.totalCount)}</td>
                          <td className="py-2 px-3 text-center font-semibold text-[#E2136E]">
                            {formatBnInt(used)}
                          </td>
                          <td className="py-2 px-3 text-center font-bold text-amber-700">
                            {formatBnInt(rem)}
                          </td>
                          <td className="py-2 px-3 text-center whitespace-nowrap">
                            <div className="inline-flex items-center gap-1">
                              <button
                                type="button"
                                onClick={() => setSelectedBookId(b.id)}
                                className="rounded border border-slate-300 bg-white px-2 py-1 text-[11px] font-bold text-slate-700 cursor-pointer"
                              >
                                পাতা দেখুন
                              </button>
                              {onUpdateDOBook && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    setEditingBookId(b.id);
                                    setBookNo(String(b.bookNo));
                                    setStartPage(String(b.startPage));
                                    setEndPage(String(b.endPage));
                                  }}
                                  className="p-1.5 text-[#E2136E] hover:bg-rose-50 rounded cursor-pointer"
                                  title="ডিও বই এডিট করুন"
                                  aria-label="ডিও বই এডিট করুন"
                                >
                                  <Pencil className="h-3.5 w-3.5" />
                                </button>
                              )}
                              <button
                                type="button"
                                onClick={() => {
                                  if (editingBookId === b.id) resetBookForm();
                                  onDeleteDOBook(b.id);
                                }}
                                className="p-1.5 text-rose-600 hover:bg-rose-50 rounded cursor-pointer"
                                title="ডিও বই মুছুন"
                                aria-label="ডিও বই মুছুন"
                              >
                                <Trash2 className="h-3.5 w-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {selectedBook && (
              <div className="rounded-2xl border border-slate-200 bg-white p-4 space-y-2.5">
                <div className="flex items-center justify-between text-xs font-bold text-slate-900">
                  <span>বই নং {toBnDigits(selectedBook.bookNo)} — সকল পাতার মানচিত্র</span>
                  <span className="text-[11px] font-normal text-slate-500">
                    রঙিন = ইস্যুকৃত · সাদা = অব্যবহৃত
                  </span>
                </div>
                <div className="grid grid-cols-5 sm:grid-cols-10 gap-1.5 text-[11px] tabular-nums">
                  {Array.from(
                    { length: Math.min(100, selectedBook.endPage - selectedBook.startPage + 1) },
                    (_, i) => selectedBook.startPage + i
                  ).map((p) => {
                    const used = issuedSet.has(normDoNo(p));
                    return (
                      <div
                        key={p}
                        className={`rounded border px-1.5 py-1 text-center font-bold ${
                          used
                            ? 'bg-rose-50 border-rose-300 text-[#E2136E]'
                            : 'bg-slate-50 border-slate-200 text-slate-600'
                        }`}
                      >
                        {toBnDigits(p)}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* TAB 2: DO ISSUE & VERIFICATION */}
      {(lockedCategory || activeTab === 'do-issue') && (
        <div className="space-y-4">
          <div className="rounded-2xl border border-slate-200 bg-white p-4 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2.5">
              <div>
                <h3 className="text-sm font-bold text-slate-900">
                  {lockedCategory ? `${lockedCategory} — ` : ''}পেন্ডিং ডিও প্রাপক তালিকা ({formatBnInt(filteredPending.length)} জন)
                </h3>
                <p className="text-xs text-slate-500">
                  খালি পাতা থেকে স্বয়ংক্রিয়ভাবে বা হাতে ডিও নম্বর বসিয়ে চালান যাচাইপূর্বক ডিও ইস্যু করুন।
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2 text-xs">
                {!lockedCategory && (
                  <select
                    value={doCatFilter}
                    onChange={(e) => setDoCatFilter(e.target.value)}
                    className="min-h-[44px] rounded-xl border border-slate-300 px-2.5 py-2 bg-white font-bold"
                  >
                    <option value="ALL">সকল খাত ({formatBnInt(pendingRecipients.length)})</option>
                    {CATEGORIES.map((c) => (
                      <option key={c.id} value={c.category}>
                        {c.category}
                      </option>
                    ))}
                  </select>
                )}

                <input
                  type="date"
                  value={doBatchDate}
                  onChange={(e) => setDoBatchDate(e.target.value)}
                  className="min-h-[44px] rounded-xl border border-slate-300 px-2.5 py-1.5"
                />

                <button
                  type="button"
                  onClick={autoAssignUnusedDos}
                  className="min-h-[44px] inline-flex items-center gap-1.5 rounded-xl border border-rose-300 bg-rose-50 hover:bg-rose-100 px-3 py-1.5 font-bold text-[#E2136E] cursor-pointer"
                >
                  <Sparkles className="h-3.5 w-3.5" />
                  <span>খালি ডিও নং বসান ({formatBnInt(unusedList.length)}টি বাকি)</span>
                </button>

                {onRelinkChallan && (
                  <button
                    type="button"
                    onClick={() => setShowChallanLocks((prev) => !prev)}
                    className={`min-h-[44px] inline-flex items-center gap-1.5 rounded-xl border px-3 py-1.5 font-bold cursor-pointer ${
                      showChallanLocks
                        ? 'border-slate-900 bg-slate-900 text-white'
                        : 'border-slate-300 bg-white hover:bg-slate-50 text-slate-800'
                    }`}
                  >
                    <Lock className="h-3.5 w-3.5" />
                    <span>চালান লক তালিকা</span>
                  </button>
                )}

                <button
                  type="button"
                  onClick={handleSaveBatchDos}
                  className="min-h-[44px] inline-flex items-center gap-1.5 rounded-xl bg-[#E2136E] hover:bg-[#B80D58] px-4 py-1.5 font-bold text-white cursor-pointer"
                >
                  <FileCheck2 className="h-3.5 w-3.5" />
                  <span>সব নির্বাচিত ডিও ইস্যু করুন</span>
                </button>
              </div>
            </div>

            {showChallanLocks && onRelinkChallan && (() => {
              const allAllocationOptions = [
                ...db.dealerSuballocations.map((s) => ({
                  id: s.id,
                  category: s.category,
                  partyName: s.dealerName,
                  date: s.unoDate,
                  memo: s.unoMemo,
                  qty: s.allocationQty,
                  label: `${s.dealerName} · ${s.unoDate ? formatBnDate(s.unoDate) : '—'} · স্মারক ${s.unoMemo ? toBnDigits(s.unoMemo) : '—'} (${formatBnQty(s.allocationQty)} মে.টন)`,
                })),
                ...db.unoSuballocations.map((u) => ({
                  id: u.id,
                  category: u.category,
                  partyName: u.recipientName,
                  date: u.unoDate,
                  memo: u.unoMemo,
                  qty: u.allocationQty,
                  label: `${u.recipientName} · ${u.unoDate ? formatBnDate(u.unoDate) : '—'} · স্মারক ${u.unoMemo ? toBnDigits(u.unoMemo) : '—'} (${formatBnQty(u.allocationQty)} মে.টন)`,
                })),
              ];
              const subLookup = new Map(allAllocationOptions.map((o) => [normMemoNo(o.id), o]));
              const qLock = challanLockSearch.trim().toLowerCase();
              const visibleChallans = db.challans.filter((c) => {
                if (doCatFilter !== 'ALL' && c.category !== doCatFilter) return false;
                if (!qLock) return true;
                const linkedSub = c.dealerAllocationId ? subLookup.get(normMemoNo(c.dealerAllocationId)) : undefined;
                return (
                  c.challanNo.toLowerCase().includes(qLock) ||
                  toBnDigits(c.challanNo).toLowerCase().includes(qLock) ||
                  (c.dealerName || '').toLowerCase().includes(qLock) ||
                  (linkedSub?.label || '').toLowerCase().includes(qLock) ||
                  (c.doNo || '').toLowerCase().includes(qLock)
                );
              });

              return (
                <div className="rounded-xl border border-slate-200 bg-slate-50/70 p-3 space-y-2.5">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                      <Lock className="h-3.5 w-3.5 text-[#E2136E]" />
                      <span>স্ক্যানকৃত চালান ও বরাদ্দ লক ব্যবস্থাপনা ({formatBnInt(visibleChallans.length)}টি)</span>
                    </div>
                    <div className="relative w-full sm:w-64">
                      <Search className="h-3.5 w-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                      <input
                        type="text"
                        value={challanLockSearch}
                        onChange={(e) => setChallanLockSearch(e.target.value)}
                        placeholder="চালান নং / ডিলার / স্মারক খুঁজুন"
                        className="w-full min-h-[34px] rounded-lg border border-slate-300 bg-white pl-8 pr-2.5 py-1 text-xs focus:outline-none focus:border-[#E2136E]"
                      />
                    </div>
                  </div>
                  <p className="text-[11px] text-slate-600">
                    যেকোনো চালান যে বরাদ্দের সাথে লক করা আছে সেখানেই স্থায়ী থাকবে। পরিবর্তন করতে চাইলে এখান থেকে ম্যানুয়ালি বরাদ্দ নির্বাচন করে লক করুন বা মুক্ত করুন।
                  </p>
                  <div className="max-h-72 overflow-y-auto divide-y divide-slate-200 rounded-lg border border-slate-200 bg-white">
                    {visibleChallans.length === 0 ? (
                      <div className="py-4 text-center text-xs text-slate-500">কোনো চালান পাওয়া যায়নি।</div>
                    ) : (
                      visibleChallans.slice(0, 80).map((c) => {
                        const lockedSub = c.dealerAllocationId ? subLookup.get(normMemoNo(c.dealerAllocationId)) : undefined;
                        const catSubs = allAllocationOptions.filter((o) => o.category === c.category);
                        const sortedSubs = [...catSubs].sort((a, b) => {
                          const sameA = c.dealerName && isSamePartyName(a.partyName, c.dealerName) ? 0 : 1;
                          const sameB = c.dealerName && isSamePartyName(b.partyName, c.dealerName) ? 0 : 1;
                          if (sameA !== sameB) return sameA - sameB;
                          return (b.date || '').localeCompare(a.date || '');
                        });
                        const draftTarget = challanTargetDrafts[c.id] ?? (c.dealerAllocationId || '');
                        return (
                          <div key={c.id} className="p-2.5 flex flex-col lg:flex-row lg:items-center justify-between gap-2 text-xs">
                            <div className="min-w-0 space-y-0.5">
                              <div className="flex flex-wrap items-center gap-1.5 font-bold text-slate-900">
                                <span className={c.challanType === 'মূল চালান' ? 'text-[#E2136E]' : 'text-sky-700'}>
                                  {c.challanType}
                                </span>
                                <span>· {toBnDigits(c.challanNo)}</span>
                                <span className="font-semibold text-slate-600">({c.challanDate ? formatBnDate(c.challanDate) : '—'})</span>
                                <span className="text-emerald-700">{formatBnMoney(c.depositedAmount)} ৳</span>
                              </div>
                              <div className="text-[11px] text-slate-600 flex flex-wrap items-center gap-1.5">
                                <span className="font-semibold">{c.dealerName || lockedSub?.partyName || 'নাম নেই'}</span>
                                <span>· {c.category}</span>
                                {c.dealerAllocationId ? (
                                  <span className="inline-flex items-center gap-1 rounded bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 text-[10px] font-bold text-emerald-800">
                                    <Lock className="h-2.5 w-2.5" />
                                    {lockedSub ? `বরাদ্দে লক: ${lockedSub.memo ? toBnDigits(lockedSub.memo) : ''} (${lockedSub.date ? formatBnDate(lockedSub.date) : ''})` : `বরাদ্দ আইডি লক: ${c.dealerAllocationId}`}
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 rounded bg-slate-100 border border-slate-200 px-1.5 py-0.5 text-[10px] font-semibold text-slate-600">
                                    <Unlock className="h-2.5 w-2.5" />
                                    উন্মুক্ত
                                  </span>
                                )}
                                {c.doNo && (
                                  <span className="inline-flex items-center rounded bg-rose-50 border border-rose-200 px-1.5 py-0.5 text-[10px] font-bold text-[#E2136E]">
                                    ডিও: {toBnDigits(c.doNo)}
                                  </span>
                                )}
                              </div>
                            </div>
                            <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                              <select
                                value={draftTarget}
                                onChange={(e) =>
                                  setChallanTargetDrafts((prev) => ({ ...prev, [c.id]: e.target.value }))
                                }
                                className="min-h-[34px] max-w-[240px] rounded-lg border border-slate-300 bg-white px-2 py-1 text-[11px]"
                              >
                                <option value="">— বরাদ্দ নির্বাচন করুন —</option>
                                {sortedSubs.map((s) => (
                                  <option key={s.id} value={s.id}>
                                    {s.label}
                                  </option>
                                ))}
                              </select>
                              <button
                                type="button"
                                disabled={!draftTarget}
                                onClick={() => {
                                  const res = onRelinkChallan(c.id, { dealerAllocationId: draftTarget });
                                  showRes(res);
                                }}
                                className="min-h-[34px] inline-flex items-center gap-1 rounded-lg bg-slate-900 hover:bg-slate-800 disabled:opacity-40 px-2.5 py-1 text-[11px] font-bold text-white cursor-pointer"
                              >
                                <Link2 className="h-3 w-3" />
                                <span>লক করুন</span>
                              </button>
                              {(c.dealerAllocationId || c.doNo) && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    const res = onRelinkChallan(c.id, { clearLink: true });
                                    setChallanTargetDrafts((prev) => ({ ...prev, [c.id]: '' }));
                                    showRes(res);
                                  }}
                                  className="min-h-[34px] inline-flex items-center gap-1 rounded-lg border border-rose-300 bg-rose-50 hover:bg-rose-100 px-2 py-1 text-[11px] font-bold text-rose-700 cursor-pointer"
                                  title="বরাদ্দ ও ডিও লিংক মুক্ত করুন"
                                >
                                  <Unlock className="h-3 w-3" />
                                  <span>মুক্ত</span>
                                </button>
                              )}
                            </div>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
              );
            })()}

            {pendingGroups.length === 0 ? (
              <div className="border border-slate-200 rounded-xl py-8 text-center text-xs text-slate-500">
                বর্তমানে এই খাতে কোনো পেন্ডিং ডিও প্রাপক নেই।
              </div>
            ) : (
              <div className="space-y-4">
                {pendingGroups.map((grp) => {
                  const allGroupSelected =
                    grp.items.length > 0 &&
                    grp.items.every((item) => Boolean(pendingInputs[item.key]?.doNo?.trim()));
                  const someGroupSelected = grp.items.some((item) =>
                    Boolean(pendingInputs[item.key]?.doNo?.trim())
                  );

                  return (
                    <div
                      key={grp.groupKey}
                      className="border border-slate-200 rounded-xl overflow-hidden bg-white"
                    >
                      <div className="bg-slate-50 border-b border-slate-200 px-3.5 py-2.5 flex flex-wrap items-center justify-between gap-2 text-xs">
                        <div className="flex flex-wrap items-center gap-x-2 gap-y-1 font-bold text-slate-900">
                          <span className="text-[#E2136E]">{grp.category}</span>
                          {grp.grains.includes('কার্নেল') && (
                            <span className="inline-flex items-center rounded-full border border-emerald-300 bg-emerald-50 px-2 py-0.5 text-[10.5px] font-bold text-emerald-900">
                              কার্নেল
                            </span>
                          )}
                          <span className="text-slate-400">·</span>
                          <span>
                            স্মারক নং {grp.unoMemo ? toBnDigits(grp.unoMemo) : '—'}
                          </span>
                          <span className="text-slate-400">·</span>
                          <span className="font-semibold text-slate-700">
                            তারিখ {grp.unoDate ? formatBnDate(grp.unoDate) : '—'}
                          </span>
                          <span className="text-slate-400">·</span>
                          <span className="font-semibold text-slate-700 tabular-nums">
                            মোট প্রাপক ({formatBnInt(grp.items.length)} জন)
                          </span>
                          <span className="text-slate-400">·</span>
                          <span className="font-bold text-[#E2136E] tabular-nums">
                            মোট বাকি ({formatBnQty(grp.totalRemaining)} মে.টন)
                          </span>
                        </div>

                        <label className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-bold text-slate-800 cursor-pointer whitespace-nowrap shrink-0">
                          <input
                            type="checkbox"
                            checked={allGroupSelected}
                            ref={(el) => {
                              if (el) el.indeterminate = !allGroupSelected && someGroupSelected;
                            }}
                            onChange={(e) => toggleGroupSelectAll(grp.items, e.target.checked)}
                            className="h-4 w-4 accent-[#E2136E] cursor-pointer"
                          />
                          <span>সব ইস্যু করুন</span>
                        </label>
                      </div>

                      <div className="divide-y divide-slate-200">
                        {grp.items.map((p) => {
                          const rowDraft = pendingInputs[p.key] || { doNo: '', date: doBatchDate };
                          const check =
                            previewChecksByKey.get(p.key) ||
                            verifyChallan(p.category, p.recipientName, p.remaining, {
                              subId: p.subId,
                              unoDate: p.unoDate,
                              unoMemo: p.unoMemo,
                              ministryMemo: p.ministryMemo,
                              issueDate: rowDraft.date || doBatchDate,
                            });
                          const isChecked = Boolean(rowDraft.doNo.trim());
                          const showVerify = verifyTargetKey === p.key;

                          return (
                            <div
                              key={p.key}
                              className={`p-3 transition-colors ${
                                isChecked ? 'bg-rose-50/40' : 'hover:bg-slate-50/70'
                              }`}
                            >
                              <div className="grid grid-cols-1 lg:grid-cols-12 lg:items-center gap-2.5 text-xs tabular-nums">
                                {/* Checkbox + Recipient & Union/SalesCenter/Project */}
                                <div className="lg:col-span-5 flex items-start gap-2.5 min-w-0">
                                  <input
                                    type="checkbox"
                                    checked={isChecked}
                                    onChange={(e) => toggleSingleRowIssue(p, e.target.checked)}
                                    aria-label={`${p.recipientName} ইস্যু করুন`}
                                    className="mt-0.5 h-5 w-5 shrink-0 accent-[#E2136E] cursor-pointer"
                                  />
                                  <div className="min-w-0 flex-1">
                                    <div className="font-bold text-slate-900 text-[13px] flex flex-wrap items-center gap-1.5 break-words">
                                      <span>{p.recipientName}</span>
                                      {resolveItemGrain(p.category, p.ministryMemo) === 'কার্নেল' && (
                                        <span className="inline-flex items-center rounded-full border border-emerald-300 bg-emerald-50 px-1.5 py-0.5 text-[10px] font-bold text-emerald-900">
                                          কার্নেল · মিলার
                                        </span>
                                      )}
                                    </div>
                                    <div className="text-[11px] text-slate-500 break-words mt-0.5">
                                      {p.unionName}
                                      {p.salesCenter ? ` · ${p.salesCenter}` : ''}
                                      {p.projectNo ? ` · প্রকল্প: ${toBnDigits(p.projectNo)}` : ''}
                                    </div>
                                  </div>
                                </div>

                                {/* Remaining Qty + Challan Verify */}
                                <div className="lg:col-span-3 flex flex-wrap items-center justify-between lg:justify-end gap-2 pl-7 lg:pl-0">
                                  <span className="font-bold text-[#E2136E] text-[13px] whitespace-nowrap">
                                    {formatBnQty(p.remaining)} মে.টন
                                  </span>
                                  {!check.applicable ? (
                                    <span className="text-[11px] text-slate-500">প্রযোজ্য নয়</span>
                                  ) : (
                                    <button
                                      type="button"
                                      onClick={() => setVerifyTargetKey(showVerify ? null : p.key)}
                                      className={`min-h-[34px] rounded-lg border px-2.5 py-1 text-[11px] font-bold cursor-pointer whitespace-nowrap inline-flex items-center gap-1 ${
                                        check.ok
                                          ? 'border-emerald-300 bg-emerald-50 text-emerald-800'
                                          : 'border-amber-300 bg-amber-50 text-amber-800'
                                      }`}
                                    >
                                      {check.ok && ((p.subId && check.mainChallan?.dealerAllocationId && normMemoNo(check.mainChallan.dealerAllocationId) === normMemoNo(p.subId)) || (p.subId && check.taxChallan?.dealerAllocationId && normMemoNo(check.taxChallan.dealerAllocationId) === normMemoNo(p.subId))) && (
                                        <Lock className="h-3 w-3 shrink-0" />
                                      )}
                                      <span>{check.ok ? 'চালান সঠিক' : 'চালান সিগন্যাল'}</span>
                                    </button>
                                  )}
                                </div>

                                {/* DO No + Issue Date + Single Issue Button */}
                                {(() => {
                                  const draftKey = normDoNo(rowDraft.doNo);
                                  const isAlreadyIssued = Boolean(draftKey && issuedSet.has(draftKey));
                                  const isDupInPending = Boolean(
                                    draftKey &&
                                      Object.entries(pendingInputs).some(
                                        ([k, v]) => k !== p.key && normDoNo(v?.doNo) === draftKey
                                      )
                                  );
                                  const hasDoConflict = isAlreadyIssued || isDupInPending;
                                  return (
                                    <div className="lg:col-span-4 pl-7 lg:pl-0 space-y-1">
                                      <div className="grid grid-cols-12 gap-1.5 items-center">
                                        <input
                                          type="text"
                                          list="unused-do-datalist"
                                          value={rowDraft.doNo}
                                          placeholder="ডিও নং"
                                          onChange={(e) =>
                                            setPendingInputs((prev) => ({
                                              ...prev,
                                              [p.key]: { ...rowDraft, doNo: e.target.value },
                                            }))
                                          }
                                          className={`col-span-5 min-h-[42px] w-full rounded-lg border px-2.5 py-1.5 text-sm font-bold text-center focus:outline-none ${
                                            hasDoConflict
                                              ? 'border-rose-500 bg-rose-50 text-rose-800 focus:border-rose-600'
                                              : 'border-slate-300 bg-white focus:border-[#E2136E]'
                                          }`}
                                        />
                                        <input
                                          type="date"
                                          value={rowDraft.date}
                                          onChange={(e) =>
                                            setPendingInputs((prev) => ({
                                              ...prev,
                                              [p.key]: { ...rowDraft, date: e.target.value },
                                            }))
                                          }
                                          className="col-span-4 min-h-[42px] w-full rounded-lg border border-slate-300 bg-white px-1.5 py-1.5 text-[11px] sm:text-xs focus:outline-none focus:border-[#E2136E]"
                                        />
                                        <button
                                          type="button"
                                          disabled={!rowDraft.doNo.trim() || hasDoConflict}
                                          onClick={() => handleIssueSingleRow(p)}
                                          className="col-span-3 min-h-[42px] w-full rounded-lg bg-slate-900 hover:bg-slate-800 disabled:opacity-40 px-2 py-1.5 text-xs font-bold text-white cursor-pointer whitespace-nowrap"
                                        >
                                          ইস্যু
                                        </button>
                                      </div>
                                      {hasDoConflict && (
                                        <div className="text-[11px] font-bold text-rose-700">
                                          {isAlreadyIssued
                                            ? `ডিও নং ${toBnDigits(rowDraft.doNo)} আগেই ইস্যু হয়েছে — ব্যবহার করা যাবে না`
                                            : `ডিও নং ${toBnDigits(rowDraft.doNo)} অন্য সারিতে লেখা আছে`}
                                        </div>
                                      )}
                                    </div>
                                  );
                                })()}
                              </div>

                              {isChecked && check.applicable && !check.ok && check.errors && check.errors.length > 0 && !showVerify && (
                                <div className="mt-2 px-3 py-1.5 bg-amber-50 border border-amber-200 rounded-lg text-[11px] font-bold text-amber-900 flex items-center gap-1.5">
                                  <AlertTriangle className="h-3.5 w-3.5 shrink-0 text-amber-600" />
                                  <span>সিগন্যাল: {check.errors.join(' · ')} (ডিও সেভ করা যাবে)</span>
                                </div>
                              )}

                              {showVerify && check.applicable && (() => {
                                const mainLockedToSub = Boolean(
                                  p.subId &&
                                    check.mainChallan?.dealerAllocationId &&
                                    normMemoNo(check.mainChallan.dealerAllocationId) === normMemoNo(p.subId)
                                );
                                const taxLockedToSub = Boolean(
                                  p.subId &&
                                    check.taxChallan?.dealerAllocationId &&
                                    normMemoNo(check.taxChallan.dealerAllocationId) === normMemoNo(p.subId)
                                );
                                const partyChallans = db.challans
                                  .filter(
                                    (c) =>
                                      c.category === p.category &&
                                      ((p.subId && c.dealerAllocationId && normMemoNo(c.dealerAllocationId) === normMemoNo(p.subId)) ||
                                        (c.dealerName && isSamePartyName(c.dealerName, p.recipientName)))
                                  )
                                  .sort((a, b) => (b.challanDate || '').localeCompare(a.challanDate || ''));
                                const selectedManualChallanId = manualRelinkPickByRow[p.key] || '';

                                return (
                                  <div className="mt-2.5 px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-xs space-y-2">
                                    <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5">
                                      <span className="inline-flex flex-wrap items-center gap-1">
                                        প্রয়োজনীয় মূল টাকা: <strong>{formatBnMoney(check.requiredAmount || 0)} ৳</strong>{' '}
                                        ({check.mainChallan ? `জমা: ${formatBnMoney(check.mainChallan.depositedAmount)} ৳ · ${toBnDigits(check.mainChallan.challanNo)}` : 'চালান নেই'})
                                        {mainLockedToSub && (
                                          <span className="inline-flex items-center gap-0.5 rounded bg-emerald-100 px-1.5 py-0.5 text-[10px] font-bold text-emerald-800">
                                            <Lock className="h-2.5 w-2.5" /> বরাদ্দে লক
                                          </span>
                                        )}
                                      </span>
                                      <span className="inline-flex flex-wrap items-center gap-1">
                                        প্রয়োজনীয় আয়কর: <strong>{formatBnMoney(check.requiredTaxAmount || 0)} ৳</strong>{' '}
                                        ({check.taxChallan ? `জমা: ${formatBnMoney(check.taxChallan.depositedAmount)} ৳ · ${toBnDigits(check.taxChallan.challanNo)}` : 'চালান নেই'})
                                        {taxLockedToSub && (
                                          <span className="inline-flex items-center gap-0.5 rounded bg-emerald-100 px-1.5 py-0.5 text-[10px] font-bold text-emerald-800">
                                            <Lock className="h-2.5 w-2.5" /> বরাদ্দে লক
                                          </span>
                                        )}
                                      </span>
                                      {check.errors && check.errors.length > 0 && (
                                        <span className="inline-flex items-center gap-1 font-bold text-amber-800">
                                          <AlertTriangle className="h-3.5 w-3.5 shrink-0 text-amber-600" />
                                          {check.errors.join(' · ')}
                                        </span>
                                      )}
                                    </div>

                                    {onRelinkChallan && p.subId && partyChallans.length > 0 && (
                                      <div className="pt-2 border-t border-slate-200 flex flex-wrap items-center gap-1.5 text-[11px]">
                                        <span className="font-bold text-slate-700">ম্যানুয়াল লিংক পরিবর্তন:</span>
                                        <select
                                          value={selectedManualChallanId}
                                          onChange={(e) =>
                                            setManualRelinkPickByRow((prev) => ({
                                              ...prev,
                                              [p.key]: e.target.value,
                                            }))
                                          }
                                          className="min-h-[32px] max-w-[300px] rounded-lg border border-slate-300 bg-white px-2 py-1"
                                        >
                                          <option value="">— চালান বেছে নিন —</option>
                                          {partyChallans.map((c) => {
                                            const lockState = c.dealerAllocationId
                                              ? normMemoNo(c.dealerAllocationId) === normMemoNo(p.subId)
                                                ? 'এই বরাদ্দে লক'
                                                : 'অন্য বরাদ্দে লক'
                                              : 'উন্মুক্ত';
                                            return (
                                              <option key={c.id} value={c.id}>
                                                {c.challanType} · {toBnDigits(c.challanNo)} ({c.challanDate ? formatBnDate(c.challanDate) : '—'}) · {formatBnMoney(c.depositedAmount)} ৳ [{lockState}]
                                              </option>
                                            );
                                          })}
                                        </select>
                                        <button
                                          type="button"
                                          disabled={!selectedManualChallanId}
                                          onClick={() => {
                                            const res = onRelinkChallan(selectedManualChallanId, {
                                              dealerAllocationId: p.subId,
                                            });
                                            showRes(res);
                                          }}
                                          className="min-h-[32px] inline-flex items-center gap-1 rounded-lg bg-slate-900 hover:bg-slate-800 disabled:opacity-40 px-2.5 py-1 font-bold text-white cursor-pointer"
                                        >
                                          <Lock className="h-3 w-3" />
                                          <span>এই বরাদ্দে লক করুন</span>
                                        </button>
                                        {mainLockedToSub && check.mainChallan && (
                                          <button
                                            type="button"
                                            onClick={() => {
                                              const res = onRelinkChallan(check.mainChallan!.id, { clearLink: true });
                                              showRes(res);
                                            }}
                                            className="min-h-[32px] inline-flex items-center gap-1 rounded-lg border border-rose-300 bg-rose-50 hover:bg-rose-100 px-2 py-1 font-bold text-rose-700 cursor-pointer"
                                          >
                                            <Unlock className="h-3 w-3" />
                                            <span>মূল মুক্ত</span>
                                          </button>
                                        )}
                                        {taxLockedToSub && check.taxChallan && (
                                          <button
                                            type="button"
                                            onClick={() => {
                                              const res = onRelinkChallan(check.taxChallan!.id, { clearLink: true });
                                              showRes(res);
                                            }}
                                            className="min-h-[32px] inline-flex items-center gap-1 rounded-lg border border-rose-300 bg-rose-50 hover:bg-rose-100 px-2 py-1 font-bold text-rose-700 cursor-pointer"
                                          >
                                            <Unlock className="h-3 w-3" />
                                            <span>আয়কর মুক্ত</span>
                                          </button>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                );
                              })()}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}

                <datalist id="unused-do-datalist">
                  {availableDatalistDos.slice(0, 120).map((u) => (
                    <option key={u.doNo} value={String(u.doNo)}>
                      বই {toBnDigits(u.bookNo)}
                    </option>
                  ))}
                </datalist>
              </div>
            )}
          </div>

          {/* Issued DO History */}
          <div className="rounded-2xl border border-slate-200 bg-white overflow-hidden">
            <div className="px-4 py-3 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
              <div className="text-xs font-bold text-slate-900">
                {lockedCategory ? `${lockedCategory} — ` : ''}ইস্যুকৃত ডিও তালিকা ({formatBnInt(filteredIssuedDos.length)}টি)
              </div>
              <div className="relative w-full sm:w-64">
                <Search className="h-3.5 w-3.5 text-slate-400 absolute left-2.5 top-3" />
                <input
                  type="text"
                  value={doSearch}
                  onChange={(e) => setDoSearch(e.target.value)}
                  placeholder="ডিও নং / ডিলার / ইউনিয়ন"
                  className="w-full min-h-[38px] rounded-lg border border-slate-300 pl-8 pr-2.5 py-1.5 text-xs focus:outline-none focus:border-[#E2136E]"
                />
              </div>
            </div>
            <div className="hidden lg:grid lg:grid-cols-12 gap-2 px-3 py-2 bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-700">
              <div className="lg:col-span-2">ডিও নং ও তারিখ</div>
              <div className="lg:col-span-2">খাত</div>
              <div className="lg:col-span-3">প্রাপক ও ইউনিয়ন</div>
              <div className="lg:col-span-1 text-right">পরিমাণ</div>
              <div className="lg:col-span-3">চালান (মূল ও আয়কর)</div>
              <div className="lg:col-span-1 text-center">কাজ</div>
            </div>
            <div className="divide-y divide-slate-200">
              {filteredIssuedDos.length === 0 ? (
                <div className="py-6 text-center text-xs text-slate-500">
                  এখনো কোনো ডিও ইস্যু করা হয়নি।
                </div>
              ) : (
                filteredIssuedDos.map((d) => {
                  const isEditing = editingDoRow?.id === d.id;
                  const validForDo = (c: ChallanRecord): boolean => {
                    if (c.category !== d.category) return false;
                    if (!isChallanOnOrBeforeDo(c.challanDate, d.issueDate, false)) return false;
                    if (c.dealerName && c.dealerName !== '—') {
                      return isSamePartyName(c.dealerName, d.recipientName);
                    }
                    return Boolean(
                      d.dealerAllocationId &&
                        c.dealerAllocationId &&
                        normMemoNo(c.dealerAllocationId) === normMemoNo(d.dealerAllocationId)
                    );
                  };
                  const mainRec = d.challanNo
                    ? db.challans.find(
                        (c) =>
                          c.challanType === 'মূল চালান' &&
                          normChallanNo(c.challanNo) === normChallanNo(d.challanNo) &&
                          validForDo(c)
                      )
                    : db.challans.find(
                        (c) =>
                          c.challanType === 'মূল চালান' &&
                          c.doNo &&
                          normDoNo(c.doNo) === normDoNo(d.doNo) &&
                          validForDo(c)
                      );
                  const taxRec = d.taxChallanNo
                    ? db.challans.find(
                        (c) =>
                          c.challanType === 'আয়কর চালান' &&
                          normChallanNo(c.challanNo) === normChallanNo(d.taxChallanNo) &&
                          validForDo(c)
                      )
                    : db.challans.find(
                        (c) =>
                          c.challanType === 'আয়কর চালান' &&
                          c.doNo &&
                          normDoNo(c.doNo) === normDoNo(d.doNo) &&
                          validForDo(c)
                      );
                  const mainNo =
                    mainRec?.challanNo ||
                    (isChallanOnOrBeforeDo(d.challanDate, d.issueDate, false)
                      ? d.challanNo
                      : undefined);
                  const taxNo =
                    taxRec?.challanNo ||
                    (isChallanOnOrBeforeDo(d.taxChallanDate, d.issueDate, false)
                      ? d.taxChallanNo
                      : undefined);
                  const mainAmt = mainNo ? mainRec?.depositedAmount || d.challanAmount || 0 : 0;
                  const taxAmt = taxNo ? taxRec?.depositedAmount || d.taxChallanAmount || 0 : 0;

                  return (
                    <div
                      key={d.id}
                      className={`p-3 text-xs tabular-nums ${isEditing ? 'bg-rose-50/50' : 'hover:bg-slate-50/60'}`}
                    >
                      <div className="grid grid-cols-1 lg:grid-cols-12 lg:items-center gap-2">
                        {/* DO No + Issue Date */}
                        <div className="lg:col-span-2 flex items-center justify-between lg:block gap-2">
                          {isEditing ? (
                            <div className="flex flex-wrap lg:flex-col gap-1.5 w-full">
                              <input
                                type="text"
                                value={editingDoRow.doNo}
                                onChange={(e) =>
                                  setEditingDoRow((prev) =>
                                    prev ? { ...prev, doNo: e.target.value } : prev
                                  )
                                }
                                className="w-28 lg:w-full rounded border border-slate-300 px-2 py-1.5 text-center font-bold text-xs bg-white"
                              />
                              <input
                                type="date"
                                value={editingDoRow.issueDate}
                                onChange={(e) =>
                                  setEditingDoRow((prev) =>
                                    prev ? { ...prev, issueDate: e.target.value } : prev
                                  )
                                }
                                className="w-36 lg:w-full rounded border border-slate-300 px-2 py-1.5 text-xs bg-white"
                              />
                            </div>
                          ) : (
                            <div>
                              <span className="font-bold text-[#E2136E] text-sm">
                                ডিও নং {toBnDigits(d.doNo)}
                              </span>
                              <div className="text-[11px] text-slate-500">
                                {formatBnDate(d.issueDate)}
                              </div>
                            </div>
                          )}

                          {/* Mobile Qty badge */}
                          <div className="lg:hidden font-bold text-slate-900">
                            {isEditing ? (
                              <input
                                type="number"
                                step="0.001"
                                value={editingDoRow.qty}
                                onChange={(e) =>
                                  setEditingDoRow((prev) =>
                                    prev ? { ...prev, qty: e.target.value } : prev
                                  )
                                }
                                className="w-24 rounded border border-slate-300 px-2 py-1.5 text-right font-bold bg-white"
                              />
                            ) : (
                              <span>{formatBnQty(d.allocationQty)} মে.টন</span>
                            )}
                          </div>
                        </div>

                        {/* Category + Kernel */}
                        <div className="lg:col-span-2 font-semibold flex flex-wrap items-center gap-1.5">
                          <span>{d.category}</span>
                          {resolveItemGrain(d.category, d.ministryMemo) === 'কার্নেল' && (
                            <span className="inline-flex items-center rounded-full border border-emerald-300 bg-emerald-50 px-1.5 py-0.5 text-[10px] font-bold text-emerald-900">
                              কার্নেল · মিলার
                            </span>
                          )}
                        </div>

                        {/* Recipient & Union */}
                        <div className="lg:col-span-3 min-w-0">
                          <div className="font-bold text-slate-900 break-words">{d.recipientName}</div>
                          <div className="text-[11px] text-slate-500 break-words">{d.unionName}</div>
                        </div>

                        {/* Desktop Qty */}
                        <div className="hidden lg:block lg:col-span-1 text-right font-bold">
                          {isEditing ? (
                            <input
                              type="number"
                              step="0.001"
                              value={editingDoRow.qty}
                              onChange={(e) =>
                                setEditingDoRow((prev) =>
                                  prev ? { ...prev, qty: e.target.value } : prev
                                )
                              }
                              className="w-full rounded border border-slate-300 px-2 py-1 text-right font-bold bg-white"
                            />
                          ) : (
                            `${formatBnQty(d.allocationQty)} মে.টন`
                          )}
                        </div>

                        {/* Challans */}
                        <div className="lg:col-span-3 text-[11px] text-slate-700 space-y-1 break-words">
                          {!mainNo && !taxNo && mainAmt <= 0 && taxAmt <= 0 ? (
                            <span className="text-slate-400">চালান: —</span>
                          ) : (
                            <>
                              {(mainNo || mainAmt > 0) && (
                                <div>
                                  মূল: <span className="font-semibold">{mainNo ? toBnDigits(mainNo) : '—'}</span>
                                  {mainAmt > 0 && (
                                    <span className="font-bold text-[#E2136E]">
                                      {' '}(জমা: {formatBnMoney(mainAmt)} ৳)
                                    </span>
                                  )}
                                </div>
                              )}
                              {(taxNo || taxAmt > 0) && (
                                <div>
                                  আয়কর: <span className="font-semibold">{taxNo ? toBnDigits(taxNo) : '—'}</span>
                                  {taxAmt > 0 && (
                                    <span className="font-bold text-[#E2136E]">
                                      {' '}(জমা: {formatBnMoney(taxAmt)} ৳)
                                    </span>
                                  )}
                                </div>
                              )}
                            </>
                          )}
                          {isEditing && onRelinkChallan && (() => {
                            const eligibleChallans = db.challans.filter(
                              (c) =>
                                c.category === d.category &&
                                (!c.dealerName || c.dealerName === '—' || isSamePartyName(c.dealerName, d.recipientName))
                            );
                            if (eligibleChallans.length === 0) return null;
                            return (
                              <div className="pt-1 flex flex-wrap items-center gap-1">
                                <select
                                  value={editingDoChallanId}
                                  onChange={(e) => setEditingDoChallanId(e.target.value)}
                                  className="min-h-[30px] max-w-[190px] rounded border border-slate-300 bg-white px-1.5 py-0.5 text-[10.5px]"
                                >
                                  <option value="">— চালান লিংক করুন —</option>
                                  {eligibleChallans.map((c) => (
                                    <option key={c.id} value={c.id}>
                                      {c.challanType} · {toBnDigits(c.challanNo)} ({formatBnMoney(c.depositedAmount)} ৳)
                                    </option>
                                  ))}
                                </select>
                                <button
                                  type="button"
                                  disabled={!editingDoChallanId}
                                  onClick={() => {
                                    const res = onRelinkChallan(editingDoChallanId, {
                                      dealerAllocationId: d.dealerAllocationId,
                                      doNo: d.doNo,
                                    });
                                    showRes(res);
                                    if (res.ok) setEditingDoChallanId('');
                                  }}
                                  className="min-h-[30px] rounded bg-slate-900 hover:bg-slate-800 disabled:opacity-40 px-2 py-0.5 text-[10.5px] font-bold text-white cursor-pointer"
                                >
                                  যুক্ত
                                </button>
                              </div>
                            );
                          })()}
                        </div>

                        {/* Actions */}
                        <div className="lg:col-span-1 flex items-center justify-end lg:justify-center gap-1 pt-1 lg:pt-0">
                          {isEditing ? (
                            <div className="inline-flex items-center gap-1.5">
                              <button
                                type="button"
                                onClick={() => {
                                  if (!onUpdateDOIssue) return;
                                  const res = onUpdateDOIssue(d.id, {
                                    doNo: editingDoRow.doNo,
                                    issueDate: editingDoRow.issueDate,
                                    allocationQty: parseNum(editingDoRow.qty),
                                  });
                                  showRes(res);
                                  if (res.ok) setEditingDoRow(null);
                                }}
                                className="min-h-[36px] min-w-[36px] inline-flex items-center justify-center text-white bg-[#E2136E] hover:bg-[#B80D58] rounded-lg cursor-pointer"
                                title="সংরক্ষণ"
                              >
                                <Check className="h-4 w-4" />
                              </button>
                              <button
                                type="button"
                                onClick={() => setEditingDoRow(null)}
                                className="min-h-[36px] min-w-[36px] inline-flex items-center justify-center text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg cursor-pointer"
                                title="বাতিল"
                              >
                                <X className="h-4 w-4" />
                              </button>
                            </div>
                          ) : (
                            <div className="inline-flex items-center gap-1.5">
                              {onUpdateDOIssue && (
                                <button
                                  type="button"
                                  onClick={() =>
                                    setEditingDoRow({
                                      id: d.id,
                                      doNo: d.doNo,
                                      issueDate: d.issueDate,
                                      qty: String(d.allocationQty),
                                    })
                                  }
                                  className="min-h-[36px] min-w-[36px] inline-flex items-center justify-center text-[#E2136E] hover:bg-rose-50 rounded-lg cursor-pointer"
                                  title="এডিট করুন"
                                >
                                  <Pencil className="h-4 w-4" />
                                </button>
                              )}
                              <button
                                type="button"
                                onClick={() => onDeleteDOIssue(d.id)}
                                className="min-h-[36px] min-w-[36px] inline-flex items-center justify-center text-rose-600 hover:bg-rose-50 rounded-lg cursor-pointer"
                                title="মুছুন"
                              >
                                <Trash2 className="h-4 w-4" />
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};



