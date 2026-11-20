import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  BarChart3,
  BookOpen,
  Coins,
  Download,
  Edit3,
  FileCheck2,
  FileSpreadsheet,
  FileWarning,
  Flame,
  Layers,
  Printer,
  RectangleHorizontal,
  Smartphone,
  Store,
  Users,
} from 'lucide-react';
import type {
  AllocOrderFormFields,
  AllocationDatabase,
  ReportTemplateId,
} from '../types/allocation';
import {
  REPORT_TEMPLATE_CATALOG,
  createDefaultTemplate,
} from '../types/allocation';
import {
  buildAllocationSummaryReport,
  buildCashSaleReport,
  buildEPRationReport,
  buildFoodAllocOrderMemos,
  buildFoodAllocOrderRows,
  buildMonthlyDistributionReport,
  buildOmsAllocOrderMemos,
  buildOmsAllocOrderRows,
  getDoBookRegisterStats,
  getUnusedDoNumbers,
} from '../shared/allocReportEngine';
import {
  BENGALI_MONTHS,
  autoDeadlineIsoDate,
  bnToEnDigits,
  formatBangabdaDate,
  formatBnDate,
  formatBnInt,
  formatBnMoney,
  formatBnQty,
  isChallanOnOrBeforeDo,
  isSamePartyName,
  normChallanNo,
  normDoNo,
  normMemoNo,
  parseMonthYear,
  round2,
  round3,
  toBnDigits,
  todayIsoDate,
} from '../shared/bnUtils';

function extractMemoSuffix(fullMemo: string, prefix: string): string {
  const raw = (fullMemo || '').trim();
  if (!raw) return '';
  const cleanPrefix = (prefix || '').trim();
  if (!cleanPrefix) return raw;
  if (raw.startsWith(cleanPrefix)) {
    return raw.slice(cleanPrefix.length).trim();
  }
  const rawNorm = bnToEnDigits(raw).replace(/\s+/g, '');
  const preNorm = bnToEnDigits(cleanPrefix).replace(/\s+/g, '');
  if (preNorm && rawNorm.startsWith(preNorm)) {
    return raw.replace(/\s+/g, '').slice(preNorm.length).trim();
  }
  return raw;
}
import {
  OfficialA4Document,
  type DocColumnGroup,
  type DocColumnLeaf,
  type DocKpiBox,
  type DocRowData,
  type DocTotalRow,
  type OfficialA4DocumentHandle,
} from './OfficialA4Document';

export type ReportSectionId =
  | 'monthly-dist'
  | 'alloc-summary'
  | 'monthly-cash'
  | 'monthly-ep'
  | 'alloc-food'
  | 'alloc-oms-gen'
  | 'alloc-oms-spc'
  | 'alloc-oms-tcb'
  | 'do-reg-book'
  | 'do-reg-issue'
  | 'do-reg-unused'
  | 'dlr-rep-food'
  | 'dlr-rep-oms';

type ReportParentGroup = 'monthly' | 'orders' | 'dealer' | 'do';

interface ReportsViewProps {
  db: AllocationDatabase;
  activeReportId: ReportSectionId;
  onSelectReport: (id: ReportSectionId) => void;
  onSaveReportMemo: (key: string, no: string, date: string) => void;
  onSaveAllocOrderForm?: (key: string, form: AllocOrderFormFields) => void;
  onEditTemplate: (tplId: ReportTemplateId) => void;
  onChangeOrientation: (tplId: ReportTemplateId, orientation: 'portrait' | 'landscape') => void;
}

const REPORT_PARENT_GROUPS: Array<{
  key: ReportParentGroup;
  label: string;
  items: Array<{
    id: ReportSectionId;
    label: string;
    icon: React.ComponentType<{ className?: string }>;
    bg: string;
    fg: string;
  }>;
}> = [
  {
    key: 'monthly',
    label: 'মাসিক রিপোর্ট',
    items: [
      { id: 'monthly-dist', label: 'মাসিক বিতরণ রিপোর্ট', icon: BarChart3, bg: 'bg-rose-50 border-rose-200', fg: 'text-[#E2136E]' },
      { id: 'alloc-summary', label: 'বরাদ্দ ও বিতরণ স্থিতি', icon: Layers, bg: 'bg-emerald-50 border-emerald-200', fg: 'text-emerald-700' },
      { id: 'monthly-cash', label: 'নগদ বিক্রয় রিপোর্ট', icon: Coins, bg: 'bg-amber-50 border-amber-200', fg: 'text-amber-700' },
      { id: 'monthly-ep', label: 'ইপি রেশন রিপোর্ট', icon: Flame, bg: 'bg-orange-50 border-orange-200', fg: 'text-orange-700' },
    ],
  },
  {
    key: 'orders',
    label: 'বরাদ্দাদেশ',
    items: [
      { id: 'alloc-food', label: 'খাদ্য বান্ধব বরাদ্দাদেশ', icon: FileSpreadsheet, bg: 'bg-teal-50 border-teal-200', fg: 'text-teal-700' },
      { id: 'alloc-oms-tcb', label: 'ওএমএস (টিসিবি) বরাদ্দাদেশ', icon: Store, bg: 'bg-violet-50 border-violet-200', fg: 'text-violet-700' },
      { id: 'alloc-oms-spc', label: 'ওএমএস (বিশেষ) বরাদ্দাদেশ', icon: FileSpreadsheet, bg: 'bg-indigo-50 border-indigo-200', fg: 'text-indigo-700' },
      { id: 'alloc-oms-gen', label: 'ওএমএস (সাধারণ) বরাদ্দাদেশ', icon: FileSpreadsheet, bg: 'bg-sky-50 border-sky-200', fg: 'text-sky-700' },
    ],
  },
  {
    key: 'dealer',
    label: 'ডিলার রিপোর্ট',
    items: [
      { id: 'dlr-rep-food', label: 'খাদ্য বান্ধব ডিলার', icon: Users, bg: 'bg-rose-50 border-rose-200', fg: 'text-[#E2136E]' },
      { id: 'dlr-rep-oms', label: 'ওএমএস ডিলার', icon: Store, bg: 'bg-sky-50 border-sky-200', fg: 'text-sky-700' },
    ],
  },
  {
    key: 'do',
    label: 'ডিউ রিপোর্ট',
    items: [
      { id: 'do-reg-book', label: 'ডিও বই রেজিস্ট্রার', icon: BookOpen, bg: 'bg-indigo-50 border-indigo-200', fg: 'text-indigo-700' },
      { id: 'do-reg-issue', label: 'ডিও ইস্যু রেজিস্টার', icon: FileCheck2, bg: 'bg-rose-50 border-rose-200', fg: 'text-[#E2136E]' },
      { id: 'do-reg-unused', label: 'অব্যবহৃত ডিও তালিকা', icon: FileWarning, bg: 'bg-amber-50 border-amber-200', fg: 'text-amber-700' },
    ],
  },
];

const DO_ISSUE_COLUMN_DEFS = [
  { key: 'doNo', label: 'ডিও নং' },
  { key: 'issueDate', label: 'ইস্যু তারিখ' },
  { key: 'category', label: 'খাত' },
  { key: 'union', label: 'ইউনিয়ন' },
  { key: 'recipient', label: 'প্রাপক / ডিলার' },
  { key: 'salesCenter', label: 'বিক্রয় কেন্দ্র' },
  { key: 'unoMemo', label: 'উপবরাদ্দ স্মারক' },
  { key: 'qty', label: 'পরিমাণ (মে.টন)' },
  { key: 'challan', label: 'চালান নং ও টাকা' },
  { key: 'signature', label: 'স্বাক্ষর' },
] as const;

const FOOD_DEALER_COLUMN_DEFS = [
  { key: 'unionName', label: 'ইউনিয়ন' },
  { key: 'dealerName', label: 'ডিলারের নাম' },
  { key: 'fatherName', label: 'পিতার নাম' },
  { key: 'nid', label: 'এনআইডি' },
  { key: 'mobile', label: 'মোবাইল' },
  { key: 'salesCenter', label: 'বিক্রয় কেন্দ্র' },
  { key: 'cardCount', label: 'কার্ড সংখ্যা' },
  { key: 'totalQty', label: 'চাল (মে.টন)' },
  { key: 'poNo', label: 'পে-অর্ডার নং' },
  { key: 'poDate', label: 'পে-অর্ডার তারিখ' },
  { key: 'bankName', label: 'ব্যাংকের নাম' },
  { key: 'amount', label: 'জামানত টাকা' },
  { key: 'supervisorName', label: 'তদারকি কর্মকর্তা' },
  { key: 'supervisorDesignation', label: 'কর্মকর্তার পদবি' },
  { key: 'supervisorMobile', label: 'কর্মকর্তার মোবাইল' },
] as const;

const OMS_DEALER_COLUMN_DEFS = [
  { key: 'areaType', label: 'এলাকার ধরন' },
  { key: 'unionName', label: 'ইউনিয়ন' },
  { key: 'municipality', label: 'পৌরসভা' },
  { key: 'dealerName', label: 'ডিলারের নাম' },
  { key: 'fatherName', label: 'পিতার নাম' },
  { key: 'nid', label: 'এনআইডি' },
  { key: 'mobile', label: 'মোবাইল' },
  { key: 'salesCenter', label: 'বিক্রয় কেন্দ্র' },
  { key: 'poNo', label: 'পে-অর্ডার নং' },
  { key: 'poDate', label: 'পে-অর্ডার তারিখ' },
  { key: 'bankName', label: 'ব্যাংকের নাম' },
  { key: 'amount', label: 'জামানত টাকা' },
  { key: 'supervisorName', label: 'তদারকি কর্মকর্তা' },
  { key: 'supervisorDesignation', label: 'কর্মকর্তার পদবি' },
  { key: 'supervisorMobile', label: 'কর্মকর্তার মোবাইল' },
] as const;

function findGroupForReport(id: ReportSectionId): ReportParentGroup {
  for (const grp of REPORT_PARENT_GROUPS) {
    if (grp.items.some((x) => x.id === id)) return grp.key;
  }
  return 'monthly';
}

export const ReportsView: React.FC<ReportsViewProps> = ({
  db,
  activeReportId,
  onSelectReport,
  onSaveReportMemo,
  onEditTemplate,
  onChangeOrientation,
}) => {
  const now = new Date();
  const docRef = useRef<OfficialA4DocumentHandle>(null);

  const [activeParent, setActiveParent] = useState<ReportParentGroup>(() =>
    findGroupForReport(activeReportId)
  );
  const [month0, setMonth0] = useState(now.getMonth());
  const [year, setYear] = useState(now.getFullYear());
  const [selectedFoodMemo, setSelectedFoodMemo] = useState('');
  const [selectedOmsMemo, setSelectedOmsMemo] = useState('');
  const [foodKernelMode, setFoodKernelMode] = useState<'without_kernel' | 'with_kernel'>('without_kernel');

  const doIssueCols = useMemo<string[]>(
    () => [
      'doNo',
      'issueDate',
      'category',
      'union',
      'recipient',
      'unoMemo',
      'qty',
      'challan',
      'signature',
    ],
    []
  );
  const foodDealerCols = useMemo<string[]>(
    () => [
      'unionName',
      'dealerName',
      'fatherName',
      'nid',
      'mobile',
      'salesCenter',
      'cardCount',
      'totalQty',
      'poNo',
      'bankName',
      'supervisorName',
    ],
    []
  );
  const omsDealerCols = useMemo<string[]>(
    () => [
      'areaType',
      'unionName',
      'municipality',
      'dealerName',
      'fatherName',
      'nid',
      'mobile',
      'salesCenter',
      'poNo',
      'bankName',
      'amount',
      'supervisorName',
    ],
    []
  );

  useEffect(() => {
    setActiveParent(findGroupForReport(activeReportId));
  }, [activeReportId]);

  const template =
    db.templates[activeReportId] ||
    createDefaultTemplate(
      REPORT_TEMPLATE_CATALOG.find((x) => x.id === activeReportId) ||
        REPORT_TEMPLATE_CATALOG[0],
      db.officer
    );
  const colLabel = (key: string, def: string) => template.columnLabels[key]?.trim() || def;

  const pendingAllocRows = useMemo(() => {
    const summary = buildAllocationSummaryReport(db, 'ALL');
    return summary.rows.filter(
      (r) =>
        r.ministryQty > r.doQty + 0.0005 ||
        (r.dcQty ?? 0) > r.doQty + 0.0005 ||
        r.subQty > r.doQty + 0.0005 ||
        r.projectQty > r.doQty + 0.0005 ||
        r.unissuedQty > 0.0005
    );
  }, [db]);

  const foodMemos = useMemo(
    () => buildFoodAllocOrderMemos(db, month0, year),
    [db, month0, year]
  );

  const omsOrderCategory = useMemo<'ওএমএস সাধারণ' | 'ওএমএস বিশেষ' | 'ওএমএস টিসিবি' | null>(() => {
    if (activeReportId === 'alloc-oms-gen') return 'ওএমএস সাধারণ';
    if (activeReportId === 'alloc-oms-spc') return 'ওএমএস বিশেষ';
    if (activeReportId === 'alloc-oms-tcb') return 'ওএমএস টিসিবি';
    return null;
  }, [activeReportId]);

  const omsMemos = useMemo(
    () => (omsOrderCategory ? buildOmsAllocOrderMemos(db, omsOrderCategory, month0, year) : []),
    [db, omsOrderCategory, month0, year]
  );

  useEffect(() => {
    if (foodMemos.length > 0 && !foodMemos.some((m) => m.memo === selectedFoodMemo)) {
      setSelectedFoodMemo(foodMemos[0].memo);
    } else if (foodMemos.length === 0 && selectedFoodMemo) {
      setSelectedFoodMemo('');
    }
  }, [foodMemos, selectedFoodMemo]);

  useEffect(() => {
    if (!omsOrderCategory) return;
    if (omsMemos.length > 0 && !omsMemos.some((m) => m.memo === selectedOmsMemo)) {
      setSelectedOmsMemo(omsMemos[0].memo);
    } else if (omsMemos.length === 0 && selectedOmsMemo) {
      setSelectedOmsMemo('');
    }
  }, [omsOrderCategory, omsMemos, selectedOmsMemo]);

  const isAllocOrder =
    activeReportId === 'alloc-food' ||
    activeReportId === 'alloc-oms-gen' ||
    activeReportId === 'alloc-oms-spc' ||
    activeReportId === 'alloc-oms-tcb';

  const memoKey =
    activeReportId === 'alloc-food'
      ? `alloc-food::${year}-${month0}::${selectedFoodMemo}`
      : omsOrderCategory
        ? `${activeReportId}::${year}-${month0}::${selectedOmsMemo}`
        : `${activeReportId}::${year}-${month0}`;

  const foodOrderData = useMemo(() => {
    if (activeReportId !== 'alloc-food') return null;
    const includeKernel = foodKernelMode === 'with_kernel';
    return buildFoodAllocOrderRows(db, month0, year, selectedFoodMemo, {
      includeKernel,
      kernelMemo: includeKernel ? 'PERCENT' : undefined,
      kernelPercent: includeKernel ? 1 : undefined,
    });
  }, [activeReportId, db, month0, year, selectedFoodMemo, foodKernelMode]);

  const omsOrderData = useMemo(() => {
    if (!omsOrderCategory) return null;
    return buildOmsAllocOrderRows(db, omsOrderCategory, month0, year, selectedOmsMemo);
  }, [omsOrderCategory, db, month0, year, selectedOmsMemo]);

  const selectedUcfMemo =
    activeReportId === 'alloc-food'
      ? selectedFoodMemo
      : omsOrderCategory
        ? selectedOmsMemo
        : '';
  const selectedUcfSuffix = extractMemoSuffix(selectedUcfMemo, template.memoPrefix);
  const defaultMemoNo = selectedUcfSuffix
    ? `${template.memoPrefix}${selectedUcfSuffix}`
    : `${template.memoPrefix}`;
  const defaultMemoDate =
    foodOrderData?.memoDate || omsOrderData?.memoDate || todayIsoDate();

  const savedMemo = db.reportMemos[memoKey];
  const currentMemoNo = savedMemo?.no?.trim() ? savedMemo.no : defaultMemoNo;
  const currentMemoDate = savedMemo?.date?.trim() ? savedMemo.date : defaultMemoDate;

  const currentAllocForm = useMemo<AllocOrderFormFields>(() => {
    const saved = db.allocOrderForms?.[memoKey];
    const defaultMinMemo =
      foodOrderData?.sourceMinistryMemo || omsOrderData?.sourceMinistryMemo || '';
    const defaultMinDate =
      foodOrderData?.sourceMinistryDate || omsOrderData?.sourceMinistryDate || '';
    const defaultDcMemo = omsOrderData?.sourceDcMemo || '';
    const defaultDcDate = omsOrderData?.sourceDcDate || '';
    const defaultDfcMemo =
      foodOrderData?.sourceDfcMemo || omsOrderData?.sourceDfcMemo || '';
    const defaultDfcDate =
      foodOrderData?.sourceDfcDate || omsOrderData?.sourceDfcDate || '';
    const monthName = BENGALI_MONTHS[month0] || '';
    const yrBn = toBnDigits(year);
    const defaultMiller =
      db.millers.length > 0
        ? db.millers
            .map(
              (m) =>
                `মেসার্স ${m.millName || m.millerName}${m.address ? `, ${m.address}` : ''}`
            )
            .join(' / ')
        : '';
    const defaultOwnSuffix =
      extractMemoSuffix(currentMemoNo, template.memoPrefix) || selectedUcfSuffix;
    const resolvedOwnDate = currentMemoDate || todayIsoDate();
    return {
      ministryDate: defaultMinDate || saved?.ministryDate?.trim() || '',
      ministryMemo: defaultMinMemo || saved?.ministryMemo?.trim() || '',
      dcDate: defaultDcDate || saved?.dcDate?.trim() || '',
      dcMemo: defaultDcMemo || saved?.dcMemo?.trim() || '',
      dfcMonth: monthName,
      dfcDate: defaultDfcDate || saved?.dfcDate?.trim() || '',
      dfcMemo: defaultDfcMemo || saved?.dfcMemo?.trim() || '',
      targetMonth: `${monthName} ${yrBn}`,
      deadlineDate: autoDeadlineIsoDate(resolvedOwnDate, year, month0),
      ownMemoSuffix: defaultOwnSuffix,
      bangabdaDate: formatBangabdaDate(resolvedOwnDate),
      ownDate: resolvedOwnDate,
      copyMillerText: defaultMiller,
      openingStockByRow: {},
    };
  }, [
    db.allocOrderForms,
    db.millers,
    memoKey,
    foodOrderData,
    omsOrderData,
    month0,
    year,
    currentMemoNo,
    currentMemoDate,
    selectedUcfSuffix,
    template.memoPrefix,
  ]);

  interface ReportPayload {
    subjectResolved: string;
    introResolved: string;
    conditionsResolved?: string;
    copyMemoResolved?: string;
    copyListResolved?: string;
    copyDateResolved?: string;
    hideTopMemoRow?: boolean;
    columns: DocColumnGroup[];
    leaves: DocColumnLeaf[];
    mergeDepth: number;
    showSerialCol: boolean;
    serialColLabel: string;
    colWidths?: string[];
    tableLayoutFixed?: boolean;
    footerLayout?: 'side_by_side' | 'sign_above_copy';
    rows: DocRowData[];
    totalRow: DocTotalRow | null;
    secondarySection?: { title: string; columns: DocColumnLeaf[]; rows: DocRowData[] } | null;
    kpiBoxes?: DocKpiBox[];
    bannerNote?: string;
    showRecipientBlock: boolean;
    showIntroBlock: boolean;
    showCopyBlock: boolean;
  }

  const reportPayload = useMemo<ReportPayload>((): ReportPayload => {
    const monthName = BENGALI_MONTHS[month0] || '';
    const yrBn = toBnDigits(year);
    const fmtDateOrBlank = (iso: string) =>
      iso && iso.trim() ? formatBnDate(iso.trim()) : '              ';
    const boxToken = (val: string | undefined, blank = '              ') =>
      `[[${(val ?? '').trim() || blank}]]`;

    const resolvePlaceholders = (s: string) => {
      if (!s) return '';
      const f = currentAllocForm;
      const effSuffix =
        f.ownMemoSuffix.trim() ||
        (currentMemoNo.startsWith(template.memoPrefix)
          ? currentMemoNo.slice(template.memoPrefix.length).trim()
          : '');
      return s
        .replace(/\{মাস\}/g, monthName)
        .replace(/\{বছর\}/g, yrBn)
        .replace(/\{মন্ত্রণালয়_তারিখ\}/g, boxToken(fmtDateOrBlank(f.ministryDate)))
        .replace(/\{মন্ত্রণালয়_স্মারক\}/g, boxToken(f.ministryMemo, '                    '))
        .replace(/\{ডিসি_তারিখ\}/g, boxToken(fmtDateOrBlank(f.dcDate)))
        .replace(/\{ডিসি_স্মারক\}/g, boxToken(f.dcMemo, '                    '))
        .replace(/\{ডিএফসি_মাস\}/g, boxToken(f.dfcMonth || monthName, '          '))
        .replace(/\{ডিএফসি_তারিখ\}/g, boxToken(fmtDateOrBlank(f.dfcDate)))
        .replace(/\{ডিএফসি_স্মারক\}/g, boxToken(f.dfcMemo, '                    '))
        .replace(/\{মাস_বক্স\}/g, boxToken(f.targetMonth || `${monthName} ${yrBn}`, '            '))
        .replace(/\{শেষ_তারিখ\}/g, boxToken(fmtDateOrBlank(f.deadlineDate)))
        .replace(/\{স্মারক_শেষ\}/g, boxToken(effSuffix, '          '))
        .replace(/\{তারিখ_বক্স\}/g, boxToken(fmtDateOrBlank(f.ownDate || currentMemoDate)))
        .replace(/\{বঙ্গাব্দ\}/g, boxToken(f.bangabdaDate, '            '))
        .replace(/\{মিল_মালিক\}/g, boxToken(f.copyMillerText, '                                        '));
    };

    if (activeReportId === 'monthly-dist') {
      const rep = buildMonthlyDistributionReport(db, month0, year, 'ALL');
      const cPrevR: DocColumnLeaf = { key: 'pr', label: colLabel('pr', 'চাল'), align: 'right' };
      const cPrevW: DocColumnLeaf = { key: 'pw', label: colLabel('pw', 'গম'), align: 'right' };
      const cPrevA: DocColumnLeaf = { key: 'pa', label: colLabel('pa', 'আটা'), align: 'right' };
      const cPrevK: DocColumnLeaf = { key: 'pk', label: colLabel('pk', 'কার্নেল'), align: 'right' };
      const cCurR: DocColumnLeaf = { key: 'cr', label: colLabel('cr', 'চাল'), align: 'right' };
      const cCurW: DocColumnLeaf = { key: 'cw', label: colLabel('cw', 'গম'), align: 'right' };
      const cCurA: DocColumnLeaf = { key: 'ca', label: colLabel('ca', 'আটা'), align: 'right' };
      const cCurK: DocColumnLeaf = { key: 'ck', label: colLabel('ck', 'কার্নেল'), align: 'right' };
      const cTotR: DocColumnLeaf = { key: 'tr', label: colLabel('tr', 'চাল'), align: 'right' };
      const cTotW: DocColumnLeaf = { key: 'tw', label: colLabel('tw', 'গম'), align: 'right' };
      const cTotA: DocColumnLeaf = { key: 'ta', label: colLabel('ta', 'আটা'), align: 'right' };
      const cTotK: DocColumnLeaf = { key: 'tk', label: colLabel('tk', 'কার্নেল'), align: 'right' };
      const cCat: DocColumnLeaf = { key: 'cat', label: colLabel('cat', 'খাতের নাম'), align: 'left' };

      const columns: DocColumnGroup[] = [
        { label: colLabel('cat', 'খাতের নাম'), leaf: cCat },
        { label: `পূর্ববর্তী ক্রমপুঞ্জিত (${rep.fiscalYear})`, children: [cPrevR, cPrevW, cPrevA, cPrevK] },
        { label: `বর্তমান মাস (${rep.monthName} ${yrBn})`, children: [cCurR, cCurW, cCurA, cCurK] },
        { label: 'বর্তমান ক্রমপুঞ্জিত', children: [cTotR, cTotW, cTotA, cTotK] },
      ];
      const leaves = [
        cCat,
        cPrevR,
        cPrevW,
        cPrevA,
        cPrevK,
        cCurR,
        cCurW,
        cCurA,
        cCurK,
        cTotR,
        cTotW,
        cTotA,
        cTotK,
      ];
      const rows: DocRowData[] = rep.rows.map((r) => ({
        id: r.category,
        cells: {
          cat: r.categoryBn,
          pr: formatBnQty(r.previousRice),
          pw: formatBnQty(r.previousWheat),
          pa: formatBnQty(r.previousAtta),
          pk: formatBnQty(r.previousKernel),
          cr: formatBnQty(r.currentRice),
          cw: formatBnQty(r.currentWheat),
          ca: formatBnQty(r.currentAtta),
          ck: formatBnQty(r.currentKernel),
          tr: formatBnQty(r.cumulativeRice),
          tw: formatBnQty(r.cumulativeWheat),
          ta: formatBnQty(r.cumulativeAtta),
          tk: formatBnQty(r.cumulativeKernel),
        },
        hasSignature: false,
      }));
      const totalRow: DocTotalRow = {
        label: template.totalLabel || 'মোট',
        labelColSpan: 1,
        values: {
          pr: formatBnQty(rep.totals.previousRice),
          pw: formatBnQty(rep.totals.previousWheat),
          pa: formatBnQty(rep.totals.previousAtta),
          pk: formatBnQty(rep.totals.previousKernel),
          cr: formatBnQty(rep.totals.currentRice),
          cw: formatBnQty(rep.totals.currentWheat),
          ca: formatBnQty(rep.totals.currentAtta),
          ck: formatBnQty(rep.totals.currentKernel),
          tr: formatBnQty(rep.totals.cumulativeRice),
          tw: formatBnQty(rep.totals.cumulativeWheat),
          ta: formatBnQty(rep.totals.cumulativeAtta),
          tk: formatBnQty(rep.totals.cumulativeKernel),
        },
      };
      return {
        subjectResolved: `${resolvePlaceholders(template.subject)} (মাস: ${monthName} ${yrBn}, অর্থবছর: ${rep.fiscalYear})`,
        introResolved: resolvePlaceholders(template.introText),
        columns,
        leaves,
        mergeDepth: 0,
        showSerialCol: true,
        serialColLabel: 'ক্রম',
        colWidths: [
          '3.6%',
          '14.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
          '6.8%',
        ],
        tableLayoutFixed: false,
        rows,
        totalRow,
        showRecipientBlock: true,
        showIntroBlock: Boolean(template.introText.trim()),
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'alloc-summary') {
      const sum = buildAllocationSummaryReport(db, 'ALL');
      const leaves: DocColumnLeaf[] = [
        { key: 'cat', label: colLabel('cat', 'খাতের নাম'), align: 'left' },
        { key: 'grain', label: colLabel('grain', 'শস্য'), align: 'center' },
        { key: 'ministry', label: colLabel('ministry', 'মন্ত্রণালয় বরাদ্দ'), align: 'right' },
        { key: 'dc', label: colLabel('dc', 'ডিসি বরাদ্দ'), align: 'right' },
        { key: 'sub', label: colLabel('sub', 'উপবরাদ্দ'), align: 'right' },
        { key: 'doQty', label: colLabel('doQty', 'ডিও ইস্যু (বিতরণ)'), align: 'right' },
        { key: 'remaining', label: colLabel('remaining', 'অবশিষ্ট / বাকি'), align: 'right' },
      ];
      return {
        subjectResolved: resolvePlaceholders(template.subject),
        introResolved: resolvePlaceholders(template.introText),
        columns: leaves.map((l) => ({ label: l.label, leaf: l })),
        leaves,
        mergeDepth: 0,
        showSerialCol: true,
        serialColLabel: 'ক্রম',
        colWidths: ['5%', '23%', '8%', '13%', '13%', '13%', '13%', '12%'],
        tableLayoutFixed: true,
        rows: sum.rows.map((r) => ({
          id: r.category,
          cells: {
            cat: r.categoryBn,
            grain: r.grain,
            ministry: formatBnQty(r.ministryQty),
            dc: r.dcQty === null ? '—' : formatBnQty(r.dcQty),
            sub: formatBnQty(r.subQty),
            doQty: formatBnQty(r.doQty),
            remaining: formatBnQty(r.unissuedQty),
          },
          hasSignature: false,
        })),
        totalRow: {
          label: template.totalLabel || 'সর্বমোট',
          labelColSpan: 2,
          values: {
            ministry: formatBnQty(sum.totals.ministryQty),
            dc: formatBnQty(sum.totals.dcQty),
            sub: formatBnQty(sum.totals.subQty),
            doQty: formatBnQty(sum.totals.doQty),
            remaining: formatBnQty(sum.totals.unissuedQty),
          },
        },
        kpiBoxes: [
          {
            label: 'মোট মন্ত্রণালয় বরাদ্দ',
            value: `${formatBnQty(sum.totals.ministryQty)} মে.টন`,
            tone: 'neutral',
          },
          {
            label: 'মোট উপবরাদ্দ',
            value: `${formatBnQty(sum.totals.subQty)} মে.টন`,
            tone: 'emerald',
          },
          {
            label: 'মোট ডিও ইস্যু',
            value: `${formatBnQty(sum.totals.doQty)} মে.টন`,
            tone: 'emerald',
          },
          {
            label: 'ডিও ইস্যু বাকি',
            value: `${formatBnQty(sum.totals.unissuedQty)} মে.টন`,
            tone: 'amber',
          },
        ],
        showRecipientBlock: false,
        showIntroBlock: Boolean(template.introText.trim()),
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'monthly-cash') {
      const rep = buildCashSaleReport(db, month0, year);
      const leaves: DocColumnLeaf[] = [
        { key: 'monthName', label: colLabel('monthName', 'মাসের নাম'), align: 'center' },
        { key: 'channel', label: colLabel('channel', 'চ্যানেলের নাম'), align: 'left' },
        { key: 'product', label: colLabel('product', 'পণ্য'), align: 'center' },
        { key: 'qty', label: colLabel('qty', 'সরবরাহকৃত পরিমাণ'), align: 'right' },
        { key: 'rate', label: colLabel('rate', 'দর'), align: 'right' },
        { key: 'challan', label: colLabel('challan', 'চালান নং / তারিখ'), align: 'left' },
        { key: 'amount', label: colLabel('amount', 'জমাকৃত টাকার পরিমাণ'), align: 'right' },
      ];
      return {
        subjectResolved: `${resolvePlaceholders(template.subject)} (মাস: ${rep.monthName} ${yrBn})`,
        introResolved: resolvePlaceholders(template.introText),
        columns: leaves.map((l) => ({ label: l.label, leaf: l })),
        leaves,
        mergeDepth: 0,
        showSerialCol: true,
        serialColLabel: 'ক্রম',
        colWidths: ['4%', '8.5%', '10.5%', '5.5%', '8.5%', '7.5%', '44.5%', '11%'],
        tableLayoutFixed: true,
        rows: rep.rows.map((r, i) => ({
          id: `cash-${i}`,
          cells: {
            monthName: `${rep.monthName} ${yrBn}`,
            channel: r.channel,
            product: r.product,
            qty: formatBnQty(r.suppliedQuantity),
            rate: formatBnMoney(r.rate),
            challan: r.challan,
            amount: formatBnMoney(r.amount),
          },
          hasSignature: false,
        })),
        totalRow: {
          label: template.totalLabel || 'মোট',
          labelColSpan: 3,
          values: {
            qty: formatBnQty(rep.totalQty),
            rate: '',
            challan: '',
            amount: formatBnMoney(rep.totalAmount),
          },
        },
        showRecipientBlock: true,
        showIntroBlock: Boolean(template.introText.trim()),
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'monthly-ep') {
      const rep = buildEPRationReport(db, month0, year);
      const cInst: DocColumnLeaf = {
        key: 'inst',
        label: colLabel('inst', 'বিশেষ জরুরী গ্রাহক (ইপি)'),
        align: 'left',
      };
      const cCr: DocColumnLeaf = { key: 'cr', label: colLabel('cr', 'চাল'), align: 'right' };
      const cCw: DocColumnLeaf = { key: 'cw', label: colLabel('cw', 'গম'), align: 'right' };
      const cTr: DocColumnLeaf = { key: 'tr', label: colLabel('tr', 'চাল'), align: 'right' };
      const cTw: DocColumnLeaf = { key: 'tw', label: colLabel('tw', 'গম'), align: 'right' };
      const columns: DocColumnGroup[] = [
        { label: cInst.label, leaf: cInst },
        { label: `চলতি মাসের বিতরণ (মে.টন) — ${rep.monthName} ${yrBn}`, children: [cCr, cCw] },
        { label: `ক্রমপুঞ্জিত মোট (মে.টন) — ${rep.fiscalYear}`, children: [cTr, cTw] },
      ];
      const leaves = [cInst, cCr, cCw, cTr, cTw];
      const tot = rep.rows.reduce(
        (a, r) => ({
          cr: round3(a.cr + r.rice),
          cw: round3(a.cw + r.wheat),
          tr: round3(a.tr + r.cumulativeRice),
          tw: round3(a.tw + r.cumulativeWheat),
        }),
        { cr: 0, cw: 0, tr: 0, tw: 0 }
      );
      return {
        subjectResolved: `${resolvePlaceholders(template.subject)} (মাস: ${rep.monthName} ${yrBn})`,
        introResolved: resolvePlaceholders(template.introText),
        columns,
        leaves,
        mergeDepth: 0,
        showSerialCol: true,
        serialColLabel: 'ক্রম',
        colWidths: ['7%', '37%', '14%', '14%', '14%', '14%'],
        tableLayoutFixed: true,
        rows: rep.rows.map((r, i) => ({
          id: `ep-${i}`,
          cells: {
            inst: r.institution,
            cr: formatBnQty(r.rice),
            cw: formatBnQty(r.wheat),
            tr: formatBnQty(r.cumulativeRice),
            tw: formatBnQty(r.cumulativeWheat),
          },
          hasSignature: false,
        })),
        totalRow: {
          label: template.totalLabel || 'সর্বমোট',
          labelColSpan: 1,
          values: {
            cr: formatBnQty(tot.cr),
            cw: formatBnQty(tot.cw),
            tr: formatBnQty(tot.tr),
            tw: formatBnQty(tot.tw),
          },
        },
        showRecipientBlock: true,
        showIntroBlock: Boolean(template.introText.trim()),
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'alloc-food' && foodOrderData) {
      const includeKernel = foodKernelMode === 'with_kernel';
      const ord = foodOrderData;

      const leaves: DocColumnLeaf[] = includeKernel
        ? [
            { key: 'union', label: colLabel('union', 'ইউনিয়ন'), align: 'center' },
            { key: 'serial', label: colLabel('serial', 'ক্রমিক'), align: 'center' },
            { key: 'center', label: colLabel('center', 'বিক্রয় কেন্দ্র'), align: 'left' },
            { key: 'dealer', label: colLabel('dealer', 'ডিলারের নাম'), align: 'left' },
            { key: 'cards', label: colLabel('cards', 'কার্ড'), align: 'center' },
            { key: 'totalRice', label: colLabel('totalRice', 'চালের পরিমাণ\n(মে.টন)'), align: 'right' },
            { key: 'openingStock', label: colLabel('openingStock', 'প্রারম্ভিক মজুদ\n(মে.টন)'), align: 'right' },
            { key: 'qtyWithoutKernel', label: colLabel('qtyWithoutKernel', 'চলতি মাসের বরাদ্দের\nপরিমাণ (মে.টন)'), align: 'right' },
            { key: 'kernelQty', label: colLabel('kernelQty', 'বরাদ্দকৃত কার্নেলের\nপরিমাণ (মে.টন)'), align: 'right' },
            { key: 'qtyWithKernel', label: colLabel('qtyWithKernel', 'কার্নেলসহ মোট\nবরাদ্দের পরিমাণ (মে.টন)'), align: 'right' },
            { key: 'amount', label: colLabel('amount', 'টাকা'), align: 'right' },
            { key: 'tax', label: colLabel('tax', 'আয়কর'), align: 'right' },
          ]
        : [
            { key: 'union', label: colLabel('union', 'ইউনিয়ন'), align: 'center' },
            { key: 'serial', label: colLabel('serial', 'ক্রমিক'), align: 'center' },
            { key: 'center', label: colLabel('center', 'বিক্রয় কেন্দ্র'), align: 'left' },
            { key: 'dealer', label: colLabel('dealer', 'ডিলারের নাম'), align: 'left' },
            { key: 'cards', label: colLabel('cards', 'কার্ড'), align: 'center' },
            { key: 'totalRice', label: colLabel('totalRice', 'চালের পরিমাণ\n(মে.টন)'), align: 'right' },
            { key: 'openingStock', label: colLabel('openingStock', 'প্রারম্ভিক মজুদ\n(মে.টন)'), align: 'right' },
            { key: 'qty', label: colLabel('qty', 'চলতি মাসের বরাদ্দের\nপরিমাণ (মে.টন)'), align: 'right' },
            { key: 'amount', label: colLabel('amount', 'টাকা'), align: 'right' },
            { key: 'tax', label: colLabel('tax', 'আয়কর'), align: 'right' },
          ];

      const rowComputed = ord.rows.map((r) => ({
        r,
        openingStock: r.openingStockQty,
        totalRice: r.totalRiceRequirement,
        qtyWithoutKernel: r.qtyWithoutKernel,
        kernelQty: r.kernelQty,
        qtyWithKernel: r.qtyWithKernel,
        amount: r.amount,
        tax: r.tax,
      }));

      const totCards = rowComputed.reduce((s, x) => s + x.r.cards, 0);
      const totTotalRice = round3(rowComputed.reduce((s, x) => s + x.totalRice, 0));
      const totOpeningStock = round3(rowComputed.reduce((s, x) => s + x.openingStock, 0));
      const totWithoutKernel = round3(rowComputed.reduce((s, x) => s + x.qtyWithoutKernel, 0));
      const totKernelQty = round3(rowComputed.reduce((s, x) => s + x.kernelQty, 0));
      const totWithKernel = round3(rowComputed.reduce((s, x) => s + x.qtyWithKernel, 0));
      const totAmt = round2(rowComputed.reduce((s, x) => s + x.amount, 0));
      const totTax = round2(rowComputed.reduce((s, x) => s + x.tax, 0));

      let prevUnion = '';
      let unionDealerIdx = 0;

      const resolvedCopyMemo = template.copyMemoText
        ? resolvePlaceholders(template.copyMemoText)
        : `স্মারক নং- ${template.memoPrefix}[[${currentAllocForm.ownMemoSuffix || '          '}]]`;
      const resolvedCopyList = resolvePlaceholders(template.copyList);
      const copyDateResolved = `${boxToken(fmtDateOrBlank(currentAllocForm.ownDate || currentMemoDate))} খ্রিঃ`;

      return {
        subjectResolved: resolvePlaceholders(template.subject),
        introResolved: resolvePlaceholders(template.introText),
        conditionsResolved: resolvePlaceholders(template.conditionsText || ''),
        copyMemoResolved: resolvedCopyMemo,
        copyListResolved: resolvedCopyList,
        copyDateResolved,
        hideTopMemoRow: template.hideTopMemoRow ?? true,
        columns: leaves.map((l) => ({ label: l.label, leaf: l })),
        leaves,
        mergeDepth: 1,
        showSerialCol: false,
        serialColLabel: 'ক্রমিক',
        colWidths: includeKernel
          ? ['8%', '4%', '10%', '16%', '6%', '8%', '8%', '8%', '8%', '9%', '8%', '7%']
          : ['10%', '5%', '12%', '19%', '7%', '9%', '9%', '10%', '10%', '9%'],
        tableLayoutFixed: true,
        footerLayout: 'sign_above_copy',
        rows: rowComputed.map(
          (
            { r, openingStock, totalRice, qtyWithoutKernel, kernelQty, qtyWithKernel, amount, tax },
            i
          ) => {
            const cleanUnion = r.union.replace(/\s*ইউনিয়ন$/g, '').trim() || r.union;
            if (cleanUnion !== prevUnion) {
              prevUnion = cleanUnion;
              unionDealerIdx = 1;
            } else {
              unionDealerIdx++;
            }
            return {
              id: `fa-${i}`,
              groupValues: [cleanUnion],
              cells: {
                union: cleanUnion,
                serial: toBnDigits(i + 1),
                center: r.salesCenter || '—',
                dealer: `${toBnDigits(unionDealerIdx)}। ${template.namePrefix || ''}${r.dealerName}`,
                cards: formatBnInt(r.cards),
                totalRice: formatBnQty(totalRice),
                openingStock: formatBnQty(openingStock),
                ...(includeKernel
                  ? {
                      qtyWithoutKernel: formatBnQty(qtyWithoutKernel),
                      kernelQty: formatBnQty(kernelQty),
                      qtyWithKernel: formatBnQty(qtyWithKernel),
                    }
                  : {
                      qty: formatBnQty(qtyWithoutKernel),
                    }),
                amount: formatBnMoney(amount),
                tax: formatBnMoney(tax),
              },
              hasSignature: false,
            };
          }
        ),
        totalRow: {
          label: template.totalLabel || 'সর্বমোট',
          labelColSpan: 4,
          values: includeKernel
            ? {
                cards: formatBnInt(totCards),
                totalRice: formatBnQty(totTotalRice),
                openingStock: formatBnQty(totOpeningStock),
                qtyWithoutKernel: formatBnQty(totWithoutKernel),
                kernelQty: formatBnQty(totKernelQty),
                qtyWithKernel: formatBnQty(totWithKernel),
                amount: formatBnMoney(totAmt),
                tax: formatBnMoney(totTax),
              }
            : {
                cards: formatBnInt(totCards),
                totalRice: formatBnQty(totTotalRice),
                openingStock: formatBnQty(totOpeningStock),
                qty: formatBnQty(totWithoutKernel),
                amount: formatBnMoney(totAmt),
                tax: formatBnMoney(totTax),
              },
        },
        showRecipientBlock: false,
        showIntroBlock: true,
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (omsOrderCategory && omsOrderData) {
      const isTCB = omsOrderCategory === 'ওএমএস টিসিবি';
      const ord = omsOrderData;

      const rowComputed = ord.rows.map((r) => ({
        r,
        openingStock: r.openingStockQty,
        grossRice: r.totalRiceRequirement,
        effQty: r.qty,
        effAmount: r.amount,
        effTax: r.tax,
        comment: r.comment || '',
      }));

      const resolvedCopyMemo = template.copyMemoText
        ? resolvePlaceholders(template.copyMemoText)
        : `স্মারক নং- ${template.memoPrefix}[[${currentAllocForm.ownMemoSuffix || '          '}]]`;
      const resolvedCopyList = resolvePlaceholders(template.copyList);
      const copyDateResolved = isTCB
        ? `${boxToken(currentAllocForm.bangabdaDate, '            ')} বঙ্গাব্দ / ${boxToken(fmtDateOrBlank(currentAllocForm.ownDate || currentMemoDate))} খ্রিঃ।`
        : `${boxToken(fmtDateOrBlank(currentAllocForm.ownDate || currentMemoDate))} খ্রিঃ`;

      if (isTCB) {
        const leaves: DocColumnLeaf[] = [
          { key: 'serial', label: colLabel('serial', 'ক্রমিক\nনং'), align: 'center' },
          { key: 'union', label: colLabel('union', 'ইউনিয়ন / পৌরসভার\nনাম'), align: 'left' },
          { key: 'dealer', label: colLabel('dealer', 'মনোনীত খাদ্যবান্ধব /\nওএমএস ডিলারের নাম'), align: 'left' },
          { key: 'cards', label: colLabel('cards', 'এক্টিভেটেড\nকার্ড সংখ্যা'), align: 'center' },
          { key: 'rice5kg', label: colLabel('rice5kg', '৫ কেজি হারে চালের\nপরিমাণ (মে. টন)'), align: 'right' },
          { key: 'openingStock', label: colLabel('openingStock', 'প্রারম্ভিক মজুদ\n(মে. টন)'), align: 'right' },
          { key: 'qty', label: colLabel('qty', 'চলতি মাসের বরাদ্দের\nপরিমাণ (মে. টন)'), align: 'right' },
          { key: 'amount', label: colLabel('amount', 'চালের মোট মূল্য\n(টাকা)'), align: 'right' },
          { key: 'tax', label: colLabel('tax', 'আয়কর\n(টাকা)'), align: 'right' },
          { key: 'comment', label: colLabel('comment', 'মন্তব্য'), align: 'center' },
        ];

        const totCards = rowComputed.reduce((s, x) => s + x.r.cards, 0);
        const totGross = round3(rowComputed.reduce((s, x) => s + x.grossRice, 0));
        const totOpen = round3(rowComputed.reduce((s, x) => s + x.openingStock, 0));
        const totQty = round3(rowComputed.reduce((s, x) => s + x.effQty, 0));
        const totAmt = round2(rowComputed.reduce((s, x) => s + x.effAmount, 0));
        const totTax = round2(rowComputed.reduce((s, x) => s + x.effTax, 0));

        return {
          subjectResolved: resolvePlaceholders(template.subject),
          introResolved: resolvePlaceholders(template.introText),
          conditionsResolved: resolvePlaceholders(template.conditionsText || ''),
          copyMemoResolved: resolvedCopyMemo,
          copyListResolved: resolvedCopyList,
          copyDateResolved,
          hideTopMemoRow: template.hideTopMemoRow ?? true,
          columns: leaves.map((l) => ({ label: l.label, leaf: l })),
          leaves,
          mergeDepth: 0,
          showSerialCol: false,
          serialColLabel: 'ক্রমিক',
          colWidths: ['5%', '12%', '17%', '9%', '10%', '9%', '10%', '11%', '9%', '8%'],
          tableLayoutFixed: true,
          footerLayout: 'sign_above_copy',
          rows: rowComputed.map(({ r, openingStock, grossRice, effQty, effAmount, effTax, comment }, i) => ({
            id: `otcb-${i}`,
            cells: {
              serial: toBnDigits(i + 1),
              union: r.union || '—',
              dealer: `${template.namePrefix || ''}${r.dealerName}${r.salesCenter && r.salesCenter !== '—' ? `\n(${r.salesCenter})` : ''}`,
              cards: formatBnInt(r.cards),
              rice5kg: formatBnQty(grossRice),
              openingStock: formatBnQty(openingStock),
              qty: formatBnQty(effQty),
              amount: formatBnMoney(effAmount),
              tax: formatBnMoney(effTax),
              comment: comment || '',
            },
            hasSignature: false,
          })),
          totalRow: {
            label: template.totalLabel || 'সর্বমোট',
            labelColSpan: 3,
            values: {
              cards: formatBnInt(totCards),
              rice5kg: formatBnQty(totGross),
              openingStock: formatBnQty(totOpen),
              qty: formatBnQty(totQty),
              amount: formatBnMoney(totAmt),
              tax: formatBnMoney(totTax),
              comment: '',
            },
          },
          showRecipientBlock: false,
          showIntroBlock: true,
          showCopyBlock: template.showCopyBlock,
        };
      }

      // alloc-oms-spc & alloc-oms-gen (9 columns matching Page 3-4)
      const leaves: DocColumnLeaf[] = [
        { key: 'union', label: colLabel('union', 'ইউনিয়ন / পৌরসভা'), align: 'center' },
        { key: 'serial', label: colLabel('serial', 'ক্রমিক'), align: 'center' },
        { key: 'center', label: colLabel('center', 'বিক্রয় কেন্দ্র'), align: 'left' },
        { key: 'dealer', label: colLabel('dealer', 'ডিলারের নাম'), align: 'left' },
        { key: 'totalRice', label: colLabel('totalRice', 'চালের পরিমাণ\n(মে.টন)'), align: 'right' },
        { key: 'openingStock', label: colLabel('openingStock', 'প্রারম্ভিক মজুদ\n(মে.টন)'), align: 'right' },
        { key: 'qty', label: colLabel('qty', 'বরাদ্দের পরিমাণ\n(মে.টন)'), align: 'right' },
        { key: 'amount', label: colLabel('amount', 'টাকা'), align: 'right' },
        { key: 'tax', label: colLabel('tax', 'আয়কর'), align: 'right' },
      ];

      const totGross = round3(rowComputed.reduce((s, x) => s + x.grossRice, 0));
      const totOpen = round3(rowComputed.reduce((s, x) => s + x.openingStock, 0));
      const totQty = round3(rowComputed.reduce((s, x) => s + x.effQty, 0));
      const totAmt = round2(rowComputed.reduce((s, x) => s + x.effAmount, 0));
      const totTax = round2(rowComputed.reduce((s, x) => s + x.effTax, 0));

      return {
        subjectResolved: resolvePlaceholders(template.subject),
        introResolved: resolvePlaceholders(template.introText),
        conditionsResolved: resolvePlaceholders(template.conditionsText || ''),
        copyMemoResolved: resolvedCopyMemo,
        copyListResolved: resolvedCopyList,
        copyDateResolved,
        hideTopMemoRow: template.hideTopMemoRow ?? true,
        columns: leaves.map((l) => ({ label: l.label, leaf: l })),
        leaves,
        mergeDepth: 1,
        showSerialCol: false,
        serialColLabel: 'ক্রমিক',
        colWidths: ['12%', '6%', '14%', '20%', '10%', '10%', '10%', '10%', '8%'],
        tableLayoutFixed: true,
        footerLayout: 'sign_above_copy',
        rows: rowComputed.map(({ r, openingStock, grossRice, effQty, effAmount, effTax }, i) => ({
          id: `ooms-${i}`,
          groupValues: [r.union || '—'],
          cells: {
            union: r.union || '—',
            serial: toBnDigits(i + 1),
            center: r.salesCenter || '—',
            dealer: `${toBnDigits(i + 1)}। ${template.namePrefix || ''}${r.dealerName}`,
            totalRice: formatBnQty(grossRice),
            openingStock: formatBnQty(openingStock),
            qty: formatBnQty(effQty),
            amount: formatBnMoney(effAmount),
            tax: formatBnMoney(effTax),
          },
          hasSignature: false,
        })),
        totalRow: {
          label: template.totalLabel || 'সর্বমোট',
          labelColSpan: 4,
          values: {
            totalRice: formatBnQty(totGross),
            openingStock: formatBnQty(totOpen),
            qty: formatBnQty(totQty),
            amount: formatBnMoney(totAmt),
            tax: formatBnMoney(totTax),
          },
        },
        showRecipientBlock: false,
        showIntroBlock: true,
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'do-reg-book') {
      const stats = getDoBookRegisterStats(db);
      const leaves: DocColumnLeaf[] = [
        { key: 'bookNo', label: colLabel('bookNo', 'বই নং'), align: 'center' },
        { key: 'startPage', label: colLabel('startPage', 'শুরু পাতা'), align: 'center' },
        { key: 'endPage', label: colLabel('endPage', 'শেষ পাতা'), align: 'center' },
        { key: 'total', label: colLabel('total', 'মোট ডিও'), align: 'center' },
        { key: 'used', label: colLabel('used', 'ব্যবহৃত ডিও'), align: 'center' },
      ];
      return {
        subjectResolved: resolvePlaceholders(template.subject),
        introResolved: '',
        columns: leaves.map((l) => ({ label: l.label, leaf: l })),
        leaves,
        mergeDepth: 0,
        showSerialCol: false,
        serialColLabel: 'ক্রম',
        colWidths: ['20%', '20%', '20%', '20%', '20%'],
        tableLayoutFixed: true,
        rows: stats.usedRows.map((r) => ({
          id: `bk-${r.bookNo}`,
          cells: {
            bookNo: toBnDigits(r.bookNo),
            startPage: toBnDigits(r.startPage),
            endPage: toBnDigits(r.endPage),
            total: formatBnInt(r.totalCount),
            used: formatBnInt(r.usedCount),
          },
          hasSignature: false,
        })),
        totalRow: {
          label: template.totalLabel || 'মোট',
          labelColSpan: 3,
          values: {
            total: formatBnInt(stats.summary.totalDO),
            used: formatBnInt(stats.summary.totalUsed),
          },
        },
        secondarySection: {
          title: '২। অব্যবহৃত ডিও পাতার হিসাব',
          columns: [
            { key: 'bookNo', label: 'বই নং', align: 'center' },
            { key: 'unusedRange', label: 'অব্যবহৃত পাতা', align: 'left' },
            { key: 'unusedCount', label: 'মোট অব্যবহৃত', align: 'center' },
          ],
          rows: stats.unusedRows.map((u) => ({
            id: `ubk-${u.bookNo}`,
            cells: {
              bookNo: toBnDigits(u.bookNo),
              unusedRange: u.unusedRange,
              unusedCount: formatBnInt(u.unusedCount),
            },
          })),
        },
        kpiBoxes: [
          { label: 'মোট ডিও', value: formatBnInt(stats.summary.totalDO), tone: 'neutral' },
          { label: 'ব্যবহৃত', value: formatBnInt(stats.summary.totalUsed), tone: 'emerald' },
          { label: 'অব্যবহৃত', value: formatBnInt(stats.summary.totalUnused), tone: 'amber' },
        ],
        showRecipientBlock: false,
        showIntroBlock: false,
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'do-reg-issue') {
      const filtered = db.doIssues
        .filter((d) => {
          const my = parseMonthYear(d.issueDate);
          return Boolean(my && my.year === year && my.month === month0);
        })
        .sort((a, b) => a.issueDate.localeCompare(b.issueDate));

      const selectedSet = new Set(doIssueCols);
      const allIssueLeaves: DocColumnLeaf[] = DO_ISSUE_COLUMN_DEFS.filter((def) =>
        selectedSet.has(def.key)
      ).map((def) => ({
        key: def.key,
        label: colLabel(def.key, def.label),
        align:
          def.key === 'qty'
            ? 'right'
            : def.key === 'doNo' || def.key === 'issueDate' || def.key === 'signature'
              ? 'center'
              : 'left',
        isSignature: def.key === 'signature',
      }));

      const totQty = round3(filtered.reduce((s, d) => s + d.allocationQty, 0));
      const qtyIdx = allIssueLeaves.findIndex((l) => l.key === 'qty');

      return {
        subjectResolved: `${resolvePlaceholders(template.subject)} (মাস: ${monthName} ${yrBn})`,
        introResolved: '',
        columns: allIssueLeaves.map((l) => ({ label: l.label, leaf: l })),
        leaves: allIssueLeaves,
        mergeDepth: 0,
        showSerialCol: true,
        serialColLabel: 'ক্রম',
        rows: filtered.map((d) => {
          const validForDo = (c: AllocationDatabase['challans'][number]): boolean =>
            c.category === d.category &&
            isSamePartyName(c.dealerName, d.recipientName) &&
            isChallanOnOrBeforeDo(c.challanDate, d.issueDate, false);
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
          const mainNo = mainRec?.challanNo || (isChallanOnOrBeforeDo(d.challanDate, d.issueDate, false) ? d.challanNo : undefined);
          const taxNo = taxRec?.challanNo || (isChallanOnOrBeforeDo(d.taxChallanDate, d.issueDate, false) ? d.taxChallanNo : undefined);
          const mainAmt = mainNo ? (mainRec?.depositedAmount || d.challanAmount || 0) : 0;
          const taxAmt = taxNo ? (taxRec?.depositedAmount || d.taxChallanAmount || 0) : 0;
          const challanParts: string[] = [];
          if (mainNo || mainAmt > 0) {
            challanParts.push(
              `মূল: ${mainNo ? toBnDigits(mainNo) : '—'}${mainAmt > 0 ? ` (জমা: ${formatBnMoney(mainAmt)} ৳)` : ''}`
            );
          }
          if (taxNo || taxAmt > 0) {
            challanParts.push(
              `আয়কর: ${taxNo ? toBnDigits(taxNo) : '—'}${taxAmt > 0 ? ` (জমা: ${formatBnMoney(taxAmt)} ৳)` : ''}`
            );
          }
          return {
            id: d.id,
            cells: {
              doNo: toBnDigits(d.doNo),
              issueDate: formatBnDate(d.issueDate),
              category: d.category,
              union: d.unionName || '—',
              recipient: d.recipientName,
              salesCenter: d.salesCenter || '—',
              unoMemo: toBnDigits(d.unoMemo || '—'),
              qty: formatBnQty(d.allocationQty),
              challan: challanParts.length > 0 ? challanParts.join('\n') : '—',
              signature: '',
            },
            hasSignature: Boolean(String(d.doNo || '').trim()),
          };
        }),
        totalRow:
          qtyIdx >= 0
            ? {
                label: template.totalLabel || 'মোট',
                labelColSpan: qtyIdx,
                values: {
                  qty: formatBnQty(totQty),
                },
              }
            : null,
        showRecipientBlock: false,
        showIntroBlock: false,
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'do-reg-unused') {
      const list = getUnusedDoNumbers(db);
      const leaves: DocColumnLeaf[] = [
        { key: 'bookNo', label: colLabel('bookNo', 'বই নং'), align: 'center' },
        { key: 'doNo', label: colLabel('doNo', 'ডিও নং'), align: 'center' },
        { key: 'startPage', label: colLabel('startPage', 'শুরু পাতা'), align: 'center' },
        { key: 'endPage', label: colLabel('endPage', 'শেষ পাতা'), align: 'center' },
      ];
      return {
        subjectResolved: resolvePlaceholders(template.subject),
        introResolved: '',
        columns: leaves.map((l) => ({ label: l.label, leaf: l })),
        leaves,
        mergeDepth: 0,
        showSerialCol: false,
        serialColLabel: 'ক্রম',
        colWidths: ['25%', '25%', '25%', '25%'],
        tableLayoutFixed: true,
        bannerNote: `মোট অব্যবহৃত ডিও: ${formatBnInt(list.length)} টি`,
        rows: list.map((u) => ({
          id: `un-${u.doNo}`,
          cells: {
            bookNo: toBnDigits(u.bookNo),
            doNo: toBnDigits(u.doNo),
            startPage: toBnDigits(u.startPage),
            endPage: toBnDigits(u.endPage),
          },
          hasSignature: false,
        })),
        totalRow: null,
        showRecipientBlock: false,
        showIntroBlock: false,
        showCopyBlock: template.showCopyBlock,
      };
    }

    if (activeReportId === 'dlr-rep-food') {
      const selectedSet = new Set(foodDealerCols);
      const leaves: DocColumnLeaf[] = FOOD_DEALER_COLUMN_DEFS.filter((def) =>
        selectedSet.has(def.key)
      ).map((def) => ({
        key: def.key,
        label: colLabel(def.key, def.label),
        align:
          def.key === 'totalQty' || def.key === 'amount'
            ? 'right'
            : def.key === 'cardCount' || def.key === 'mobile' || def.key === 'poDate'
              ? 'center'
              : 'left',
      }));
      const totCards = db.foodDealers.reduce((s, d) => s + (d.cardCount || 0), 0);
      const totQty = round3(db.foodDealers.reduce((s, d) => s + (d.totalQty || 0), 0));
      const cardsIdx = leaves.findIndex((l) => l.key === 'cardCount');
      const qtyIdx = leaves.findIndex((l) => l.key === 'totalQty');
      const firstTotIdx =
        cardsIdx >= 0 && qtyIdx >= 0
          ? Math.min(cardsIdx, qtyIdx)
          : Math.max(cardsIdx, qtyIdx);

      return {
        subjectResolved: resolvePlaceholders(template.subject),
        introResolved: '',
        columns: leaves.map((l) => ({ label: l.label, leaf: l })),
        leaves,
        mergeDepth: leaves[0]?.key === 'unionName' ? 1 : 0,
        showSerialCol: true,
        serialColLabel: 'ক্রম',
        rows: db.foodDealers.map((d) => ({
          id: d.id,
          groupValues: leaves[0]?.key === 'unionName' ? [d.unionName] : [],
          cells: {
            unionName: d.unionName,
            dealerName: d.dealerName,
            fatherName: d.fatherName || '—',
            nid: d.nid ? toBnDigits(d.nid) : '—',
            mobile: d.mobile ? toBnDigits(d.mobile) : '—',
            salesCenter: d.salesCenter || '—',
            cardCount: formatBnInt(d.cardCount),
            totalQty: formatBnQty(d.totalQty),
            poNo: d.poNo ? toBnDigits(d.poNo) : '—',
            poDate: d.poDate ? formatBnDate(d.poDate) : '—',
            bankName: d.bankName || '—',
            amount: d.amount ? formatBnMoney(d.amount) : '—',
            supervisorName: d.supervisorName || '—',
            supervisorDesignation: d.supervisorDesignation || '—',
            supervisorMobile: d.supervisorMobile ? toBnDigits(d.supervisorMobile) : '—',
          },
          hasSignature: false,
        })),
        totalRow:
          firstTotIdx >= 0
            ? {
                label: template.totalLabel || 'সর্বমোট',
                labelColSpan: firstTotIdx,
                values: {
                  cardCount: formatBnInt(totCards),
                  totalQty: formatBnQty(totQty),
                },
              }
            : null,
        showRecipientBlock: false,
        showIntroBlock: false,
        showCopyBlock: template.showCopyBlock,
      };
    }

    // dlr-rep-oms
    const selectedOmsSet = new Set(omsDealerCols);
    const leaves: DocColumnLeaf[] = OMS_DEALER_COLUMN_DEFS.filter((def) =>
      selectedOmsSet.has(def.key)
    ).map((def) => ({
      key: def.key,
      label: colLabel(def.key, def.label),
      align:
        def.key === 'amount'
          ? 'right'
          : def.key === 'areaType' || def.key === 'mobile' || def.key === 'poDate'
            ? 'center'
            : 'left',
    }));
    const totAmount = round2(db.omsDealers.reduce((s, d) => s + (Number(d.amount) || 0), 0));
    const amtIdx = leaves.findIndex((l) => l.key === 'amount');

    return {
      subjectResolved: resolvePlaceholders(template.subject),
      introResolved: '',
      columns: leaves.map((l) => ({ label: l.label, leaf: l })),
      leaves,
      mergeDepth: 0,
      showSerialCol: true,
      serialColLabel: 'ক্রম',
      rows: db.omsDealers.map((d) => ({
        id: d.id,
        cells: {
          areaType: d.areaType || '—',
          unionName: d.unionName || '—',
          municipality: d.municipality || '—',
          dealerName: d.dealerName,
          fatherName: d.fatherName || '—',
          nid: d.nid ? toBnDigits(d.nid) : '—',
          mobile: d.mobile ? toBnDigits(d.mobile) : '—',
          salesCenter: d.salesCenter || '—',
          poNo: d.poNo ? toBnDigits(d.poNo) : '—',
          poDate: d.poDate ? formatBnDate(d.poDate) : '—',
          bankName: d.bankName || '—',
          amount: d.amount ? formatBnMoney(d.amount) : '—',
          supervisorName: d.supervisorName || '—',
          supervisorDesignation: d.supervisorDesignation || '—',
          supervisorMobile: d.supervisorMobile ? toBnDigits(d.supervisorMobile) : '—',
        },
        hasSignature: false,
      })),
      totalRow:
        amtIdx >= 0 && totAmount > 0
          ? {
              label: template.totalLabel || 'সর্বমোট',
              labelColSpan: amtIdx,
              values: {
                amount: formatBnMoney(totAmount),
              },
            }
          : null,
      showRecipientBlock: false,
      showIntroBlock: false,
      showCopyBlock: template.showCopyBlock,
    };
  }, [
    activeReportId,
    db,
    month0,
    year,
    foodOrderData,
    omsOrderCategory,
    omsOrderData,
    foodKernelMode,
    currentAllocForm,
    currentMemoNo,
    currentMemoDate,
    template,
    doIssueCols,
    foodDealerCols,
    omsDealerCols,
  ]);

  const yearOptions = useMemo(() => {
    const cur = new Date().getFullYear();
    return [cur - 2, cur - 1, cur, cur + 1];
  }, []);

  const currentGroupConfig = useMemo(
    () => REPORT_PARENT_GROUPS.find((g) => g.key === activeParent) || REPORT_PARENT_GROUPS[0],
    [activeParent]
  );

  const showMonthYearRow =
    activeReportId === 'monthly-dist' ||
    activeReportId === 'monthly-cash' ||
    activeReportId === 'monthly-ep' ||
    activeReportId === 'do-reg-issue' ||
    isAllocOrder;

  const currentOrientation = template?.orientation || 'portrait';

  return (
    <div className="space-y-4 font-nikosh">
      {/* Top 4 Parent Report Options + Circular Sub-Menu Icons */}
      <div className="no-print rounded-2xl border border-slate-200 bg-white p-3.5 space-y-3.5">
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
          {REPORT_PARENT_GROUPS.map((grp) => {
            const active = grp.key === activeParent;
            return (
              <button
                key={grp.key}
                type="button"
                onClick={() => {
                  setActiveParent(grp.key);
                  if (!grp.items.some((x) => x.id === activeReportId)) {
                    onSelectReport(grp.items[0].id);
                  }
                }}
                className={`min-h-[44px] rounded-xl px-2.5 py-2 text-xs font-bold border transition-colors cursor-pointer ${
                  active
                    ? 'bg-[#E2136E] text-white border-[#E2136E]'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                {grp.label}
              </button>
            );
          })}
        </div>

        {/* Sub-menu circular icons beneath selected parent */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 pt-2 border-t border-slate-100">
          {currentGroupConfig.items.map((item) => {
            const Icon = item.icon;
            const active = item.id === activeReportId;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => onSelectReport(item.id)}
                className={`flex items-center sm:flex-col sm:justify-center gap-2.5 sm:gap-1.5 rounded-2xl border p-2.5 sm:p-3 text-left sm:text-center transition-all cursor-pointer active:scale-95 ${
                  active
                    ? 'border-[#E2136E] bg-rose-50/50 shadow-xs'
                    : 'border-slate-100 bg-white hover:border-rose-200 hover:bg-slate-50/70'
                }`}
              >
                <div
                  className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-full border ${item.bg} ${item.fg}`}
                >
                  <Icon className="h-4 w-4" />
                </div>
                <span className="text-xs font-bold text-slate-800 leading-snug">
                  {item.label}
                </span>
              </button>
            );
          })}
        </div>

        {/* Minimal Selection Controls: ONLY Month, Year, and Suballocation Memo */}
        {(showMonthYearRow || isAllocOrder) && (
          <div className="space-y-2.5 pt-2 border-t border-slate-100 text-xs">
            {showMonthYearRow && (
              <div className="grid grid-cols-2 gap-2.5">
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 mb-1">মাস</label>
                  <select
                    value={month0}
                    onChange={(e) => setMonth0(Number(e.target.value))}
                    className="w-full min-h-[42px] rounded-xl border border-slate-300 px-3 py-1.5 bg-white font-bold"
                  >
                    {BENGALI_MONTHS.map((m, idx) => (
                      <option key={m} value={idx}>
                        {m}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-600 mb-1">বছর</label>
                  <select
                    value={year}
                    onChange={(e) => setYear(Number(e.target.value))}
                    className="w-full min-h-[42px] rounded-xl border border-slate-300 px-3 py-1.5 bg-white font-bold"
                  >
                    {yearOptions.map((y) => (
                      <option key={y} value={y}>
                        {toBnDigits(y)}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            )}

            {isAllocOrder && (
              <div>
                {activeReportId === 'alloc-food' ? (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                    <div>
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">
                        উপ-বরাদ্দের স্মারক নম্বর
                      </label>
                      <select
                        value={selectedFoodMemo}
                        onChange={(e) => setSelectedFoodMemo(e.target.value)}
                        className="w-full min-h-[42px] rounded-xl border border-slate-300 px-3 py-1.5 bg-white font-bold"
                      >
                        {foodMemos.length === 0 ? (
                          <option value="">নির্বাচিত মাসে কোনো বরাদ্দ নেই</option>
                        ) : (
                          foodMemos.map((m) => (
                            <option key={m.memo} value={m.memo}>
                              {toBnDigits(m.memo)} ({formatBnDate(m.date)} · {formatBnQty(m.totalQty)} মে.টন · {m.grain})
                            </option>
                          ))
                        )}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">
                        প্রিন্ট ধরন
                      </label>
                      <div className="grid grid-cols-2 gap-2">
                        <button
                          type="button"
                          onClick={() => setFoodKernelMode('without_kernel')}
                          className={`min-h-[42px] rounded-xl border px-3 py-1.5 text-xs font-bold cursor-pointer transition-colors ${
                            foodKernelMode === 'without_kernel'
                              ? 'bg-[#E2136E] text-white border-[#E2136E] shadow-xs'
                              : 'bg-white text-slate-800 border-slate-300 hover:bg-rose-50'
                          }`}
                        >
                          কার্নেল ছাড়া
                        </button>
                        <button
                          type="button"
                          onClick={() => setFoodKernelMode('with_kernel')}
                          className={`min-h-[42px] rounded-xl border px-3 py-1.5 text-xs font-bold cursor-pointer transition-colors ${
                            foodKernelMode === 'with_kernel'
                              ? 'bg-emerald-700 text-white border-emerald-700 shadow-xs'
                              : 'bg-white text-slate-800 border-slate-300 hover:bg-emerald-50'
                          }`}
                        >
                          কার্নেল সহ (১%)
                        </button>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div>
                    <label className="block text-[11px] font-bold text-slate-600 mb-1">
                      উপ-বরাদ্দের স্মারক নম্বর
                    </label>
                    <select
                      value={selectedOmsMemo}
                      onChange={(e) => setSelectedOmsMemo(e.target.value)}
                      className="w-full sm:max-w-lg min-h-[42px] rounded-xl border border-slate-300 px-3 py-1.5 bg-white font-bold"
                    >
                      {omsMemos.length === 0 ? (
                        <option value="">নির্বাচিত মাসে কোনো উপবরাদ্দ নেই</option>
                      ) : (
                        omsMemos.map((m) => (
                          <option key={m.memo || m.date} value={m.memo}>
                            {m.memo ? toBnDigits(m.memo) : 'স্মারকহীন'} ({formatBnDate(m.date)} · {formatBnInt(m.count)} জন · {formatBnQty(m.totalQty)} মে.টন)
                          </option>
                        ))
                      )}
                    </select>
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </div>

      {activeReportId === 'monthly-dist' && pendingAllocRows.length > 0 && (
        <div className="no-print rounded-2xl border border-amber-200 bg-amber-50/90 px-4 py-3 text-xs text-amber-950 space-y-2">
          <div className="flex flex-wrap items-start justify-between gap-2">
            <div className="flex items-start gap-2">
              <AlertCircle className="h-4 w-4 text-amber-600 shrink-0 mt-0.5" />
              <div className="space-y-1">
                <div className="font-bold text-amber-900">
                  মাসিক বিতরণ রিপোর্টে কেবল &quot;ডিও ইস্যু&quot; হওয়া পরিমাণ হিসাব হয়
                </div>
                <p className="text-[11px] leading-relaxed text-amber-800">
                  নিচের খাতগুলোতে বরাদ্দ/উপবরাদ্দ এন্ট্রি করা আছে কিন্তু এখনো সম্পূর্ণ ডিও ইস্যু হয়নি, তাই ফরম ৩.১-এ বিতরণ ০ দেখাচ্ছে:
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={() => onSelectReport('alloc-summary')}
              className="shrink-0 rounded-xl bg-amber-600 hover:bg-amber-700 text-white px-3 py-1.5 text-[11px] font-bold cursor-pointer transition-colors"
            >
              বরাদ্দ ও বিতরণ স্থিতি দেখুন
            </button>
          </div>
          <div className="flex flex-wrap gap-1.5 pl-6">
            {pendingAllocRows.map((r) => {
              const allocVal = Math.max(r.ministryQty, r.dcQty ?? 0, r.subQty, r.projectQty);
              return (
                <span
                  key={r.category}
                  className="inline-flex items-center gap-1 rounded-lg border border-amber-200 bg-white px-2.5 py-1 text-[11px] font-semibold text-slate-800"
                >
                  <span className="font-bold text-[#E2136E]">{r.category}:</span>
                  <span>মোট বরাদ্দ {formatBnQty(allocVal)}</span>
                  <span className="text-slate-400">·</span>
                  <span>ডিও ইস্যু {formatBnQty(r.doQty)} মে.টন</span>
                </span>
              );
            })}
          </div>
        </div>
      )}

      {/* Rendered A4 Official Report + Single-Line Floating Export Bar */}
      <div className="relative">
        <OfficialA4Document
          ref={docRef}
          hideDefaultToolbar={true}
          template={template}
          officer={db.officer}
          subjectResolved={reportPayload.subjectResolved}
          introResolved={reportPayload.introResolved}
          conditionsResolved={reportPayload.conditionsResolved}
          copyMemoResolved={reportPayload.copyMemoResolved}
          copyListResolved={reportPayload.copyListResolved}
          copyDateResolved={reportPayload.copyDateResolved}
          hideTopMemoRow={reportPayload.hideTopMemoRow}
          memoNo={currentMemoNo}
          memoDate={currentMemoDate}
          onSaveMemo={(no, dt) => onSaveReportMemo(memoKey, no, dt)}
          columns={reportPayload.columns}
          leaves={reportPayload.leaves}
          mergeDepth={reportPayload.mergeDepth}
          showSerialCol={reportPayload.showSerialCol}
          serialColLabel={reportPayload.serialColLabel}
          colWidths={reportPayload.colWidths}
          tableLayoutFixed={reportPayload.tableLayoutFixed}
          footerLayout={reportPayload.footerLayout}
          rows={reportPayload.rows}
          totalRow={reportPayload.totalRow}
          secondarySection={reportPayload.secondarySection || null}
          kpiBoxes={reportPayload.kpiBoxes}
          bannerNote={reportPayload.bannerNote}
          showRecipientBlock={reportPayload.showRecipientBlock}
          showIntroBlock={reportPayload.showIntroBlock}
          showCopyBlock={reportPayload.showCopyBlock}
          showOfficerSignatureBlock={true}
          compactSinglePage={activeReportId === 'monthly-dist' && currentOrientation === 'portrait'}
          pdfFileName={`${activeReportId}-${new Date().toISOString().slice(0, 10)}.pdf`}
          onEditTemplate={() => onEditTemplate(activeReportId)}
        />

        {/* Lightly Floating Single-Line Symbol-Only Export Bar Over PDF Bottom */}
        <div className="no-print sticky bottom-[70px] lg:bottom-4 z-30 -mt-14 pt-2 pb-1.5 px-1 flex justify-center pointer-events-none">
          <div className="pointer-events-auto inline-flex items-center flex-nowrap gap-1.5 rounded-full border border-rose-200/80 bg-white/85 backdrop-blur-md shadow-[0_6px_22px_rgba(226,19,110,0.18)] px-2 py-1.5">
            <button
              type="button"
              onClick={() => onChangeOrientation(activeReportId, 'portrait')}
              title="পোর্ট্রেট মোড"
              aria-label="পোর্ট্রেট মোড"
              className={`h-9 w-10 sm:h-10 sm:w-11 shrink-0 inline-flex items-center justify-center rounded-full cursor-pointer transition-colors ${
                currentOrientation === 'portrait'
                  ? 'bg-[#E2136E] text-white shadow-xs'
                  : 'bg-slate-100/80 text-slate-700 hover:bg-slate-200/80'
              }`}
            >
              <Smartphone className="h-4 w-4 shrink-0" />
            </button>

            <button
              type="button"
              onClick={() => onChangeOrientation(activeReportId, 'landscape')}
              title="ল্যান্ডস্কেপ মোড"
              aria-label="ল্যান্ডস্কেপ মোড"
              className={`h-9 w-10 sm:h-10 sm:w-11 shrink-0 inline-flex items-center justify-center rounded-full cursor-pointer transition-colors ${
                currentOrientation === 'landscape'
                  ? 'bg-[#E2136E] text-white shadow-xs'
                  : 'bg-slate-100/80 text-slate-700 hover:bg-slate-200/80'
              }`}
            >
              <RectangleHorizontal className="h-4 w-4 shrink-0" />
            </button>

            <button
              type="button"
              onClick={() => onEditTemplate(activeReportId)}
              title="এডিট করুন"
              aria-label="এডিট করুন"
              className="h-9 w-10 sm:h-10 sm:w-11 shrink-0 inline-flex items-center justify-center rounded-full border border-slate-200/80 bg-white/85 hover:bg-rose-50 text-[#E2136E] cursor-pointer transition-colors"
            >
              <Edit3 className="h-4 w-4 shrink-0" />
            </button>

            <button
              type="button"
              onClick={() => docRef.current?.printDoc()}
              title="প্রিন্ট / Save as PDF"
              aria-label="প্রিন্ট / Save as PDF"
              className="h-9 w-10 sm:h-10 sm:w-11 shrink-0 inline-flex items-center justify-center rounded-full border border-slate-200/80 bg-white/85 hover:bg-rose-50 text-slate-800 cursor-pointer transition-colors"
            >
              <Printer className="h-4 w-4 shrink-0" />
            </button>

            <button
              type="button"
              onClick={() => void docRef.current?.exportPdf()}
              title="ডাউনলোড করুন"
              aria-label="ডাউনলোড করুন"
              className="h-9 w-10 sm:h-10 sm:w-11 shrink-0 inline-flex items-center justify-center rounded-full bg-[#E2136E] hover:bg-[#B80D58] text-white shadow-xs cursor-pointer transition-colors"
            >
              <Download className="h-4 w-4 shrink-0" />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};



