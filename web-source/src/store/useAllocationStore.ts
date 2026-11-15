import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { User } from 'firebase/auth';
import {
  CATEGORY_BY_NAME,
  CHALLAN_CATEGORIES,
  DEFAULT_OFFICER_PROFILE,
  createDefaultTemplate,
  createDefaultTemplatesMap,
  REPORT_TEMPLATE_CATALOG,
  type AllocationCategory,
  type AllocationDatabase,
  type AllocationSheetsConfig,
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
  type OfficerProfileSettings,
  type ProjectEntry,
  type ReportTemplateConfig,
  type ReportTemplateId,
  type UNOSuballocation,
} from '../types/allocation';
import type { DriveSpreadsheetFile, SheetWriteConfirmRequest } from '../types/grain';
import { createEmptyAllocationDatabase } from '../shared/allocSampleData';
import {
  ALLOC_SHEET_TABS,
  autoDetectAllocationSpreadsheets,
  countAllocationDatabaseRecords,
  fetchAllocationFromAppsScript,
  fetchAllocationFromOAuthSheet,
  fetchAllocationFromPublicSheet,
  fetchExternalChallansFromAppsScript,
  fetchExternalChallansFromSheet,
  reconcileChallansAndDOs,
  writeAllocationDatabaseToAppsScript,
  writeAllocationDatabaseToOAuthSheet,
} from '../shared/allocSheetsSync';
import { getUnusedDoNumbers } from '../shared/allocReportEngine';
import { triggerBlobDownload } from '../shared/downloadUtils';
import {
  bnToEnDigits,
  fiscalYearLabel,
  isChallanOnOrBeforeDo,
  isSamePartyName,
  memosLikelyMatch,
  namesFuzzyMatch,
  namesLikelyMatch,
  normMemoNo,
  parseMonthYear,
  round2,
  round3,
  toBnDigits,
  toIsoDateStr,
  uid,
} from '../shared/bnUtils';
import {
  cancelPendingAndroidSignIn,
  clearGoogleAccessToken,
  getGoogleAccessToken,
  googleLogout,
  googleSignIn,
  initGoogleAuth,
} from '../shared/googleAuth';
import { extractSpreadsheetId, GoogleAuthExpiredError } from '../shared/googleSheetsSync';

const STORAGE_DB_KEY = 'khaddo_boraddo_db_v1';
const STORAGE_SHEETS_KEY = 'khaddo_boraddo_sheets_v1';
const STORAGE_PENDING_PUSH_KEY = 'khaddo_boraddo_pending_push_v1';
const LEGACY_GRAIN_SHEETS_KEY = 'khaddo_stock_gsheets_config_v1';
const PORTRAIT_MIGRATION_KEY = 'khaddo_boraddo_portrait_migrated_v1';
const PURGE_SAMPLE_MIGRATION_KEY = 'khaddo_boraddo_purge_sample_v1';
const CLEAN_TITLES_MIGRATION_KEY = 'khaddo_boraddo_clean_titles_v1';
const MONTHLY_NO_COPY_MIGRATION_KEY = 'khaddo_boraddo_monthly_no_copy_v1';
const CHALLAN_SIGN_EQUAL_MIGRATION_KEY = 'khaddo_boraddo_challan_sign_equal_v1';

function getWriteTargetKey(spreadsheetId: string, gasUrl: string): string {
  const cleanId = extractSpreadsheetId(spreadsheetId);
  const cleanGas = String(gasUrl || '').trim();
  if (cleanId) return `sheet:${cleanId}`;
  if (cleanGas) return `gas:${cleanGas}`;
  return '';
}

function loadPendingPushFlag(): boolean {
  try {
    return typeof localStorage !== 'undefined' && localStorage.getItem(STORAGE_PENDING_PUSH_KEY) === '1';
  } catch {
    return false;
  }
}

function savePendingPushFlag(pending: boolean): void {
  try {
    if (typeof localStorage === 'undefined') return;
    if (pending) {
      localStorage.setItem(STORAGE_PENDING_PUSH_KEY, '1');
    } else {
      localStorage.removeItem(STORAGE_PENDING_PUSH_KEY);
    }
  } catch {
    // ignore
  }
}

const SAMPLE_ID_PREFIXES = new Set([
  'min-1', 'min-2', 'min-3', 'min-4', 'min-5',
  'dc-1', 'dc-2',
  'ds-1', 'ds-2', 'ds-3', 'ds-4', 'ds-5',
  'prj-1', 'prj-2',
  'uno-1', 'uno-2', 'uno-3',
  'book-1',
  'chln-1', 'chln-2', 'chln-3', 'chln-4',
  'do-1', 'do-2', 'do-3', 'do-4', 'do-5', 'do-6',
  'fd-1', 'fd-2', 'fd-3', 'fd-4',
  'od-1', 'od-2',
  'ch-1', 'ch-2', 'ch-3', 'ch-4',
  'fs-1',
  'rt-1', 'rt-2', 'rt-3', 'rt-4', 'rt-5', 'rt-6',
]);

const SAMPLE_REPORT_MEMO_KEYS = new Set([
  'allocFood|2026-11',
  'monthly-dist|2026-11',
  'monthly-cash|2026-11',
  'monthly-ep|2026-11',
]);

function stripSampleIds<T extends { id: string }>(arr: T[] | undefined): T[] {
  return (arr || []).filter((item) => !SAMPLE_ID_PREFIXES.has(item.id));
}

function loadInitialDatabase(): AllocationDatabase {
  try {
    const raw = localStorage.getItem(STORAGE_DB_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<AllocationDatabase>;
      const emptyBase = createEmptyAllocationDatabase();
      const officer: OfficerProfileSettings = {
        ...DEFAULT_OFFICER_PROFILE,
        ...(parsed.officer || {}),
      };
      const defaultTpls = createDefaultTemplatesMap(officer);
      const mergedTemplates = { ...defaultTpls };
      for (const item of REPORT_TEMPLATE_CATALOG) {
        if (parsed.templates?.[item.id]) {
          mergedTemplates[item.id] = {
            ...defaultTpls[item.id],
            ...parsed.templates[item.id],
          };
        }
      }

      if (localStorage.getItem(PORTRAIT_MIGRATION_KEY) !== '1') {
        mergedTemplates['monthly-dist'] = {
          ...mergedTemplates['monthly-dist'],
          orientation: 'portrait',
          fontSizePt: mergedTemplates['monthly-dist'].fontSizePt < 9 ? 9.5 : mergedTemplates['monthly-dist'].fontSizePt,
        };
        localStorage.setItem(PORTRAIT_MIGRATION_KEY, '1');
      }

      if (localStorage.getItem(CLEAN_TITLES_MIGRATION_KEY) !== '1') {
        const cleanMonthly = (id: 'monthly-dist' | 'monthly-cash' | 'monthly-ep', cleanLabel: string) => {
          mergedTemplates[id] = {
            ...mergedTemplates[id],
            label: cleanLabel,
            reportTitle: (mergedTemplates[id].reportTitle || cleanLabel).replace(/\s*\(৩\.[১২৩]\)\s*/g, '').trim() || cleanLabel,
          };
        };
        cleanMonthly('monthly-dist', 'মাসিক বিতরণ রিপোর্ট');
        cleanMonthly('monthly-cash', 'নগদ বিক্রয় রিপোর্ট');
        cleanMonthly('monthly-ep', 'ইপি রেশন রিপোর্ট');
        localStorage.setItem(CLEAN_TITLES_MIGRATION_KEY, '1');
      }

      if (localStorage.getItem(MONTHLY_NO_COPY_MIGRATION_KEY) !== '1') {
        (['monthly-dist', 'monthly-cash', 'monthly-ep'] as const).forEach((id) => {
          mergedTemplates[id] = {
            ...mergedTemplates[id],
            showCopyBlock: false,
          };
        });
        localStorage.setItem(MONTHLY_NO_COPY_MIGRATION_KEY, '1');
      }

      if (localStorage.getItem(CHALLAN_SIGN_EQUAL_MIGRATION_KEY) !== '1') {
        mergedTemplates['reg-challan'] = {
          ...mergedTemplates['reg-challan'],
          challanSignRule: 'excess_or_equal',
        };
        localStorage.setItem(CHALLAN_SIGN_EQUAL_MIGRATION_KEY, '1');
      }

      let dbObj: AllocationDatabase = {
        ...emptyBase,
        ...parsed,
        officer,
        templates: mergedTemplates,
      };

      if (localStorage.getItem(PURGE_SAMPLE_MIGRATION_KEY) !== '1') {
        const cleanedMemos: Record<string, { no: string; date: string }> = {};
        for (const [k, v] of Object.entries(dbObj.reportMemos || {})) {
          if (!SAMPLE_REPORT_MEMO_KEYS.has(k)) cleanedMemos[k] = v;
        }
        dbObj = {
          ...dbObj,
          ministryAllocations: stripSampleIds(dbObj.ministryAllocations),
          dcAllocations: stripSampleIds(dbObj.dcAllocations),
          dealerSuballocations: stripSampleIds(dbObj.dealerSuballocations),
          projects: stripSampleIds(dbObj.projects),
          unoSuballocations: stripSampleIds(dbObj.unoSuballocations),
          doBooks: stripSampleIds(dbObj.doBooks),
          doIssues: stripSampleIds(dbObj.doIssues),
          challans: stripSampleIds(dbObj.challans),
          foodDealers: stripSampleIds(dbObj.foodDealers),
          omsDealers: stripSampleIds(dbObj.omsDealers),
          chairmen: stripSampleIds(dbObj.chairmen),
          fireServices: stripSampleIds(dbObj.fireServices),
          rates: stripSampleIds(dbObj.rates),
          reportMemos: cleanedMemos,
        };
        localStorage.setItem(PURGE_SAMPLE_MIGRATION_KEY, '1');
      }

      return reconcileChallansAndDOs(dbObj);
    }
  } catch {
    // ignore storage parse failure
  }
  const initial = createEmptyAllocationDatabase();
  try {
    localStorage.setItem(PORTRAIT_MIGRATION_KEY, '1');
    localStorage.setItem(CLEAN_TITLES_MIGRATION_KEY, '1');
    localStorage.setItem(PURGE_SAMPLE_MIGRATION_KEY, '1');
    localStorage.setItem(MONTHLY_NO_COPY_MIGRATION_KEY, '1');
    localStorage.setItem(CHALLAN_SIGN_EQUAL_MIGRATION_KEY, '1');
    localStorage.setItem(STORAGE_DB_KEY, JSON.stringify(initial));
  } catch {
    // ignore quota error
  }
  return initial;
}

function loadInitialSheetsConfig(): AllocationSheetsConfig {
  const defaults: AllocationSheetsConfig = {
    enabled: true,
    spreadsheetId: '',
    spreadsheetName: '',
    challanSpreadsheetId: '',
    challanSpreadsheetName: '',
    challanAppsScriptUrl: '',
    appsScriptUrl: '',
    apiSecret: '',
    autoSync: true,
    lastSyncStatus: 'idle',
  };

  let parsed: Partial<AllocationSheetsConfig> | null = null;
  try {
    const raw = localStorage.getItem(STORAGE_SHEETS_KEY);
    if (raw) parsed = JSON.parse(raw) as Partial<AllocationSheetsConfig>;
  } catch {
    // ignore
  }

  let legacySpreadsheetId = '';
  let legacySpreadsheetName = '';
  let legacyChallanId = '';
  let legacyChallanName = '';
  let legacyAppsScriptUrl = '';
  try {
    const legacyRaw = localStorage.getItem(LEGACY_GRAIN_SHEETS_KEY);
    if (legacyRaw) {
      const legacy = JSON.parse(legacyRaw) as Record<string, unknown>;
      const extDoId = extractSpreadsheetId(String(legacy.externalDoSpreadsheetId || ''));
      const mainId = extractSpreadsheetId(String(legacy.spreadsheetId || ''));
      if (extDoId) {
        legacySpreadsheetId = extDoId;
        legacySpreadsheetName = String(legacy.externalDoSpreadsheetName || '');
      } else if (mainId) {
        legacySpreadsheetId = mainId;
        legacySpreadsheetName = String(legacy.spreadsheetName || '');
      }
      if (extDoId && mainId && extDoId !== mainId) {
        legacyChallanId = mainId;
        legacyChallanName = String(legacy.spreadsheetName || '');
      }
      legacyAppsScriptUrl = String(legacy.appsScriptUrl || '');
    }
  } catch {
    // ignore
  }

  return {
    ...defaults,
    ...(parsed || {}),
    spreadsheetId: parsed?.spreadsheetId || legacySpreadsheetId,
    spreadsheetName: parsed?.spreadsheetName || legacySpreadsheetName,
    challanSpreadsheetId: parsed?.challanSpreadsheetId || legacyChallanId,
    challanSpreadsheetName: parsed?.challanSpreadsheetName || legacyChallanName,
    appsScriptUrl: parsed?.appsScriptUrl || legacyAppsScriptUrl,
    autoSync: parsed?.autoSync ?? true,
  };
}

export interface ActionResult {
  ok: boolean;
  message?: string;
  error?: string;
}

export type SheetSyncUiState =
  | 'unlinked'
  | 'syncing'
  | 'synced'
  | 'pending_push'
  | 'needs_auth'
  | 'needs_write_confirm'
  | 'error';

export function useAllocationStore() {
  const [db, setDb] = useState<AllocationDatabase>(loadInitialDatabase);
  const [sheetsConfig, setSheetsConfig] = useState<AllocationSheetsConfig>(loadInitialSheetsConfig);
  const [googleUser, setGoogleUser] = useState<User | null>(null);
  const [googleAccessToken, setGoogleAccessToken] = useState<string | null>(() =>
    getGoogleAccessToken()
  );
  const [isGoogleSigningIn, setIsGoogleSigningIn] = useState(false);
  const [googleAuthError, setGoogleAuthError] = useState<string | null>(null);
  const [isSyncingSheets, setIsSyncingSheets] = useState(false);
  const [driveFiles, setDriveFiles] = useState<DriveSpreadsheetFile[]>([]);
  const [hasPendingSheetSync, setHasPendingSheetSync] = useState<boolean>(loadPendingPushFlag);
  const [writeConfirmState, setWriteConfirmState] = useState<{
    request: SheetWriteConfirmRequest;
    targetKey: string;
    overrideGasUrl?: string;
    overrideSheetId?: string;
    resolveResult?: (res: ActionResult) => void;
  } | null>(null);

  const dbRef = useRef(db);
  const sheetsConfigRef = useRef(sheetsConfig);
  const tokenRef = useRef(googleAccessToken);
  const initialAutoSyncDoneRef = useRef(false);
  const lastSyncedDbJsonRef = useRef<string>(
    loadPendingPushFlag() ? '' : JSON.stringify(db)
  );
  const autoPushTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const isPushingRef = useRef(false);
  const pendingAutoPushRef = useRef(false);
  const touchedMainChallansRef = useRef(false);
  const hadNonEmptyDbRef = useRef(countAllocationDatabaseRecords(db) > 0);

  useEffect(() => {
    dbRef.current = db;
    if (countAllocationDatabaseRecords(db) > 0) {
      hadNonEmptyDbRef.current = true;
    }
  }, [db]);

  useEffect(() => {
    sheetsConfigRef.current = sheetsConfig;
  }, [sheetsConfig]);

  useEffect(() => {
    tokenRef.current = googleAccessToken;
  }, [googleAccessToken]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_DB_KEY, JSON.stringify(db));
    } catch {
      // ignore storage quota errors
    }
  }, [db]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_SHEETS_KEY, JSON.stringify(sheetsConfig));
    } catch {
      // ignore storage quota errors
    }
  }, [sheetsConfig]);

  useEffect(() => {
    const unsub = initGoogleAuth(
      (user, token) => {
        setGoogleUser(user);
        setGoogleAccessToken(token);
      },
      () => {
        setGoogleUser(null);
        setGoogleAccessToken(null);
      }
    );
    return () => unsub();
  }, []);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    if (localStorage.getItem('khaddo_boraddo_seed_loaded') === '1') return;
    fetch('/web-seed.json')
      .then((r) => (r.ok ? r.json() : null))
      .then((seed) => {
        if (!seed || typeof seed !== 'object') return;
        if (seed.db && typeof seed.db === 'object') {
          setDb((prev) => {
            const merged = { ...prev, ...seed.db };
            lastSyncedDbJsonRef.current = JSON.stringify(merged);
            return merged;
          });
        }
        if (seed.sheetsConfig && typeof seed.sheetsConfig === 'object') {
          setSheetsConfig((prev) => ({ ...prev, ...seed.sheetsConfig }));
        }
        localStorage.setItem('khaddo_boraddo_seed_loaded', '1');
      })
      .catch(() => {
        // optional
      });
  }, []);

  const usageIndex = useMemo(() => {
    const minById = new Map<string, { total: number; used: number; remaining: number }>();
    const dcById = new Map<string, { total: number; used: number; remaining: number }>();

    const fyKeyOf = (dateStr?: string): string => {
      const parsed = dateStr ? parseMonthYear(dateStr) : null;
      return parsed ? fiscalYearLabel(parsed.year, parsed.month) : '';
    };

    const allocatePool = <T extends { id: string; allocationQty: number }>(
      upstreams: Array<{ rec: T; date: string }>,
      downstreams: Array<{ memoKey: string; qty: number; date: string }>,
      getUpstreamMemo: (r: T) => string,
      outMap: Map<string, { total: number; used: number; remaining: number }>
    ) => {
      const directUsed = new Map<string, number>();
      const unlinked: Array<{ qty: number; date: string }> = [];

      for (const item of downstreams) {
        if (item.qty <= 0) continue;
        const key = (item.memoKey || '').trim();
        let matched: T | undefined;
        if (key) {
          const nk = normMemoNo(key);
          matched =
            upstreams.find((u) => normMemoNo(getUpstreamMemo(u.rec)) === nk)?.rec ||
            upstreams.find((u) => memosLikelyMatch(getUpstreamMemo(u.rec), key))?.rec;
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

      for (const u of upstreams) {
        const total = round3(Number(u.rec.allocationQty) || 0);
        const used = round3((directUsed.get(u.rec.id) || 0) + (fifoUsed.get(u.rec.id) || 0));
        outMap.set(u.rec.id, {
          total,
          used,
          remaining: round3(total - used),
        });
      }
    };

    const categories = new Set<AllocationCategory>([
      ...db.ministryAllocations.map((m) => m.category),
      ...db.dcAllocations.map((d) => d.category),
    ]);

    for (const category of categories) {
      const meta = CATEGORY_BY_NAME[category];
      if (!meta) continue;

      const mins = db.ministryAllocations
        .filter((m) => m.category === category)
        .map((rec) => ({ rec, date: rec.memoDate || rec.createdAt }));

      const minDownstreams = meta.requiresDC
        ? db.dcAllocations
            .filter((d) => d.category === category)
            .map((d) => ({
              memoKey: d.ministryMemo,
              qty: Number(d.allocationQty) || 0,
              date: d.dcMemoDate || d.createdAt,
            }))
        : meta.flowType === 'vwb'
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

      allocatePool(mins, minDownstreams, (m) => m.memoNo, minById);

      const dcs = db.dcAllocations
        .filter((d) => d.category === category)
        .map((rec) => ({ rec, date: rec.dcMemoDate || rec.createdAt }));

      const dcDownstreams =
        meta.flowType === 'dealer_oms'
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

      allocatePool(dcs, dcDownstreams, (d) => d.dcMemo, dcById);
    }

    return { minById, dcById };
  }, [db.ministryAllocations, db.dcAllocations, db.unoSuballocations, db.dealerSuballocations]);

  const getMinistryUsage = useCallback(
    (category: AllocationCategory, memoNo: string) => {
      const target = (memoNo || '').trim();
      const normTarget = normMemoNo(target);
      const minRec =
        db.ministryAllocations.find((m) => m.category === category && m.memoNo === memoNo) ||
        (normTarget
          ? db.ministryAllocations.find(
              (m) => m.category === category && normMemoNo(m.memoNo) === normTarget
            ) ||
            db.ministryAllocations.find(
              (m) => m.category === category && memosLikelyMatch(m.memoNo, target)
            )
          : undefined);
      if (!minRec) return { total: 0, used: 0, remaining: 0 };
      return usageIndex.minById.get(minRec.id) || {
        total: round3(Number(minRec.allocationQty) || 0),
        used: 0,
        remaining: round3(Number(minRec.allocationQty) || 0),
      };
    },
    [db.ministryAllocations, usageIndex]
  );

  const getDCUsage = useCallback(
    (category: AllocationCategory, dcMemo: string) => {
      const target = (dcMemo || '').trim();
      const normTarget = normMemoNo(target);
      const dcRec =
        db.dcAllocations.find((d) => d.category === category && d.dcMemo === dcMemo) ||
        (normTarget
          ? db.dcAllocations.find(
              (d) => d.category === category && normMemoNo(d.dcMemo) === normTarget
            ) ||
            db.dcAllocations.find(
              (d) => d.category === category && memosLikelyMatch(d.dcMemo, target)
            )
          : undefined);
      if (!dcRec) return { total: 0, used: 0, remaining: 0 };
      return usageIndex.dcById.get(dcRec.id) || {
        total: round3(Number(dcRec.allocationQty) || 0),
        used: 0,
        remaining: round3(Number(dcRec.allocationQty) || 0),
      };
    },
    [db.dcAllocations, usageIndex]
  );

  const addMinistryAllocation = useCallback(
    (input: Omit<MinistryAllocation, 'id' | 'createdAt'>): ActionResult => {
      const memo = input.memoNo.trim();
      if (!memo) return { ok: false, error: 'মন্ত্রণালয়ের স্মারক নং দিন।' };
      if ((Number(input.allocationQty) || 0) <= 0) {
        return { ok: false, error: 'বরাদ্দের পরিমাণ ০-এর বেশি হতে হবে।' };
      }
      const dup = db.ministryAllocations.some(
        (m) => m.category === input.category && m.memoNo.trim() === memo
      );
      if (dup) {
        return { ok: false, error: `স্মারক নং "${memo}" এই খাতে আগেই এন্ট্রি আছে।` };
      }
      const row: MinistryAllocation = {
        ...input,
        memoNo: memo,
        allocationQty: round3(input.allocationQty),
        id: uid('min'),
        createdAt: new Date().toISOString(),
      };
      setDb((prev) => ({
        ...prev,
        ministryAllocations: [row, ...prev.ministryAllocations],
      }));
      return { ok: true, message: 'মন্ত্রণালয় বরাদ্দ সংরক্ষিত হয়েছে।' };
    },
    [db.ministryAllocations]
  );

  const updateMinistryAllocation = useCallback(
    (
      id: string,
      patch: Partial<Omit<MinistryAllocation, 'id' | 'createdAt' | 'category'>>
    ): ActionResult => {
      const existing = db.ministryAllocations.find((m) => m.id === id);
      if (!existing) return { ok: false, error: 'মন্ত্রণালয় বরাদ্দ খুঁজে পাওয়া যায়নি।' };
      const nextMemo = (patch.memoNo ?? existing.memoNo).trim();
      const nextQty = round3(patch.allocationQty ?? existing.allocationQty);
      if (!nextMemo) return { ok: false, error: 'মন্ত্রণালয়ের স্মারক নং দিন।' };
      if (nextQty <= 0) return { ok: false, error: 'বরাদ্দের পরিমাণ ০-এর বেশি হতে হবে।' };
      if (
        nextMemo !== existing.memoNo &&
        db.ministryAllocations.some(
          (m) => m.id !== id && m.category === existing.category && m.memoNo.trim() === nextMemo
        )
      ) {
        return { ok: false, error: `স্মারক নং "${nextMemo}" এই খাতে আগেই আছে।` };
      }
      const usage = getMinistryUsage(existing.category, existing.memoNo);
      if (nextQty + 0.0005 < usage.used) {
        return {
          ok: false,
          error: `ইতিমধ্যে ${toBnDigits(usage.used.toFixed(3))} মে.টন ব্যবহৃত হয়েছে; এর চেয়ে কম পরিমাণ দেওয়া যাবে না।`,
        };
      }
      const oldMemo = existing.memoNo;
      const memoChanged = nextMemo !== oldMemo;
      const meta = CATEGORY_BY_NAME[existing.category];
      setDb((prev) => ({
        ...prev,
        ministryAllocations: prev.ministryAllocations.map((m) =>
          m.id === id
            ? {
                ...m,
                ...patch,
                memoNo: nextMemo,
                allocationQty: nextQty,
              }
            : m
        ),
        dcAllocations: memoChanged
          ? prev.dcAllocations.map((d) =>
              d.category === existing.category && memosLikelyMatch(d.ministryMemo, oldMemo)
                ? { ...d, ministryMemo: nextMemo }
                : d
            )
          : prev.dcAllocations,
        unoSuballocations: memoChanged
          ? prev.unoSuballocations.map((u) =>
              u.category === existing.category && memosLikelyMatch(u.ministryMemo, oldMemo)
                ? { ...u, ministryMemo: nextMemo }
                : u
            )
          : prev.unoSuballocations,
        dealerSuballocations:
          memoChanged && !meta?.requiresDC
            ? prev.dealerSuballocations.map((s) =>
                s.category === existing.category && memosLikelyMatch(s.sourceMemo, oldMemo)
                  ? { ...s, sourceMemo: nextMemo }
                  : s
              )
            : prev.dealerSuballocations,
        doIssues: memoChanged
          ? prev.doIssues.map((d) =>
              d.category === existing.category && memosLikelyMatch(d.ministryMemo, oldMemo)
                ? { ...d, ministryMemo: nextMemo }
                : d
            )
          : prev.doIssues,
      }));
      return { ok: true, message: 'মন্ত্রণালয় বরাদ্দ হালনাগাদ হয়েছে।' };
    },
    [db.ministryAllocations, getMinistryUsage]
  );

  const deleteMinistryAllocation = useCallback((id: string) => {
    setDb((prev) => ({
      ...prev,
      ministryAllocations: prev.ministryAllocations.filter((m) => m.id !== id),
    }));
  }, []);

  const addDCAllocation = useCallback(
    (
      input: Omit<DCAllocation, 'id' | 'createdAt'>,
      attachedProjects?: Array<Omit<ProjectEntry, 'id' | 'category' | 'dcMemo' | 'createdAt'>>
    ): ActionResult => {
      const mMemo = input.ministryMemo.trim();
      const dcMemo = input.dcMemo.trim();
      const qty = round3(input.allocationQty);
      if (!mMemo || !dcMemo || !input.dcMemoDate || qty <= 0) {
        return { ok: false, error: 'মন্ত্রণালয় স্মারক, ডিসি স্মারক, তারিখ ও পরিমাণ সবগুলো পূরণ করুন।' };
      }
      const dup = db.dcAllocations.some(
        (d) => d.category === input.category && d.dcMemo.trim() === dcMemo
      );
      if (dup) {
        return { ok: false, error: `ডিসি স্মারক "${dcMemo}" এই খাতে আগেই আছে।` };
      }
      const usage = getMinistryUsage(input.category, mMemo);
      if (qty > usage.remaining + 0.0005) {
        return {
          ok: false,
          error: `মন্ত্রণালয় অবশিষ্ট ব্যালেন্স (${toBnDigits(usage.remaining.toFixed(3))} মে.টন) এর চেয়ে বেশি ডিসি বরাদ্দ দেওয়া যাবে না।`,
        };
      }
      const now = new Date().toISOString();
      const dcRow: DCAllocation = {
        ...input,
        ministryMemo: mMemo,
        dcMemo,
        allocationQty: qty,
        id: uid('dc'),
        createdAt: now,
      };
      const newProjects: ProjectEntry[] = (attachedProjects || []).map((p) => ({
        ...p,
        id: uid('prj'),
        category: input.category,
        dcMemo,
        allocationQty: round3(p.allocationQty),
        createdAt: now,
      }));
      setDb((prev) => ({
        ...prev,
        dcAllocations: [dcRow, ...prev.dcAllocations],
        projects: [...newProjects, ...prev.projects],
      }));
      return {
        ok: true,
        message:
          newProjects.length > 0
            ? `ডিসি বরাদ্দ ও ${toBnDigits(newProjects.length)}টি প্রকল্প সংরক্ষিত হয়েছে।`
            : 'ডিসি বরাদ্দ সংরক্ষিত হয়েছে।',
      };
    },
    [db.dcAllocations, getMinistryUsage]
  );

  const updateDCAllocation = useCallback(
    (
      id: string,
      patch: Partial<Omit<DCAllocation, 'id' | 'createdAt' | 'category'>>
    ): ActionResult => {
      const existing = db.dcAllocations.find((d) => d.id === id);
      if (!existing) return { ok: false, error: 'ডিসি বরাদ্দ খুঁজে পাওয়া যায়নি।' };
      const nextMinMemo = (patch.ministryMemo ?? existing.ministryMemo).trim();
      const nextDcMemo = (patch.dcMemo ?? existing.dcMemo).trim();
      const nextDate = (patch.dcMemoDate ?? existing.dcMemoDate).trim();
      const nextQty = round3(patch.allocationQty ?? existing.allocationQty);
      if (!nextMinMemo || !nextDcMemo || !nextDate || nextQty <= 0) {
        return { ok: false, error: 'মন্ত্রণালয় স্মারক, ডিসি স্মারক, তারিখ ও পরিমাণ সবগুলো পূরণ করুন।' };
      }
      if (
        nextDcMemo !== existing.dcMemo &&
        db.dcAllocations.some(
          (d) => d.id !== id && d.category === existing.category && d.dcMemo.trim() === nextDcMemo
        )
      ) {
        return { ok: false, error: `ডিসি স্মারক "${nextDcMemo}" এই খাতে আগেই আছে।` };
      }
      const minUsage = getMinistryUsage(existing.category, nextMinMemo);
      const effectiveRemaining = memosLikelyMatch(nextMinMemo, existing.ministryMemo)
        ? round3(minUsage.remaining + existing.allocationQty)
        : minUsage.remaining;
      if (nextQty > effectiveRemaining + 0.0005) {
        return {
          ok: false,
          error: `মন্ত্রণালয় অবশিষ্ট ব্যালেন্স (${toBnDigits(effectiveRemaining.toFixed(3))} মে.টন) এর চেয়ে বেশি ডিসি বরাদ্দ দেওয়া যাবে না।`,
        };
      }
      const dcUsage = getDCUsage(existing.category, existing.dcMemo);
      if (nextQty + 0.0005 < dcUsage.used) {
        return {
          ok: false,
          error: `ইতিমধ্যে ${toBnDigits(dcUsage.used.toFixed(3))} মে.টন উপবরাদ্দ হয়েছে; এর চেয়ে কম পরিমাণ দেওয়া যাবে না।`,
        };
      }
      const oldDcMemo = existing.dcMemo;
      const dcChanged = nextDcMemo !== oldDcMemo;
      const meta = CATEGORY_BY_NAME[existing.category];
      setDb((prev) => ({
        ...prev,
        dcAllocations: prev.dcAllocations.map((d) =>
          d.id === id
            ? {
                ...d,
                ministryMemo: nextMinMemo,
                dcMemo: nextDcMemo,
                dcMemoDate: nextDate,
                allocationQty: nextQty,
              }
            : d
        ),
        projects: dcChanged
          ? prev.projects.map((p) =>
              p.category === existing.category && memosLikelyMatch(p.dcMemo, oldDcMemo)
                ? { ...p, dcMemo: nextDcMemo }
                : p
            )
          : prev.projects,
        unoSuballocations: dcChanged
          ? prev.unoSuballocations.map((u) =>
              u.category === existing.category && memosLikelyMatch(u.dcMemo, oldDcMemo)
                ? { ...u, dcMemo: nextDcMemo, ministryMemo: nextMinMemo }
                : u
            )
          : prev.unoSuballocations,
        dealerSuballocations:
          dcChanged && meta?.requiresDC
            ? prev.dealerSuballocations.map((s) =>
                s.category === existing.category && memosLikelyMatch(s.sourceMemo, oldDcMemo)
                  ? { ...s, sourceMemo: nextDcMemo }
                  : s
              )
            : prev.dealerSuballocations,
        doIssues: dcChanged
          ? prev.doIssues.map((d) =>
              d.category === existing.category && memosLikelyMatch(d.dcMemo, oldDcMemo)
                ? { ...d, dcMemo: nextDcMemo, ministryMemo: nextMinMemo }
                : d
            )
          : prev.doIssues,
      }));
      return { ok: true, message: 'ডিসি বরাদ্দ হালনাগাদ হয়েছে।' };
    },
    [db.dcAllocations, getDCUsage, getMinistryUsage]
  );

  const deleteDCAllocation = useCallback((id: string) => {
    setDb((prev) => {
      const target = prev.dcAllocations.find((d) => d.id === id);
      return {
        ...prev,
        dcAllocations: prev.dcAllocations.filter((d) => d.id !== id),
        projects: target
          ? prev.projects.filter((p) => !(p.category === target.category && p.dcMemo === target.dcMemo))
          : prev.projects,
      };
    });
  }, []);

  const addProjectsToDC = useCallback(
    (
      category: AllocationCategory,
      dcMemo: string,
      entries: Array<Omit<ProjectEntry, 'id' | 'category' | 'dcMemo' | 'createdAt'>>
    ): ActionResult => {
      if (!dcMemo) return { ok: false, error: 'ডিসি স্মারক নির্বাচন করুন।' };
      if (!entries.length) return { ok: false, error: 'কমপক্ষে ১টি প্রকল্প যোগ করুন।' };
      const now = new Date().toISOString();
      const rows: ProjectEntry[] = entries.map((p) => ({
        ...p,
        id: uid('prj'),
        category,
        dcMemo,
        allocationQty: round3(p.allocationQty),
        createdAt: now,
      }));
      setDb((prev) => ({
        ...prev,
        projects: [...rows, ...prev.projects],
      }));
      return { ok: true, message: `${toBnDigits(rows.length)}টি প্রকল্প যোগ হয়েছে।` };
    },
    []
  );

  const updateProject = useCallback(
    (
      id: string,
      patch: Partial<
        Pick<
          ProjectEntry,
          'projectNo' | 'projectName' | 'presidentName' | 'unionName' | 'allocationQty'
        >
      >
    ): ActionResult => {
      const existing = db.projects.find((p) => p.id === id);
      if (!existing) return { ok: false, error: 'প্রকল্প খুঁজে পাওয়া যায়নি।' };
      const nextQty = round3(patch.allocationQty ?? existing.allocationQty);
      if (nextQty <= 0) return { ok: false, error: 'প্রকল্পের পরিমাণ ০-এর বেশি হতে হবে।' };
      const nextNo = (patch.projectNo ?? existing.projectNo).trim() || existing.projectNo;
      const nextName = (patch.projectName ?? existing.projectName).trim();
      const nextPres = (patch.presidentName ?? existing.presidentName).trim();
      const nextUnion = (patch.unionName ?? existing.unionName).trim() || existing.unionName;
      const matchesOldProj = (cat: AllocationCategory, dcMemo: string, projNo?: string) =>
        cat === existing.category &&
        memosLikelyMatch(dcMemo, existing.dcMemo) &&
        normMemoNo(projNo || '') === normMemoNo(existing.projectNo);

      setDb((prev) => ({
        ...prev,
        projects: prev.projects.map((p) =>
          p.id === id
            ? {
                ...p,
                projectNo: nextNo,
                projectName: nextName,
                presidentName: nextPres,
                unionName: nextUnion,
                allocationQty: nextQty,
              }
            : p
        ),
        unoSuballocations: prev.unoSuballocations.map((u) =>
          matchesOldProj(u.category, u.dcMemo, u.projectNo)
            ? {
                ...u,
                projectNo: nextNo,
                projectName: nextName || u.projectName,
              }
            : u
        ),
        doIssues: prev.doIssues.map((d) =>
          matchesOldProj(d.category, d.dcMemo, d.projectNo)
            ? {
                ...d,
                projectNo: nextNo,
              }
            : d
        ),
      }));
      return { ok: true, message: 'প্রকল্পের তথ্য হালনাগাদ হয়েছে।' };
    },
    [db.projects]
  );

  const deleteProject = useCallback((id: string) => {
    setDb((prev) => ({
      ...prev,
      projects: prev.projects.filter((p) => p.id !== id),
    }));
  }, []);

  const transferProject = useCallback(
    (
      id: string,
      target: { category?: AllocationCategory; dcMemo: string }
    ): ActionResult => {
      const existing = db.projects.find((p) => p.id === id);
      if (!existing) return { ok: false, error: 'প্রকল্প খুঁজে পাওয়া যায়নি।' };
      const targetCat = target.category || existing.category;
      const targetMeta = CATEGORY_BY_NAME[targetCat];
      if (!targetMeta?.supportsProjects) {
        return { ok: false, error: 'নির্বাচিত খাতে প্রকল্প স্থানান্তরযোগ্য নয়।' };
      }
      const nextDcMemo = target.dcMemo.trim();
      if (!nextDcMemo) {
        return { ok: false, error: 'লক্ষ্য ডিসি স্মারক নির্বাচন করুন।' };
      }
      const hasDc = db.dcAllocations.some(
        (d) =>
          d.category === targetCat &&
          (normMemoNo(d.dcMemo) === normMemoNo(nextDcMemo) || memosLikelyMatch(d.dcMemo, nextDcMemo))
      );
      if (!hasDc) {
        return { ok: false, error: `${targetCat} খাতে "${toBnDigits(nextDcMemo)}" ডিসি স্মারক পাওয়া যায়নি।` };
      }
      setDb((prev) => ({
        ...prev,
        projects: prev.projects.map((p) =>
          p.id === id ? { ...p, category: targetCat, dcMemo: nextDcMemo } : p
        ),
      }));
      return {
        ok: true,
        message: `প্রকল্প ${toBnDigits(existing.projectNo)} সফলভাবে "${targetCat} (${toBnDigits(nextDcMemo)})"-এ স্থানান্তরিত হয়েছে।`,
      };
    },
    [db.dcAllocations, db.projects]
  );

  const saveDealerSuballocationBatch = useCallback(
    (payload: {
      category: AllocationCategory;
      sourceMemo: string;
      unoMemo: string;
      unoDate: string;
      rows: Array<{
        dealerName: string;
        unionName: string;
        salesCenter: string;
        cardCount: number;
        unusedCards: number;
        allocationQty: number;
        totalAmount: number;
        taxAmount: number;
      }>;
    }): ActionResult => {
      const { category, sourceMemo, unoMemo, unoDate, rows } = payload;
      if (!sourceMemo || !unoDate) {
        return { ok: false, error: 'উৎস স্মারক ও তারিখ দিন।' };
      }
      const validRows = rows.filter((r) => Number(r.allocationQty) > 0);
      if (!validRows.length) {
        return { ok: false, error: 'কোনো প্রাপক বা পরিমাণ ০-এর বেশি নয়।' };
      }
      const totalRequested = round3(validRows.reduce((s, r) => s + Number(r.allocationQty), 0));
      const meta = CATEGORY_BY_NAME[category];
      const upstream = meta.requiresDC
        ? getDCUsage(category, sourceMemo)
        : getMinistryUsage(category, sourceMemo);

      if (totalRequested > upstream.remaining + 0.0005) {
        return {
          ok: false,
          error: `মোট চাহিদা (${toBnDigits(totalRequested.toFixed(3))} মে.টন) অবশিষ্ট বরাদ্দের (${toBnDigits(upstream.remaining.toFixed(3))} মে.টন) চেয়ে বেশি!`,
        };
      }

      const now = new Date().toISOString();
      const created: DealerSuballocation[] = validRows.map((r) => ({
        id: uid('ds'),
        category,
        sourceMemo,
        unoMemo: unoMemo.trim() || sourceMemo,
        unoDate,
        dealerName: r.dealerName,
        unionName: r.unionName,
        salesCenter: r.salesCenter,
        cardCount: Math.round(r.cardCount || 0),
        unusedCards: Math.round(r.unusedCards || 0),
        allocationQty: round3(r.allocationQty),
        totalAmount: round2(r.totalAmount),
        taxAmount: round2(r.taxAmount),
        createdAt: now,
      }));
      setDb((prev) => ({
        ...prev,
        dealerSuballocations: [...created, ...prev.dealerSuballocations],
      }));
      return {
        ok: true,
        message: `${toBnDigits(created.length)} জন প্রাপকের উপবরাদ্দ (${toBnDigits(totalRequested.toFixed(3))} মে.টন) সংরক্ষিত হয়েছে।`,
      };
    },
    [getDCUsage, getMinistryUsage]
  );

  const updateDealerSuballocation = useCallback(
    (
      id: string,
      patch: Partial<
        Pick<
          DealerSuballocation,
          'unoMemo' | 'unoDate' | 'dealerName' | 'unionName' | 'allocationQty'
        >
      >
    ): ActionResult => {
      const existing = db.dealerSuballocations.find((r) => r.id === id);
      if (!existing) return { ok: false, error: 'উপবরাদ্দ রেকর্ড পাওয়া যায়নি।' };
      const nextQty = round3(patch.allocationQty ?? existing.allocationQty);
      if (nextQty <= 0) return { ok: false, error: 'বরাদ্দ পরিমাণ ০-এর বেশি হতে হবে।' };
      const meta = CATEGORY_BY_NAME[existing.category];
      const upstream = meta.requiresDC
        ? getDCUsage(existing.category, existing.sourceMemo)
        : getMinistryUsage(existing.category, existing.sourceMemo);
      const effectiveRemaining = round3(upstream.remaining + existing.allocationQty);
      if (nextQty > effectiveRemaining + 0.0005) {
        return {
          ok: false,
          error: `অবশিষ্ট বরাদ্দ (${toBnDigits(effectiveRemaining.toFixed(3))} মে.টন) এর চেয়ে বেশি পরিমাণ দেওয়া যাবে না।`,
        };
      }
      const rate = db.rates.find((r) => r.category === existing.category);
      const nextTotal =
        rate && rate.unitPrice > 0
          ? round2(nextQty * rate.unitPrice)
          : existing.allocationQty > 0
            ? round2((existing.totalAmount / existing.allocationQty) * nextQty)
            : existing.totalAmount;
      const nextTax =
        rate && rate.taxRate > 0
          ? round2(nextQty * rate.taxRate)
          : existing.allocationQty > 0
            ? round2((existing.taxAmount / existing.allocationQty) * nextQty)
            : existing.taxAmount;
      setDb((prev) => ({
        ...prev,
        dealerSuballocations: prev.dealerSuballocations.map((r) =>
          r.id === id
            ? {
                ...r,
                unoMemo: (patch.unoMemo ?? r.unoMemo).trim() || r.sourceMemo,
                unoDate: (patch.unoDate ?? r.unoDate).trim() || r.unoDate,
                dealerName: (patch.dealerName ?? r.dealerName).trim() || r.dealerName,
                unionName: (patch.unionName ?? r.unionName).trim() || r.unionName,
                allocationQty: nextQty,
                totalAmount: nextTotal,
                taxAmount: nextTax,
              }
            : r
        ),
      }));
      return { ok: true, message: 'উপবরাদ্দ হালনাগাদ হয়েছে।' };
    },
    [db.dealerSuballocations, db.rates, getDCUsage, getMinistryUsage]
  );

  const deleteDealerSuballocation = useCallback((id: string) => {
    setDb((prev) => ({
      ...prev,
      dealerSuballocations: prev.dealerSuballocations.filter((r) => r.id !== id),
    }));
  }, []);

  const saveUNOSuballocationBatch = useCallback(
    (payload: {
      category: AllocationCategory;
      sourceMemo: string;
      unoMemo: string;
      unoDate: string;
      rows: Array<{
        unionName: string;
        recipientName: string;
        projectNo?: string;
        projectName?: string;
        allocationQty: number;
      }>;
    }): ActionResult => {
      const { category, sourceMemo, unoMemo, unoDate, rows } = payload;
      if (!sourceMemo || !unoMemo.trim() || !unoDate) {
        return { ok: false, error: 'স্মারক নং, ইউএনও স্মারক ও তারিখ সবগুলো দিন।' };
      }
      const validRows = rows.filter((r) => Number(r.allocationQty) > 0);
      if (!validRows.length) {
        return { ok: false, error: 'কোনো বরাদ্দ পরিমাণ ০-এর বেশি নয়।' };
      }
      const totalReq = round3(validRows.reduce((s, r) => s + Number(r.allocationQty), 0));
      const meta = CATEGORY_BY_NAME[category];
      const upstream = meta.requiresDC
        ? getDCUsage(category, sourceMemo)
        : getMinistryUsage(category, sourceMemo);

      if (totalReq > upstream.remaining + 0.0005) {
        return {
          ok: false,
          error: `মোট উপবরাদ্দ (${toBnDigits(totalReq.toFixed(3))} মে.টন) অবশিষ্ট বরাদ্দের (${toBnDigits(upstream.remaining.toFixed(3))} মে.টন) চেয়ে বেশি!`,
        };
      }

      const dcRec = meta.requiresDC
        ? db.dcAllocations.find(
            (d) => d.category === category && memosLikelyMatch(d.dcMemo, sourceMemo)
          )
        : undefined;
      const now = new Date().toISOString();
      const created: UNOSuballocation[] = validRows.map((r) => ({
        id: uid('uno'),
        category,
        ministryMemo: meta.requiresDC ? dcRec?.ministryMemo || '' : sourceMemo,
        dcMemo: meta.requiresDC ? sourceMemo : '',
        unoMemo: unoMemo.trim(),
        unoDate,
        unionName: r.unionName,
        recipientName: r.recipientName,
        projectNo: r.projectNo,
        projectName: r.projectName,
        allocationQty: round3(r.allocationQty),
        createdAt: now,
      }));
      setDb((prev) => ({
        ...prev,
        unoSuballocations: [...created, ...prev.unoSuballocations],
      }));
      return {
        ok: true,
        message: `${toBnDigits(created.length)} টি ইউএনও উপবরাদ্দ (${toBnDigits(totalReq.toFixed(3))} মে.টন) সংরক্ষিত হয়েছে।`,
      };
    },
    [db.dcAllocations, getDCUsage, getMinistryUsage]
  );

  const updateUNOSuballocation = useCallback(
    (
      id: string,
      patch: Partial<
        Pick<
          UNOSuballocation,
          'unoMemo' | 'unoDate' | 'unionName' | 'recipientName' | 'allocationQty'
        >
      >
    ): ActionResult => {
      const existing = db.unoSuballocations.find((u) => u.id === id);
      if (!existing) return { ok: false, error: 'উপবরাদ্দ রেকর্ড পাওয়া যায়নি।' };
      const nextQty = round3(patch.allocationQty ?? existing.allocationQty);
      if (nextQty <= 0) return { ok: false, error: 'বরাদ্দ পরিমাণ ০-এর বেশি হতে হবে।' };
      const meta = CATEGORY_BY_NAME[existing.category];
      const sourceMemo = meta.requiresDC ? existing.dcMemo : existing.ministryMemo;
      const upstream = meta.requiresDC
        ? getDCUsage(existing.category, sourceMemo)
        : getMinistryUsage(existing.category, sourceMemo);
      const effectiveRemaining = round3(upstream.remaining + existing.allocationQty);
      if (nextQty > effectiveRemaining + 0.0005) {
        return {
          ok: false,
          error: `অবশিষ্ট বরাদ্দ (${toBnDigits(effectiveRemaining.toFixed(3))} মে.টন) এর চেয়ে বেশি পরিমাণ দেওয়া যাবে না।`,
        };
      }
      setDb((prev) => ({
        ...prev,
        unoSuballocations: prev.unoSuballocations.map((u) =>
          u.id === id
            ? {
                ...u,
                unoMemo: (patch.unoMemo ?? u.unoMemo).trim() || u.unoMemo,
                unoDate: (patch.unoDate ?? u.unoDate).trim() || u.unoDate,
                unionName: (patch.unionName ?? u.unionName).trim() || u.unionName,
                recipientName: (patch.recipientName ?? u.recipientName).trim() || u.recipientName,
                allocationQty: nextQty,
              }
            : u
        ),
      }));
      return { ok: true, message: 'উপবরাদ্দ হালনাগাদ হয়েছে।' };
    },
    [db.unoSuballocations, getDCUsage, getMinistryUsage]
  );

  const deleteUNOSuballocation = useCallback((id: string) => {
    setDb((prev) => ({
      ...prev,
      unoSuballocations: prev.unoSuballocations.filter((u) => u.id !== id),
    }));
  }, []);

  const saveDOBook = useCallback(
    (input: { bookNo: number; startPage: number; endPage: number }): ActionResult => {
      const bookNo = Math.round(input.bookNo);
      const startPage = Math.round(input.startPage);
      const endPage = Math.round(input.endPage);
      if (bookNo <= 0 || startPage <= 0 || endPage < startPage) {
        return { ok: false, error: 'সঠিক বই নং, শুরু পাতা ও শেষ পাতা দিন (শেষ পাতা >= শুরু পাতা)।' };
      }
      const overlap = db.doBooks.find(
        (b) => startPage <= b.endPage && endPage >= b.startPage
      );
      if (overlap) {
        return {
          ok: false,
          error: `বই নং ${toBnDigits(overlap.bookNo)}-এর পাতার (${toBnDigits(overlap.startPage)}-${toBnDigits(overlap.endPage)}) সাথে মিল আছে!`,
        };
      }
      const row: DOBook = {
        id: uid('book'),
        bookNo,
        startPage,
        endPage,
        totalCount: endPage - startPage + 1,
        createdAt: new Date().toISOString(),
      };
      setDb((prev) => ({ ...prev, doBooks: [...prev.doBooks, row] }));
      return { ok: true, message: `বই নং ${toBnDigits(bookNo)} (${toBnDigits(row.totalCount)}টি ডিও) সংরক্ষিত হয়েছে।` };
    },
    [db.doBooks]
  );

  const updateDOBook = useCallback(
    (
      id: string,
      input: { bookNo: number; startPage: number; endPage: number }
    ): ActionResult => {
      const existing = db.doBooks.find((b) => b.id === id);
      if (!existing) return { ok: false, error: 'ডিও বই খুঁজে পাওয়া যায়নি।' };
      const bookNo = Math.round(input.bookNo);
      const startPage = Math.round(input.startPage);
      const endPage = Math.round(input.endPage);
      if (bookNo <= 0 || startPage <= 0 || endPage < startPage) {
        return { ok: false, error: 'সঠিক বই নং, শুরু পাতা ও শেষ পাতা দিন (শেষ পাতা >= শুরু পাতা)।' };
      }
      const overlap = db.doBooks.find(
        (b) => b.id !== id && startPage <= b.endPage && endPage >= b.startPage
      );
      if (overlap) {
        return {
          ok: false,
          error: `বই নং ${toBnDigits(overlap.bookNo)}-এর পাতার (${toBnDigits(overlap.startPage)}-${toBnDigits(overlap.endPage)}) সাথে মিল আছে!`,
        };
      }
      const totalCount = endPage - startPage + 1;
      setDb((prev) => ({
        ...prev,
        doBooks: prev.doBooks.map((b) =>
          b.id === id ? { ...b, bookNo, startPage, endPage, totalCount } : b
        ),
      }));
      return { ok: true, message: `বই নং ${toBnDigits(bookNo)} (${toBnDigits(totalCount)}টি ডিও) হালনাগাদ হয়েছে।` };
    },
    [db.doBooks]
  );

  const deleteDOBook = useCallback((id: string) => {
    setDb((prev) => ({ ...prev, doBooks: prev.doBooks.filter((b) => b.id !== id) }));
  }, []);

  const saveChallan = useCallback(
    (input: Omit<ChallanRecord, 'id' | 'createdAt'>, existingId?: string): ActionResult => {
      if (!input.challanNo.trim() || !input.dealerName.trim()) {
        return { ok: false, error: 'চালান নং ও ডিলারের নাম দিন।' };
      }
      const normNo = normMemoNo(input.challanNo);
      if (!normNo) {
        return { ok: false, error: 'সঠিক চালান নম্বর দিন।' };
      }

      // একবার ব্যবহৃত চালান পুনরায় এন্ট্রি নেওয়া যাবে না (একই ধরন ও নম্বরের চালান ডুপ্লিকেট বা ইতিমধ্যে ব্যবহৃত হলে বাতিল)
      const dupChallan = db.challans.find(
        (c) =>
          c.id !== existingId &&
          c.challanType === input.challanType &&
          normMemoNo(c.challanNo) === normNo
      );
      if (dupChallan) {
        return {
          ok: false,
          error: `${input.challanType} নং ${toBnDigits(input.challanNo.trim())} ইতিমধ্যে ব্যবহৃত হয়েছে (${dupChallan.dealerName})! একবার ব্যবহৃত চালান আর ব্যবহার করা যাবে না।`,
        };
      }

      const existingRec = existingId ? db.challans.find((c) => c.id === existingId) : undefined;
      const usedInDo = db.doIssues.find((d) => {
        const matchMain =
          input.challanType === 'মূল চালান' && d.challanNo && normMemoNo(d.challanNo) === normNo;
        const matchTax =
          input.challanType === 'আয়কর চালান' && d.taxChallanNo && normMemoNo(d.taxChallanNo) === normNo;
        if (!matchMain && !matchTax) return false;
        if (
          existingRec &&
          existingRec.doNo &&
          normMemoNo(d.doNo) === normMemoNo(existingRec.doNo) &&
          normMemoNo(existingRec.challanNo) === normNo
        ) {
          return false;
        }
        return true;
      });
      if (usedInDo) {
        return {
          ok: false,
          error: `${input.challanType} নং ${toBnDigits(input.challanNo.trim())} ইতিমধ্যে ডিও নং ${toBnDigits(usedInDo.doNo)} (${usedInDo.recipientName})-এ ব্যবহৃত হয়েছে! একবার ব্যবহৃত চালান আর ব্যবহার করা যাবে না।`,
        };
      }

      const targetDoNo = (input.doNo || '').trim();
      if (targetDoNo) {
        const targetDo = db.doIssues.find((d) => normMemoNo(d.doNo) === normMemoNo(targetDoNo));
        if (targetDo) {
          if (targetDo.category !== input.category) {
            return {
              ok: false,
              error: `এক খাতের চালান (${input.category}) অন্য খাতের ডিও নং ${toBnDigits(targetDo.doNo)} (${targetDo.category})-এ যুক্ত করা যাবে না!`,
            };
          }
          if (!isSamePartyName(targetDo.recipientName, input.dealerName)) {
            return {
              ok: false,
              error: `এক ডিলার/ব্যক্তির চালান (${input.dealerName}) অন্য ডিলার/ব্যক্তির ডিও নং ${toBnDigits(targetDo.doNo)} (${targetDo.recipientName})-এ যুক্ত করা যাবে না!`,
            };
          }
          if (!isChallanOnOrBeforeDo(input.challanDate, targetDo.issueDate, true)) {
            return {
              ok: false,
              error: `চালানের তারিখ অবশ্যই ডিও ইস্যুর একই তারিখ বা তার আগের তারিখ হতে হবে (ডিও তারিখ: ${targetDo.issueDate})!`,
            };
          }
          const alreadyAssigned =
            input.challanType === 'মূল চালান' ? targetDo.challanNo : targetDo.taxChallanNo;
          if (alreadyAssigned && normMemoNo(alreadyAssigned) !== normNo) {
            return {
              ok: false,
              error: `ডিও নং ${toBnDigits(targetDo.doNo)}-এ ইতিমধ্যে ${input.challanType} নং ${toBnDigits(alreadyAssigned)} যুক্ত আছে!`,
            };
          }
        }
      }

      touchedMainChallansRef.current = true;
      if (existingId) {
        setDb((prev) =>
          reconcileChallansAndDOs({
            ...prev,
            challans: prev.challans.map((c) =>
              c.id === existingId
                ? {
                    ...c,
                    ...input,
                    challanNo: input.challanNo.trim(),
                    recordedAmount: round2(input.recordedAmount),
                    depositedAmount: round2(input.depositedAmount),
                  }
                : c
            ),
          })
        );
        return { ok: true, message: 'চালান তথ্য আপডেট হয়েছে।' };
      }
      const row: ChallanRecord = {
        ...input,
        challanNo: input.challanNo.trim(),
        recordedAmount: round2(input.recordedAmount),
        depositedAmount: round2(input.depositedAmount),
        id: uid('chln'),
        createdAt: new Date().toISOString(),
      };
      setDb((prev) =>
        reconcileChallansAndDOs({
          ...prev,
          challans: [row, ...prev.challans],
        })
      );
      return { ok: true, message: 'নতুন চালান সংরক্ষিত হয়েছে।' };
    },
    [db.challans, db.doIssues]
  );

  const deleteChallan = useCallback((id: string) => {
    touchedMainChallansRef.current = true;
    setDb((prev) => {
      const target = prev.challans.find((c) => c.id === id);
      const targetNorm = target ? normMemoNo(target.challanNo) : '';
      const cleanedDos = target
        ? prev.doIssues.map((d) => {
            if (
              target.challanType === 'মূল চালান' &&
              d.challanNo &&
              normMemoNo(d.challanNo) === targetNorm
            ) {
              return { ...d, challanNo: undefined, challanDate: undefined, challanAmount: undefined };
            }
            if (
              target.challanType === 'আয়কর চালান' &&
              d.taxChallanNo &&
              normMemoNo(d.taxChallanNo) === targetNorm
            ) {
              return {
                ...d,
                taxChallanNo: undefined,
                taxChallanDate: undefined,
                taxChallanAmount: undefined,
              };
            }
            return d;
          })
        : prev.doIssues;
      return reconcileChallansAndDOs({
        ...prev,
        challans: prev.challans.filter((c) => c.id !== id),
        doIssues: cleanedDos,
      });
    });
  }, []);

  const verifyChallanForRecipient = useCallback(
    (
      category: AllocationCategory,
      dealerName: string,
      allocationQty: number,
      opts?: {
        subId?: string;
        unoDate?: string;
        unoMemo?: string;
        issueDate?: string;
        excludeDoId?: string;
        excludeChallanIds?: Set<string>;
      }
    ) => {
      if (!CHALLAN_CATEGORIES.includes(category)) {
        return { applicable: false, ok: true, message: 'চালান প্রযোজ্য নয় (বিনামূল্য খাত)' };
      }
      const rate = db.rates.find((r) => r.category === category);
      const effUnitPrice = rate?.unitPrice || (category === 'খাদ্য বান্ধব' ? 13000 : 0);
      const effTaxRate = rate?.taxRate || (category === 'খাদ্য বান্ধব' ? 200 : 0);
      const reqMain = round2(allocationQty * effUnitPrice);
      const reqTax = round2(allocationQty * effTaxRate);

      // ইতিমধ্যে অন্য ডিও-তে ব্যবহৃত সব চালান নং ও চালান আইডি চিহ্নিত করি যাতে একবার ব্যবহৃত চালান আর ব্যবহার করা না যায়
      const usedMainNos = new Set<string>();
      const usedTaxNos = new Set<string>();
      const usedDoNos = new Set<string>();
      for (const d of db.doIssues) {
        if (opts?.excludeDoId && d.id === opts.excludeDoId) continue;
        if (d.doNo) usedDoNos.add(normMemoNo(d.doNo));
        if (d.challanNo) usedMainNos.add(normMemoNo(d.challanNo));
        if (d.taxChallanNo) usedTaxNos.add(normMemoNo(d.taxChallanNo));
      }

      const isChallanAlreadyUsed = (c: ChallanRecord): boolean => {
        if (opts?.excludeChallanIds?.has(c.id)) return true;
        const cNo = normMemoNo(c.challanNo);
        if (c.challanType === 'মূল চালান' && cNo && usedMainNos.has(cNo)) return true;
        if (c.challanType === 'আয়কর চালান' && cNo && usedTaxNos.has(cNo)) return true;
        if (c.doNo && usedDoNos.has(normMemoNo(c.doNo))) return true;
        return false;
      };

      const targetDoDate = (opts?.issueDate || '').trim();
      const challanNumSortVal = (no: string): number => {
        const digits = bnToEnDigits(no || '').replace(/\D/g, '');
        return digits ? Number(digits) : Number.MAX_SAFE_INTEGER;
      };

      const pickChallan = (
        type: ChallanRecord['challanType']
      ): { chosen?: ChallanRecord; hadUsedOnly: boolean; hadPostDatedOnly: boolean } => {
        // Strict: same category AND same dealer/institution/person ONLY (never cross-category or cross-dealer)
        const samePartyAndCat = db.challans.filter(
          (c) =>
            c.challanType === type &&
            c.category === category &&
            isSamePartyName(c.dealerName, dealerName)
        );
        if (samePartyAndCat.length === 0) {
          return { chosen: undefined, hadUsedOnly: false, hadPostDatedOnly: false };
        }

        const unused = samePartyAndCat.filter((c) => !isChallanAlreadyUsed(c));
        if (unused.length === 0) {
          return { chosen: undefined, hadUsedOnly: true, hadPostDatedOnly: false };
        }

        // Strict date rule: challan date must be <= DO issue date (when DO issue date is known)
        const dateEligible = targetDoDate
          ? unused.filter((c) => isChallanOnOrBeforeDo(c.challanDate, targetDoDate, true))
          : unused.filter((c) => Boolean(toIsoDateStr(c.challanDate)));

        if (dateEligible.length === 0) {
          return { chosen: undefined, hadUsedOnly: false, hadPostDatedOnly: true };
        }

        // Earliest challan maps to earliest DO for the same dealer & category
        const sorted = [...dateEligible].sort((a, b) => {
          const isoA = toIsoDateStr(a.challanDate) || '9999-12-31';
          const isoB = toIsoDateStr(b.challanDate) || '9999-12-31';
          if (isoA !== isoB) return isoA.localeCompare(isoB);

          // Same date tie-breaker: direct subId link -> qty proximity -> challan number ASC
          if (opts?.subId) {
            const subA = a.dealerAllocationId && normMemoNo(a.dealerAllocationId) === normMemoNo(opts.subId) ? 0 : 1;
            const subB = b.dealerAllocationId && normMemoNo(b.dealerAllocationId) === normMemoNo(opts.subId) ? 0 : 1;
            if (subA !== subB) return subA - subB;
          }
          const qDiffA = a.allocationQty ? Math.abs(a.allocationQty - allocationQty) : 999;
          const qDiffB = b.allocationQty ? Math.abs(b.allocationQty - allocationQty) : 999;
          if (Math.abs(qDiffA - qDiffB) > 0.001) return qDiffA - qDiffB;
          return challanNumSortVal(a.challanNo) - challanNumSortVal(b.challanNo);
        });

        return { chosen: sorted[0], hadUsedOnly: false, hadPostDatedOnly: false };
      };

      const mainRes = pickChallan('মূল চালান');
      const taxRes = pickChallan('আয়কর চালান');
      const mainC = mainRes.chosen;
      const taxC = taxRes.chosen;

      const errs: string[] = [];
      if (!mainC && reqMain > 0) {
        if (mainRes.hadUsedOnly) {
          errs.push('এই ডিলারের মূল চালান ইতিমধ্যে অন্য ডিও-তে ব্যবহৃত হয়েছে');
        } else if (mainRes.hadPostDatedOnly) {
          errs.push('মূল চালানের তারিখ ডিও ইস্যুর তারিখের পরে হওয়ায় সংযুক্ত করা যায়নি');
        } else {
          errs.push('একই খাত ও ডিলারের মূল চালান পাওয়া যায়নি');
        }
      } else if (mainC && mainC.depositedAmount + 0.01 < reqMain) {
        errs.push(`মূল চালানে কম জমা (${toBnDigits(round2(reqMain - mainC.depositedAmount))} ৳)`);
      }

      if (reqTax > 0) {
        if (!taxC) {
          if (taxRes.hadUsedOnly) {
            errs.push('এই ডিলারের আয়কর চালান ইতিমধ্যে অন্য ডিও-তে ব্যবহৃত হয়েছে');
          } else if (taxRes.hadPostDatedOnly) {
            errs.push('আয়কর চালানের তারিখ ডিও ইস্যুর তারিখের পরে হওয়ায় সংযুক্ত করা যায়নি');
          } else {
            errs.push('একই খাত ও ডিলারের আয়কর চালান পাওয়া যায়নি');
          }
        } else if (taxC.depositedAmount + 0.01 < reqTax) {
          errs.push(`আয়কর চালানে কম জমা (${toBnDigits(round2(reqTax - taxC.depositedAmount))} ৳)`);
        }
      }

      return {
        applicable: true,
        ok: errs.length === 0,
        requiredAmount: reqMain,
        requiredTaxAmount: reqTax,
        mainChallan: mainC,
        taxChallan: taxC,
        errors: errs,
      };
    },
    [db.challans, db.doIssues, db.rates]
  );

  const issueDOBatch = useCallback(
    (
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
    ): ActionResult => {
      if (!items.length) return { ok: false, error: 'কোনো ডিও নম্বর দেওয়া হয়নি।' };
      const existingSet = new Set(db.doIssues.map((d) => String(d.doNo).trim()));
      const batchSet = new Set<string>();
      for (const it of items) {
        const dNo = it.doNo.trim();
        if (!dNo || !it.issueDate) {
          return { ok: false, error: 'প্রতিটি নির্বাচিত প্রাপকের জন্য ডিও নং এবং তারিখ দিন।' };
        }
        if (existingSet.has(dNo)) {
          return { ok: false, error: `ডিও নং ${toBnDigits(dNo)} আগেই ইস্যু করা হয়েছে!` };
        }
        if (batchSet.has(dNo)) {
          return { ok: false, error: `একই ডিও নং ${toBnDigits(dNo)} দুইবার লেখা হয়েছে!` };
        }
        batchSet.add(dNo);
      }

      // একাধিক ডিও ইস্যুর ক্ষেত্রে একই ডিলারের সবচেয়ে আগের ডিও-তে সবচেয়ে আগের চালান ১-থেকে-১ বরাদ্দ হবে
      const doNumVal = (no: string): number => {
        const digits = bnToEnDigits(no || '').replace(/\D/g, '');
        return digits ? Number(digits) : Number.MAX_SAFE_INTEGER;
      };
      const sortedItems = [...items].sort((a, b) => {
        const isoA = toIsoDateStr(a.issueDate) || '9999-12-31';
        const isoB = toIsoDateStr(b.issueDate) || '9999-12-31';
        if (isoA !== isoB) return isoA.localeCompare(isoB);
        return doNumVal(a.doNo) - doNumVal(b.doNo);
      });

      const batchUsedChallanIds = new Set<string>();
      const now = new Date().toISOString();
      const created: DOIssue[] = sortedItems.map((it) => {
        const check = verifyChallanForRecipient(it.category, it.recipientName, it.allocationQty, {
          subId: it.subId,
          unoDate: it.unoDate,
          unoMemo: it.unoMemo,
          issueDate: it.issueDate,
          excludeChallanIds: batchUsedChallanIds,
        });
        if (check.mainChallan) batchUsedChallanIds.add(check.mainChallan.id);
        if (check.taxChallan) batchUsedChallanIds.add(check.taxChallan.id);

        return {
          id: uid('do'),
          doNo: it.doNo.trim(),
          issueDate: it.issueDate,
          category: it.category,
          unionName: it.unionName,
          recipientName: it.recipientName,
          salesCenter: it.salesCenter,
          allocationQty: round3(it.allocationQty),
          ministryMemo: it.ministryMemo,
          dcMemo: it.dcMemo,
          unoMemo: it.unoMemo,
          dealerAllocationId: it.subId,
          projectNo: it.projectNo,
          challanNo: check.mainChallan?.challanNo,
          challanDate: check.mainChallan?.challanDate,
          challanAmount: check.mainChallan?.depositedAmount,
          taxChallanNo: check.taxChallan?.challanNo,
          taxChallanDate: check.taxChallan?.challanDate,
          taxChallanAmount: check.taxChallan?.depositedAmount,
          requiredAmount: check.requiredAmount,
          requiredTaxAmount: check.requiredTaxAmount,
          createdAt: now,
        };
      });

      setDb((prev) =>
        reconcileChallansAndDOs({
          ...prev,
          doIssues: [...created, ...prev.doIssues],
        })
      );
      return {
        ok: true,
        message: `${toBnDigits(created.length)} টি ডিও সফলভাবে ইস্যু হয়েছে।`,
      };
    },
    [db.doIssues, verifyChallanForRecipient]
  );

  const updateDOIssue = useCallback(
    (
      id: string,
      patch: Partial<Pick<DOIssue, 'doNo' | 'issueDate' | 'allocationQty'>>
    ): ActionResult => {
      const existing = db.doIssues.find((d) => d.id === id);
      if (!existing) return { ok: false, error: 'ডিও রেকর্ড পাওয়া যায়নি।' };
      const nextDoNo = (patch.doNo ?? existing.doNo).trim();
      const nextDate = (patch.issueDate ?? existing.issueDate).trim();
      const nextQty = round3(patch.allocationQty ?? existing.allocationQty);
      if (!nextDoNo || !nextDate) {
        return { ok: false, error: 'ডিও নং এবং তারিখ পূরণ করুন।' };
      }
      if (nextQty <= 0) {
        return { ok: false, error: 'ডিও পরিমাণ ০-এর বেশি হতে হবে।' };
      }
      if (
        nextDoNo !== existing.doNo &&
        db.doIssues.some((d) => d.id !== id && String(d.doNo).trim() === nextDoNo)
      ) {
        return { ok: false, error: `ডিও নং ${toBnDigits(nextDoNo)} আগেই ইস্যু করা হয়েছে!` };
      }
      setDb((prev) =>
        reconcileChallansAndDOs({
          ...prev,
          doIssues: prev.doIssues.map((d) => {
            if (d.id !== id) return d;
            const mainDateOk = isChallanOnOrBeforeDo(d.challanDate, nextDate, false);
            const taxDateOk = isChallanOnOrBeforeDo(d.taxChallanDate, nextDate, false);
            return {
              ...d,
              doNo: nextDoNo,
              issueDate: nextDate,
              allocationQty: nextQty,
              challanNo: mainDateOk ? d.challanNo : undefined,
              challanDate: mainDateOk ? d.challanDate : undefined,
              challanAmount: mainDateOk ? d.challanAmount : undefined,
              taxChallanNo: taxDateOk ? d.taxChallanNo : undefined,
              taxChallanDate: taxDateOk ? d.taxChallanDate : undefined,
              taxChallanAmount: taxDateOk ? d.taxChallanAmount : undefined,
            };
          }),
        })
      );
      return { ok: true, message: 'ডিও তথ্য হালনাগাদ হয়েছে।' };
    },
    [db.doIssues]
  );

  const deleteDOIssue = useCallback((id: string) => {
    setDb((prev) =>
      reconcileChallansAndDOs({
        ...prev,
        doIssues: prev.doIssues.filter((d) => d.id !== id),
      })
    );
  }, []);

  const saveFoodDealer = useCallback((dealer: Omit<FoodDealer, 'id'>, existingId?: string) => {
    const qty = round3((Number(dealer.cardCount) || 0) * 0.03);
    setDb((prev) => {
      if (existingId) {
        return {
          ...prev,
          foodDealers: prev.foodDealers.map((d) =>
            d.id === existingId ? { ...dealer, totalQty: qty, id: existingId } : d
          ),
        };
      }
      return {
        ...prev,
        foodDealers: [...prev.foodDealers, { ...dealer, totalQty: qty, id: uid('fd') }],
      };
    });
  }, []);

  const deleteFoodDealer = useCallback((id: string) => {
    setDb((prev) => ({ ...prev, foodDealers: prev.foodDealers.filter((d) => d.id !== id) }));
  }, []);

  const saveOMSDealer = useCallback((dealer: Omit<OMSDealer, 'id'>, existingId?: string) => {
    setDb((prev) => {
      if (existingId) {
        return {
          ...prev,
          omsDealers: prev.omsDealers.map((d) =>
            d.id === existingId ? { ...dealer, id: existingId } : d
          ),
        };
      }
      return {
        ...prev,
        omsDealers: [...prev.omsDealers, { ...dealer, id: uid('od') }],
      };
    });
  }, []);

  const deleteOMSDealer = useCallback((id: string) => {
    setDb((prev) => ({ ...prev, omsDealers: prev.omsDealers.filter((d) => d.id !== id) }));
  }, []);

  const saveChairman = useCallback((item: Omit<ChairmanEntry, 'id'>, existingId?: string) => {
    setDb((prev) => {
      if (existingId) {
        return {
          ...prev,
          chairmen: prev.chairmen.map((c) => (c.id === existingId ? { ...item, id: existingId } : c)),
        };
      }
      return { ...prev, chairmen: [...prev.chairmen, { ...item, id: uid('ch') }] };
    });
  }, []);

  const deleteChairman = useCallback((id: string) => {
    setDb((prev) => ({ ...prev, chairmen: prev.chairmen.filter((c) => c.id !== id) }));
  }, []);

  const saveFireService = useCallback((item: Omit<FireServiceEntry, 'id'>, existingId?: string) => {
    setDb((prev) => {
      if (existingId) {
        return {
          ...prev,
          fireServices: prev.fireServices.map((f) =>
            f.id === existingId ? { ...item, id: existingId } : f
          ),
        };
      }
      return { ...prev, fireServices: [...prev.fireServices, { ...item, id: uid('fs') }] };
    });
  }, []);

  const deleteFireService = useCallback((id: string) => {
    setDb((prev) => ({ ...prev, fireServices: prev.fireServices.filter((f) => f.id !== id) }));
  }, []);

  const saveRate = useCallback((category: AllocationCategory, unitPrice: number, taxRate: number) => {
    setDb((prev) => {
      const existing = prev.rates.find((r) => r.category === category);
      if (existing) {
        return {
          ...prev,
          rates: prev.rates.map((r) =>
            r.category === category ? { ...r, unitPrice, taxRate } : r
          ),
        };
      }
      const next: CategoryRate = { id: uid('rt'), category, unitPrice, taxRate };
      return { ...prev, rates: [...prev.rates, next] };
    });
  }, []);

  const deleteRate = useCallback((id: string) => {
    setDb((prev) => ({ ...prev, rates: prev.rates.filter((r) => r.id !== id) }));
  }, []);

  const updateOfficerProfile = useCallback((patch: Partial<OfficerProfileSettings>) => {
    setDb((prev) => ({
      ...prev,
      officer: { ...prev.officer, ...patch },
    }));
  }, []);

  const updateReportTemplate = useCallback(
    (id: ReportTemplateId, patch: Partial<ReportTemplateConfig>) => {
      setDb((prev) => {
        const catalogItem = REPORT_TEMPLATE_CATALOG.find((x) => x.id === id);
        const baseTemplate =
          prev.templates[id] ||
          (catalogItem
            ? createDefaultTemplate(catalogItem, prev.officer)
            : prev.templates['monthly-dist']);
        return {
          ...prev,
          templates: {
            ...prev.templates,
            [id]: { ...baseTemplate, ...patch },
          },
        };
      });
    },
    []
  );

  const resetReportTemplate = useCallback((id: ReportTemplateId) => {
    setDb((prev) => {
      const item = REPORT_TEMPLATE_CATALOG.find((x) => x.id === id);
      if (!item) return prev;
      return {
        ...prev,
        templates: {
          ...prev.templates,
          [id]: createDefaultTemplate(item, prev.officer),
        },
      };
    });
  }, []);

  const saveReportMemo = useCallback((key: string, no: string, date: string) => {
    setDb((prev) => ({
      ...prev,
      reportMemos: { ...prev.reportMemos, [key]: { no, date } },
    }));
  }, []);

  const syncFromGoogleSheet = useCallback(
    async (
      rawIdOrUrl?: string,
      rawChallanIdOrUrl?: string,
      rawChallanGasUrl?: string,
      tokenOverride?: string | null,
      rawMainGasUrl?: string
    ): Promise<ActionResult> => {
      const cfg = sheetsConfigRef.current;
      const curDb = dbRef.current;
      const targetId = extractSpreadsheetId(rawIdOrUrl ?? cfg.spreadsheetId);
      const challanTargetId = extractSpreadsheetId(
        rawChallanIdOrUrl ?? cfg.challanSpreadsheetId ?? ''
      );
      const challanGasUrl = String(rawChallanGasUrl ?? cfg.challanAppsScriptUrl ?? '').trim();
      const mainGasUrl = String(rawMainGasUrl ?? cfg.appsScriptUrl ?? '').trim();

      if (!targetId && !challanTargetId && !challanGasUrl && !mainGasUrl) {
        return { ok: false, error: 'সেটিংসে Google Sheet ID বা Apps Script লিংক যুক্ত করুন।' };
      }

      setIsSyncingSheets(true);
      setGoogleAuthError(null);
      try {
        const token = tokenOverride ?? tokenRef.current ?? getGoogleAccessToken();
        let nextDb = curDb;
        let mainTitle = cfg.spreadsheetName;
        let tabsCount = 0;
        let resolvedSheetId = targetId;

        if (targetId || mainGasUrl) {
          let res: {
            db: AllocationDatabase;
            title: string;
            tabsFound: string[];
            matchedTabNames: string[];
            spreadsheetId?: string;
          };
          if (token && targetId) {
            res = await fetchAllocationFromOAuthSheet(token, targetId, curDb);
          } else if (targetId) {
            try {
              res = await fetchAllocationFromPublicSheet(targetId, curDb);
            } catch (pubErr) {
              if (mainGasUrl) {
                res = await fetchAllocationFromAppsScript(mainGasUrl, targetId, curDb);
              } else {
                throw pubErr;
              }
            }
          } else {
            res = await fetchAllocationFromAppsScript(mainGasUrl, '', curDb);
          }
          nextDb = res.db;
          mainTitle = res.title;
          tabsCount = res.matchedTabNames.length || res.tabsFound.length;
          if (!resolvedSheetId && res.spreadsheetId) {
            resolvedSheetId = extractSpreadsheetId(res.spreadsheetId);
          }
        }

        let challanTitle = cfg.challanSpreadsheetName || '';
        let challanSyncedCount = 0;

        if (challanTargetId) {
          const extChallan = await fetchExternalChallansFromSheet(challanTargetId, token);
          challanTitle = extChallan.title || challanTitle;
          if (extChallan.challans.length > 0) {
            challanSyncedCount = extChallan.challans.length;
            nextDb = { ...nextDb, challans: extChallan.challans };
          }
        } else if (challanGasUrl) {
          const extGas = await fetchExternalChallansFromAppsScript(challanGasUrl);
          challanTitle = 'চালান ওয়েব অ্যাপ';
          if (extGas.challans.length > 0) {
            challanSyncedCount = extGas.challans.length;
            nextDb = { ...nextDb, challans: extGas.challans };
          }
        }

        nextDb = reconcileChallansAndDOs(nextDb);
        if (countAllocationDatabaseRecords(nextDb) > 0) {
          hadNonEmptyDbRef.current = true;
        }
        lastSyncedDbJsonRef.current = JSON.stringify(nextDb);
        touchedMainChallansRef.current = false;
        savePendingPushFlag(false);
        setHasPendingSheetSync(false);
        setDb(nextDb);

        const summaryDetails = `মন্ত্রণালয় ${toBnDigits(nextDb.ministryAllocations.length)}, উপবরাদ্দ ${toBnDigits(
          nextDb.dealerSuballocations.length + nextDb.unoSuballocations.length
        )}, ডিও ${toBnDigits(nextDb.doIssues.length)}, চালান ${toBnDigits(nextDb.challans.length)}`;

        const hasMainSource = Boolean(resolvedSheetId || mainGasUrl);
        const syncMsg = hasMainSource
          ? `${mainTitle} (${toBnDigits(tabsCount)}টি ট্যাব · ${summaryDetails}${
              challanTitle ? ` · চালান উৎস: ${challanTitle}` : ''
            })`
          : `${challanTitle || 'চালান শিট'} থেকে ${toBnDigits(challanSyncedCount)}টি চালান সিঙ্ক হয়েছে`;

        setSheetsConfig((prev) => {
          const effectiveGas = mainGasUrl || String(prev.appsScriptUrl || '').trim();
          const confirmed = prev.writeConfirmedSheetIds || [];
          const syncedKey =
            (resolvedSheetId || effectiveGas) && tabsCount > 0
              ? getWriteTargetKey(resolvedSheetId, effectiveGas)
              : '';
          const nextConfirmed =
            syncedKey && !confirmed.includes(syncedKey) ? [...confirmed, syncedKey] : confirmed;
          return {
            ...prev,
            enabled: true,
            spreadsheetId: resolvedSheetId || prev.spreadsheetId,
            spreadsheetName: mainTitle || prev.spreadsheetName,
            appsScriptUrl: effectiveGas || prev.appsScriptUrl,
            writeConfirmedSheetIds: nextConfirmed,
            challanSpreadsheetId: challanTargetId || prev.challanSpreadsheetId,
            challanSpreadsheetName: challanTitle || prev.challanSpreadsheetName,
            challanAppsScriptUrl: challanGasUrl || prev.challanAppsScriptUrl,
            lastSyncAt: new Date().toISOString(),
            lastSyncStatus: 'ok',
            lastSyncMessage: syncMsg,
          };
        });

        return {
          ok: true,
          message: hasMainSource
            ? `"${mainTitle}" থেকে সফলভাবে সিঙ্ক হয়েছে (${summaryDetails})।`
            : `${challanTitle || 'চালান শিট'} থেকে ${toBnDigits(challanSyncedCount)}টি চালান সিঙ্ক হয়েছে!`,
        };
      } catch (err) {
        if (err instanceof GoogleAuthExpiredError) {
          clearGoogleAccessToken();
          setGoogleAccessToken(null);
        }
        const msg = err instanceof Error ? err.message : 'Google Sheet সিঙ্ক ব্যর্থ হয়েছে।';
        setSheetsConfig((prev) => ({
          ...prev,
          spreadsheetId: targetId || prev.spreadsheetId,
          lastSyncStatus: 'error',
          lastSyncMessage: msg,
        }));
        return { ok: false, error: msg };
      } finally {
        setIsSyncingSheets(false);
      }
    },
    []
  );

  const pushToGoogleSheet = useCallback(
    async (
      overrideGasUrl?: string,
      overrideSheetId?: string,
      tokenOverride?: string | null
    ): Promise<ActionResult> => {
      const cfg = sheetsConfigRef.current;
      const targetId = extractSpreadsheetId(overrideSheetId ?? cfg.spreadsheetId);
      const gasUrl = String(overrideGasUrl ?? cfg.appsScriptUrl ?? '').trim();
      const token = tokenOverride ?? tokenRef.current ?? getGoogleAccessToken();

      if (!targetId && !gasUrl) {
        return {
          ok: false,
          error: 'প্রথমে মূল Google Sheet ID অথবা Apps Script Web App URL যুক্ত করুন।',
        };
      }
      if (!token && !gasUrl) {
        savePendingPushFlag(true);
        setHasPendingSheetSync(true);
        return {
          ok: false,
          error:
            'শিটে ডেটা লিখতে প্রথমে Google Sign-In করুন অথবা নিচে "মূল শিট Save/Push Apps Script Web App URL" যুক্ত করুন।',
        };
      }

      const curDb = dbRef.current;
      if (countAllocationDatabaseRecords(curDb) === 0 && !hadNonEmptyDbRef.current) {
        return {
          ok: false,
          error:
            'অ্যাপে এখনো কোনো বরাদ্দ বা ডিও ডেটা নেই। খালি অ্যাপ থেকে Push করলে Google Sheet-এর বিদ্যমান ডেটা মুছে যেতে পারে—তাই আগে «লিংক সেভ ও শিট থেকে ডেটা আনুন» বাটনে ক্লিক করে শিট থেকে ডেটা আনুন।',
        };
      }

      const hasExternalChallanSource = Boolean(
        extractSpreadsheetId(cfg.challanSpreadsheetId || '') ||
          String(cfg.challanAppsScriptUrl || '').trim()
      );
      const hasManualChallans = curDb.challans.some((c) => c.id.startsWith('chln-'));
      const writeOpts = {
        skipChallansTab:
          hasExternalChallanSource && !hasManualChallans && !touchedMainChallansRef.current,
        allowEmptyOverwrite: hadNonEmptyDbRef.current,
      };
      const snapshotJson = JSON.stringify(curDb);
      isPushingRef.current = true;
      setIsSyncingSheets(true);
      let pushSucceeded = false;
      try {
        let tabsUpdated = 0;
        let modeLabel = 'Google OAuth';

        if (token && targetId) {
          try {
            const res = await writeAllocationDatabaseToOAuthSheet(
              token,
              targetId,
              curDb,
              writeOpts
            );
            tabsUpdated = res.tabsUpdated;
          } catch (oauthErr) {
            if (oauthErr instanceof GoogleAuthExpiredError) {
              clearGoogleAccessToken();
              setGoogleAccessToken(null);
              tokenRef.current = null;
            }
            if (gasUrl) {
              const gasRes = await writeAllocationDatabaseToAppsScript(
                gasUrl,
                targetId,
                curDb,
                writeOpts
              );
              tabsUpdated = gasRes.tabsUpdated;
              modeLabel = 'Apps Script';
            } else {
              throw oauthErr;
            }
          }
        } else if (gasUrl) {
          const gasRes = await writeAllocationDatabaseToAppsScript(
            gasUrl,
            targetId,
            curDb,
            writeOpts
          );
          tabsUpdated = gasRes.tabsUpdated;
          modeLabel = 'Apps Script';
        }

        pushSucceeded = true;
        touchedMainChallansRef.current = false;
        lastSyncedDbJsonRef.current = snapshotJson;
        savePendingPushFlag(false);
        setHasPendingSheetSync(false);
        const targetKey = getWriteTargetKey(targetId, gasUrl);
        setSheetsConfig((prev) => {
          const confirmed = prev.writeConfirmedSheetIds || [];
          const nextConfirmed =
            targetKey && !confirmed.includes(targetKey) ? [...confirmed, targetKey] : confirmed;
          return {
            ...prev,
            enabled: true,
            spreadsheetId: targetId || prev.spreadsheetId,
            appsScriptUrl: gasUrl || prev.appsScriptUrl,
            writeConfirmedSheetIds: nextConfirmed,
            lastSyncAt: new Date().toISOString(),
            lastSyncStatus: 'ok',
            lastSyncMessage: `সকল ডেটা Google Sheet-এ (${toBnDigits(tabsUpdated)}টি ট্যাব · ${modeLabel}) সেভ হয়েছে`,
          };
        });
        return {
          ok: true,
          message: `সকল বরাদ্দ, ডিও, চালান ও সেটিংস Google Sheet-এ (${modeLabel}) সেভ হয়েছে!`,
        };
      } catch (err) {
        if (err instanceof GoogleAuthExpiredError) {
          clearGoogleAccessToken();
          setGoogleAccessToken(null);
          tokenRef.current = null;
        }
        savePendingPushFlag(true);
        setHasPendingSheetSync(true);
        const msg = err instanceof Error ? err.message : 'শিটে সংরক্ষণ ব্যর্থ হয়েছে।';
        setSheetsConfig((prev) => ({
          ...prev,
          lastSyncStatus: 'error',
          lastSyncMessage: `শিটে সেভ ব্যর্থ: ${msg}`,
        }));
        return { ok: false, error: msg };
      } finally {
        isPushingRef.current = false;
        setIsSyncingSheets(false);
        if (pushSucceeded && pendingAutoPushRef.current) {
          pendingAutoPushRef.current = false;
          if (JSON.stringify(dbRef.current) !== lastSyncedDbJsonRef.current) {
            setTimeout(() => {
              void pushToGoogleSheet();
            }, 50);
          }
        } else if (!pushSucceeded) {
          pendingAutoPushRef.current = false;
        }
      }
    },
    []
  );

  const buildWriteConfirmRequest = useCallback(
    (overrideGasUrl?: string, overrideSheetId?: string): SheetWriteConfirmRequest => {
      const cfg = sheetsConfigRef.current;
      const curDb = dbRef.current;
      const targetId = extractSpreadsheetId(overrideSheetId ?? cfg.spreadsheetId);
      const gasUrl = String(overrideGasUrl ?? cfg.appsScriptUrl ?? '').trim();
      const targetName =
        cfg.spreadsheetName ||
        (targetId ? `Google Sheet (${targetId.slice(0, 12)}...)` : '') ||
        (gasUrl ? 'সংযুক্ত Apps Script শিট' : 'সংযুক্ত Google Sheet');

      return {
        title: 'Google Sheet-এ সকল ডেটা সেভ নিশ্চিত করুন',
        description:
          'অ্যাপের সকল বরাদ্দ, উপবরাদ্দ, প্রকল্প, ডিও ইস্যু, চালান, ডিলার/মাস্টার তালিকা এবং টেমপ্লেট সেটিংস আপনার সংযুক্ত Google Sheet-এ হালনাগাদ ও সংরক্ষণ করা হবে।',
        spreadsheetName: targetName,
        targetSheets: Object.values(ALLOC_SHEET_TABS).map((t) => t.name),
        details: [
          `মন্ত্রণালয় ও ডিসি বরাদ্দ: ${toBnDigits(curDb.ministryAllocations.length + curDb.dcAllocations.length)}টি`,
          `ডিলার ও ইউএনও উপবরাদ্দ: ${toBnDigits(curDb.dealerSuballocations.length + curDb.unoSuballocations.length)}টি (প্রকল্প ${toBnDigits(curDb.projects.length)}টি)`,
          `ডিও বই ও ইস্যু: ${toBnDigits(curDb.doBooks.length)}টি বই, ${toBnDigits(curDb.doIssues.length)}টি ডিও`,
          `ট্রেজারি চালান ও ডিলার তালিকা: ${toBnDigits(curDb.challans.length)}টি চালান, ${toBnDigits(curDb.foodDealers.length + curDb.omsDealers.length)} জন ডিলার`,
        ],
        isDestructive: true,
        confirmLabel: 'নিশ্চিত ও Google Sheet-এ সেভ করুন',
      };
    },
    []
  );

  const requestPushToGoogleSheet = useCallback(
    (overrideGasUrl?: string, overrideSheetId?: string): Promise<ActionResult> => {
      const cfg = sheetsConfigRef.current;
      const targetId = extractSpreadsheetId(overrideSheetId ?? cfg.spreadsheetId);
      const gasUrl = String(overrideGasUrl ?? cfg.appsScriptUrl ?? '').trim();
      const token = tokenRef.current ?? getGoogleAccessToken();

      if (!targetId && !gasUrl) {
        return Promise.resolve({
          ok: false,
          error: 'প্রথমে মূল Google Sheet ID অথবা Apps Script Web App URL যুক্ত করুন।',
        });
      }
      if (!token && !gasUrl) {
        savePendingPushFlag(true);
        setHasPendingSheetSync(true);
        return Promise.resolve({
          ok: false,
          error:
            'শিটে ডেটা লিখতে প্রথমে Google Sign-In করুন অথবা নিচে "মূল শিট Save/Push Apps Script Web App URL" যুক্ত করুন।',
        });
      }

      if (countAllocationDatabaseRecords(dbRef.current) === 0 && !hadNonEmptyDbRef.current) {
        return Promise.resolve({
          ok: false,
          error:
            'অ্যাপে এখনো কোনো বরাদ্দ বা ডিও ডেটা নেই। খালি অ্যাপ থেকে Push করলে Google Sheet-এর বিদ্যমান ডেটা মুছে যেতে পারে—তাই আগে «লিংক সেভ ও শিট থেকে ডেটা আনুন» বাটনে ক্লিক করে শিট থেকে ডেটা আনুন।',
        });
      }

      const targetKey = getWriteTargetKey(targetId, gasUrl);
      const request = buildWriteConfirmRequest(overrideGasUrl, overrideSheetId);

      return new Promise<ActionResult>((resolve) => {
        setWriteConfirmState({
          request,
          targetKey,
          overrideGasUrl,
          overrideSheetId,
          resolveResult: resolve,
        });
      });
    },
    [buildWriteConfirmRequest]
  );

  const confirmPendingSheetWrite = useCallback(async () => {
    const pending = writeConfirmState;
    if (!pending) return;
    if (pending.targetKey) {
      setSheetsConfig((prev) => {
        const list = prev.writeConfirmedSheetIds || [];
        return list.includes(pending.targetKey)
          ? prev
          : { ...prev, writeConfirmedSheetIds: [...list, pending.targetKey] };
      });
    }
    const res = await pushToGoogleSheet(pending.overrideGasUrl, pending.overrideSheetId);
    setWriteConfirmState(null);
    pending.resolveResult?.(res);
  }, [pushToGoogleSheet, writeConfirmState]);

  const cancelPendingSheetWrite = useCallback(() => {
    const pending = writeConfirmState;
    setWriteConfirmState(null);
    pending?.resolveResult?.({
      ok: false,
      error: 'ব্যবহারকারী Google Sheet সেভ বাতিল করেছেন।',
    });
  }, [writeConfirmState]);

  const autoDetectAndSyncDriveSheets = useCallback(
    async (tokenOverride?: string | null): Promise<ActionResult> => {
      const token = tokenOverride ?? tokenRef.current ?? getGoogleAccessToken();
      if (!token) {
        return { ok: false, error: 'Google Drive স্ক্যান করতে প্রথমে Google Sign-In করুন।' };
      }
      setIsSyncingSheets(true);
      try {
        const res = await autoDetectAllocationSpreadsheets(token);
        setDriveFiles(res.files);
        const curCfg = sheetsConfigRef.current;
        const detectedMainId = res.main?.file.id || curCfg.spreadsheetId;
        const detectedMainName = res.main?.meta.title || curCfg.spreadsheetName;
        const detectedChallanId = res.challan?.file.id || curCfg.challanSpreadsheetId || '';
        const detectedChallanName = res.challan?.meta.title || curCfg.challanSpreadsheetName || '';

        if (!detectedMainId && !detectedChallanId) {
          return {
            ok: true,
            message: `Google Drive-এ ${toBnDigits(res.files.length)}টি শিট পাওয়া গেছে; ড্রপডাউন থেকে আপনার শিট নির্বাচন করুন।`,
          };
        }

        setSheetsConfig((prev) => ({
          ...prev,
          enabled: true,
          spreadsheetId: detectedMainId,
          spreadsheetName: detectedMainName,
          challanSpreadsheetId: detectedChallanId,
          challanSpreadsheetName: detectedChallanName,
        }));

        const syncRes = await syncFromGoogleSheet(
          detectedMainId,
          detectedChallanId,
          curCfg.challanAppsScriptUrl,
          token
        );
        return syncRes.ok
          ? {
              ok: true,
              message: `Drive থেকে অটো-ডিটেক্ট ও সিঙ্ক সম্পন্ন (${detectedMainName || detectedChallanName})!`,
            }
          : syncRes;
      } catch (err) {
        if (err instanceof GoogleAuthExpiredError) {
          clearGoogleAccessToken();
          setGoogleAccessToken(null);
        }
        return {
          ok: false,
          error: err instanceof Error ? err.message : 'Google Drive স্ক্যান ব্যর্থ হয়েছে।',
        };
      } finally {
        setIsSyncingSheets(false);
      }
    },
    [syncFromGoogleSheet]
  );

  const handleGoogleSignIn = useCallback(async (): Promise<ActionResult> => {
    setIsGoogleSigningIn(true);
    setGoogleAuthError(null);
    try {
      const { user, accessToken } = await googleSignIn();
      setGoogleUser(user);
      setGoogleAccessToken(accessToken);
      tokenRef.current = accessToken;

      const cfg = sheetsConfigRef.current;
      const targetId = extractSpreadsheetId(cfg.spreadsheetId);
      const gasUrl = String(cfg.appsScriptUrl || '').trim();
      const pendingLocal =
        loadPendingPushFlag() || JSON.stringify(dbRef.current) !== lastSyncedDbJsonRef.current;

      if (pendingLocal && targetId) {
        const targetKey = getWriteTargetKey(targetId, gasUrl);
        const isConfirmed = Boolean(targetKey && (cfg.writeConfirmedSheetIds || []).includes(targetKey));
        if (isConfirmed) {
          const pushRes = await pushToGoogleSheet(gasUrl, targetId, accessToken);
          return pushRes.ok
            ? { ok: true, message: 'Google সাইন-ইন ও অপেক্ষমাণ সকল এন্ট্রি শিটে সেভ হয়েছে!' }
            : pushRes;
        }
        return {
          ok: true,
          message: 'Google সাইন-ইন সম্পন্ন! অপেক্ষমাণ এন্ট্রি শিটে পাঠাতে "শিটে সেভ করুন" নিশ্চিত করুন।',
        };
      }

      if (
        cfg.spreadsheetId ||
        cfg.appsScriptUrl ||
        cfg.challanSpreadsheetId ||
        cfg.challanAppsScriptUrl
      ) {
        await syncFromGoogleSheet(
          cfg.spreadsheetId,
          cfg.challanSpreadsheetId,
          cfg.challanAppsScriptUrl,
          accessToken,
          cfg.appsScriptUrl
        );
      } else {
        await autoDetectAndSyncDriveSheets(accessToken);
      }
      return { ok: true, message: 'Google সাইন-ইন ও শিট সংযোগ সম্পন্ন হয়েছে।' };
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Google সাইন-ইন ব্যর্থ হয়েছে।';
      setGoogleAuthError(msg);
      return { ok: false, error: msg };
    } finally {
      setIsGoogleSigningIn(false);
    }
  }, [autoDetectAndSyncDriveSheets, pushToGoogleSheet, syncFromGoogleSheet]);

  const handleGoogleSignOut = useCallback(async () => {
    await googleLogout();
    setGoogleUser(null);
    setGoogleAccessToken(null);
    setDriveFiles([]);
  }, []);

  const cancelGoogleSignIn = useCallback(() => {
    cancelPendingAndroidSignIn();
    setIsGoogleSigningIn(false);
    setGoogleAuthError(
      'Google সাইন-ইন বাতিল করা হয়েছে। চাইলে লগইন ছাড়াই নিচে Apps Script URL দিয়ে শিটে সেভ করতে পারবেন।'
    );
  }, []);

  useEffect(() => {
    if (initialAutoSyncDoneRef.current) return;
    initialAutoSyncDoneRef.current = true;
    const cfg = sheetsConfigRef.current;
    // Never overwrite unsynced local entries on startup!
    if (loadPendingPushFlag()) {
      setHasPendingSheetSync(true);
      return;
    }
    if (
      cfg.autoSync !== false &&
      (cfg.spreadsheetId ||
        cfg.appsScriptUrl ||
        cfg.challanSpreadsheetId ||
        cfg.challanAppsScriptUrl)
    ) {
      void syncFromGoogleSheet(
        cfg.spreadsheetId,
        cfg.challanSpreadsheetId,
        cfg.challanAppsScriptUrl,
        undefined,
        cfg.appsScriptUrl
      );
    }
  }, [syncFromGoogleSheet]);

  const scheduleAutoPush = useCallback(() => {
    const cfg = sheetsConfigRef.current;
    const targetId = extractSpreadsheetId(cfg.spreadsheetId);
    const gasUrl = String(cfg.appsScriptUrl || '').trim();
    const hasWriteTarget = Boolean(targetId || gasUrl);

    if (countAllocationDatabaseRecords(dbRef.current) === 0 && !hadNonEmptyDbRef.current) {
      return;
    }

    const currentJson = JSON.stringify(dbRef.current);
    const isDirty = currentJson !== lastSyncedDbJsonRef.current;

    if (hasWriteTarget && isDirty) {
      savePendingPushFlag(true);
      setHasPendingSheetSync(true);
    } else if (!isDirty) {
      savePendingPushFlag(false);
      setHasPendingSheetSync(false);
    }

    if (cfg.autoSync === false || !hasWriteTarget) return;

    const token = tokenRef.current || getGoogleAccessToken();
    const canWrite = Boolean(gasUrl || (token && targetId));
    const targetKey = getWriteTargetKey(targetId, gasUrl);
    const isConfirmed = Boolean(targetKey && (cfg.writeConfirmedSheetIds || []).includes(targetKey));

    if (!canWrite || !isConfirmed || !isDirty) return;

    if (autoPushTimerRef.current) {
      clearTimeout(autoPushTimerRef.current);
    }
    autoPushTimerRef.current = setTimeout(() => {
      autoPushTimerRef.current = null;
      const latestJson = JSON.stringify(dbRef.current);
      if (latestJson === lastSyncedDbJsonRef.current) return;
      if (isPushingRef.current) {
        pendingAutoPushRef.current = true;
        return;
      }
      void pushToGoogleSheet();
    }, 450);
  }, [pushToGoogleSheet]);

  useEffect(() => {
    scheduleAutoPush();
    return () => {
      if (autoPushTimerRef.current) {
        clearTimeout(autoPushTimerRef.current);
      }
    };
  }, [
    db,
    googleAccessToken,
    sheetsConfig.appsScriptUrl,
    sheetsConfig.spreadsheetId,
    sheetsConfig.autoSync,
    sheetsConfig.writeConfirmedSheetIds,
    scheduleAutoPush,
  ]);

  useEffect(() => {
    if (typeof window === 'undefined' || typeof document === 'undefined') return;
    const flushPendingPush = () => {
      const cfg = sheetsConfigRef.current;
      if (cfg.autoSync === false) return;
      const targetId = extractSpreadsheetId(cfg.spreadsheetId);
      const gasUrl = String(cfg.appsScriptUrl || '').trim();
      const token = tokenRef.current || getGoogleAccessToken();
      const canWrite = Boolean(gasUrl || (token && targetId));
      const targetKey = getWriteTargetKey(targetId, gasUrl);
      const isConfirmed = Boolean(targetKey && (cfg.writeConfirmedSheetIds || []).includes(targetKey));
      const isDirty = JSON.stringify(dbRef.current) !== lastSyncedDbJsonRef.current;
      if (countAllocationDatabaseRecords(dbRef.current) === 0 && !hadNonEmptyDbRef.current) return;
      if (!canWrite || !isConfirmed || !isDirty) return;
      if (autoPushTimerRef.current) {
        clearTimeout(autoPushTimerRef.current);
        autoPushTimerRef.current = null;
      }
      if (isPushingRef.current) {
        pendingAutoPushRef.current = true;
        return;
      }
      void pushToGoogleSheet();
    };
    const onVisibilityChange = () => {
      if (document.visibilityState === 'hidden') flushPendingPush();
    };
    window.addEventListener('pagehide', flushPendingPush);
    document.addEventListener('visibilitychange', onVisibilityChange);
    return () => {
      window.removeEventListener('pagehide', flushPendingPush);
      document.removeEventListener('visibilitychange', onVisibilityChange);
    };
  }, [pushToGoogleSheet]);

  const sheetSyncState: SheetSyncUiState = useMemo(() => {
    const targetId = extractSpreadsheetId(sheetsConfig.spreadsheetId);
    const gasUrl = String(sheetsConfig.appsScriptUrl || '').trim();
    if (!targetId && !gasUrl) return 'unlinked';
    if (isSyncingSheets) return 'syncing';
    const token = googleAccessToken || getGoogleAccessToken();
    const canWrite = Boolean(gasUrl || (token && targetId));
    if (hasPendingSheetSync) {
      if (!canWrite) return 'needs_auth';
      const targetKey = getWriteTargetKey(targetId, gasUrl);
      const isConfirmed = Boolean(
        targetKey && (sheetsConfig.writeConfirmedSheetIds || []).includes(targetKey)
      );
      if (!isConfirmed) return 'needs_write_confirm';
    }
    if (sheetsConfig.lastSyncStatus === 'error') return 'error';
    if (hasPendingSheetSync) return 'pending_push';
    return 'synced';
  }, [
    googleAccessToken,
    hasPendingSheetSync,
    isSyncingSheets,
    sheetsConfig.appsScriptUrl,
    sheetsConfig.lastSyncStatus,
    sheetsConfig.spreadsheetId,
    sheetsConfig.writeConfirmedSheetIds,
  ]);

  const exportBackupJson = useCallback(() => {
    const blob = new Blob([JSON.stringify({ db, sheetsConfig }, null, 2)], {
      type: 'application/json;charset=utf-8',
    });
    triggerBlobDownload(blob, `khaddo-boraddo-backup-${new Date().toISOString().slice(0, 10)}.json`);
  }, [db, sheetsConfig]);

  const importBackupJson = useCallback((jsonText: string): ActionResult => {
    try {
      const parsed = JSON.parse(jsonText) as {
        db?: Partial<AllocationDatabase>;
        sheetsConfig?: Partial<AllocationSheetsConfig>;
      } & Partial<AllocationDatabase>;
      const source = parsed.db ? parsed.db : parsed;
      const emptyBase = createEmptyAllocationDatabase();
      const nextOfficer = { ...DEFAULT_OFFICER_PROFILE, ...(source.officer || {}) };
      const defaultTpls = createDefaultTemplatesMap(nextOfficer);
      setDb(
        reconcileChallansAndDOs({
          ...emptyBase,
          ...source,
          officer: nextOfficer,
          templates: { ...defaultTpls, ...(source.templates || {}) },
        })
      );
      if (parsed.sheetsConfig) {
        setSheetsConfig((prev) => ({ ...prev, ...parsed.sheetsConfig }));
      }
      return { ok: true, message: 'ব্যাকআপ ডেটা সফলভাবে রিস্টোর হয়েছে!' };
    } catch (err) {
      return {
        ok: false,
        error: err instanceof Error ? err.message : 'সঠিক JSON ব্যাকআপ ফাইল নয়।',
      };
    }
  }, []);

  const clearAllLocalData = useCallback(() => {
    hadNonEmptyDbRef.current = false;
    const emptyBase = createEmptyAllocationDatabase();
    setDb((prev) => {
      const cleared = {
        ...emptyBase,
        officer: prev.officer,
        templates: prev.templates,
      };
      lastSyncedDbJsonRef.current = JSON.stringify(cleared);
      return cleared;
    });
    savePendingPushFlag(false);
    setHasPendingSheetSync(false);
  }, []);

  const unusedDoList = useMemo(() => getUnusedDoNumbers(db), [db]);

  return {
    db,
    sheetsConfig,
    setSheetsConfig,
    googleUser,
    hasGoogleToken: Boolean(googleAccessToken),
    isGoogleSigningIn,
    googleAuthError,
    isSyncingSheets,
    hasPendingSheetSync,
    sheetSyncState,
    pendingWriteConfirmRequest: writeConfirmState?.request || null,
    confirmPendingSheetWrite,
    cancelPendingSheetWrite,
    driveFiles,
    unusedDoList,
    getMinistryUsage,
    getDCUsage,
    addMinistryAllocation,
    updateMinistryAllocation,
    deleteMinistryAllocation,
    addDCAllocation,
    updateDCAllocation,
    deleteDCAllocation,
    addProjectsToDC,
    updateProject,
    deleteProject,
    transferProject,
    saveDealerSuballocationBatch,
    updateDealerSuballocation,
    deleteDealerSuballocation,
    saveUNOSuballocationBatch,
    updateUNOSuballocation,
    deleteUNOSuballocation,
    saveDOBook,
    updateDOBook,
    deleteDOBook,
    saveChallan,
    deleteChallan,
    verifyChallanForRecipient,
    issueDOBatch,
    updateDOIssue,
    deleteDOIssue,
    saveFoodDealer,
    deleteFoodDealer,
    saveOMSDealer,
    deleteOMSDealer,
    saveChairman,
    deleteChairman,
    saveFireService,
    deleteFireService,
    saveRate,
    deleteRate,
    updateOfficerProfile,
    updateReportTemplate,
    resetReportTemplate,
    saveReportMemo,
    handleGoogleSignIn,
    handleGoogleSignOut,
    cancelGoogleSignIn,
    syncFromGoogleSheet,
    autoDetectAndSyncDriveSheets,
    pushToGoogleSheet,
    requestPushToGoogleSheet,
    exportBackupJson,
    importBackupJson,
    clearAllLocalData,
  };
}

