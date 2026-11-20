const BN_DIGITS = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];

export const BENGALI_MONTHS = [
  'জানুয়ারি',
  'ফেব্রুয়ারি',
  'মার্চ',
  'এপ্রিল',
  'মে',
  'জুন',
  'জুলাই',
  'আগস্ট',
  'সেপ্টেম্বর',
  'অক্টোবর',
  'নভেম্বর',
  'ডিসেম্বর',
];

export function toBnDigits(input: string | number | null | undefined): string {
  if (input === null || input === undefined) return '';
  return String(input).replace(/\d/g, (d) => BN_DIGITS[Number(d)] ?? d);
}

export function bnToEnDigits(input: string | number | null | undefined): string {
  if (input === null || input === undefined) return '';
  return String(input)
    .normalize('NFC')
    .replace(/[০-৯]/g, (ch) => {
      const idx = BN_DIGITS.indexOf(ch);
      return idx >= 0 ? String(idx) : ch;
    });
}

export function parseNum(val: unknown): number {
  if (typeof val === 'number') return Number.isFinite(val) ? val : 0;
  const stripped = bnToEnDigits(String(val ?? '')).replace(/[,\s৳]/g, '');
  const m = stripped.match(/-?\d+(?:\.\d+)?/) || stripped.match(/(?:^|[^a-zA-Z])(-?\.\d+)/);
  if (!m) return 0;
  const token = m[1] ?? m[0];
  const n = Number(token);
  return Number.isFinite(n) ? n : 0;
}

export function round3(n: number): number {
  return Math.round((Number(n) || 0) * 1000) / 1000;
}

export function round2(n: number): number {
  return Math.round((Number(n) || 0) * 100) / 100;
}

export function formatBnQty(val: number | null | undefined): string {
  const n = Number(val) || 0;
  return toBnDigits(n.toFixed(3));
}

export function formatBnMoney(val: number | null | undefined): string {
  const n = Number(val) || 0;
  return toBnDigits(
    n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  );
}

export function formatBnInt(val: number | null | undefined): string {
  const n = Math.round(Number(val) || 0);
  return toBnDigits(n);
}

export function todayIsoDate(): string {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

export function formatBnDate(isoOrRaw: string | null | undefined): string {
  const s = bnToEnDigits(String(isoOrRaw || '')).trim();
  if (!s) return '';
  const mIso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (mIso) {
    return toBnDigits(
      `${mIso[3].padStart(2, '0')}/${mIso[2].padStart(2, '0')}/${mIso[1]}`
    );
  }
  const mDmy = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (mDmy) {
    return toBnDigits(
      `${mDmy[1].padStart(2, '0')}/${mDmy[2].padStart(2, '0')}/${mDmy[3]}`
    );
  }
  return toBnDigits(s);
}

export function parseMonthYear(dateStr: string, preferMdy = false): { year: number; month: number } | null {
  const s = bnToEnDigits(String(dateStr || '')).trim();
  if (/^\d{5}(?:\.\d+)?$/.test(s)) {
    const serial = Number(s);
    if (Number.isFinite(serial) && serial > 20000 && serial < 80000) {
      const wholeDays = Math.floor(serial + 1e-4);
      const d = new Date((wholeDays - 25569) * 86400 * 1000);
      if (!Number.isNaN(d.getTime())) {
        return { year: d.getUTCFullYear(), month: d.getUTCMonth() };
      }
    }
  }
  const mIso = s.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:$|[^\d])/);
  if (mIso) {
    const yr = Number(mIso[1]);
    const mon = Number(mIso[2]);
    if (mon >= 1 && mon <= 12) return { year: yr, month: mon - 1 };
  }
  const mDmy = s.match(/^(\d{1,2})([/-])(\d{1,2})\2(\d{4})(?:$|[^\d])/);
  if (mDmy) {
    const a = Number(mDmy[1]);
    const b = Number(mDmy[3]);
    const yr = Number(mDmy[4]);
    const mon = a > 12 ? b : b > 12 ? a : preferMdy ? a : b;
    if (mon >= 1 && mon <= 12) return { year: yr, month: mon - 1 };
  }
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return { year: d.getFullYear(), month: d.getMonth() };
  }
  return null;
}

export function fiscalYearLabel(year: number, monthIndex0: number): string {
  const startYr = monthIndex0 >= 6 ? year : year - 1;
  return toBnDigits(`${startYr}-${String((startYr + 1) % 100).padStart(2, '0')}`);
}

export function fiscalYearBounds(year: number, monthIndex0: number): {
  fyStartYear: number;
  fyStartMonth0: number;
} {
  const fyStartYear = monthIndex0 >= 6 ? year : year - 1;
  return { fyStartYear, fyStartMonth0: 6 };
}

export function isMonthInFiscalUpTo(
  entryYear: number,
  entryMonth0: number,
  targetYear: number,
  targetMonth0: number
): 'before' | 'current' | 'after' | 'outside' {
  const { fyStartYear } = fiscalYearBounds(targetYear, targetMonth0);
  const fyEndYear = fyStartYear + 1;
  const entryKey = entryYear * 12 + entryMonth0;
  const startKey = fyStartYear * 12 + 6;
  const endKey = fyEndYear * 12 + 5;
  const targetKey = targetYear * 12 + targetMonth0;
  if (entryKey < startKey || entryKey > endKey) return 'outside';
  if (entryKey < targetKey) return 'before';
  if (entryKey === targetKey) return 'current';
  return 'after';
}

export function uid(prefix: string): string {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

export function normMemoNo(input: string | number | null | undefined): string {
  return bnToEnDigits(String(input ?? ''))
    .trim()
    .replace(/\s+/g, '')
    .toLowerCase();
}

export function normChallanNo(input: string | number | null | undefined): string {
  let s = bnToEnDigits(String(input ?? ''))
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .replace(/^'+/, '')
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, '')
    .toLowerCase();
  if (!s) return '';
  if (
    /^[-_.]+$/.test(s) ||
    s === 'n/a' ||
    s === 'na' ||
    s === 'none' ||
    s === 'nil' ||
    s === 'null' ||
    s === 'undefined' ||
    s === 'নেই' ||
    s === 'নাই'
  ) {
    return '';
  }
  const floatInt = s.match(/^(\d+)\.0+$/);
  if (floatInt) s = floatInt[1];
  if (/^0+\d+$/.test(s)) {
    s = s.replace(/^0+(?=\d)/, '');
  }
  if (/^0+$/.test(s)) return '';
  return s;
}

export function normDoNo(input: string | number | null | undefined): string {
  let s = bnToEnDigits(String(input ?? ''))
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .replace(/^'+/, '')
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, '')
    .toLowerCase();
  if (!s) return '';
  if (/^[\d,]+$/.test(s)) {
    s = s.replace(/,/g, '');
  }
  const floatInt = s.match(/^(\d+)\.0+$/);
  if (floatInt) s = floatInt[1];
  if (/^0+\d+$/.test(s)) {
    s = s.replace(/^0+(?=\d)/, '');
  }
  return s;
}

export function isFireServiceLabel(input: string | null | undefined): boolean {
  const s = bnToEnDigits(String(input ?? ''))
    .normalize('NFC')
    .trim()
    .toLowerCase();
  if (!s) return false;
  const hasFire = s.includes('ফায়ার') || s.includes('ফায়ার') || s.includes('fire');
  const hasService =
    s.includes('সার্ভিস') ||
    s.includes('service') ||
    s.includes('সিভিল') ||
    s.includes('civil') ||
    s.includes('স্টেশন') ||
    s.includes('station');
  return hasFire && hasService;
}

export function normPersonName(input: string | null | undefined): string {
  let s = bnToEnDigits(String(input ?? ''))
    .normalize('NFC')
    .replace(/\u09AF\u09BC/g, '\u09DF')
    .replace(/\u09A1\u09BC/g, '\u09DC')
    .replace(/\u09A2\u09BC/g, '\u09DD')
    .trim()
    .replace(/[.,।_]+/g, ' ')
    .replace(/\s+/g, ' ');
  for (let pass = 0; pass < 3; pass++) {
    s = s
      .replace(
        /^(জনাব|মেসার্স|মেসাস|মোঃ|মো:|মোহঃ|মোহ:|মোছাঃ|মোছা:|মুঃ|মুহাঃ|মুহাম্মাদ|মুহম্মদ|মুহাম্মদ|মোহাম্মাদ|মোহম্মদ|মোহাম্মদ|মহাঃ|শ্রী|হাজী|হাজি|আলহাজ্ব|আলহাজ|এডভোকেট|ডাঃ|md|mohd|mohammad|muhammad|mohammed|muhammed|mohamad|muhamad|mst|mr|mrs|ms|messrs|janab|hazi|haji|alhaj|shri)(?=\s+\S)/i,
        ''
      )
      .replace(/^মো(?=\s+\S)/, '')
      .trim();
  }
  return s
    .replace(/[-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function splitPartyAliases(raw: string | null | undefined): string[] {
  const s = bnToEnDigits(String(raw ?? ''))
    .normalize('NFC')
    .replace(/\u09AF\u09BC/g, '\u09DF')
    .replace(/\u09A1\u09BC/g, '\u09DC')
    .replace(/\u09A2\u09BC/g, '\u09DD')
    .trim();
  if (!s) return [];
  const parts = s
    .split(/[,/|]+/)
    .map((p) => p.trim())
    .filter(Boolean);
  const out: string[] = [];
  const seen = new Set<string>();
  const addAlias = (val: string) => {
    const clean = val.trim();
    if (!clean) return;
    const key = clean.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      out.push(clean);
    }
  };
  for (const p of parts) {
    addAlias(p);
    // Also support name without parenthetical nickname (e.g. "SAROAR HOSSAIN ( SANTU)" -> "SAROAR HOSSAIN")
    const noParens = p.replace(/\([^)]*\)/g, ' ').replace(/\s+/g, ' ').trim();
    if (noParens && noParens !== p) {
      addAlias(noParens);
    }
  }
  return out;
}

const BN_TO_LATIN_VOWELS: Record<string, string> = {
  'অ': 'o',
  'আ': 'a',
  'ই': 'i',
  'ঈ': 'i',
  'উ': 'u',
  'ঊ': 'u',
  'ঋ': 'ri',
  'এ': 'e',
  'ঐ': 'oi',
  'ও': 'o',
  'ঔ': 'ou',
  'া': 'a',
  'ি': 'i',
  'ী': 'i',
  'ু': 'u',
  'ূ': 'u',
  'ৃ': 'ri',
  'ে': 'e',
  'ৈ': 'oi',
  'ো': 'o',
  'ৌ': 'ou',
};

const BN_TO_LATIN_CONS: Record<string, string> = {
  'ক': 'k',
  'খ': 'kh',
  'গ': 'g',
  'ঘ': 'gh',
  'ঙ': 'ng',
  'চ': 'ch',
  'ছ': 's',
  'জ': 'j',
  'ঝ': 'jh',
  'ঞ': 'n',
  'ট': 't',
  'ঠ': 'th',
  'ড': 'd',
  'ঢ': 'dh',
  'ণ': 'n',
  'ত': 't',
  'থ': 'th',
  'দ': 'd',
  'ধ': 'dh',
  'ন': 'n',
  'প': 'p',
  'ফ': 'f',
  'ব': 'b',
  'ভ': 'bh',
  'ম': 'm',
  'য': 'j',
  'র': 'r',
  'ল': 'l',
  'শ': 's',
  'ষ': 's',
  'স': 's',
  'হ': 'h',
  'ড়': 'r',
  'ড়': 'r',
  'ঢ়': 'rh',
  'ঢ়': 'rh',
  'য়': 'y',
  'য়': 'y',
  'ৎ': 't',
  'ং': 'ng',
  'ঃ': '',
  'ঁ': '',
};

function bnWordToLatin(word: string): string {
  let res = '';
  for (let i = 0; i < word.length; i++) {
    const ch = word[i];
    if (ch === '\u09CD') {
      const next = word[i + 1];
      if (next === 'য') {
        if (word[i + 2] === 'া') {
          res += 'a';
          i += 2;
        } else {
          i += 1;
        }
      }
      continue;
    }
    if (BN_TO_LATIN_VOWELS[ch] !== undefined) {
      res += BN_TO_LATIN_VOWELS[ch];
      continue;
    }
    if (BN_TO_LATIN_CONS[ch] !== undefined) {
      res += BN_TO_LATIN_CONS[ch];
      const next = word[i + 1];
      // Bare Bengali consonant immediately before independent vowel (e.g. ম+ই in মইনুল) carries inherent 'o'
      if (next && 'ইঈউঊও'.includes(next)) {
        res += 'o';
      }
      continue;
    }
    if (/[a-z0-9]/.test(ch)) {
      res += ch;
    }
  }
  return res;
}

function normalizeLatinToken(rawTok: string): { full: string; skel: string; highVowels: string } {
  const hasBn = /[\u0980-\u09FF]/.test(rawTok);
  let t = (hasBn ? bnWordToLatin(rawTok) : rawTok.toLowerCase()).replace(/[^a-z0-9]/g, '');

  // Harmonize Bengali/English phonetic digraphs and spellings
  t = t
    .replace(/ph/g, 'f')
    .replace(/bh/g, 'v')
    .replace(/sh/g, 's')
    .replace(/z/g, 'j')
    .replace(/q/g, 'k')
    .replace(/w/g, 'o')
    .replace(/ee/g, 'i')
    .replace(/oo/g, 'u')
    .replace(/ie/g, 'i')
    .replace(/ey$/g, 'i')
    .replace(/y$/g, 'i')
    .replace(/(^|[bcdfghjklmnpqrstvx])my/g, '$1mai')
    .replace(/iy/g, 'i')
    .replace(/ya/g, 'ia')
    .replace(/ye/g, 'ie')
    .replace(/y/g, 'i')
    .replace(/c(?!h)/g, 's')
    .replace(/ch/g, 's')
    .replace(/ao|au|ou/g, 'o')
    .replace(/ai|ei|oi/g, 'e')
    .replace(/([a-z])\1+/g, '$1');

  // Remove vowel noise around h-digraphs / trailing silent letters cleanly in consonant skeleton
  const skel = t
    .replace(/([kgtdpb])h/g, '$1')
    .replace(/[aeiou]/g, '')
    .replace(/([a-z])\1+/g, '$1');

  // Track primary high/front vowels (i/e/u) inside token to separate pairs like JAHIDUL vs ZAHEDUL
  const highVowels = t.replace(/[^ieu]/g, '');
  return { full: t, skel, highVowels };
}

function editDist(a: string, b: string): number {
  if (a === b) return 0;
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = new Array<number>(n + 1);
  for (let j = 0; j <= n; j++) dp[j] = j;
  for (let i = 1; i <= m; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const temp = dp[j];
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + cost);
      prev = temp;
    }
  }
  return dp[n];
}

function tokensPhoneticallyMatch(rawA: string, rawB: string): boolean {
  const a = normalizeLatinToken(rawA);
  const b = normalizeLatinToken(rawB);
  if (!a.full || !b.full) return false;
  if (a.full === b.full) return true;

  // Compare consonant skeletons first
  if (!a.skel || !b.skel) return false;
  const skelDist = editDist(a.skel, b.skel);
  const minSkel = Math.min(a.skel.length, b.skel.length);

  // Skeleton must be identical, or at most 1 edit when >= 4 consonants
  if (skelDist > 1) return false;
  if (skelDist === 1 && minSkel < 4) return false;

  // If skeletons match identically, guard against real different names like JAHIDUL vs ZAHEDUL
  if (skelDist === 0) {
    if (a.highVowels && b.highVowels && a.highVowels !== b.highVowels) {
      // Allow trailing i/e difference on long surnames or known pairs like hossain/hosen, howladar/hawlader
      // Reject stem vowel mismatch when lengths match closely and both have same # of high vowels
      if (a.highVowels.length === b.highVowels.length && a.highVowels.length >= 2) {
        // e.g. JAHIDUL (iu) vs ZAHEDUL (eu) -> different first high vowel!
        if (a.highVowels[0] !== b.highVowels[0]) return false;
      } else if (a.highVowels.length === 1 && b.highVowels.length === 1) {
        // e.g. JAHID (i) vs ZAHED (e) -> 100% different name!
        return false;
      }
    }
    return true;
  }

  // skelDist === 1 and minSkel >= 4 (e.g. khndkr vs khdkr, or slight transliteration drop)
  const fullDist = editDist(a.full, b.full);
  return fullDist <= 2 && Math.min(a.full.length, b.full.length) >= 5;
}

function partyTokensOf(raw: string): string[] {
  const n = normPersonName(raw);
  if (!n) return [];
  return n
    .replace(/[()]/g, ' ')
    .split(/\s+/)
    .map((t) => t.trim())
    .filter((t) => t.length > 0);
}

function singleAliasMatch(a: string, b: string): boolean {
  const na = normPersonName(a);
  const nb = normPersonName(b);
  if (!na || !nb) return false;
  if (na === nb) return true;

  const toksA = partyTokensOf(a);
  const toksB = partyTokensOf(b);
  if (toksA.length === 0 || toksB.length === 0) return false;

  // Compound join match without spaces (e.g. "নুরুল ইসলাম" vs "নুরুলইসলাম")
  const joinA = normalizeLatinToken(toksA.join(''));
  const joinB = normalizeLatinToken(toksB.join(''));
  if (joinA.full && joinA.full === joinB.full && joinA.full.length >= 5) {
    return true;
  }

  // Strictly require identical non-honorific token count so "MD. JAMAL" NEVER matches "MD. JAMAL HAWLADER"
  if (toksA.length !== toksB.length) return false;

  return toksA.every((ta, idx) => tokensPhoneticallyMatch(ta, toksB[idx]));
}

export function namesLikelyMatch(a: string | null | undefined, b: string | null | undefined): boolean {
  const aliasesA = splitPartyAliases(a);
  const aliasesB = splitPartyAliases(b);
  if (aliasesA.length === 0 || aliasesB.length === 0) return false;
  for (const alA of aliasesA) {
    for (const alB of aliasesB) {
      const sa = bnToEnDigits(alA).normalize('NFC').trim().replace(/\s+/g, ' ').toLowerCase();
      const sb = bnToEnDigits(alB).normalize('NFC').trim().replace(/\s+/g, ' ').toLowerCase();
      if (sa === sb) return true;
      if (isFireServiceLabel(sa) && isFireServiceLabel(sb)) return true;
      const na = normPersonName(sa);
      const nb = normPersonName(sb);
      if (na && nb && na === nb) return true;
    }
  }
  return false;
}

export function namesFuzzyMatch(a: string | null | undefined, b: string | null | undefined): boolean {
  if (namesLikelyMatch(a, b)) return true;
  const aliasesA = splitPartyAliases(a);
  const aliasesB = splitPartyAliases(b);
  for (const alA of aliasesA) {
    for (const alB of aliasesB) {
      if (singleAliasMatch(alA, alB)) return true;
    }
  }
  return false;
}

export function unionsLikelyMatch(a: string | null | undefined, b: string | null | undefined): boolean {
  const clean = (s: string | null | undefined) =>
    bnToEnDigits(String(s ?? ''))
      .normalize('NFC')
      .trim()
      .toLowerCase()
      .replace(/(ইউনিয়ন|ইউনিয়ন|ইউপি|পৌরসভা|উপজেলা|সদর|union|pourashava|municipality|upazila)/g, '')
      .replace(/[.,।\-_\s]+/g, '');
  const ca = clean(a);
  const cb = clean(b);
  if (!ca || !cb) return false;
  return ca === cb || (Math.min(ca.length, cb.length) >= 3 && (ca.includes(cb) || cb.includes(ca)));
}

export function memosLikelyMatch(
  a: string | number | null | undefined,
  b: string | number | null | undefined
): boolean {
  const na = normMemoNo(a);
  const nb = normMemoNo(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  const shorter = na.length <= nb.length ? na : nb;
  const longer = na.length <= nb.length ? nb : na;
  if (shorter.length >= 2) {
    if (
      longer.endsWith(`-${shorter}`) ||
      longer.endsWith(`/${shorter}`) ||
      longer.endsWith(`.${shorter}`) ||
      longer.endsWith(`–${shorter}`)
    ) {
      return true;
    }
  }
  return false;
}

export function toIsoDateStr(raw: string | null | undefined, preferMdy = false): string {
  const s = bnToEnDigits(String(raw ?? '')).trim();
  if (!s) return '';
  if (/^\d{5}(?:\.\d+)?$/.test(s)) {
    const serial = Number(s);
    if (Number.isFinite(serial) && serial > 20000 && serial < 80000) {
      const wholeDays = Math.floor(serial + 1e-4);
      const d = new Date((wholeDays - 25569) * 86400 * 1000);
      if (!Number.isNaN(d.getTime())) {
        return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
      }
    }
  }
  const mIso = s.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:$|[^\d])/);
  if (mIso) {
    const mon = Number(mIso[2]);
    const day = Number(mIso[3]);
    if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) {
      return `${mIso[1]}-${String(mon).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }
  const mDmy = s.match(/^(\d{1,2})([/-])(\d{1,2})\2(\d{4})(?:$|[^\d])/);
  if (mDmy) {
    const a = Number(mDmy[1]);
    const b = Number(mDmy[3]);
    const yr = mDmy[4];
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
    if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) {
      return `${yr}-${String(mon).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }
  const parsed = new Date(s);
  if (!Number.isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return '';
}

export function isChallanOnOrBeforeDo(
  challanDate: string | null | undefined,
  doIssueDate: string | null | undefined,
  requireChallanDate = true
): boolean {
  const cIso = toIsoDateStr(challanDate);
  const dIso = toIsoDateStr(doIssueDate);
  if (!dIso) return true;
  if (!cIso) return !requireChallanDate;
  return cIso <= dIso;
}

export function isSamePartyName(
  a: string | null | undefined,
  b: string | null | undefined
): boolean {
  return namesLikelyMatch(a, b) || namesFuzzyMatch(a, b);
}

const BANGABDA_MONTHS = [
  'বৈশাখ',
  'জ্যৈষ্ঠ',
  'আষাঢ়',
  'শ্রাবণ',
  'ভাদ্র',
  'আশ্বিন',
  'কার্তিক',
  'অগ্রহায়ণ',
  'পৌষ',
  'মাঘ',
  'ফাল্গুন',
  'চৈত্র',
];

function isGregorianLeapYear(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

/**
 * Converts a Gregorian date (ISO YYYY-MM-DD or DD/MM/YYYY) into the official
 * revised Bangladesh Bangabda calendar date string (e.g. "০৫ অগ্রহায়ণ ১৪৩৩").
 */
export function formatBangabdaDate(isoOrRaw: string | null | undefined): string {
  const iso = toIsoDateStr(isoOrRaw);
  if (!iso) return '';
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  const gYear = Number(m[1]);
  const gMonth = Number(m[2]); // 1..12
  const gDay = Number(m[3]);   // 1..31

  const beforePahelaBaishakh = gMonth < 4 || (gMonth === 4 && gDay < 14);
  const bYear = beforePahelaBaishakh ? gYear - 594 : gYear - 593;
  const baishakhGregYear = bYear + 593;
  const falgunGregYear = bYear + 594;

  const targetUtc = Date.UTC(gYear, gMonth - 1, gDay);
  const baishakhUtc = Date.UTC(baishakhGregYear, 3, 14);
  let remainingDays = Math.floor((targetUtc - baishakhUtc) / 86400000);
  if (remainingDays < 0) return '';

  const monthLengths = [
    31,
    31,
    31,
    31,
    31,
    31,
    30,
    30,
    30,
    30,
    isGregorianLeapYear(falgunGregYear) ? 30 : 29,
    30,
  ];

  let bMonthIdx = 0;
  while (bMonthIdx < 11 && remainingDays >= monthLengths[bMonthIdx]) {
    remainingDays -= monthLengths[bMonthIdx];
    bMonthIdx++;
  }
  const bDay = remainingDays + 1;
  const dayStr = toBnDigits(String(bDay).padStart(2, '0'));
  const monthStr = BANGABDA_MONTHS[bMonthIdx] || '';
  const yearStr = toBnDigits(bYear);
  return `${dayStr} ${monthStr} ${yearStr}`;
}

/**
 * Computes a default deadline ISO date (YYYY-MM-DD) for allocation orders:
 * last day of the target month, or +7 days from ownDate if less than 5 days remain.
 */
export function autoDeadlineIsoDate(
  ownDateIso: string | null | undefined,
  year: number,
  month0: number
): string {
  const endDay = new Date(Date.UTC(year, month0 + 1, 0));
  const endIso = `${endDay.getUTCFullYear()}-${String(endDay.getUTCMonth() + 1).padStart(2, '0')}-${String(endDay.getUTCDate()).padStart(2, '0')}`;
  const ownIso = toIsoDateStr(ownDateIso);
  if (!ownIso) return endIso;
  const ownMatch = ownIso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!ownMatch) return endIso;
  const ownUtc = Date.UTC(Number(ownMatch[1]), Number(ownMatch[2]) - 1, Number(ownMatch[3]));
  const endUtc = endDay.getTime();
  const diffDays = Math.floor((endUtc - ownUtc) / 86400000);
  if (diffDays >= 5) return endIso;
  const plus7 = new Date(ownUtc + 7 * 86400000);
  return `${plus7.getUTCFullYear()}-${String(plus7.getUTCMonth() + 1).padStart(2, '0')}-${String(plus7.getUTCDate()).padStart(2, '0')}`;
}
