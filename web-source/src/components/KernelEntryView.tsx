import React, { useMemo, useState } from 'react';
import { ArrowLeft, BookOpen, Pencil, Trash2, Wheat } from 'lucide-react';
import {
  CATEGORIES,
  CATEGORY_BY_NAME,
  GRAIN_OPTIONS,
  type AllocationCategory,
  type AllocationDatabase,
  type GrainType,
} from '../types/allocation';
import {
  formatBnDate,
  formatBnInt,
  formatBnQty,
  parseNum,
  toBnDigits,
  todayIsoDate,
} from '../shared/bnUtils';
import type { ActionResult } from '../store/useAllocationStore';
import type { EntryStageKey } from './CategoryEntryView';

interface KernelEntryViewProps {
  db: AllocationDatabase;
  onBack: () => void;
  getMinistryUsage: (
    cat: AllocationCategory,
    memo: string
  ) => { total: number; used: number; remaining: number };
  onAddMinistry: (input: {
    category: AllocationCategory;
    memoNo: string;
    memoDate: string;
    allocationQty: number;
    grain: GrainType;
    phase?: number;
    comment?: string;
  }) => ActionResult;
  onUpdateMinistry?: (
    id: string,
    patch: {
      memoNo: string;
      memoDate: string;
      allocationQty: number;
      grain: GrainType;
      phase?: number;
      comment?: string;
    }
  ) => ActionResult;
  onDeleteMinistry: (id: string) => void;
  onJumpToCategory: (cat: AllocationCategory, stage: EntryStageKey, memoNo: string) => void;
  onOpenKernelRegister?: () => void;
}

const KERNEL_QUICK_CATEGORIES: AllocationCategory[] = ['খাদ্য বান্ধব', 'ভিডাব্লিউভি'];

export const KernelEntryView: React.FC<KernelEntryViewProps> = ({
  db,
  onBack,
  getMinistryUsage,
  onAddMinistry,
  onUpdateMinistry,
  onDeleteMinistry,
  onJumpToCategory,
  onOpenKernelRegister,
}) => {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [category, setCategory] = useState<AllocationCategory>('খাদ্য বান্ধব');
  const [memoNo, setMemoNo] = useState('');
  const [memoDate, setMemoDate] = useState(todayIsoDate());
  const [allocationQty, setAllocationQty] = useState('');
  const [grain, setGrain] = useState<GrainType>('কার্নেল');
  const [phase, setPhase] = useState('1');
  const [comment, setComment] = useState('');
  const [feedback, setFeedback] = useState<{ ok: boolean; text: string } | null>(null);

  const kernelList = useMemo(() => {
    return db.ministryAllocations
      .filter((m) => m.grain === 'কার্নেল')
      .map((m) => ({
        item: m,
        meta: CATEGORY_BY_NAME[m.category],
        usage: getMinistryUsage(m.category, m.memoNo),
      }));
  }, [db.ministryAllocations, getMinistryUsage]);

  const totalKernelQty = useMemo(
    () => kernelList.reduce((sum, row) => sum + (Number(row.item.allocationQty) || 0), 0),
    [kernelList]
  );

  const resetForm = () => {
    setEditingId(null);
    setMemoNo('');
    setMemoDate(todayIsoDate());
    setAllocationQty('');
    setGrain('কার্নেল');
    setPhase('1');
    setComment('');
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const catMeta = CATEGORY_BY_NAME[category];
    const qtyNum = parseNum(allocationQty);

    if (editingId && onUpdateMinistry) {
      const res = onUpdateMinistry(editingId, {
        memoNo,
        memoDate,
        allocationQty: qtyNum,
        grain,
        phase: catMeta?.hasPhase ? Number(phase) || 1 : undefined,
        comment: comment.trim() || undefined,
      });
      setFeedback({
        ok: res.ok,
        text: res.ok ? res.message || 'কার্নেল বরাদ্দ হালনাগাদ হয়েছে।' : res.error || 'ত্রুটি হয়েছে।',
      });
      if (res.ok) resetForm();
      return;
    }

    const res = onAddMinistry({
      category,
      memoNo,
      memoDate,
      allocationQty: qtyNum,
      grain,
      phase: catMeta?.hasPhase ? Number(phase) || 1 : undefined,
      comment: comment.trim() || undefined,
    });
    setFeedback({
      ok: res.ok,
      text: res.ok ? res.message || 'কার্নেল বরাদ্দ সংরক্ষিত হয়েছে।' : res.error || 'ত্রুটি হয়েছে।',
    });
    if (res.ok) resetForm();
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2 rounded-2xl border border-emerald-200 bg-emerald-50/70 px-4 py-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <button
            type="button"
            onClick={onBack}
            className="min-h-[40px] inline-flex items-center gap-1.5 rounded-xl border border-slate-300 bg-white hover:bg-slate-50 px-3 py-1.5 text-xs font-bold text-slate-800 cursor-pointer shrink-0"
          >
            <ArrowLeft className="h-4 w-4" />
            <span>ফিরে যান</span>
          </button>
          <div className="min-w-0">
            <h2 className="text-sm sm:text-base font-extrabold text-slate-900 flex items-center gap-1.5 truncate">
              <Wheat className="h-4 w-4 text-emerald-700 shrink-0" />
              <span>কার্নেল বরাদ্দ এন্ট্রি (মন্ত্রণালয় পর্যায়)</span>
            </h2>
            <p className="text-[11px] text-slate-600">
              খাদ্য বান্ধব, ভিডাব্লিউভি বা যেকোনো খাতের কার্নেল বরাদ্দের স্মারক, তারিখ ও পরিমাণ (মে.টন) এন্ট্রি করুন।
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0">
          {onOpenKernelRegister && (
            <button
              type="button"
              onClick={onOpenKernelRegister}
              className="min-h-[38px] inline-flex items-center gap-1.5 rounded-xl bg-emerald-700 hover:bg-emerald-800 px-3 py-1.5 text-xs font-bold text-white cursor-pointer shadow-xs"
            >
              <BookOpen className="h-3.5 w-3.5" />
              <span>কার্নেল রেজিস্টার</span>
            </button>
          )}
          <div className="inline-flex items-center gap-2 rounded-xl border border-emerald-300 bg-white px-3 py-1.5 text-xs font-bold text-emerald-900 tabular-nums">
            <span>মোট কার্নেল: {formatBnQty(totalKernelQty)} মে.টন</span>
            <span className="text-slate-300">|</span>
            <span>{formatBnInt(kernelList.length)}টি স্মারক</span>
          </div>
        </div>
      </div>

      {feedback && (
        <div
          className={`rounded-xl border px-3.5 py-2.5 text-xs font-bold ${
            feedback.ok
              ? 'border-emerald-300 bg-emerald-50 text-emerald-900'
              : 'border-rose-300 bg-rose-50 text-rose-900'
          }`}
        >
          {feedback.text}
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
        <form
          onSubmit={handleSubmit}
          className="lg:col-span-5 rounded-2xl border border-slate-200 bg-white p-4 space-y-3"
        >
          <div className="flex items-center justify-between gap-2 text-sm font-bold text-slate-900">
            <span>{editingId ? 'কার্নেল বরাদ্দ সম্পাদনা (এডিট)' : 'নতুন কার্নেল বরাদ্দ এন্ট্রি'}</span>
            {editingId && (
              <button
                type="button"
                onClick={resetForm}
                className="rounded-lg border border-slate-300 bg-slate-50 px-2.5 py-1 text-[11px] font-bold text-slate-700 cursor-pointer"
              >
                বাতিল
              </button>
            )}
          </div>

          <div className="space-y-2.5 text-xs">
            <div>
              <label className="block font-bold text-slate-700 mb-1.5">খাত নির্বাচন করুন *</label>
              <div className="flex flex-wrap gap-1.5 mb-2">
                {KERNEL_QUICK_CATEGORIES.map((quickCat) => {
                  const active = category === quickCat;
                  return (
                    <button
                      key={quickCat}
                      type="button"
                      disabled={Boolean(editingId)}
                      onClick={() => setCategory(quickCat)}
                      className={`rounded-lg border px-2.5 py-1 text-[11px] font-bold transition-colors cursor-pointer disabled:opacity-60 ${
                        active
                          ? 'bg-emerald-700 text-white border-emerald-700'
                          : 'bg-emerald-50 text-emerald-900 border-emerald-200 hover:bg-emerald-100'
                      }`}
                    >
                      {quickCat}
                    </button>
                  );
                })}
              </div>
              <select
                value={category}
                disabled={Boolean(editingId)}
                onChange={(e) => setCategory(e.target.value as AllocationCategory)}
                className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white font-bold text-slate-900 disabled:bg-slate-100 focus:outline-none focus:border-[#E2136E]"
              >
                {CATEGORIES.map((c) => (
                  <option key={c.id} value={c.category}>
                    {c.category}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block font-bold text-slate-700 mb-1">মন্ত্রণালয়ের স্মারক নং *</label>
              <input
                type="text"
                required
                value={memoNo}
                onChange={(e) => setMemoNo(e.target.value)}
                placeholder="যেমন: ১৩.০১.০০০০.০৩১.৩৩"
                className="w-full rounded-lg border border-slate-300 px-3 py-2 focus:outline-none focus:border-[#E2136E]"
              />
            </div>

            <div className="grid grid-cols-2 gap-2.5">
              <div>
                <label className="block font-bold text-slate-700 mb-1">স্মারকের তারিখ *</label>
                <input
                  type="date"
                  required
                  value={memoDate}
                  onChange={(e) => setMemoDate(e.target.value)}
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 focus:outline-none focus:border-[#E2136E]"
                />
              </div>
              <div>
                <label className="block font-bold text-slate-700 mb-1">পরিমাণ (মে.টন) *</label>
                <input
                  type="number"
                  step="0.001"
                  required
                  value={allocationQty}
                  onChange={(e) => setAllocationQty(e.target.value)}
                  placeholder="০.০০০"
                  className="w-full rounded-lg border border-slate-300 px-3 py-2 focus:outline-none focus:border-[#E2136E]"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2.5">
              <div>
                <label className="block font-bold text-slate-700 mb-1">পণ্যের ক্যাটাগরি *</label>
                <select
                  value={grain}
                  onChange={(e) => setGrain(e.target.value as GrainType)}
                  className="w-full rounded-lg border border-emerald-400 px-3 py-2 bg-emerald-50/50 font-bold text-emerald-950 focus:outline-none focus:border-[#E2136E]"
                >
                  {GRAIN_OPTIONS.map((g) => (
                    <option key={g} value={g}>
                      {g}
                    </option>
                  ))}
                </select>
              </div>
              {CATEGORY_BY_NAME[category]?.hasPhase && (
                <div>
                  <label className="block font-bold text-slate-700 mb-1">পর্যায়</label>
                  <select
                    value={phase}
                    onChange={(e) => setPhase(e.target.value)}
                    className="w-full rounded-lg border border-slate-300 px-3 py-2 bg-white"
                  >
                    <option value="1">১ম পর্যায়</option>
                    <option value="2">২য় পর্যায়</option>
                    <option value="3">৩য় পর্যায়</option>
                    <option value="4">৪র্থ পর্যায়</option>
                  </select>
                </div>
              )}
            </div>

            <div>
              <label className="block font-bold text-slate-700 mb-1">মন্তব্য / ট্রাক / বিবরণ (ঐচ্ছিক)</label>
              <input
                type="text"
                value={comment}
                onChange={(e) => setComment(e.target.value)}
                placeholder="যেমন: কার্নেল বরাদ্দ / মাস / ট্রাক নং"
                className="w-full rounded-lg border border-slate-300 px-3 py-2 focus:outline-none focus:border-[#E2136E]"
              />
            </div>
          </div>

          <div className="flex items-center gap-2 pt-1">
            <button
              type="submit"
              className="flex-1 min-h-[44px] rounded-xl bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold cursor-pointer"
            >
              {editingId ? 'কার্নেল বরাদ্দ হালনাগাদ করুন' : 'কার্নেল বরাদ্দ সংরক্ষণ করুন'}
            </button>
            {editingId && (
              <button
                type="button"
                onClick={resetForm}
                className="min-h-[44px] rounded-xl border border-slate-300 bg-white hover:bg-slate-50 px-4 text-xs font-bold text-slate-700 cursor-pointer"
              >
                বাতিল
              </button>
            )}
          </div>
        </form>

        <div className="lg:col-span-7 rounded-2xl border border-slate-200 bg-white overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-200 text-xs font-bold text-slate-900 flex items-center justify-between">
            <span>সংরক্ষিত কার্নেল বরাদ্দ তালিকা ({formatBnInt(kernelList.length)}টি)</span>
            <span className="text-[11px] font-semibold text-emerald-800">
              পণ্য: কার্নেল
            </span>
          </div>
          <div className="overflow-x-auto max-h-[540px] overflow-y-auto">
            <table className="w-full text-xs border-collapse tabular-nums">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-700 font-bold sticky top-0">
                  <th className="py-2 px-3 text-left">খাত ও স্মারক</th>
                  <th className="py-2 px-3 text-center">পণ্য</th>
                  <th className="py-2 px-3 text-right">মোট</th>
                  <th className="py-2 px-3 text-right">অবশিষ্ট</th>
                  <th className="py-2 px-3 text-center">কাজ</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {kernelList.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="py-10 text-center text-slate-500">
                      এখনো কোনো কার্নেল বরাদ্দ এন্ট্রি করা হয়নি। বাম পাশের ফর্ম থেকে এন্ট্রি দিন।
                    </td>
                  </tr>
                ) : (
                  kernelList.map(({ item: m, meta: mMeta, usage }) => {
                    const hasRem = usage.remaining > 0.0005;
                    const nextStage: EntryStageKey = mMeta?.requiresDC
                      ? 'dc'
                      : mMeta?.flowType.startsWith('dealer') || mMeta?.flowType === 'ep'
                        ? 'ucf'
                        : 'uno';
                    const nextStageLabel = mMeta?.requiresDC
                      ? 'ডিসি বরাদ্দ'
                      : nextStage === 'ucf'
                        ? 'ইউসিএফ উপবরাদ্দ'
                        : 'ইউএনও উপবরাদ্দ';

                    return (
                      <tr
                        key={m.id}
                        className={editingId === m.id ? 'bg-emerald-50/70' : 'hover:bg-slate-50'}
                      >
                        <td className="py-2 px-3">
                          <div className="font-bold text-[#E2136E]">{m.category}</div>
                          <div className="font-semibold text-slate-900">{toBnDigits(m.memoNo)}</div>
                          <div className="text-[11px] text-slate-500">
                            {formatBnDate(m.memoDate)}
                            {m.phase ? ` · পর্যায়-${toBnDigits(m.phase)}` : ''}
                            {m.comment ? ` · ${m.comment}` : ''}
                          </div>
                        </td>
                        <td className="py-2 px-3 text-center">
                          <span className="inline-flex items-center rounded-full border border-emerald-300 bg-emerald-50 px-2 py-0.5 text-[11px] font-bold text-emerald-900">
                            {m.grain}
                          </span>
                        </td>
                        <td className="py-2 px-3 text-right font-semibold">{formatBnQty(usage.total)}</td>
                        <td
                          className={`py-2 px-3 text-right font-bold ${
                            hasRem ? 'text-emerald-700' : 'text-slate-400'
                          }`}
                        >
                          {hasRem ? formatBnQty(usage.remaining) : '০.০০০ (শেষ)'}
                        </td>
                        <td className="py-2 px-3 text-center whitespace-nowrap">
                          <div className="inline-flex items-center gap-1">
                            <button
                              type="button"
                              onClick={() => onJumpToCategory(m.category, nextStage, m.memoNo)}
                              className="rounded border border-emerald-300 bg-emerald-50 hover:bg-emerald-100 px-2 py-1 text-[11px] font-bold text-emerald-900 cursor-pointer"
                            >
                              {hasRem ? nextStageLabel : 'তালিকা দেখুন'}
                            </button>
                            {onUpdateMinistry && (
                              <button
                                type="button"
                                onClick={() => {
                                  setEditingId(m.id);
                                  setCategory(m.category);
                                  setMemoNo(m.memoNo);
                                  setMemoDate(m.memoDate);
                                  setAllocationQty(String(m.allocationQty));
                                  setGrain(m.grain);
                                  setPhase(String(m.phase || 1));
                                  setComment(m.comment || '');
                                }}
                                className="p-1.5 text-[#E2136E] hover:bg-rose-50 rounded cursor-pointer"
                                title="এডিট করুন"
                                aria-label="এডিট করুন"
                              >
                                <Pencil className="h-3.5 w-3.5" />
                              </button>
                            )}
                            <button
                              type="button"
                              onClick={() => {
                                if (editingId === m.id) resetForm();
                                onDeleteMinistry(m.id);
                              }}
                              className="p-1.5 text-rose-600 hover:bg-rose-50 rounded cursor-pointer"
                              title="মুছে ফেলুন"
                              aria-label="মুছে ফেলুন"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};
