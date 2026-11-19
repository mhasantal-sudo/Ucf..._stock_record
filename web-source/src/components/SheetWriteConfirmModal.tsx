import React from 'react';
import { AlertTriangle, CloudUpload, FileSpreadsheet, ShieldCheck, X } from 'lucide-react';
import type { SheetWriteConfirmRequest } from '../types/grain';
import { toBnDigits } from '../shared/bnUtils';

interface SheetWriteConfirmModalProps {
  request: SheetWriteConfirmRequest | null;
  isBusy?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}

export const SheetWriteConfirmModal: React.FC<SheetWriteConfirmModalProps> = ({
  request,
  isBusy = false,
  onConfirm,
  onCancel,
}) => {
  if (!request) return null;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={request.title}
      className="no-print fixed inset-0 z-[70] bg-slate-900/55 backdrop-blur-xs flex items-center justify-center p-4"
      onClick={() => {
        if (!isBusy) onCancel();
      }}
    >
      <div
        className="w-full max-w-lg rounded-3xl bg-white shadow-2xl border border-rose-200 overflow-hidden font-nikosh"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-3 px-5 py-4 bg-gradient-to-r from-rose-50 via-white to-amber-50/70 border-b border-rose-100">
          <div className="flex items-center gap-2.5 min-w-0">
            <span
              className={`h-10 w-10 rounded-2xl flex items-center justify-center shrink-0 ${
                request.isDestructive
                  ? 'bg-amber-100 text-amber-800'
                  : 'bg-[#E2136E]/10 text-[#E2136E]'
              }`}
            >
              {request.isDestructive ? (
                <AlertTriangle className="h-5 w-5" />
              ) : (
                <CloudUpload className="h-5 w-5" />
              )}
            </span>
            <div className="min-w-0">
              <h3 className="text-sm sm:text-base font-bold text-slate-900 leading-snug">
                {request.title}
              </h3>
              <p className="text-[11px] font-semibold text-[#E2136E] flex items-center gap-1 mt-0.5 truncate">
                <FileSpreadsheet className="h-3.5 w-3.5 shrink-0" />
                <span className="truncate">{request.spreadsheetName}</span>
              </p>
            </div>
          </div>

          <button
            type="button"
            disabled={isBusy}
            onClick={onCancel}
            aria-label="বন্ধ করুন"
            className="h-8 w-8 rounded-full bg-white text-slate-500 hover:bg-slate-100 hover:text-slate-800 flex items-center justify-center shrink-0 cursor-pointer disabled:opacity-50"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="p-5 space-y-3.5 text-xs text-slate-700 max-h-[70vh] overflow-y-auto">
          <p className="leading-relaxed text-slate-800 font-medium">{request.description}</p>

          {request.details.length > 0 && (
            <div className="rounded-2xl border border-slate-200 bg-slate-50/80 p-3 space-y-1.5">
              <div className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                <ShieldCheck className="h-3.5 w-3.5 text-emerald-700" />
                <span>যেসকল ডেটা Google Sheet-এ হালনাগাদ হবে:</span>
              </div>
              <ul className="space-y-1 text-xs font-semibold text-slate-800 pl-1">
                {request.details.map((line, idx) => (
                  <li key={idx} className="flex items-start gap-1.5">
                    <span className="text-[#E2136E] mt-0.5">•</span>
                    <span>{line}</span>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {request.targetSheets.length > 0 && (
            <div className="space-y-1.5">
              <div className="text-[11px] font-bold text-slate-600">
                সংশ্লিষ্ট Google Sheet ট্যাবসমূহ ({toBnDigits(request.targetSheets.length)}টি):
              </div>
              <div className="flex flex-wrap gap-1.5">
                {request.targetSheets.map((tab) => (
                  <code
                    key={tab}
                    className="rounded-lg border border-slate-200 bg-slate-50 px-2 py-0.5 text-[10.5px] font-semibold text-slate-800"
                  >
                    {tab}
                  </code>
                ))}
              </div>
            </div>
          )}
        </div>

        <div className="flex flex-wrap items-center justify-end gap-2.5 px-5 py-3.5 bg-slate-50 border-t border-slate-200">
          <button
            type="button"
            disabled={isBusy}
            onClick={onCancel}
            className="min-h-[42px] rounded-xl border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-50 px-4 py-2 text-xs font-bold text-slate-700 cursor-pointer"
          >
            বাতিল করুন
          </button>
          <button
            type="button"
            disabled={isBusy}
            onClick={onConfirm}
            className="min-h-[42px] inline-flex items-center gap-1.5 rounded-xl bg-[#E2136E] hover:bg-[#B80D58] disabled:opacity-50 px-4 py-2 text-xs font-bold text-white shadow-xs cursor-pointer"
          >
            <CloudUpload className="h-4 w-4" />
            <span>{isBusy ? 'শিটে সেভ হচ্ছে...' : request.confirmLabel || 'নিশ্চিত ও সেভ করুন'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};


