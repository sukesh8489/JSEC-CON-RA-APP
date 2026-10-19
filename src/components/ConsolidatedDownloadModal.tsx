import React, { useState } from 'react';
import {
  ConsolidatedData,
  ExportDesignTheme,
} from '../utils/consolidatedExport';
import CollegeLogo from './CollegeLogo';
import {
  X,
  FileSpreadsheet,
  FileText,
  Download,
  CheckCircle2,
  BookOpen,
} from 'lucide-react';

interface ConsolidatedDownloadModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: ConsolidatedData | null;
  onDownloadExcel: (theme?: ExportDesignTheme) => Promise<void>;
  onDownloadPdf: (theme?: ExportDesignTheme) => Promise<void>;
  onDownloadWord: (theme?: ExportDesignTheme) => Promise<void>;
}

export default function ConsolidatedDownloadModal({
  isOpen,
  onClose,
  data,
  onDownloadExcel,
  onDownloadPdf,
  onDownloadWord,
}: ConsolidatedDownloadModalProps) {
  const [downloadingType, setDownloadingType] = useState<'excel' | 'pdf' | 'word' | null>(null);
  const [downloadSuccessNotice, setDownloadSuccessNotice] = useState<string | null>(null);

  if (!isOpen || !data) return null;

  const handleDownload = async (type: 'excel' | 'pdf' | 'word') => {
    try {
      setDownloadingType(type);
      if (type === 'excel') {
        await onDownloadExcel();
        setDownloadSuccessNotice('Consolidated Excel (.xlsx) file downloaded successfully!');
      } else if (type === 'pdf') {
        await onDownloadPdf();
        setDownloadSuccessNotice('Consolidated PDF report downloaded successfully!');
      } else if (type === 'word') {
        await onDownloadWord();
        setDownloadSuccessNotice('Consolidated Word (.docx) file downloaded successfully!');
      }
      setTimeout(() => setDownloadSuccessNotice(null), 4000);
    } catch (err) {
      console.error('Error downloading consolidated file:', err);
    } finally {
      setDownloadingType(null);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="consolidated-download-modal-title"
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-950/70 backdrop-blur-xs animate-in fade-in duration-200"
    >
      <div
        className="bg-white w-full max-w-2xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[92vh] animate-in zoom-in-95 duration-200"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Modal Top Header (Institutional Dark Navy) */}
        <div className="bg-[#0b1e36] px-5 py-4 flex items-center justify-between text-white relative">
          <div className="flex items-center gap-3">
            <div className="p-1 rounded-full bg-white/10 border border-white/20">
              <CollegeLogo size={36} />
            </div>
            <div>
              <div className="text-[10px] font-bold tracking-widest text-[#f59e0b] uppercase">
                CENTRAL EXAM CELL • ALL SUBJECTS MARKSHEET
              </div>
              <h2
                id="consolidated-download-modal-title"
                className="text-base sm:text-lg font-black tracking-wide text-white"
              >
                Download Consolidated Marksheet &amp; Analysis
              </h2>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-300 hover:text-white rounded-lg hover:bg-white/10 transition cursor-pointer"
            title="Close"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Content */}
        <div className="p-5 sm:p-6 overflow-y-auto space-y-4">
          {/* Metadata Card */}
          <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Department
              </span>
              <span className="font-bold text-slate-800 line-clamp-1" title={data.department}>
                {data.department.split(' - ')[0]}
              </span>
            </div>
            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Semester
              </span>
              <span className="font-bold text-slate-800">{data.semester}</span>
            </div>
            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Assessment
              </span>
              <span className="font-bold text-[#0b1e36] uppercase">{data.examType}</span>
            </div>
            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Total Strength
              </span>
              <span className="font-bold text-slate-800">{data.totalStudentStrength} Students</span>
            </div>
          </div>

          {/* Submitted Subjects Included */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wide flex items-center gap-1.5">
                <BookOpen className="w-3.5 h-3.5 text-blue-900" />
                <span>Submitted Subjects Included ({data.subjects.length}):</span>
              </span>
              <span className="text-[11px] font-semibold text-slate-700 bg-slate-100 px-2 py-0.5 rounded-full border border-slate-200">
                All Marks Linked
              </span>
            </div>
            <div className="flex flex-wrap gap-2 max-h-32 overflow-y-auto p-2 bg-slate-50/80 rounded-xl border border-slate-200/80">
              {data.subjects.map((sub, idx) => (
                <div
                  key={sub.code}
                  className="flex items-center gap-1.5 px-2.5 py-1 bg-white rounded-lg border border-slate-200 text-xs shadow-2xs"
                >
                  <span className="w-4 h-4 rounded-full bg-slate-100 text-[#0b1e36] text-[10px] font-bold flex items-center justify-center shrink-0">
                    {idx + 1}
                  </span>
                  <span className="font-mono font-bold text-slate-800">{sub.code}</span>
                  <span className="text-slate-400">•</span>
                  <span className="text-slate-700 font-semibold text-[11px]">{sub.shortName}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Quick Stats Overview */}
          <div className="grid grid-cols-3 gap-2.5 text-center">
            <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-xl">
              <div className="text-lg font-black text-[#0b1e36]">{data.passInAllSubjectsCount}</div>
              <div className="text-[10px] font-bold text-slate-600 uppercase tracking-wide">
                Pass in all Subjects
              </div>
            </div>
            <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-xl">
              <div className="text-lg font-black text-slate-800">
                {data.totalStudentStrength - data.passInAllSubjectsCount}
              </div>
              <div className="text-[10px] font-bold text-slate-600 uppercase tracking-wide">
                Arrear Students
              </div>
            </div>
            <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-xl">
              <div className="text-lg font-black text-[#0b1e36]">
                {data.overallPassPercentage.toFixed(1)}%
              </div>
              <div className="text-[10px] font-bold text-slate-600 uppercase tracking-wide">
                Overall Pass Rate
              </div>
            </div>
          </div>

          {/* Subject Wise Analysis Table Preview - Clean & Professional Design */}
          <div className="rounded-xl border border-slate-200 overflow-hidden text-xs">
            <div className="bg-[#0b1e36] text-white px-3 py-2.5 font-bold uppercase tracking-wide text-[11px] flex items-center justify-between">
              <span>Subject Wise Result &amp; Performance Analysis</span>
              <span className="text-slate-300 font-normal">{data.examType} Evaluation</span>
            </div>
            <div className="overflow-x-auto max-h-64">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-100 text-[10px] text-slate-700 font-bold border-b border-slate-200 uppercase">
                    <th className="p-2 text-center w-8">#</th>
                    <th className="p-2">Subject Code &amp; Title</th>
                    <th className="p-2">Handled By</th>
                    <th className="p-2 text-center">Str</th>
                    <th className="p-2 text-center">App</th>
                    <th className="p-2 text-center">Pass</th>
                    <th className="p-2 text-center">Fail</th>
                    <th className="p-2 text-center">Pass %</th>
                    <th className="p-2 text-center">Highest</th>
                    <th className="p-2 text-center">Rating</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-[11px]">
                  {data.subjects.map((s, idx) => {
                    const m = data.subjectMetrics[s.code];
                    const appPct = m?.passPercentageApp ?? 0;
                    let rating = 'Satisfactory';
                    let ratingClass = 'text-blue-800 bg-blue-50';
                    if (appPct >= 90) {
                      rating = 'Exemplary';
                      ratingClass = 'text-emerald-800 bg-emerald-50';
                    } else if (appPct >= 80) {
                      rating = 'Very Good';
                      ratingClass = 'text-emerald-800 bg-emerald-50';
                    } else if (appPct >= 75) {
                      rating = 'Target Met';
                      ratingClass = 'text-blue-800 bg-blue-50';
                    } else if (appPct < 60) {
                      rating = 'Remedial Req.';
                      ratingClass = 'text-rose-800 bg-rose-50';
                    }

                    return (
                      <tr key={s.code} className="hover:bg-slate-50/80 transition-colors">
                        <td className="p-2 text-center text-slate-400 font-mono">{idx + 1}</td>
                        <td className="p-2">
                          <span className="font-mono font-bold text-slate-900 mr-1.5">{s.code}</span>
                          <span className="text-slate-600 font-medium">{s.name}</span>
                        </td>
                        <td className="p-2 text-slate-600 font-medium">
                          {s.facultyInCharge || m?.facultyInCharge || 'Faculty'}
                        </td>
                        <td className="p-2 text-center font-bold text-slate-800">
                          {m?.totalStrength ?? data.totalStudentStrength}
                        </td>
                        <td className="p-2 text-center font-bold text-slate-800">
                          {m?.totalPresent ?? 0}
                        </td>
                        <td className="p-2 text-center font-black text-emerald-700">
                          {m?.totalPass ?? 0}
                        </td>
                        <td className="p-2 text-center font-bold text-slate-700">
                          {m?.totalFail ?? 0}
                        </td>
                        <td className="p-2 text-center font-black text-blue-900">
                          {appPct}%
                        </td>
                        <td className="p-2 text-center font-bold text-slate-900">
                          {m?.highestMarks ?? 0}
                        </td>
                        <td className="p-2 text-center">
                          <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${ratingClass}`}>
                            {rating}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-100 font-bold text-slate-900 border-t-2 border-slate-300 text-xs">
                    <td colSpan={3} className="p-2 text-right uppercase tracking-wider font-extrabold text-[#0b1e36]">
                      Dept. Average
                    </td>
                    <td className="p-2 text-center font-black">
                      {data.totalStudentStrength}
                    </td>
                    <td className="p-2 text-center font-black">
                      {data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.totalPresent ?? 0), 0)}
                    </td>
                    <td className="p-2 text-center font-black text-emerald-800">
                      {data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.totalPass ?? 0), 0)}
                    </td>
                    <td className="p-2 text-center font-black">
                      {data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.totalFail ?? 0), 0)}
                    </td>
                    <td className="p-2 text-center font-black text-blue-900">
                      {data.subjects.length > 0
                        ? Math.round(
                            data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.passPercentageApp ?? 0), 0) /
                              data.subjects.length
                          )
                        : 0}%
                    </td>
                    <td colSpan={2} className="p-2 text-center font-black text-[#0b1e36]">
                      Overall Pass: {data.overallPassPercentage.toFixed(1)}%
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>

          {/* Download Notification Banner */}
          {downloadSuccessNotice && (
            <div className="p-3 bg-emerald-100/90 text-emerald-900 border border-emerald-300 rounded-xl text-xs font-bold flex items-center gap-2 animate-in fade-in">
              <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
              <span>{downloadSuccessNotice}</span>
            </div>
          )}

          {/* 3 Download Action Options */}
          <div className="space-y-2.5 pt-2">
            <div className="text-xs font-bold text-slate-700 uppercase tracking-wide">
              Select Download Format:
            </div>

            {/* 1. PDF Consolidated Report (.pdf) */}
            <button
              type="button"
              id="btn-consolidated-download-pdf"
              disabled={downloadingType !== null}
              onClick={() => handleDownload('pdf')}
              className="w-full flex items-center justify-between p-3.5 rounded-xl border border-rose-200 bg-rose-50/40 hover:bg-rose-50 hover:border-rose-300 transition-all text-left group cursor-pointer disabled:opacity-50 shadow-2xs"
            >
              <div className="flex items-center gap-3">
                <div className="p-2.5 rounded-lg bg-rose-600 text-white shadow-2xs">
                  <FileText className="w-5 h-5" />
                </div>
                <div>
                  <div className="font-bold text-slate-900 text-sm flex items-center gap-2">
                    <span>PDF Document</span>
                    <span className="text-[10px] font-extrabold uppercase px-1.5 py-0.5 rounded-md bg-rose-100 text-rose-700 border border-rose-200">
                      .pdf
                    </span>
                  </div>
                  <div className="text-xs text-slate-500 mt-0.5">
                    Official Landscape Marksheet + Subject Wise Analysis (2-Page Report)
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2 text-xs font-bold text-rose-700 bg-white px-3 py-1.5 rounded-lg border border-rose-200 group-hover:bg-rose-600 group-hover:text-white transition">
                {downloadingType === 'pdf' ? (
                  <span>Generating...</span>
                ) : (
                  <>
                    <Download className="w-4 h-4" />
                    <span>Download PDF</span>
                  </>
                )}
              </div>
            </button>

            {/* 2. Excel Consolidated Spreadsheet (.xlsx) */}
            <button
              type="button"
              id="btn-consolidated-download-excel"
              disabled={downloadingType !== null}
              onClick={() => handleDownload('excel')}
              className="w-full flex items-center justify-between p-3.5 rounded-xl border border-emerald-200 bg-emerald-50/40 hover:bg-emerald-50 hover:border-emerald-300 transition-all text-left group cursor-pointer disabled:opacity-50 shadow-2xs"
            >
              <div className="flex items-center gap-3">
                <div className="p-2.5 rounded-lg bg-emerald-600 text-white shadow-2xs">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <div className="font-bold text-slate-900 text-sm flex items-center gap-2">
                    <span>Excel Spreadsheet</span>
                    <span className="text-[10px] font-extrabold uppercase px-1.5 py-0.5 rounded-md bg-emerald-100 text-emerald-700 border border-emerald-200">
                      .xlsx
                    </span>
                  </div>
                  <div className="text-xs text-slate-500 mt-0.5">
                    Sheet 1: Consolidated Marksheet • Sheet 2: Subwise Result Analysis
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2 text-xs font-bold text-emerald-700 bg-white px-3 py-1.5 rounded-lg border border-emerald-200 group-hover:bg-emerald-600 group-hover:text-white transition">
                {downloadingType === 'excel' ? (
                  <span>Generating...</span>
                ) : (
                  <>
                    <Download className="w-4 h-4" />
                    <span>Download Excel</span>
                  </>
                )}
              </div>
            </button>

            {/* 3. Word Document (.docx) */}
            <button
              type="button"
              id="btn-consolidated-download-word"
              disabled={downloadingType !== null}
              onClick={() => handleDownload('word')}
              className="w-full flex items-center justify-between p-3.5 rounded-xl border border-blue-200 bg-blue-50/40 hover:bg-blue-50 hover:border-blue-300 transition-all text-left group cursor-pointer disabled:opacity-50 shadow-2xs"
            >
              <div className="flex items-center gap-3">
                <div className="p-2.5 rounded-lg bg-blue-600 text-white shadow-2xs">
                  <FileText className="w-5 h-5" />
                </div>
                <div>
                  <div className="font-bold text-slate-900 text-sm flex items-center gap-2">
                    <span>Word Document</span>
                    <span className="text-[10px] font-extrabold uppercase px-1.5 py-0.5 rounded-md bg-blue-100 text-blue-700 border border-blue-200">
                      .docx
                    </span>
                  </div>
                  <div className="text-xs text-slate-500 mt-0.5">
                    Editable Microsoft Word Table with full Institutional Letterhead
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2 text-xs font-bold text-blue-700 bg-white px-3 py-1.5 rounded-lg border border-blue-200 group-hover:bg-blue-600 group-hover:text-white transition">
                {downloadingType === 'word' ? (
                  <span>Generating...</span>
                ) : (
                  <>
                    <Download className="w-4 h-4" />
                    <span>Download Word</span>
                  </>
                )}
              </div>
            </button>
          </div>
        </div>

        {/* Modal Bottom Footer */}
        <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
          <div>Central Examination Cell • Jaya Sakthi Engineering College, Thiruninravur, Chennai - 602024</div>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold rounded-xl transition cursor-pointer"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}
