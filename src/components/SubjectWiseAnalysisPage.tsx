import React, { useMemo } from 'react';
import {
  ConsolidatedData,
  getYearSemRoman,
} from '../utils/consolidatedExport';
import CollegeLogo from './CollegeLogo';
import {
  ArrowLeft,
  Download,
  FileSpreadsheet,
  FileText,
  Printer,
  Award,
  TrendingUp,
  Users,
  CheckCircle2,
  XCircle,
  AlertCircle,
  Clock,
  BookOpen,
  Building2,
  Calendar,
  Layers,
} from 'lucide-react';

interface SubjectWiseAnalysisPageProps {
  data: ConsolidatedData;
  onBack?: () => void;
  onDownloadPdf?: () => void;
  onDownloadExcel?: () => void;
  onDownloadWord?: () => void;
}

export default function SubjectWiseAnalysisPage({
  data,
  onBack,
  onDownloadPdf,
  onDownloadExcel,
  onDownloadWord,
}: SubjectWiseAnalysisPageProps) {
  // Compute detailed statistics per subject
  const detailedSubjectStats = useMemo(() => {
    return data.subjects.map((sub, idx) => {
      const metric = data.subjectMetrics[sub.code];
      const strength = metric?.totalStrength ?? data.totalStudentStrength;
      const appeared = metric?.totalPresent ?? 0;
      const absent = metric?.totalAbsent ?? 0;
      const pass = metric?.totalPass ?? 0;
      const fail = metric?.totalFail ?? 0;
      const passStr = metric?.passPercentageStr ?? (strength > 0 ? Math.round((pass / strength) * 100) : 0);
      const passApp = metric?.passPercentageApp ?? (appeared > 0 ? Math.round((pass / appeared) * 100) : 0);
      const highest = metric?.highestMarks ?? 0;

      // Grade breakdown & average mark calculation
      let totalMarks = 0;
      let markCount = 0;
      let lowestMark = 100;
      const gradeCounts: Record<string, number> = {
        O: 0,
        'A+': 0,
        A: 0,
        'B+': 0,
        B: 0,
        C: 0,
        U: 0,
        AB: 0,
      };

      data.students.forEach((st) => {
        const sm = st.subjects[sub.code];
        if (sm) {
          if (typeof sm.marksObtained === 'number') {
            totalMarks += sm.marksObtained;
            markCount++;
            if (sm.marksObtained < lowestMark) lowestMark = sm.marksObtained;
          }
          if (sm.grade) {
            gradeCounts[sm.grade] = (gradeCounts[sm.grade] || 0) + 1;
          }
        }
      });

      const averageMark = markCount > 0 ? (totalMarks / markCount).toFixed(1) : '0.0';

      // Academic performance category
      let performanceRating = 'Satisfactory';
      let ratingVariant: 'emerald' | 'blue' | 'amber' | 'rose' = 'blue';

      if (passApp >= 90) {
        performanceRating = 'Exemplary (≥90%)';
        ratingVariant = 'emerald';
      } else if (passApp >= 80) {
        performanceRating = 'Very Good (80–89%)';
        ratingVariant = 'emerald';
      } else if (passApp >= 75) {
        performanceRating = 'Target Met (75–79%)';
        ratingVariant = 'blue';
      } else if (passApp >= 60) {
        performanceRating = 'Average (60–74%)';
        ratingVariant = 'amber';
      } else {
        performanceRating = 'Remedial Required (<60%)';
        ratingVariant = 'rose';
      }

      return {
        sNo: idx + 1,
        code: sub.code,
        name: sub.name,
        shortName: sub.shortName,
        faculty: sub.facultyInCharge || metric?.facultyInCharge || 'Faculty In-Charge',
        strength,
        appeared,
        absent,
        pass,
        fail,
        passStr,
        passApp,
        highest,
        lowest: markCount > 0 ? lowestMark : 0,
        averageMark,
        gradeCounts,
        performanceRating,
        ratingVariant,
      };
    });
  }, [data]);

  // Overall totals across all subjects
  const aggregates = useMemo(() => {
    const totalSubjects = detailedSubjectStats.length;
    const avgPassApp =
      totalSubjects > 0
        ? Math.round(
            detailedSubjectStats.reduce((acc, s) => acc + s.passApp, 0) / totalSubjects
          )
        : 0;

    const totalPassedEntries = detailedSubjectStats.reduce((acc, s) => acc + s.pass, 0);
    const totalFailedEntries = detailedSubjectStats.reduce((acc, s) => acc + s.fail, 0);
    const totalAppearedEntries = detailedSubjectStats.reduce((acc, s) => acc + s.appeared, 0);
    const totalAbsentEntries = detailedSubjectStats.reduce((acc, s) => acc + s.absent, 0);

    const highestPerformingSubject = [...detailedSubjectStats].sort(
      (a, b) => b.passApp - a.passApp
    )[0];

    return {
      totalSubjects,
      avgPassApp,
      totalPassedEntries,
      totalFailedEntries,
      totalAppearedEntries,
      totalAbsentEntries,
      highestPerformingSubject,
    };
  }, [detailedSubjectStats]);

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="bg-slate-50 min-h-screen py-8 px-4 sm:px-6 lg:px-8 space-y-6 print:p-0 print:bg-white">
      {/* Top Action & Navigation Bar */}
      <div className="max-w-7xl mx-auto flex flex-col sm:flex-row sm:items-center justify-between gap-4 print:hidden">
        <div className="flex items-center gap-3">
          {onBack && (
            <button
              type="button"
              id="btn-subwise-back"
              onClick={onBack}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-bold border border-slate-200 shadow-2xs transition cursor-pointer"
            >
              <ArrowLeft className="w-4 h-4 text-slate-600" />
              <span>Back to Marksheets</span>
            </button>
          )}
          <div className="text-xs text-slate-500 font-medium">
            <span>Central Examination Cell</span>
            <span className="mx-1.5">·</span>
            <span className="font-bold text-slate-800">Subject-Wise Result Audit</span>
          </div>
        </div>

        {/* Export Buttons */}
        <div className="flex flex-wrap items-center gap-2">
          {onDownloadPdf && (
            <button
              type="button"
              id="btn-subwise-download-pdf"
              onClick={onDownloadPdf}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded-xl text-xs font-bold transition shadow-2xs border border-amber-600 cursor-pointer"
              title="Download Subject Wise Analysis as PDF"
            >
              <Download className="w-4 h-4" />
              <span>PDF (.pdf)</span>
            </button>
          )}

          {onDownloadExcel && (
            <button
              type="button"
              id="btn-subwise-download-excel"
              onClick={onDownloadExcel}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-700 hover:bg-emerald-600 text-white rounded-xl text-xs font-bold transition shadow-2xs border border-emerald-800 cursor-pointer"
              title="Download Subject Wise Analysis as Excel"
            >
              <FileSpreadsheet className="w-4 h-4 text-emerald-200" />
              <span>Excel (.xlsx)</span>
            </button>
          )}

          {onDownloadWord && (
            <button
              type="button"
              id="btn-subwise-download-word"
              onClick={onDownloadWord}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-blue-700 hover:bg-blue-600 text-white rounded-xl text-xs font-bold transition shadow-2xs border border-blue-800 cursor-pointer"
              title="Download Subject Wise Analysis as Word"
            >
              <FileText className="w-4 h-4 text-blue-200" />
              <span>Word (.docx)</span>
            </button>
          )}

          <button
            type="button"
            id="btn-subwise-print"
            onClick={handlePrint}
            className="flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-bold border border-slate-200 shadow-2xs transition cursor-pointer"
            title="Print Subject Wise Analysis"
          >
            <Printer className="w-4 h-4 text-slate-600" />
            <span>Print Report</span>
          </button>
        </div>
      </div>

      {/* Main Analysis Document Sheet */}
      <div className="max-w-7xl mx-auto bg-white rounded-2xl border border-slate-200/90 shadow-sm overflow-hidden print:border-none print:shadow-none print:rounded-none">
        {/* ========================================================================= */}
        {/* 1. INSTITUTIONAL HEADER                                                   */}
        {/* ========================================================================= */}
        <div className="p-6 sm:p-8 bg-white border-b border-slate-200 text-center relative">
          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
            {/* Left College Logo */}
            <div className="shrink-0 flex items-center justify-center">
              <CollegeLogo size={68} />
            </div>

            {/* Center Institution Info */}
            <div className="flex-1 space-y-1">
              <h1 className="text-xl sm:text-2xl font-serif font-black text-[#0b1e36] tracking-tight uppercase">
                JAYA SAKTHI ENGINEERING COLLEGE
              </h1>
              <p className="text-xs sm:text-sm text-slate-700 font-bold">
                St. Mary&apos;s Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu
              </p>
              <p className="text-[11px] sm:text-xs text-slate-500 font-medium">
                Approved by AICTE, New Delhi &amp; Affiliated to Anna University, Chennai
              </p>
              <div className="pt-2">
                <span className="inline-block px-3 py-1 bg-slate-100 text-[#0b1e36] rounded-md text-xs font-extrabold uppercase tracking-widest border border-slate-200">
                  {data.examType} RESULT ANALYSIS
                </span>
                <h2 className="text-sm sm:text-base font-extrabold text-slate-900 tracking-wide mt-1 underline decoration-slate-400 underline-offset-4 uppercase">
                  SUBJECT WISE ANALYSIS
                </h2>
              </div>
            </div>

            {/* Right Badge / Academic Cell */}
            <div className="shrink-0 hidden sm:flex flex-col items-center justify-center p-3 rounded-xl bg-slate-50 border border-slate-200 text-center w-28">
              <Award className="w-6 h-6 text-blue-900 mb-1" />
              <div className="text-[10px] font-black uppercase text-[#0b1e36]">Central Exam</div>
              <div className="text-[9px] text-slate-500 font-semibold">Quality Cell</div>
            </div>
          </div>

          {/* Institutional Metadata Grid */}
          <div className="mt-6 pt-4 border-t border-slate-200 grid grid-cols-2 sm:grid-cols-4 gap-4 text-left text-xs bg-slate-50/70 p-4 rounded-xl border">
            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Department
              </span>
              <span className="font-extrabold text-slate-900 line-clamp-1" title={data.department}>
                {data.department}
              </span>
            </div>

            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Year / Semester
              </span>
              <span className="font-extrabold text-slate-900">
                {getYearSemRoman(data.semester)} ({data.semester})
              </span>
            </div>

            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Assessment Date
              </span>
              <span className="font-extrabold text-slate-900 font-mono">
                {data.assessmentDate}
              </span>
            </div>

            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                Section &amp; Batch
              </span>
              <span className="font-extrabold text-slate-900">
                Section {data.section} · {data.academicYear}
              </span>
            </div>
          </div>
        </div>

        {/* ========================================================================= */}
        {/* 2. EXECUTIVE PERFORMANCE METRICS BAR                                      */}
        {/* ========================================================================= */}
        <div className="p-6 bg-slate-50/50 border-b border-slate-200">
          <div className="grid grid-cols-2 lg:grid-cols-5 gap-3.5">
            {/* 1. Subjects Evaluated */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                <BookOpen className="w-3.5 h-3.5 text-blue-900" />
                <span>Subjects Analyzed</span>
              </div>
              <div className="text-2xl font-black text-slate-900 mt-1 tabular-nums">
                {aggregates.totalSubjects}
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5 font-medium">
                All submitted courses
              </div>
            </div>

            {/* 2. Total Student Strength */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                <Users className="w-3.5 h-3.5 text-slate-600" />
                <span>Total Class Strength</span>
              </div>
              <div className="text-2xl font-black text-slate-900 mt-1 tabular-nums">
                {data.totalStudentStrength}
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5 font-medium">
                Enrolled students in batch
              </div>
            </div>

            {/* 3. Passed in All Subjects */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                <span>All-Clear Students</span>
              </div>
              <div className="text-2xl font-black text-emerald-700 mt-1 tabular-nums">
                {data.passInAllSubjectsCount}
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5 font-medium">
                Zero arrear students
              </div>
            </div>

            {/* 4. Overall Class Pass Rate */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                <TrendingUp className="w-3.5 h-3.5 text-blue-800" />
                <span>Overall Pass Rate</span>
              </div>
              <div className="text-2xl font-black text-[#0b1e36] mt-1 tabular-nums">
                {data.overallPassPercentage.toFixed(1)}%
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5 font-medium">
                All-subjects pass ratio
              </div>
            </div>

            {/* 5. Subject Average Pass Rate */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs col-span-2 lg:col-span-1">
              <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                <Award className="w-3.5 h-3.5 text-amber-600" />
                <span>Dept. Avg Pass Rate</span>
              </div>
              <div className="text-2xl font-black text-slate-900 mt-1 tabular-nums">
                {aggregates.avgPassApp}%
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5 font-medium">
                Top: {aggregates.highestPerformingSubject?.code || 'N/A'} ({aggregates.highestPerformingSubject?.passApp || 0}%)
              </div>
            </div>
          </div>
        </div>

        {/* ========================================================================= */}
        {/* 3. MAIN SUBJECT WISE ANALYSIS TABLE                                      */}
        {/* ========================================================================= */}
        <div className="p-6 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div>
              <h3 className="text-base font-extrabold text-slate-900 uppercase tracking-tight">
                Subject Performance &amp; Evaluation Audit Table
              </h3>
              <p className="text-xs text-slate-500">
                Detailed comparison of student appearances, pass/fail counts, pass rates, and evaluation remarks.
              </p>
            </div>
            <div className="text-xs text-slate-500 font-medium">
              Institutional Benchmark: <strong className="text-slate-800">75% Pass Rate Target</strong>
            </div>
          </div>

          <div className="overflow-x-auto rounded-xl border border-slate-200 shadow-2xs">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                {/* Tier 1 Header */}
                <tr className="bg-[#0b1e36] text-white font-bold uppercase text-[11px] border-b border-slate-700">
                  <th rowSpan={2} className="py-3 px-3 text-center w-12 border-r border-white/10">
                    S.No
                  </th>
                  <th rowSpan={2} className="py-3 px-3.5 min-w-[200px] border-r border-white/10">
                    Subject Code &amp; Title
                  </th>
                  <th rowSpan={2} className="py-3 px-3 min-w-[140px] border-r border-white/10">
                    Handled By (Faculty)
                  </th>
                  <th rowSpan={2} className="py-3 px-2.5 text-center border-r border-white/10">
                    Strength
                  </th>
                  <th rowSpan={2} className="py-3 px-2.5 text-center border-r border-white/10">
                    Appeared
                  </th>
                  <th rowSpan={2} className="py-3 px-2 text-center border-r border-white/10">
                    Absent
                  </th>
                  <th rowSpan={2} className="py-3 px-2.5 text-center border-r border-white/10">
                    Passed
                  </th>
                  <th rowSpan={2} className="py-3 px-2 text-center border-r border-white/10">
                    Failed
                  </th>
                  <th colSpan={2} className="py-1.5 px-2 text-center border-b border-white/20 border-r border-white/10">
                    Pass %
                  </th>
                  <th rowSpan={2} className="py-3 px-2.5 text-center border-r border-white/10">
                    Highest Mark
                  </th>
                  <th rowSpan={2} className="py-3 px-2.5 text-center border-r border-white/10">
                    Class Avg
                  </th>
                  <th rowSpan={2} className="py-3 px-3 text-center">
                    Evaluation Performance Rating
                  </th>
                </tr>
                {/* Tier 2 Header for Pass % */}
                <tr className="bg-[#0f284e] text-white text-[10px] font-bold uppercase border-b border-slate-300">
                  <th className="py-1.5 px-2 text-center border-r border-white/10" title="Based on total enrolled strength">
                    Str
                  </th>
                  <th className="py-1.5 px-2 text-center border-r border-white/10" title="Based on appeared students">
                    App
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 bg-white">
                {detailedSubjectStats.map((row) => (
                  <tr key={row.code} className="hover:bg-slate-50/80 transition-colors">
                    {/* S.No */}
                    <td className="py-3 px-3 text-center text-slate-400 font-mono font-medium">
                      {row.sNo}
                    </td>

                    {/* Subject Code & Name */}
                    <td className="py-3 px-3.5">
                      <div className="flex items-center gap-2">
                        <span className="font-mono font-bold text-slate-950 bg-slate-100 px-2 py-0.5 rounded text-[11px] border border-slate-300">
                          {row.code}
                        </span>
                        <div className="font-bold text-slate-900 leading-tight">
                          {row.name}
                        </div>
                      </div>
                    </td>

                    {/* Faculty */}
                    <td className="py-3 px-3 font-medium text-slate-700">
                      {row.faculty}
                    </td>

                    {/* Strength */}
                    <td className="py-3 px-2.5 text-center font-bold text-slate-800 tabular-nums">
                      {row.strength}
                    </td>

                    {/* Appeared */}
                    <td className="py-3 px-2.5 text-center font-bold text-slate-800 tabular-nums">
                      {row.appeared}
                    </td>

                    {/* Absent */}
                    <td className="py-3 px-2 text-center tabular-nums">
                      {row.absent > 0 ? (
                        <span className="text-slate-600 font-bold">{row.absent}</span>
                      ) : (
                        <span className="text-slate-400">0</span>
                      )}
                    </td>

                    {/* Passed */}
                    <td className="py-3 px-2.5 text-center tabular-nums font-black text-emerald-700">
                      {row.pass}
                    </td>

                    {/* Failed */}
                    <td className="py-3 px-2 text-center tabular-nums">
                      {row.fail > 0 ? (
                        <span className="font-black text-slate-700">{row.fail}</span>
                      ) : (
                        <span className="text-slate-400">0</span>
                      )}
                    </td>

                    {/* Pass % (Str) */}
                    <td className="py-3 px-2 text-center tabular-nums font-semibold text-slate-700">
                      {row.passStr}%
                    </td>

                    {/* Pass % (App) */}
                    <td className="py-3 px-2 text-center tabular-nums font-black text-blue-900 text-[13px]">
                      {row.passApp}%
                    </td>

                    {/* Highest Mark */}
                    <td className="py-3 px-2.5 text-center tabular-nums font-black text-slate-900">
                      {row.highest}
                    </td>

                    {/* Class Avg */}
                    <td className="py-3 px-2.5 text-center tabular-nums font-semibold text-slate-600">
                      {row.averageMark}
                    </td>

                    {/* Performance Rating */}
                    <td className="py-3 px-3 text-center">
                      <span
                        className={`inline-block px-2.5 py-1 rounded-md text-[11px] font-bold ${
                          row.ratingVariant === 'emerald'
                            ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                            : row.ratingVariant === 'blue'
                            ? 'bg-blue-50 text-blue-800 border border-blue-200'
                            : row.ratingVariant === 'amber'
                            ? 'bg-amber-50 text-amber-800 border border-amber-200'
                            : 'bg-rose-50 text-rose-800 border border-rose-200'
                        }`}
                      >
                        {row.performanceRating}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
              {/* Summary Totals Row */}
              <tfoot>
                <tr className="bg-slate-100/90 font-bold text-slate-900 border-t-2 border-slate-300 text-xs">
                  <td colSpan={3} className="py-3 px-3.5 text-right uppercase tracking-wider font-extrabold text-[#0b1e36]">
                    Total / Department Average
                  </td>
                  <td className="py-3 px-2.5 text-center font-black tabular-nums">
                    {data.totalStudentStrength}
                  </td>
                  <td className="py-3 px-2.5 text-center font-black tabular-nums">
                    {aggregates.totalAppearedEntries}
                  </td>
                  <td className="py-3 px-2 text-center font-black tabular-nums">
                    {aggregates.totalAbsentEntries}
                  </td>
                  <td className="py-3 px-2.5 text-center font-black text-emerald-800 tabular-nums">
                    {aggregates.totalPassedEntries}
                  </td>
                  <td className="py-3 px-2 text-center font-black tabular-nums">
                    {aggregates.totalFailedEntries}
                  </td>
                  <td className="py-3 px-2 text-center tabular-nums font-bold text-slate-700">
                    {data.totalStudentStrength > 0
                      ? Math.round((aggregates.totalPassedEntries / (data.totalStudentStrength * aggregates.totalSubjects)) * 100)
                      : 0}%
                  </td>
                  <td className="py-3 px-2 text-center tabular-nums font-black text-blue-900 text-sm">
                    {aggregates.avgPassApp}%
                  </td>
                  <td colSpan={3} className="py-3 px-3 text-center text-slate-600 font-semibold text-[11px]">
                    Overall Result: <strong className="text-slate-900">{data.overallPassPercentage.toFixed(1)}% All-Clear</strong>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>

        {/* ========================================================================= */}
        {/* 4. VISUAL SUBJECT PERFORMANCE PROGRESS BARS                               */}
        {/* ========================================================================= */}
        <div className="p-6 bg-slate-50/60 border-t border-slate-200 space-y-4">
          <div className="flex items-center justify-between">
            <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <Layers className="w-4 h-4 text-blue-900" />
              <span>Subject Pass Percentage Breakdown &amp; Benchmark Comparison</span>
            </h4>
            <span className="text-xs text-slate-500 font-medium">
              Target Threshold: 75%
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3.5">
            {detailedSubjectStats.map((sub) => (
              <div
                key={sub.code}
                className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs space-y-2.5"
              >
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <span className="font-mono text-xs font-bold text-blue-900 bg-blue-50 px-1.5 py-0.5 rounded border border-blue-200">
                      {sub.code}
                    </span>
                    <h5 className="font-extrabold text-slate-900 text-xs mt-1 line-clamp-1" title={sub.name}>
                      {sub.name}
                    </h5>
                    <div className="text-[11px] text-slate-500 font-medium">
                      Faculty: {sub.faculty}
                    </div>
                  </div>
                  <div className="text-right shrink-0">
                    <div className="text-base font-black text-[#0b1e36] tabular-nums">
                      {sub.passApp}%
                    </div>
                    <div className="text-[10px] text-slate-400 uppercase font-semibold">
                      Pass Rate
                    </div>
                  </div>
                </div>

                {/* Progress Bar with Benchmark Marker at 75% */}
                <div className="relative pt-1">
                  <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden">
                    <div
                      className={`h-full rounded-full transition-all duration-300 ${
                        sub.passApp >= 85
                          ? 'bg-emerald-600'
                          : sub.passApp >= 75
                          ? 'bg-blue-600'
                          : sub.passApp >= 60
                          ? 'bg-amber-500'
                          : 'bg-rose-500'
                      }`}
                      style={{ width: `${Math.min(100, Math.max(0, sub.passApp))}%` }}
                    />
                  </div>
                  {/* Benchmark 75% indicator line */}
                  <div
                    className="absolute top-0 bottom-0 w-0.5 bg-slate-400/60"
                    style={{ left: '75%' }}
                    title="75% Target"
                  />
                </div>

                {/* Quick Numbers Bar */}
                <div className="flex items-center justify-between text-[11px] text-slate-600 pt-1 border-t border-slate-100">
                  <span>
                    Passed: <strong className="text-emerald-700 font-bold">{sub.pass}</strong> / {sub.appeared}
                  </span>
                  <span>
                    Highest: <strong className="text-slate-900 font-bold">{sub.highest}</strong>
                  </span>
                  <span>
                    Avg: <strong className="text-slate-900 font-bold">{sub.averageMark}</strong>
                  </span>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* ========================================================================= */}
        {/* 5. OFFICIAL SIGNATURES & VERIFICATION BLOCK                               */}
        {/* ========================================================================= */}
        <div className="p-6 sm:p-8 bg-white border-t border-slate-200">
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-6 sm:gap-4 text-center text-xs">
            {/* 1. CC */}
            <div className="space-y-10">
              <div className="h-10 flex items-end justify-center">
                <span className="text-[11px] text-slate-400 italic">Signature</span>
              </div>
              <div>
                <div className="w-24 sm:w-28 h-0.5 bg-slate-300 mx-auto mb-1.5" />
                <p className="font-extrabold text-slate-800 uppercase tracking-wide">
                  CC
                </p>
                <p className="text-[11px] text-slate-500 font-medium">(Class Coordinator)</p>
              </div>
            </div>

            {/* 2. HOD */}
            <div className="space-y-10">
              <div className="h-10 flex items-end justify-center">
                <span className="text-[11px] text-slate-400 italic">Signature</span>
              </div>
              <div>
                <div className="w-24 sm:w-28 h-0.5 bg-slate-300 mx-auto mb-1.5" />
                <p className="font-extrabold text-slate-800 uppercase tracking-wide">
                  HOD
                </p>
                <p className="text-[11px] text-slate-500 font-medium">({data.department.split(' - ')[0]})</p>
              </div>
            </div>

            {/* 3. AC */}
            <div className="space-y-10">
              <div className="h-10 flex items-end justify-center">
                <span className="text-[11px] text-slate-400 italic">Signature</span>
              </div>
              <div>
                <div className="w-24 sm:w-28 h-0.5 bg-slate-300 mx-auto mb-1.5" />
                <p className="font-extrabold text-slate-800 uppercase tracking-wide">
                  AC
                </p>
                <p className="text-[11px] text-slate-500 font-medium">(Academic Coordinator)</p>
              </div>
            </div>

            {/* 4. EXAM CELL COORDINATOR */}
            <div className="space-y-10">
              <div className="h-10 flex items-end justify-center">
                <span className="text-[11px] text-slate-400 italic">Signature</span>
              </div>
              <div>
                <div className="w-24 sm:w-28 h-0.5 bg-slate-300 mx-auto mb-1.5" />
                <p className="font-extrabold text-slate-800 uppercase tracking-wide leading-tight">
                  EXAM CELL COORDINATOR
                </p>
                <p className="text-[11px] text-slate-500 font-medium">(Exam Cell)</p>
              </div>
            </div>

            {/* 5. PRINCIPAL */}
            <div className="space-y-10">
              <div className="h-10 flex items-end justify-center">
                <span className="text-[11px] text-slate-400 italic">Signature</span>
              </div>
              <div>
                <div className="w-24 sm:w-28 h-0.5 bg-slate-300 mx-auto mb-1.5" />
                <p className="font-extrabold text-slate-800 uppercase tracking-wide">
                  PRINCIPAL
                </p>
                <p className="text-[11px] text-slate-500 font-medium">(JSEC)</p>
              </div>
            </div>

            {/* 6. DIRECTOR */}
            <div className="space-y-10">
              <div className="h-10 flex items-end justify-center">
                <span className="text-[11px] text-slate-400 italic">Signature</span>
              </div>
              <div>
                <div className="w-24 sm:w-28 h-0.5 bg-slate-300 mx-auto mb-1.5" />
                <p className="font-extrabold text-slate-800 uppercase tracking-wide">
                  DIRECTOR
                </p>
                <p className="text-[11px] text-slate-500 font-medium">(Jaya Group)</p>
              </div>
            </div>
          </div>

          {/* Footer note */}
          <div className="mt-8 pt-4 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between text-[11px] text-slate-500 gap-2">
            <div>
              Central Examination Cell · Jaya Sakthi Engineering College, Thiruninravur, Chennai - 602024
            </div>
            <div>
              Report Generated for {data.examType} · Academic Session {data.academicYear}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
