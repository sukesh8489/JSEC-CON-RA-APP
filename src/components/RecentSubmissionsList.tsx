import React, { useState, useMemo } from 'react';
import { AssessmentMeta, StudentMarkRecord } from '../types';
import { SubmittedSheetRecord } from '../utils/submittedSheets';
import { computeMetrics } from '../utils/gradeCalculator';
import {
  Search,
  Filter,
  ArrowUpDown,
  CheckCircle2,
  Clock,
  Lock,
  Unlock,
  FileText,
  FileSpreadsheet,
  Download,
  Printer,
  Trash2,
  Eye,
  Building2,
  GraduationCap,
  Calendar,
  User,
  CheckCircle,
  LayoutGrid,
  List,
  ChevronRight,
  RotateCcw,
  AlertCircle,
  BarChart3,
} from 'lucide-react';

interface RecentSubmissionsListProps {
  sheets: SubmittedSheetRecord[];
  latestSubmittedSheet?: { meta: AssessmentMeta; students: StudentMarkRecord[] } | null;
  onSelectSheet: (sheet: SubmittedSheetRecord) => void;
  onDeleteSheet: (sheet: SubmittedSheetRecord) => void;
  onApproveSheet: (sheetId: string) => void;
  onDownloadPDF: (sheet: SubmittedSheetRecord) => void;
  onDownloadXLS: (sheet: SubmittedSheetRecord) => void;
  onDownloadWord: (sheet: SubmittedSheetRecord) => void;
  onPrintPreview: (sheet: SubmittedSheetRecord) => void;
  onToggleLock: (sheetId: string) => void;
  onRefreshFirestore: () => Promise<void>;
  isRefreshing?: boolean;
  departmentOptions: string[];
}

export default function RecentSubmissionsList({
  sheets,
  latestSubmittedSheet,
  onSelectSheet,
  onDeleteSheet,
  onApproveSheet,
  onDownloadPDF,
  onDownloadXLS,
  onDownloadWord,
  onPrintPreview,
  onToggleLock,
  onRefreshFirestore,
  isRefreshing = false,
  departmentOptions,
}: RecentSubmissionsListProps) {
  // Filter states
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedDept, setSelectedDept] = useState<string>('');
  const [selectedSem, setSelectedSem] = useState<string>('');
  const [selectedExamType, setSelectedExamType] = useState<string>('');
  const [selectedStatus, setSelectedStatus] = useState<string>('');
  const [selectedLockStatus, setSelectedLockStatus] = useState<string>('');
  const [sortBy, setSortBy] = useState<
    'newest' | 'oldest' | 'subject_asc' | 'subject_desc' | 'pass_high' | 'pass_low' | 'students_high'
  >('newest');
  const [viewLayout, setViewLayout] = useState<'table' | 'grid'>('table');

  const EXAM_TYPE_OPTIONS = ['CAT-1', 'CAT-2', 'CAT-3', 'Model-1', 'Model-2', 'Model-3', 'Laboratory'];

  const SEMESTER_OPTIONS = [
    'Semester 1',
    'Semester 2',
    'Semester 3',
    'Semester 4',
    'Semester 5',
    'Semester 6',
    'Semester 7',
    'Semester 8',
  ];

  // Helper for matching exam type
  const matchesExamType = (sheetExamType: string, filterExamType: string): boolean => {
    if (!filterExamType) return true;
    const s = (sheetExamType || '').trim().toUpperCase();
    const f = filterExamType.trim().toUpperCase();
    if (s === f) return true;
    if (s.startsWith(f)) return true;
    if (
      (f === 'MODEL-1' || f === 'M-1') &&
      (s.includes('MODEL-1') || s.includes('MODEL 1') || s.includes('M-1') || s === 'MODEL EXAMINATION')
    ) {
      return true;
    }
    if (
      (f === 'MODEL-2' || f === 'M-2') &&
      (s.includes('MODEL-2') || s.includes('MODEL 2') || s.includes('M-2') || s.includes('MODEL EXAMINATION 2'))
    ) {
      return true;
    }
    if (
      (f === 'MODEL-3' || f === 'M-3') &&
      (s.includes('MODEL-3') || s.includes('MODEL 3') || s.includes('M-3') || s.includes('MODEL EXAMINATION 3'))
    ) {
      return true;
    }
    if ((f === 'LABORATORY' || f === 'LAB') && (s.includes('LABORATORY') || s.includes('LAB'))) {
      return true;
    }
    return s.includes(f);
  };

  // Filtered and sorted sheets
  const filteredSheets = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();

    const filtered = sheets.filter((sheet) => {
      // Search term
      if (term) {
        const matchesSearch =
          sheet.meta.subjectCode.toLowerCase().includes(term) ||
          sheet.meta.subjectName.toLowerCase().includes(term) ||
          sheet.meta.department.toLowerCase().includes(term) ||
          sheet.meta.semester.toLowerCase().includes(term) ||
          (sheet.meta.examType || '').toLowerCase().includes(term) ||
          (sheet.meta.facultyInCharge || '').toLowerCase().includes(term) ||
          (sheet.meta.evaluatorName || '').toLowerCase().includes(term) ||
          (sheet.meta.academicYear || '').toLowerCase().includes(term);

        if (!matchesSearch) return false;
      }

      // Department
      if (selectedDept && sheet.meta.department.trim().toLowerCase() !== selectedDept.trim().toLowerCase()) {
        return false;
      }

      // Semester
      if (selectedSem && sheet.meta.semester.trim().toLowerCase() !== selectedSem.trim().toLowerCase()) {
        return false;
      }

      // Exam Type
      if (selectedExamType && !matchesExamType(sheet.meta.examType, selectedExamType)) {
        return false;
      }

      // Status
      if (selectedStatus && sheet.status !== selectedStatus) {
        return false;
      }

      // Lock Status
      if (selectedLockStatus === 'LOCKED' && sheet.meta.isUnlocked) {
        return false;
      }
      if (selectedLockStatus === 'UNLOCKED' && !sheet.meta.isUnlocked) {
        return false;
      }

      return true;
    });

    // Sorting
    return filtered.sort((a, b) => {
      switch (sortBy) {
        case 'newest':
          return (b.submittedAt || '').localeCompare(a.submittedAt || '');
        case 'oldest':
          return (a.submittedAt || '').localeCompare(b.submittedAt || '');
        case 'subject_asc':
          return a.meta.subjectCode.localeCompare(b.meta.subjectCode);
        case 'subject_desc':
          return b.meta.subjectCode.localeCompare(a.meta.subjectCode);
        case 'pass_high': {
          const passA = computeMetrics(a.students, a.meta.passMark).passPercentage;
          const passB = computeMetrics(b.students, b.meta.passMark).passPercentage;
          return passB - passA;
        }
        case 'pass_low': {
          const passA = computeMetrics(a.students, a.meta.passMark).passPercentage;
          const passB = computeMetrics(b.students, b.meta.passMark).passPercentage;
          return passA - passB;
        }
        case 'students_high':
          return b.students.length - a.students.length;
        default:
          return 0;
      }
    });
  }, [sheets, searchTerm, selectedDept, selectedSem, selectedExamType, selectedStatus, selectedLockStatus, sortBy]);

  const hasActiveFilters = Boolean(
    searchTerm || selectedDept || selectedSem || selectedExamType || selectedStatus || selectedLockStatus
  );

  const resetFilters = () => {
    setSearchTerm('');
    setSelectedDept('');
    setSelectedSem('');
    setSelectedExamType('');
    setSelectedStatus('');
    setSelectedLockStatus('');
    setSortBy('newest');
  };

  // Aggregated quick stats across all stored sheets
  const stats = useMemo(() => {
    const total = sheets.length;
    const approved = sheets.filter((s) => s.status === 'APPROVED').length;
    const pending = sheets.filter((s) => s.status === 'PENDING_VERIFICATION').length;
    let totalStudentsAssessed = 0;
    let totalPassed = 0;

    sheets.forEach((s) => {
      totalStudentsAssessed += s.students.length;
      const m = computeMetrics(s.students, s.meta.passMark);
      totalPassed += m.passCount;
    });

    const avgPassRate =
      totalStudentsAssessed > 0 ? (totalPassed / totalStudentsAssessed) * 100 : 0;

    return { total, approved, pending, totalStudentsAssessed, avgPassRate };
  }, [sheets]);

  return (
    <div className="space-y-6 animate-in fade-in duration-150">
      {/* ========================================================================= */}
      {/* 1. TOP HEADER & FIRESTORE CLOUD PERSISTENCE STATUS BAR                   */}
      {/* ========================================================================= */}
      <div className="bg-white rounded-2xl p-5 sm:p-6 border border-slate-200/90 shadow-2xs">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <span className="p-2 rounded-xl bg-[#0f284e] text-white shadow-xs">
                <Eye className="w-5 h-5 text-amber-400" />
              </span>
              <h2 className="text-lg sm:text-xl font-black text-slate-900 tracking-tight">
                View Submitted Marksheets
              </h2>
            </div>
            <p className="text-xs text-slate-600 max-w-2xl leading-relaxed">
              Browse, filter, and inspect submitted marksheets across all departments and semesters.
            </p>
          </div>
        </div>

        {/* Aggregate Stats Cards */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3.5 mt-5 pt-4 border-t border-slate-100">
          <div className="bg-slate-50/90 rounded-xl p-3 border border-slate-200/80">
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider block">
              Total Sheets
            </span>
            <span className="text-xl sm:text-2xl font-black text-slate-900 mt-0.5 block">
              {stats.total}
            </span>
          </div>

          <div className="bg-emerald-50/80 rounded-xl p-3 border border-emerald-200/80">
            <span className="text-[11px] font-bold text-emerald-700 uppercase tracking-wider block">
              Approved
            </span>
            <span className="text-xl sm:text-2xl font-black text-emerald-900 mt-0.5 block">
              {stats.approved}
            </span>
          </div>

          <div className="bg-amber-50/80 rounded-xl p-3 border border-amber-200/80">
            <span className="text-[11px] font-bold text-amber-700 uppercase tracking-wider block">
              Pending Review
            </span>
            <span className="text-xl sm:text-2xl font-black text-amber-900 mt-0.5 block">
              {stats.pending}
            </span>
          </div>

          <div className="bg-blue-50/80 rounded-xl p-3 border border-blue-200/80">
            <span className="text-[11px] font-bold text-blue-700 uppercase tracking-wider block">
              Avg Pass Rate
            </span>
            <span className="text-xl sm:text-2xl font-black text-blue-950 mt-0.5 block">
              {stats.avgPassRate.toFixed(1)}%
            </span>
          </div>
        </div>
      </div>

      {/* ========================================================================= */}
      {/* 2. SEARCH, MULTI-DIMENSIONAL FILTERS & SORT SUITE                         */}
      {/* ========================================================================= */}
      <div className="bg-white rounded-2xl p-5 sm:p-6 border border-slate-200/90 shadow-xs space-y-4">
        {/* Controls Header */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100">
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-blue-900" />
            <h3 className="font-extrabold text-slate-900 text-sm sm:text-base">
              Filter &amp; Search Submissions
            </h3>
          </div>

          <div className="flex items-center gap-2">
            {/* View Layout Toggle: Table vs Grid */}
            <div className="flex items-center bg-slate-100 p-0.5 rounded-xl border border-slate-200">
              <button
                type="button"
                id="btn-layout-table"
                onClick={() => setViewLayout('table')}
                className={`p-1.5 rounded-lg text-xs font-bold transition cursor-pointer ${
                  viewLayout === 'table' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-500 hover:text-slate-900'
                }`}
                title="Detailed Table View"
              >
                <List className="w-4 h-4" />
              </button>
              <button
                type="button"
                id="btn-layout-grid"
                onClick={() => setViewLayout('grid')}
                className={`p-1.5 rounded-lg text-xs font-bold transition cursor-pointer ${
                  viewLayout === 'grid' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-500 hover:text-slate-900'
                }`}
                title="Card Grid View"
              >
                <LayoutGrid className="w-4 h-4" />
              </button>
            </div>

            {hasActiveFilters && (
              <button
                type="button"
                id="btn-reset-filters"
                onClick={resetFilters}
                className="flex items-center gap-1 px-3 py-1.5 text-xs text-rose-600 hover:text-rose-700 font-bold hover:bg-rose-50 rounded-xl transition cursor-pointer border border-rose-200"
              >
                <RotateCcw className="w-3 h-3" />
                <span>Reset Filters</span>
              </button>
            )}
          </div>
        </div>

        {/* Search Bar */}
        <div className="relative">
          <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
          <input
            type="text"
            id="input-search-submissions"
            placeholder="Search by subject code, name, department, evaluator, faculty in-charge, batch year..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-xs sm:text-sm font-medium text-slate-900 placeholder:text-slate-400 focus:ring-2 focus:ring-blue-900 focus:bg-white focus:outline-hidden transition"
          />
          {searchTerm && (
            <button
              type="button"
              onClick={() => setSearchTerm('')}
              className="absolute right-3.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
            >
              <span className="text-xs font-bold">✕</span>
            </button>
          )}
        </div>

        {/* Filter Dropdowns Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
          {/* 1. Department */}
          <div>
            <label className="block text-[10px] font-extrabold text-slate-600 uppercase tracking-wider mb-1">
              Department
            </label>
            <select
              id="filter-select-department"
              value={selectedDept}
              onChange={(e) => setSelectedDept(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-blue-900 focus:outline-hidden cursor-pointer"
            >
              <option value="">All Departments</option>
              {departmentOptions.map((dept) => (
                <option key={dept} value={dept}>
                  {dept}
                </option>
              ))}
            </select>
          </div>

          {/* 2. Semester */}
          <div>
            <label className="block text-[10px] font-extrabold text-slate-600 uppercase tracking-wider mb-1">
              Semester
            </label>
            <select
              id="filter-select-semester"
              value={selectedSem}
              onChange={(e) => setSelectedSem(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-blue-900 focus:outline-hidden cursor-pointer"
            >
              <option value="">All Semesters</option>
              {SEMESTER_OPTIONS.map((sem) => (
                <option key={sem} value={sem}>
                  {sem}
                </option>
              ))}
            </select>
          </div>

          {/* 3. Exam Type */}
          <div>
            <label className="block text-[10px] font-extrabold text-slate-600 uppercase tracking-wider mb-1">
              Exam Type
            </label>
            <select
              id="filter-select-exam-type"
              value={selectedExamType}
              onChange={(e) => setSelectedExamType(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-blue-900 focus:outline-hidden cursor-pointer"
            >
              <option value="">All Exam Types</option>
              {EXAM_TYPE_OPTIONS.map((et) => (
                <option key={et} value={et}>
                  {et}
                </option>
              ))}
            </select>
          </div>

          {/* 4. Approval Status */}
          <div>
            <label className="block text-[10px] font-extrabold text-slate-600 uppercase tracking-wider mb-1">
              Approval Status
            </label>
            <select
              id="filter-select-status"
              value={selectedStatus}
              onChange={(e) => setSelectedStatus(e.target.value)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-blue-900 focus:outline-hidden cursor-pointer"
            >
              <option value="">All Statuses</option>
              <option value="APPROVED">Approved by Exam Cell</option>
              <option value="PENDING_VERIFICATION">Pending Verification</option>
            </select>
          </div>

          {/* 5. Sort By */}
          <div>
            <label className="block text-[10px] font-extrabold text-slate-600 uppercase tracking-wider mb-1">
              Sort By
            </label>
            <select
              id="filter-select-sort-by"
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value as any)}
              className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-800 focus:ring-2 focus:ring-blue-900 focus:outline-hidden cursor-pointer"
            >
              <option value="newest">🕒 Newest Submitted First</option>
              <option value="oldest">⏳ Oldest Submitted First</option>
              <option value="subject_asc">🔤 Subject Code (A-Z)</option>
              <option value="subject_desc">🔡 Subject Code (Z-A)</option>
              <option value="pass_high">📈 Pass Rate (Highest First)</option>
              <option value="pass_low">📉 Pass Rate (Lowest First)</option>
              <option value="students_high">👥 Total Students (Highest)</option>
            </select>
          </div>
        </div>
      </div>

      {/* ========================================================================= */}
      {/* 4. RESULTS DISPLAY: TABLE VIEW OR GRID VIEW                              */}
      {/* ========================================================================= */}
      {filteredSheets.length === 0 ? (
        <div className="bg-white rounded-2xl p-12 border border-slate-200 text-center space-y-4 shadow-2xs">
          <div className="w-14 h-14 rounded-2xl bg-slate-100 text-slate-400 flex items-center justify-center mx-auto">
            <Search className="w-7 h-7" />
          </div>
          <div className="max-w-md mx-auto">
            <h4 className="text-base font-extrabold text-slate-900">No marksheets match your search</h4>
            <p className="text-xs text-slate-500 mt-1 leading-relaxed">
              We couldn&apos;t find any marksheets matching the current filter criteria. Try clearing some filters or searching with a different term.
            </p>
            {hasActiveFilters && (
              <button
                type="button"
                onClick={resetFilters}
                className="mt-4 px-4 py-2 bg-[#0f284e] hover:bg-blue-900 text-white rounded-xl text-xs font-bold transition cursor-pointer"
              >
                Clear All Filters
              </button>
            )}
          </div>
        </div>
      ) : viewLayout === 'table' ? (
        /* Detailed Table Layout */
        <div className="bg-white rounded-2xl border border-slate-200 shadow-2xs overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-[#0b1e36] text-white font-bold uppercase text-[11px] tracking-wider border-b border-slate-800">
                  <th className="py-3 px-3.5">Subject</th>
                  <th className="py-3 px-3">Dept &amp; Semester</th>
                  <th className="py-3 px-3">Exam Type</th>
                  <th className="py-3 px-3">Faculty / Evaluator</th>
                  <th className="py-3 px-3 text-center">Students &amp; Pass %</th>
                  <th className="py-3 px-3">Submitted At</th>
                  <th className="py-3 px-3 text-center">Status</th>
                  <th className="py-3 px-3.5 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredSheets.map((sheet) => {
                  const metrics = computeMetrics(sheet.students, sheet.meta.passMark);
                  const isApproved = sheet.status === 'APPROVED';
                  const isUnlocked = sheet.meta.isUnlocked;

                  return (
                    <tr key={sheet.id} className="hover:bg-slate-50/90 transition group">
                      {/* Subject Info */}
                      <td className="py-3 px-3.5">
                        <div className="flex items-center gap-1.5">
                          <span className="font-mono font-bold text-blue-950 bg-blue-50 px-2 py-0.5 rounded border border-blue-200 text-xs">
                            {sheet.meta.subjectCode}
                          </span>
                          {sheet.meta.isUnlocked && (
                            <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold border border-emerald-300">
                              Unlocked
                            </span>
                          )}
                        </div>
                        <div className="font-bold text-slate-900 mt-1 max-w-[220px] truncate" title={sheet.meta.subjectName}>
                          {sheet.meta.subjectName}
                        </div>
                      </td>

                      {/* Dept & Sem */}
                      <td className="py-3 px-3">
                        <div className="font-bold text-slate-800">{sheet.meta.department}</div>
                        <div className="text-[11px] text-blue-900 font-semibold">{sheet.meta.semester}</div>
                        <div className="text-[10px] text-slate-400">Batch {sheet.meta.academicYear}</div>
                      </td>

                      {/* Exam Type */}
                      <td className="py-3 px-3">
                        <span className="bg-amber-100 text-amber-950 border border-amber-300/80 px-2.5 py-0.5 rounded-md text-[11px] font-bold">
                          {sheet.meta.examType}
                        </span>
                      </td>

                      {/* Faculty / Evaluator */}
                      <td className="py-3 px-3">
                        <div className="font-semibold text-slate-800">
                          {sheet.meta.facultyInCharge || <span className="text-slate-400 italic">Faculty</span>}
                        </div>
                        <div className="text-[11px] text-slate-500">
                          Eval: {sheet.meta.evaluatorName || <span className="text-slate-400 italic">Pending</span>}
                        </div>
                      </td>

                      {/* Students & Pass Rate */}
                      <td className="py-3 px-3 text-center">
                        <div className="font-black text-slate-900">{sheet.students.length} students</div>
                        <div className="mt-0.5">
                          <span
                            className={`text-[10px] font-black px-2 py-0.5 rounded-full inline-block ${
                              metrics.passPercentage >= 75
                                ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                : metrics.passPercentage >= 50
                                ? 'bg-blue-100 text-blue-800 border border-blue-300'
                                : 'bg-rose-100 text-rose-800 border border-rose-300'
                            }`}
                          >
                            {metrics.passPercentage.toFixed(1)}% Pass
                          </span>
                        </div>
                      </td>

                      {/* Submitted At */}
                      <td className="py-3 px-3 text-slate-600 text-[11px] whitespace-nowrap">
                        <div className="font-medium">{sheet.submittedAt}</div>
                        {sheet.approvedAt && (
                          <div className="text-[10px] text-emerald-700">Appr: {sheet.approvedAt}</div>
                        )}
                      </td>

                      {/* Status */}
                      <td className="py-3 px-3 text-center">
                        <span
                          className={`text-[10px] px-2.5 py-1 rounded-full font-bold uppercase tracking-wider inline-flex items-center gap-1 ${
                            isApproved
                              ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                              : 'bg-amber-100 text-amber-800 border border-amber-300'
                          }`}
                        >
                          {isApproved ? (
                            <>
                              <CheckCircle className="w-3 h-3 text-emerald-600" />
                              <span>Approved</span>
                            </>
                          ) : (
                            <>
                              <Clock className="w-3 h-3 text-amber-600" />
                              <span>Pending</span>
                            </>
                          )}
                        </span>
                      </td>

                      {/* Action Buttons */}
                      <td className="py-3 px-3.5 text-right whitespace-nowrap">
                        <div className="flex items-center justify-end gap-1.5">
                          {/* Inspect / View Full Details */}
                          <button
                            type="button"
                            id={`btn-view-submission-${sheet.id}`}
                            onClick={() => onSelectSheet(sheet)}
                            className="px-2.5 py-1.5 bg-[#0f284e] hover:bg-blue-900 text-white rounded-lg text-xs font-bold transition cursor-pointer shadow-2xs flex items-center gap-1"
                            title="Inspect full marksheet with all student rows"
                          >
                            <Eye className="w-3.5 h-3.5" />
                            <span>View</span>
                          </button>

                          {/* Print / Preview */}
                          <button
                            type="button"
                            id={`btn-print-submission-${sheet.id}`}
                            onClick={() => onPrintPreview(sheet)}
                            className="p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-bold transition cursor-pointer border border-slate-300"
                            title="Print Preview Modal"
                          >
                            <Printer className="w-3.5 h-3.5" />
                          </button>

                          {/* Quick PDF Download */}
                          <button
                            type="button"
                            id={`btn-pdf-submission-${sheet.id}`}
                            onClick={() => onDownloadPDF(sheet)}
                            className="p-1.5 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded-lg text-xs font-bold transition cursor-pointer"
                            title="Download PDF"
                          >
                            <Download className="w-3.5 h-3.5" />
                          </button>

                          {/* Quick XLS Download */}
                          <button
                            type="button"
                            id={`btn-xls-submission-${sheet.id}`}
                            onClick={() => onDownloadXLS(sheet)}
                            className="p-1.5 bg-emerald-700 hover:bg-emerald-600 text-white rounded-lg text-xs font-bold transition cursor-pointer"
                            title="Download Excel (.xlsx)"
                          >
                            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-100" />
                          </button>

                          {/* Quick Word Download */}
                          <button
                            type="button"
                            id={`btn-word-submission-${sheet.id}`}
                            onClick={() => onDownloadWord(sheet)}
                            className="p-1.5 bg-blue-700 hover:bg-blue-600 text-white rounded-lg text-xs font-bold transition cursor-pointer"
                            title="Download Word (.docx)"
                          >
                            <FileText className="w-3.5 h-3.5 text-blue-100" />
                          </button>

                          {/* Direct Approve (if not yet approved) */}
                          {!isApproved && (
                            <button
                              type="button"
                              id={`btn-approve-submission-${sheet.id}`}
                              onClick={() => onApproveSheet(sheet.id)}
                              className="p-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-lg text-xs font-bold transition cursor-pointer border border-emerald-300"
                              title="Approve Marksheet"
                            >
                              <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                            </button>
                          )}

                          {/* Delete Marksheet */}
                          <button
                            type="button"
                            id={`btn-delete-submission-${sheet.id}`}
                            onClick={() => onDeleteSheet(sheet)}
                            className="p-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-xs font-bold transition cursor-pointer border border-rose-200"
                            title="Delete marksheet from Firestore"
                          >
                            <Trash2 className="w-3.5 h-3.5 text-rose-600" />
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
      ) : (
        /* Card Grid Layout */
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {filteredSheets.map((sheet) => {
            const metrics = computeMetrics(sheet.students, sheet.meta.passMark);
            const isApproved = sheet.status === 'APPROVED';

            return (
              <div
                key={sheet.id}
                className="bg-white rounded-2xl p-5 border border-slate-200 shadow-2xs hover:shadow-md transition flex flex-col justify-between space-y-4"
              >
                <div>
                  {/* Card Header */}
                  <div className="flex items-start justify-between gap-2 pb-3 border-b border-slate-100">
                    <div>
                      <div className="flex items-center gap-1.5">
                        <span className="font-mono font-bold text-blue-950 bg-blue-50 px-2 py-0.5 rounded border border-blue-200 text-xs">
                          {sheet.meta.subjectCode}
                        </span>
                        <span className="bg-amber-100 text-amber-950 border border-amber-300/80 px-2 py-0.5 rounded text-[10px] font-bold">
                          {sheet.meta.examType}
                        </span>
                      </div>
                      <h4 className="font-black text-slate-900 text-sm mt-1.5 leading-snug">
                        {sheet.meta.subjectName}
                      </h4>
                    </div>

                    <span
                      className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase tracking-wider shrink-0 ${
                        isApproved
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                          : 'bg-amber-100 text-amber-800 border border-amber-300'
                      }`}
                    >
                      {isApproved ? 'Approved' : 'Pending'}
                    </span>
                  </div>

                  {/* Metadata Matrix */}
                  <div className="grid grid-cols-2 gap-2 text-xs text-slate-600 py-3 border-b border-slate-100">
                    <div>
                      <span className="text-[10px] text-slate-400 block uppercase font-bold">Department</span>
                      <strong className="text-slate-800 truncate block">{sheet.meta.department}</strong>
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-400 block uppercase font-bold">Semester</span>
                      <strong className="text-blue-900 block">{sheet.meta.semester}</strong>
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-400 block uppercase font-bold">Faculty In-Charge</span>
                      <span className="text-slate-800 truncate block">{sheet.meta.facultyInCharge || 'Faculty'}</span>
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-400 block uppercase font-bold">Submitted Date</span>
                      <span className="text-slate-700 block text-[11px]">{sheet.submittedAt}</span>
                    </div>
                  </div>

                  {/* Quick Stats Pill */}
                  <div className="flex items-center justify-between bg-slate-50 rounded-xl p-2.5 mt-3 text-xs">
                    <div>
                      <span className="text-[10px] text-slate-500 uppercase font-bold block">Assessed</span>
                      <strong className="text-slate-900 font-black">{sheet.students.length} Students</strong>
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500 uppercase font-bold block">Pass Rate</span>
                      <strong
                        className={`font-black ${
                          metrics.passPercentage >= 75
                            ? 'text-emerald-700'
                            : metrics.passPercentage >= 50
                            ? 'text-blue-700'
                            : 'text-rose-700'
                        }`}
                      >
                        {metrics.passPercentage.toFixed(1)}%
                      </strong>
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500 uppercase font-bold block">Pass / Fail</span>
                      <strong className="text-slate-900 font-bold">
                        {metrics.passCount} / {metrics.failCount}
                      </strong>
                    </div>
                  </div>
                </div>

                {/* Card Action Buttons */}
                <div className="flex items-center justify-between gap-1.5 pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => onSelectSheet(sheet)}
                    className="flex-1 py-1.5 px-3 bg-[#0f284e] hover:bg-blue-900 text-white rounded-xl text-xs font-bold transition cursor-pointer flex items-center justify-center gap-1.5 shadow-2xs"
                  >
                    <Eye className="w-3.5 h-3.5" />
                    <span>View Sheet</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => onDownloadPDF(sheet)}
                    className="p-2 bg-amber-500 hover:bg-amber-400 text-slate-950 rounded-xl transition cursor-pointer"
                    title="Download PDF"
                  >
                    <Download className="w-3.5 h-3.5" />
                  </button>

                  <button
                    type="button"
                    onClick={() => onDownloadXLS(sheet)}
                    className="p-2 bg-emerald-700 hover:bg-emerald-600 text-white rounded-xl transition cursor-pointer"
                    title="Download Excel"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" />
                  </button>

                  <button
                    type="button"
                    onClick={() => onDownloadWord(sheet)}
                    className="p-2 bg-blue-700 hover:bg-blue-600 text-white rounded-xl transition cursor-pointer"
                    title="Download Word (.docx)"
                  >
                    <FileText className="w-3.5 h-3.5" />
                  </button>

                  <button
                    type="button"
                    onClick={() => onDeleteSheet(sheet)}
                    className="p-2 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-xl transition cursor-pointer border border-rose-200"
                    title="Delete Marksheet"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
