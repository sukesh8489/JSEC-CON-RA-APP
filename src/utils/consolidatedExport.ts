import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  VerticalAlign,
  BorderStyle,
  ImageRun,
  ShadingType,
  PageOrientation,
} from 'docx';
import JSZip from 'jszip';
import { SubmittedSheetRecord } from './submittedSheets';
import { BatchStudent } from './studentBatches';
import { COLLEGE_LOGO_BASE64 } from '../assets/logoBase64';
import { NAAC_LOGO_BASE64 } from '../assets/naacLogoBase64';
import { renderSingleSubjectMarksheetToPDF } from './pdfExport';
import { buildSubjectMarksheetWorksheet, sanitizeExcelFor2007 } from './excelExport';
import { computeMetrics } from './gradeCalculator';

function base64ToUint8Array(base64Str: string): Uint8Array {
  try {
    const clean = base64Str.replace(/^data:image\/\w+;base64,/, '').replace(/\s+/g, '');
    const binaryString = atob(clean);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
  } catch (err) {
    console.warn('Could not decode base64 image to Uint8Array:', err);
    return new Uint8Array();
  }
}

export interface ConsolidatedSubjectInfo {
  code: string;
  name: string;
  shortName: string;
  maxMarks: number;
  passMark: number;
  examType: string;
  sheetId: string;
  facultyInCharge: string;
}

export interface ConsolidatedStudentSubjectMark {
  attendance: string;
  marksObtained: number | string;
  grade: string;
  result: string;
  maxMarks: number;
  passMark: number;
  isPass: boolean;
}

export interface ConsolidatedStudentRow {
  sno: number;
  regNo: string;
  name: string;
  subjects: Record<string, ConsolidatedStudentSubjectMark>;
  totalObtained: number;
  totalMax: number;
  percentage: number;
  passedSubjectsCount: number;
  failedSubjectsCount: number;
  overallResult: 'P' | 'F';
}

export interface SubjectAnalysisMetrics {
  totalStrength: number;
  totalPresent: number;
  totalAbsent: number;
  totalPass: number;
  totalFail: number;
  highestMarks: number;
  passPercentageStr: number; // Based on strength: (Pass / Strength) * 100
  passPercentageApp: number; // Based on appeared: (Pass / Appeared) * 100
  facultyInCharge: string;
}

export interface ConsolidatedData {
  department: string;
  semester: string;
  academicYear: string;
  examType: string;
  assessmentDate: string;
  section: string;
  subjects: ConsolidatedSubjectInfo[];
  students: ConsolidatedStudentRow[];
  totalStudentStrength: number;
  totalPresent: number;
  totalAbsent: number;
  passInAllSubjectsCount: number;
  overallPassPercentage: number;
  subjectMetrics: Record<string, SubjectAnalysisMetrics>;
  matchingSheets?: SubmittedSheetRecord[];
}

export type ExportDesignTheme = 'dark-navy';

export interface ExportThemeConfig {
  id: 'dark-navy';
  name: string;
  badge: string;
  isDark: boolean;
  bannerBg: [number, number, number];
  bannerText: [number, number, number];
  headFill: [number, number, number];
  headText: [number, number, number];
  summaryLabelBg: [number, number, number];
  summaryLabelText: [number, number, number];
  borderColor: [number, number, number];
  alternateRowBg: [number, number, number];
  excelHeaderFill: string;
  excelAccentFill: string;
  wordHeaderBg: string;
}

export const EXPORT_DESIGN_THEMES: Record<string, ExportThemeConfig> = {
  'dark-navy': {
    id: 'dark-navy',
    name: 'Official Institutional Dark & Light',
    badge: 'Standard Theme',
    isDark: true,
    bannerBg: [11, 30, 54], // #0B1E36 (Exact match to normal marksheet)
    bannerText: [255, 255, 255],
    headFill: [11, 30, 54], // #0B1E36
    headText: [255, 255, 255],
    summaryLabelBg: [11, 30, 54], // #0B1E36
    summaryLabelText: [255, 255, 255],
    borderColor: [203, 213, 225], // #CBD5E1 Slate-300
    alternateRowBg: [248, 250, 252], // #F8FAFC Slate-50
    excelHeaderFill: 'FF0B1E36',
    excelAccentFill: 'FF0F284E',
    wordHeaderBg: '0B1E36',
  },
};

/**
 * Convert Semester string to Roman Numeral "Year / Semester" representation
 * E.g., "Semester 3" -> "II / III", "Semester 7" -> "IV / VII" (matching institutional format in photo)
 */
export function getYearSemRoman(semStr: string): string {
  const num = parseInt(String(semStr).replace(/[^0-9]/g, '') || '1', 10);
  const romans = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII'];
  const yearNum = Math.ceil(num / 2);
  const yearRoman = romans[yearNum] || 'I';
  const semRoman = romans[num] || 'I';
  return `${yearRoman} / ${semRoman}`;
}

/**
 * Generate short acronym or abbreviation for subject name
 * E.g., "Principles of Management" -> "POM", "Democracy and Good Governance" -> "DGG"
 */
export function getSubjectAcronym(name: string, code: string): string {
  if (!name || name.trim() === code.trim()) return code;
  const clean = name.replace(/\(.*?\)/g, '').trim();
  const words = clean
    .split(/[\s_-]+/)
    .filter((w) => w.length > 0 && !['and', '&', 'for', 'the', 'in', 'to'].includes(w.toLowerCase()));

  if (words.length >= 2 && words.length <= 5) {
    return words.map((w) => w[0].toUpperCase()).join('');
  }
  if (words.length === 1) {
    return words[0].substring(0, 4).toUpperCase();
  }
  if (words.length > 5) {
    return words.slice(0, 4).map((w) => w[0].toUpperCase()).join('');
  }
  return code;
}

/**
 * Build consolidated dataset across all subjects with full result analysis metrics
 */
export function buildConsolidatedData(
  department: string,
  semester: string,
  examTypeFilter: string,
  matchingSheets: SubmittedSheetRecord[],
  batchStudents: BatchStudent[] = []
): ConsolidatedData {
  // 1. Identify all unique subjects in order
  const subjectsMap = new Map<string, ConsolidatedSubjectInfo>();
  matchingSheets.forEach((s) => {
    const code = (s.meta.subjectCode || '').trim().toUpperCase();
    if (code && !subjectsMap.has(code)) {
      const subName = (s.meta.subjectName || '').trim() || code;
      const shortName = getSubjectAcronym(subName, code);
      subjectsMap.set(code, {
        code,
        name: subName,
        shortName,
        maxMarks: Number(s.meta.maxMarks) || 100,
        passMark: Number(s.meta.passMark) || 50,
        examType: s.meta.examType || 'Continuous Assessment',
        sheetId: s.id,
        facultyInCharge: s.meta.facultyInCharge || s.meta.evaluatorName || 'Staff Member',
      });
    }
  });

  const subjects = Array.from(subjectsMap.values());

  // 2. Identify all students (from batch roster + union with any marksheets)
  const studentMap = new Map<string, { regNo: string; name: string }>();

  // Add from batch roster first to preserve official order
  batchStudents.forEach((bs) => {
    const reg = (bs.regNo || '').trim();
    if (reg) {
      studentMap.set(reg.toUpperCase(), {
        regNo: reg,
        name: (bs.name || '').trim().toUpperCase(),
      });
    }
  });

  // Union with any student records in matching marksheets
  matchingSheets.forEach((s) => {
    (s.students || []).forEach((st) => {
      const reg = (st.regNo || (st as any).registerNumber || '').trim();
      if (reg) {
        const key = reg.toUpperCase();
        if (!studentMap.has(key)) {
          studentMap.set(key, {
            regNo: reg,
            name: (st.name || (st as any).studentName || '').trim().toUpperCase(),
          });
        }
      }
    });
  });

  const studentList = Array.from(studentMap.values());
  // Sort naturally by Register Number
  studentList.sort((a, b) => a.regNo.localeCompare(b.regNo, undefined, { numeric: true }));

  // 3. For each student, compute mark in each subject and totals
  const studentRows: ConsolidatedStudentRow[] = [];
  let passInAllCount = 0;

  studentList.forEach((st, idx) => {
    const subjectMarks: Record<string, ConsolidatedStudentSubjectMark> = {};
    let totalObt = 0;
    let totalMax = 0;
    let passedCount = 0;
    let failedCount = 0;

    subjects.forEach((sub) => {
      // Find marksheet for this subject
      const sheet = matchingSheets.find(
        (ms) => (ms.meta.subjectCode || '').trim().toUpperCase() === sub.code
      );

      const studentInSheet = sheet?.students?.find(
        (s) => (s.regNo || (s as any).registerNumber || '').trim().toUpperCase() === st.regNo.toUpperCase()
      );

      if (studentInSheet) {
        const attendance = String(studentInSheet.attendance || 'PRESENT').toUpperCase();
        const isAbsent = attendance === 'ABSENT';
        const rawMark = studentInSheet.marksObtained;
        const numMark = typeof rawMark === 'number' ? rawMark : parseFloat(String(rawMark));
        const hasNumMark = !isNaN(numMark) && !isAbsent;

        const maxM = Number(studentInSheet.maxMarks ?? sub.maxMarks) || 100;
        const passM = Number(sheet?.meta?.passMark ?? sub.passMark) || 50;
        const isPass = hasNumMark && numMark >= passM && studentInSheet.result === 'PASS';

        if (isAbsent) {
          subjectMarks[sub.code] = {
            attendance: 'ABSENT',
            marksObtained: 'AB',
            grade: 'AB',
            result: 'FAIL',
            maxMarks: maxM,
            passMark: passM,
            isPass: false,
          };
          failedCount++;
          totalMax += maxM;
        } else if (hasNumMark) {
          subjectMarks[sub.code] = {
            attendance: 'PRESENT',
            marksObtained: numMark,
            grade: studentInSheet.grade || (isPass ? 'B' : 'RA'),
            result: isPass ? 'PASS' : 'FAIL',
            maxMarks: maxM,
            passMark: passM,
            isPass,
          };
          totalObt += numMark;
          totalMax += maxM;
          if (isPass) {
            passedCount++;
          } else {
            failedCount++;
          }
        } else {
          subjectMarks[sub.code] = {
            attendance: 'PRESENT',
            marksObtained: '-',
            grade: '-',
            result: '-',
            maxMarks: maxM,
            passMark: passM,
            isPass: false,
          };
        }
      } else {
        subjectMarks[sub.code] = {
          attendance: '-',
          marksObtained: '-',
          grade: '-',
          result: '-',
          maxMarks: sub.maxMarks,
          passMark: sub.passMark,
          isPass: false,
        };
      }
    });

    const percentage = totalMax > 0 ? (totalObt / totalMax) * 100 : 0;
    // Anna University Rule: P (Pass) only if all registered subjects are passed, else F (Fail / Arrear)
    const overallResult: 'P' | 'F' =
      subjects.length > 0 && passedCount === subjects.length && failedCount === 0 ? 'P' : 'F';

    if (overallResult === 'P') {
      passInAllCount++;
    }

    studentRows.push({
      sno: idx + 1,
      regNo: st.regNo,
      name: st.name,
      subjects: subjectMarks,
      totalObtained: totalObt,
      totalMax,
      percentage,
      passedSubjectsCount: passedCount,
      failedSubjectsCount: failedCount,
      overallResult,
    });
  });

  const totalStrength = studentRows.length;

  // 4. Calculate Detailed Subject-wise Result Analysis Metrics (as shown in images)
  const subjectMetrics: ConsolidatedData['subjectMetrics'] = {};
  subjects.forEach((sub) => {
    let appeared = 0;
    let absent = 0;
    let pass = 0;
    let fail = 0;
    let highest = 0;

    studentRows.forEach((r) => {
      const sm = r.subjects[sub.code];
      if (sm) {
        if (sm.attendance === 'ABSENT' || sm.marksObtained === 'AB') {
          absent++;
          fail++;
        } else if (typeof sm.marksObtained === 'number') {
          appeared++;
          if (sm.marksObtained > highest) {
            highest = sm.marksObtained;
          }
          if (sm.isPass) {
            pass++;
          } else {
            fail++;
          }
        }
      }
    });

    // Pass % (Str) based on Strength: (Total Pass / Total Strength) * 100
    const passPercentageStr = totalStrength > 0 ? Math.round((pass / totalStrength) * 100) : 0;
    // Pass % (App) based on Appeared: (Total Pass / Appeared) * 100
    const passPercentageApp = appeared > 0 ? Math.round((pass / appeared) * 100) : 0;

    subjectMetrics[sub.code] = {
      totalStrength,
      totalPresent: appeared,
      totalAbsent: absent,
      totalPass: pass,
      totalFail: fail,
      highestMarks: highest,
      passPercentageStr,
      passPercentageApp,
      facultyInCharge: sub.facultyInCharge,
    };
  });

  const overallPassPercentage = totalStrength > 0 ? (passInAllCount / totalStrength) * 100 : 0;

  const firstSheetMeta = matchingSheets[0]?.meta;
  const academicYear = firstSheetMeta?.academicYear || '2024-2028';
  const examType = examTypeFilter || firstSheetMeta?.examType || 'Continuous Assessment';
  const assessmentDate = firstSheetMeta?.assessmentDate || new Date().toISOString().split('T')[0];

  return {
    department,
    semester,
    academicYear,
    examType,
    assessmentDate,
    section: 'A',
    subjects,
    students: studentRows,
    totalStudentStrength: totalStrength,
    totalPresent: studentRows.filter((s) => Object.values(s.subjects).some((m) => m.attendance === 'PRESENT')).length,
    totalAbsent: studentRows.filter((s) => Object.values(s.subjects).every((m) => m.attendance === 'ABSENT')).length,
    passInAllSubjectsCount: passInAllCount,
    overallPassPercentage,
    subjectMetrics,
    matchingSheets,
  };
}

/**
 * Format Excel Tab Name matching the image format: e.g. "AD-3-S-CRA" or "AIDS-7 SEM"
 */
function getSheetTabName(dept: string, sem: string): string {
  const shortDept = (dept.split(' - ')[0] || dept).replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  const semNum = sem.replace(/[^0-9]/g, '') || '1';
  return `${shortDept}-${semNum} SEM`.substring(0, 31);
}

/**
 * 1. EXPORT CONSOLIDATED MARKSHEET TO EXCEL (.xlsx)
 * - Sheet 1: Consolidated Marksheet (Hairline borders, rich navy headers, tight aligned columns)
 * - Sheet 2: "Subwise" (The exact Subject Wise Analysis table shown in uploaded 2nd page photo)
 */
export async function exportConsolidatedToExcel(
  data: ConsolidatedData,
  theme: ExportDesignTheme = 'dark-navy',
  matchingSheetsParam?: SubmittedSheetRecord[]
): Promise<void> {
  const themeConfig = EXPORT_DESIGN_THEMES[theme] || EXPORT_DESIGN_THEMES['dark-navy'];
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Jaya Sakthi Engineering College';
  workbook.lastModifiedBy = 'Central Examination Cell';
  workbook.created = new Date();
  workbook.modified = new Date();

  // Thin hairline border
  const hairlineBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  const darkNavy = themeConfig.excelHeaderFill;
  const darkSlate = themeConfig.excelAccentFill;
  const redFailColor = 'FFDC2626';
  const greenPassColor = 'FF15803D';

  // =========================================================================
  // SHEET 1: CONSOLIDATED STUDENT MARKSHEET (COMPACT & PROPORTIONATE)
  // =========================================================================
  const sheetName1 = getSheetTabName(data.department, data.semester);
  const worksheet1 = workbook.addWorksheet(sheetName1, {
    views: [{ showGridLines: true }],
    pageSetup: {
      paperSize: 9, // A4
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.3, right: 0.3, top: 0.35, bottom: 0.35, header: 0.2, footer: 0.2 },
    },
  });

  const subCount = data.subjects.length;
  const totalCols1 = 3 + subCount + 3;
  const lastColLetter1 = getColumnLetter(totalCols1);

  // FULL-SCREEN RESPONSIVE COLUMN SIZING (Fills landscape screen gracefully)
  let rollColW = 20;
  let nameColW = 34;
  let subColW = 14;
  let totColW = 12;
  let pctColW = 11;
  let resColW = 10;

  if (subCount <= 1) {
    rollColW = 24;
    nameColW = 42;
    subColW = 28;
    totColW = 14;
    pctColW = 12;
    resColW = 12;
  } else if (subCount === 2) {
    rollColW = 22;
    nameColW = 38;
    subColW = 20;
    totColW = 13;
    pctColW = 12;
    resColW = 11;
  } else if (subCount === 3) {
    rollColW = 21;
    nameColW = 35;
    subColW = 17;
    totColW = 12;
    pctColW = 11;
    resColW = 10;
  } else if (subCount === 4) {
    rollColW = 20;
    nameColW = 32;
    subColW = 15;
    totColW = 11;
    pctColW = 10;
    resColW = 10;
  } else {
    rollColW = 19;
    nameColW = Math.max(26, 30 - (subCount - 5));
    subColW = Math.max(12, Math.floor(65 / subCount));
    totColW = 10.5;
    pctColW = 10;
    resColW = 9.5;
  }

  worksheet1.getColumn(1).width = 7.5; // A: S.NO
  worksheet1.getColumn(2).width = rollColW; // B: ROLL NO
  worksheet1.getColumn(3).width = nameColW; // C: STUDENT NAME
  for (let i = 0; i < subCount; i++) {
    worksheet1.getColumn(4 + i).width = subColW; // D..: Subject marks
  }
  worksheet1.getColumn(4 + subCount).width = totColW; // TOTAL
  worksheet1.getColumn(5 + subCount).width = pctColW; // %
  worksheet1.getColumn(6 + subCount).width = resColW; // RESULT (P/F)

  // Embed College Logo (Left) and NAAC Logo (Right) in Sheet 1
  try {
    if (COLLEGE_LOGO_BASE64) {
      const cleanBase64 = COLLEGE_LOGO_BASE64.replace(/^data:image\/\w+;base64,/, '');
      const collegeImgId = workbook.addImage({
        base64: cleanBase64,
        extension: 'png',
      });
      worksheet1.addImage(collegeImgId, {
        tl: { col: 0.1, row: 0.1 },
        ext: { width: 62, height: 62 },
      });
    }
    if (NAAC_LOGO_BASE64) {
      const cleanNaacBase64 = NAAC_LOGO_BASE64.replace(/^data:image\/\w+;base64,/, '');
      const naacExt = NAAC_LOGO_BASE64.startsWith('data:image/jpeg') ? 'jpeg' : 'png';
      const naacImgId = workbook.addImage({
        base64: cleanNaacBase64,
        extension: naacExt,
      });
      worksheet1.addImage(naacImgId, {
        tl: { col: totalCols1 - 1.15, row: 0.1 },
        ext: { width: 75, height: 60 },
      });
    }
  } catch (err) {
    console.warn('Could not add logos to Excel Sheet 1:', err);
  }

  // Institutional Header (Rich Dark Navy & Gold)
  worksheet1.mergeCells(`A1:${lastColLetter1}1`);
  const c1_1 = worksheet1.getCell('A1');
  c1_1.value = 'JAYA SAKTHI ENGINEERING COLLEGE';
  c1_1.font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: darkNavy } };
  c1_1.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet1.getRow(1).height = 25;

  worksheet1.mergeCells(`A2:${lastColLetter1}2`);
  const c1_2 = worksheet1.getCell('A2');
  c1_2.value = "St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu";
  c1_2.font = { name: 'Arial', size: 8.5, bold: true, color: { argb: 'FF334155' } };
  c1_2.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet1.getRow(2).height = 15;

  worksheet1.mergeCells(`A3:${lastColLetter1}3`);
  const c1_3 = worksheet1.getCell('A3');
  c1_3.value = 'Approved by AICTE, New Delhi & Affiliated to Anna University, Chennai';
  c1_3.font = { name: 'Arial', size: 8, color: { argb: 'FF64748B' } };
  c1_3.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet1.getRow(3).height = 15;

  worksheet1.getRow(4).height = 4;

  // Row 4: Decorative dual rule (Navy top border + Golden amber bottom border)
  worksheet1.getRow(4).height = 6;
  for (let c = 1; c <= totalCols1; c++) {
    const cell = worksheet1.getRow(4).getCell(c);
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0F284E' } },
      bottom: { style: 'thin', color: { argb: 'FFD97706' } },
    };
  }

  // Row 5: Title Banner (Dark Navy with White Bold Text)
  worksheet1.mergeCells(`A5:${lastColLetter1}5`);
  const c1_5 = worksheet1.getCell('A5');
  c1_5.value = `ALL SUBJECTS CONSOLIDATED ASSESSMENT MARKSHEET — ${data.examType.toUpperCase()}`;
  c1_5.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  c1_5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: darkNavy } };
  c1_5.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet1.getRow(5).height = 24;

  // Row 6: Spacer
  worksheet1.getRow(6).height = 4;

  // Rows 7 - 9: Official 3-Row Metadata Matrix Box (Matching uploaded reference design)
  const midCol1 = Math.floor(totalCols1 / 2);
  const midColLetter1 = getColumnLetter(midCol1);
  const midNextColLetter1 = getColumnLetter(midCol1 + 1);

  // Row 7: Department & Assessment Date
  worksheet1.getRow(7).height = 18;
  worksheet1.mergeCells(`A7:${midColLetter1}7`);
  const metaR7_1 = worksheet1.getCell('A7');
  metaR7_1.value = `Department:  ${data.department}`;
  metaR7_1.font = { name: 'Arial', size: 9, bold: true, color: { argb: darkSlate } };
  metaR7_1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  metaR7_1.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  worksheet1.mergeCells(`${midNextColLetter1}7:${lastColLetter1}7`);
  const metaR7_2 = worksheet1.getCell(`${midNextColLetter1}7`);
  metaR7_2.value = `Assessment Date:  ${data.assessmentDate}`;
  metaR7_2.font = { name: 'Arial', size: 9, bold: true, color: { argb: darkSlate } };
  metaR7_2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  metaR7_2.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // Row 8: Semester & Section
  worksheet1.getRow(8).height = 18;
  worksheet1.mergeCells(`A8:${midColLetter1}8`);
  const metaR8_1 = worksheet1.getCell('A8');
  metaR8_1.value = `Semester:  ${data.semester} (${data.academicYear})`;
  metaR8_1.font = { name: 'Arial', size: 9, bold: true, color: { argb: darkSlate } };
  metaR8_1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  metaR8_1.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  worksheet1.mergeCells(`${midNextColLetter1}8:${lastColLetter1}8`);
  const metaR8_2 = worksheet1.getCell(`${midNextColLetter1}8`);
  metaR8_2.value = `Section:  ${data.section || 'A'}`;
  metaR8_2.font = { name: 'Arial', size: 9, bold: true, color: { argb: darkSlate } };
  metaR8_2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  metaR8_2.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // Row 9: Total Strength & Pass in all subjects
  worksheet1.getRow(9).height = 18;
  worksheet1.mergeCells(`A9:${midColLetter1}9`);
  const metaR9_1 = worksheet1.getCell('A9');
  metaR9_1.value = `Total Strength:  ${data.totalStudentStrength} (Present: ${data.totalPresent}  |  Absent: ${data.totalAbsent})`;
  metaR9_1.font = { name: 'Arial', size: 9, bold: true, color: { argb: darkSlate } };
  metaR9_1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  metaR9_1.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  worksheet1.mergeCells(`${midNextColLetter1}9:${lastColLetter1}9`);
  const metaR9_2 = worksheet1.getCell(`${midNextColLetter1}9`);
  metaR9_2.value = `Pass in All Subjects:  ${data.passInAllSubjectsCount} / ${data.totalStudentStrength} (${data.overallPassPercentage.toFixed(1)}%)`;
  metaR9_2.font = { name: 'Arial', size: 9, bold: true, color: { argb: darkSlate } };
  metaR9_2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  metaR9_2.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  for (let r = 7; r <= 9; r++) {
    for (let c = 1; c <= totalCols1; c++) {
      worksheet1.getRow(r).getCell(c).border = hairlineBorder;
    }
  }

  // Row 10: Spacer
  worksheet1.getRow(10).height = 6;

  // Table Headers (Rows 11, 12, 13 - Dark Navy Fill with Hairline Borders)
  const styleDarkHeader = (cell: ExcelJS.Cell, fontSize = 9) => {
    cell.font = { name: 'Arial', size: fontSize, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: darkNavy } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = hairlineBorder;
  };

  worksheet1.mergeCells('A11:A13');
  worksheet1.getCell('A11').value = 'S.NO';
  styleDarkHeader(worksheet1.getCell('A11'));

  worksheet1.mergeCells('B11:B13');
  worksheet1.getCell('B11').value = 'ROLL NO';
  styleDarkHeader(worksheet1.getCell('B11'));

  worksheet1.mergeCells('C11:C13');
  worksheet1.getCell('C11').value = 'STUDENT NAME';
  styleDarkHeader(worksheet1.getCell('C11'));

  data.subjects.forEach((sub, sIdx) => {
    const colL = getColumnLetter(4 + sIdx);
    const cCode = worksheet1.getCell(`${colL}11`);
    cCode.value = sub.code;
    styleDarkHeader(cCode);

    const cShort = worksheet1.getCell(`${colL}12`);
    cShort.value = sub.shortName;
    styleDarkHeader(cShort, 8.5);

    const cMax = worksheet1.getCell(`${colL}13`);
    cMax.value = `[${sub.maxMarks}]`;
    styleDarkHeader(cMax, 8);
  });

  const totCol = getColumnLetter(4 + subCount);
  worksheet1.mergeCells(`${totCol}11:${totCol}13`);
  worksheet1.getCell(`${totCol}11`).value = 'TOTAL';
  styleDarkHeader(worksheet1.getCell(`${totCol}11`));

  const pctCol = getColumnLetter(5 + subCount);
  worksheet1.mergeCells(`${pctCol}11:${pctCol}13`);
  worksheet1.getCell(`${pctCol}11`).value = '%';
  styleDarkHeader(worksheet1.getCell(`${pctCol}11`));

  const resCol = getColumnLetter(6 + subCount);
  worksheet1.mergeCells(`${resCol}11:${resCol}13`);
  worksheet1.getCell(`${resCol}11`).value = 'RESULT';
  styleDarkHeader(worksheet1.getCell(`${resCol}11`));

  worksheet1.getRow(11).height = 18;
  worksheet1.getRow(12).height = 16;
  worksheet1.getRow(13).height = 16;

  // Student Rows (Starting row 14 - Compact 17pt height, neat fonts)
  let rowIdx1 = 14;
  data.students.forEach((st, sIdx) => {
    const row = worksheet1.getRow(rowIdx1);
    row.height = 17;
    const isAlt = sIdx % 2 === 1;

    row.getCell(1).value = st.sno;
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(1).font = { name: 'Arial', size: 8.5, color: { argb: 'FF64748B' } };

    row.getCell(2).value = st.regNo;
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).font = { name: 'Courier New', size: 9, bold: true, color: { argb: darkSlate } };

    row.getCell(3).value = st.name;
    row.getCell(3).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    row.getCell(3).font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF1E293B' } };

    data.subjects.forEach((sub, subIdx) => {
      const sm = st.subjects[sub.code];
      const cell = row.getCell(4 + subIdx);

      const isAbsent = sm?.attendance === 'ABSENT' || sm?.marksObtained === 'AB';
      const markVal = sm?.marksObtained;
      const isFail = typeof markVal === 'number' && markVal < sub.passMark;

      cell.value = isAbsent ? 'AB' : markVal !== undefined && markVal !== '' ? markVal : '-';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      if (isAbsent || isFail) {
        cell.font = { name: 'Arial', size: 9, bold: true, color: { argb: redFailColor } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
      } else {
        cell.font = { name: 'Arial', size: 9, color: { argb: darkSlate } };
      }
    });

    // Total
    const cTotal = row.getCell(4 + subCount);
    cTotal.value = st.totalObtained;
    cTotal.alignment = { horizontal: 'center', vertical: 'middle' };
    cTotal.font = { name: 'Arial', size: 9, bold: true, color: { argb: darkSlate } };

    // %
    const cPct = row.getCell(5 + subCount);
    cPct.value = `${st.percentage.toFixed(1)}%`;
    cPct.alignment = { horizontal: 'center', vertical: 'middle' };
    cPct.font = { name: 'Arial', size: 8.5, bold: true, color: { argb: 'FF2563EB' } };

    // Result (P / F)
    const cRes = row.getCell(6 + subCount);
    cRes.value = st.overallResult;
    cRes.alignment = { horizontal: 'center', vertical: 'middle' };
    cRes.font = {
      name: 'Arial',
      size: 9.5,
      bold: true,
      color: { argb: st.overallResult === 'P' ? greenPassColor : redFailColor },
    };
    cRes.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: st.overallResult === 'P' ? 'FFDCFCE7' : 'FFFEE2E2' },
    };

    for (let c = 1; c <= totalCols1; c++) {
      row.getCell(c).border = hairlineBorder;
      if (c <= 3 && isAlt) {
        row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      }
    }

    rowIdx1++;
  });

  // 1. Space below students table
  rowIdx1 += 2;

  // 2. TABLE 1: SUBJECT-WISE EVALUATION SUMMARY (Without OVERALL column)
  const sumHeaderRow = worksheet1.getRow(rowIdx1);
  sumHeaderRow.height = 22;

  worksheet1.mergeCells(`A${rowIdx1}:C${rowIdx1}`);
  const sumTitleCell = sumHeaderRow.getCell(1);
  sumTitleCell.value = 'EVALUATION SUMMARY (SUBJECT-WISE)';
  sumTitleCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  sumTitleCell.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
  sumTitleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F2744' } };

  data.subjects.forEach((sub, sIdx) => {
    const sHCell = sumHeaderRow.getCell(4 + sIdx);
    sHCell.value = sub.code;
    sHCell.alignment = { horizontal: 'center', vertical: 'middle' };
    sHCell.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
    sHCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F2744' } };
  });

  const t1LastCol = 3 + subCount;
  for (let c = 1; c <= t1LastCol; c++) {
    sumHeaderRow.getCell(c).border = hairlineBorder;
  }
  rowIdx1++;

  // Subject-wise analysis rows in Excel Sheet 1
  const addExcelSubjectRow = (
    label: string,
    getVal: (sub: ConsolidatedSubjectInfo) => number | string,
    valColorArgb: string = 'FF0F172A',
    rowBgArgb: string = 'FFFFFFFF',
    labelBgArgb: string = 'FFF1F5F9'
  ) => {
    const aRow = worksheet1.getRow(rowIdx1);
    aRow.height = 19;

    worksheet1.mergeCells(`A${rowIdx1}:C${rowIdx1}`);
    const lCell = aRow.getCell(1);
    lCell.value = label;
    lCell.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
    lCell.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: labelBgArgb } };

    data.subjects.forEach((sub, sIdx) => {
      const vCell = aRow.getCell(4 + sIdx);
      vCell.value = getVal(sub);
      vCell.alignment = { horizontal: 'center', vertical: 'middle' };
      vCell.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: valColorArgb } };
      vCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBgArgb } };
    });

    for (let c = 1; c <= t1LastCol; c++) {
      aRow.getCell(c).border = hairlineBorder;
    }
    rowIdx1++;
  };

  // Subject-wise Metrics (NO OVERALL column here)
  addExcelSubjectRow('Total Student Strength', (s) => data.subjectMetrics[s.code]?.totalStrength ?? data.totalStudentStrength, 'FF0F172A', 'FFFFFFFF', 'FFF1F5F9');
  addExcelSubjectRow('Total Present', (s) => data.subjectMetrics[s.code]?.totalPresent ?? 0, 'FF15803D', 'FFFFFFFF', 'FFF1F5F9');
  addExcelSubjectRow('Total Absent', (s) => data.subjectMetrics[s.code]?.totalAbsent ?? 0, 'FFDC2626', 'FFFFFFFF', 'FFF1F5F9');
  addExcelSubjectRow('Total Pass', (s) => data.subjectMetrics[s.code]?.totalPass ?? 0, 'FF15803D', 'FFFFFFFF', 'FFF1F5F9');
  addExcelSubjectRow('Total Fail', (s) => data.subjectMetrics[s.code]?.totalFail ?? 0, 'FFDC2626', 'FFFFFFFF', 'FFF1F5F9');
  addExcelSubjectRow('Highest marks', (s) => data.subjectMetrics[s.code]?.highestMarks ?? 0, 'FFB45309', 'FFFFFFFF', 'FFF1F5F9');
  addExcelSubjectRow('Pass (%)', (s) => `${data.subjectMetrics[s.code]?.passPercentageStr ?? 0}%`, 'FF1D4ED8', 'FFEFF6FF', 'FFEFF6FF');

  // Staff Name row
  const rowStaff = worksheet1.getRow(rowIdx1);
  rowStaff.height = 20;
  worksheet1.mergeCells(`A${rowIdx1}:C${rowIdx1}`);
  rowStaff.getCell(1).value = 'Staff Name';
  rowStaff.getCell(1).alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
  rowStaff.getCell(1).font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF334155' } };
  rowStaff.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };

  data.subjects.forEach((sub, sIdx) => {
    const sCell = rowStaff.getCell(4 + sIdx);
    sCell.value = sub.facultyInCharge || 'Faculty';
    sCell.alignment = { horizontal: 'center', vertical: 'middle' };
    sCell.font = { name: 'Arial', size: 8.5, bold: true, color: { argb: 'FF0F172A' } };
    sCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  });

  for (let c = 1; c <= t1LastCol; c++) {
    rowStaff.getCell(c).border = hairlineBorder;
  }
  rowIdx1 += 2;

  // =========================================================================
  // 3. TABLE 2: SEPARATE DETAILED OVERALL RESULT & PERFORMANCE ANALYSIS TABLE
  // =========================================================================
  const appearedCount = data.students.filter((st) =>
    Object.values(st.subjects || {}).some(
      (s) => s && (s.attendance === 'Present' || (typeof s.marksObtained === 'number' && s.marksObtained >= 0))
    )
  ).length || data.totalPresent || data.totalStudentStrength;
  const failedAllCount = Math.max(0, data.totalStudentStrength - data.passInAllSubjectsCount);
  const highestClassTotal = data.students.length > 0 ? Math.max(...data.students.map((s) => s.totalObtained || 0)) : 0;
  const maxClassTotal = data.subjects.length * 100;
  const passAllPctStr = `${((data.passInAllSubjectsCount / (data.totalStudentStrength || 1)) * 100).toFixed(1)}%`;
  const classStatusStr = data.overallPassPercentage >= 75 ? 'First Class / Exemplary' : data.overallPassPercentage >= 50 ? 'Satisfactory' : 'Needs Academic Attention';

  // Table 2 Header: Spans the full width
  const t2HeaderRow = worksheet1.getRow(rowIdx1);
  t2HeaderRow.height = 22;
  worksheet1.mergeCells(`A${rowIdx1}:${lastColLetter1}${rowIdx1}`);
  const t2HCell = t2HeaderRow.getCell(1);
  t2HCell.value = 'OVERALL RESULT & PERFORMANCE ANALYSIS';
  t2HCell.alignment = { horizontal: 'center', vertical: 'middle' };
  t2HCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  t2HCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F2744' } };
  for (let c = 1; c <= totalCols1; c++) {
    t2HeaderRow.getCell(c).border = hairlineBorder;
  }
  rowIdx1++;

  // 4 balanced, non-overlapping column zones across totalCols1
  let s1 = Math.max(1, Math.round(totalCols1 * 0.30));
  let s2 = Math.max(1, Math.round(totalCols1 * 0.20));
  let s3 = Math.max(1, Math.round(totalCols1 * 0.30));
  let s4 = totalCols1 - (s1 + s2 + s3);
  if (s4 < 1) s4 = 1;
  let curSum = s1 + s2 + s3 + s4;
  while (curSum > totalCols1) {
    if (s1 > 1) { s1--; curSum--; }
    else if (s3 > 1) { s3--; curSum--; }
    else if (s2 > 1) { s2--; curSum--; }
    else if (s4 > 1) { s4--; curSum--; }
    else break;
  }
  while (curSum < totalCols1) {
    s3++;
    curSum++;
  }

  const c1End = s1;
  const c2Start = c1End + 1;
  const c2End = c2Start + s2 - 1;
  const c3Start = c2End + 1;
  const c3End = c3Start + s3 - 1;
  const c4Start = c3End + 1;
  const c4End = totalCols1;

  const addOverallDetailRow = (
    lbl1: string,
    val1: string | number,
    val1Color: string,
    lbl2: string,
    val2: string | number,
    val2Color: string,
    val2Bg: string = 'FFFFFFFF'
  ) => {
    const oRow = worksheet1.getRow(rowIdx1);
    oRow.height = 20;

    // Col 1 Label
    if (c1End > 1) {
      worksheet1.mergeCells(`${getColumnLetter(1)}${rowIdx1}:${getColumnLetter(c1End)}${rowIdx1}`);
    }
    const l1 = oRow.getCell(1);
    l1.value = lbl1;
    l1.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
    l1.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    l1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

    // Col 2 Value
    if (c2End > c2Start) {
      worksheet1.mergeCells(`${getColumnLetter(c2Start)}${rowIdx1}:${getColumnLetter(c2End)}${rowIdx1}`);
    }
    const v1 = oRow.getCell(c2Start);
    v1.value = val1;
    v1.alignment = { horizontal: 'center', vertical: 'middle' };
    v1.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: val1Color } };
    v1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    // Col 3 Label
    if (c3End > c3Start) {
      worksheet1.mergeCells(`${getColumnLetter(c3Start)}${rowIdx1}:${getColumnLetter(c3End)}${rowIdx1}`);
    }
    const l2 = oRow.getCell(c3Start);
    l2.value = lbl2;
    l2.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
    l2.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0F172A' } };
    l2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

    // Col 4 Value
    if (c4End > c4Start) {
      worksheet1.mergeCells(`${getColumnLetter(c4Start)}${rowIdx1}:${getColumnLetter(c4End)}${rowIdx1}`);
    }
    const v2 = oRow.getCell(c4Start);
    v2.value = val2;
    v2.alignment = { horizontal: 'center', vertical: 'middle' };
    v2.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: val2Color } };
    v2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: val2Bg } };

    for (let c = 1; c <= totalCols1; c++) {
      oRow.getCell(c).border = hairlineBorder;
    }
    rowIdx1++;
  };

  // Row 1: Strength & Passed All
  addOverallDetailRow(
    'Total Students Enrolled:',
    data.totalStudentStrength,
    'FF0F172A',
    'Passed in All Subjects (All Clear):',
    `${data.passInAllSubjectsCount} (${passAllPctStr})`,
    'FF15803D',
    'FFF0FDF4'
  );

  // Row 2: Appeared & Failed
  addOverallDetailRow(
    'Students Appeared (Present):',
    appearedCount,
    'FF0F172A',
    'Failed in One or More Subjects:',
    failedAllCount,
    failedAllCount > 0 ? 'FFDC2626' : 'FF0F172A'
  );

  // Row 3: Highest Total & Performance Classification
  addOverallDetailRow(
    'Class Highest Total Score:',
    `${highestClassTotal} / ${maxClassTotal}`,
    'FFB45309',
    'Academic Performance Status:',
    classStatusStr,
    'FF0F172A'
  );

  // Row 4: Big Highlight KPI Banner for Overall Pass %
  const rowOverallBanner = worksheet1.getRow(rowIdx1);
  rowOverallBanner.height = 24;
  worksheet1.mergeCells(`A${rowIdx1}:${lastColLetter1}${rowIdx1}`);
  const bCell = rowOverallBanner.getCell(1);
  bCell.value = `OVERALL CLASS PASS PERCENTAGE:  ${data.overallPassPercentage.toFixed(1)}%`;
  bCell.alignment = { horizontal: 'center', vertical: 'middle' };
  bCell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF1D4ED8' } };
  bCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
  for (let c = 1; c <= totalCols1; c++) {
    rowOverallBanner.getCell(c).border = hairlineBorder;
  }
  // Leave signing space above CC | HOD | AC | EXAM CELL COORDINATOR | PRINCIPAL | DIRECTOR
  rowIdx1 += 1;
  const gapRow1 = worksheet1.getRow(rowIdx1);
  gapRow1.height = 14;

  rowIdx1 += 1;
  const signSpaceRow1 = worksheet1.getRow(rowIdx1);
  signSpaceRow1.height = 36; // Generous height for physical/digital signatures above the lines

  rowIdx1 += 1;
  // Signatures on Sheet 1 (CC | HOD | AC | EXAM CELL COORDINATOR | PRINCIPAL | DIRECTOR)
  // Dedicated separate sign lines row for each role
  const sigLineRow1 = worksheet1.getRow(rowIdx1);
  sigLineRow1.height = 16;
  const sigTitles1 = ['CC', 'HOD', 'AC', 'EXAM CELL COORDINATOR', 'PRINCIPAL', 'DIRECTOR'];
  const sigStep1 = Math.max(1, Math.floor(totalCols1 / 6));

  for (let i = 0; i < 6; i++) {
    const startC = i * sigStep1 + 1;
    const endC = i === 5 ? totalCols1 : (i + 1) * sigStep1;
    const sL = getColumnLetter(startC);
    const eL = getColumnLetter(endC);

    if (endC > startC) {
      worksheet1.mergeCells(`${sL}${rowIdx1}:${eL}${rowIdx1}`);
    }
    const lCell = worksheet1.getCell(`${sL}${rowIdx1}`);
    lCell.value = '__________________';
    lCell.alignment = { horizontal: 'center', vertical: 'bottom' };
    lCell.font = { name: 'Arial', size: 9, color: { argb: 'FF94A3B8' } };
  }

  rowIdx1 += 1;
  const sigTitleRow1 = worksheet1.getRow(rowIdx1);
  sigTitleRow1.height = 24;

  for (let i = 0; i < 6; i++) {
    const startC = i * sigStep1 + 1;
    const endC = i === 5 ? totalCols1 : (i + 1) * sigStep1;
    const sL = getColumnLetter(startC);
    const eL = getColumnLetter(endC);

    if (endC > startC) {
      worksheet1.mergeCells(`${sL}${rowIdx1}:${eL}${rowIdx1}`);
    }
    const sCell = worksheet1.getCell(`${sL}${rowIdx1}`);
    sCell.value = sigTitles1[i];
    sCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    sCell.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF0B1E36' } };
  }

  // =========================================================================
  // SHEET 2: "Subwise" — SUBJECT WISE ANALYSIS (EXACT REPLICA OF UPLOADED IMAGE)
  // =========================================================================
  const worksheet2 = workbook.addWorksheet('Subwise', {
    views: [{ showGridLines: true }],
    pageSetup: {
      paperSize: 9, // A4
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.35, right: 0.35, top: 0.35, bottom: 0.35, header: 0.2, footer: 0.2 },
    },
  });

  worksheet2.getColumn(1).width = 2;   // A (spacer)
  worksheet2.getColumn(2).width = 6;   // B: S.No
  worksheet2.getColumn(3).width = 34;  // C: Subject Code & Name
  worksheet2.getColumn(4).width = 20;  // D: Handled By
  worksheet2.getColumn(5).width = 10;  // E: Strength
  worksheet2.getColumn(6).width = 14;  // F: No. Students appeared
  worksheet2.getColumn(7).width = 14;  // G: No. of Passed (widened for EXAM CELL COORDINATOR visibility)
  worksheet2.getColumn(8).width = 14;  // H: No. of Failed (widened for EXAM CELL COORDINATOR visibility)
  worksheet2.getColumn(9).width = 9;   // I: Pass % Str
  worksheet2.getColumn(10).width = 9;  // J: Pass % App
  worksheet2.getColumn(11).width = 13; // K: Previous Pass %
  worksheet2.getColumn(12).width = 18; // L: Comments on Comparison

  // Embed College Logo (Left) and NAAC Logo (Right) in Sheet 2
  try {
    if (COLLEGE_LOGO_BASE64) {
      const cleanBase64 = COLLEGE_LOGO_BASE64.replace(/^data:image\/\w+;base64,/, '');
      const collegeImgId = workbook.addImage({
        base64: cleanBase64,
        extension: 'png',
      });
      worksheet2.addImage(collegeImgId, {
        tl: { col: 0.15, row: 0.1 },
        ext: { width: 60, height: 60 },
      });
    }
    if (NAAC_LOGO_BASE64) {
      const cleanNaacBase64 = NAAC_LOGO_BASE64.replace(/^data:image\/\w+;base64,/, '');
      const naacExt = NAAC_LOGO_BASE64.startsWith('data:image/jpeg') ? 'jpeg' : 'png';
      const naacImgId = workbook.addImage({
        base64: cleanNaacBase64,
        extension: naacExt,
      });
      worksheet2.addImage(naacImgId, {
        tl: { col: 11.2, row: 0.1 },
        ext: { width: 72, height: 58 },
      });
    }
  } catch (err) {
    console.warn('Could not add logos to Excel Sheet 2:', err);
  }

  // Row 1: Jaya Sakthi Engineering College
  worksheet2.mergeCells('B1:L1');
  const sw_c1 = worksheet2.getCell('B1');
  sw_c1.value = 'Jaya Sakthi Engineering College';
  sw_c1.font = { name: 'Times New Roman', size: 14, bold: true, color: { argb: 'FF0B1E36' } };
  sw_c1.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet2.getRow(1).height = 24;

  // Row 2: College Address
  worksheet2.mergeCells('B2:L2');
  const sw_addr = worksheet2.getCell('B2');
  sw_addr.value = "St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu";
  sw_addr.font = { name: 'Arial', size: 8.5, color: { argb: 'FF475569' } };
  sw_addr.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet2.getRow(2).height = 15;

  // Row 3: [Exam Type] Result Analysis
  worksheet2.mergeCells('B3:L3');
  const sw_c2 = worksheet2.getCell('B3');
  sw_c2.value = `${data.examType} Result Analysis`;
  sw_c2.font = { name: 'Arial', size: 10.5, bold: true, color: { argb: darkSlate } };
  sw_c2.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet2.getRow(3).height = 16;

  // Row 4: Subject Wise Analysis
  worksheet2.mergeCells('B4:L4');
  const sw_c4 = worksheet2.getCell('B4');
  sw_c4.value = 'Subject Wise Analysis';
  sw_c4.font = { name: 'Arial', size: 10.5, bold: true, underline: true, color: { argb: darkSlate } };
  sw_c4.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet2.getRow(4).height = 17;

  // Row 5: spacer
  worksheet2.getRow(5).height = 4;

  // Metadata Box
  worksheet2.getCell('B6').value = 'Department:';
  worksheet2.getCell('B6').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF334155' } };
  worksheet2.mergeCells('C6:G6');
  worksheet2.getCell('C6').value = data.department;
  worksheet2.getCell('C6').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: darkSlate } };

  worksheet2.mergeCells('H6:I6');
  worksheet2.getCell('H6').value = 'Date:';
  worksheet2.getCell('H6').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF334155' } };
  worksheet2.getCell('H6').alignment = { horizontal: 'right' };
  worksheet2.mergeCells('J6:L6');
  worksheet2.getCell('J6').value = data.assessmentDate;
  worksheet2.getCell('J6').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: darkSlate } };

  worksheet2.getCell('B7').value = 'Year/ Semester:';
  worksheet2.getCell('B7').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF334155' } };
  worksheet2.mergeCells('C7:G7');
  worksheet2.getCell('C7').value = getYearSemRoman(data.semester);
  worksheet2.getCell('C7').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: darkSlate } };

  worksheet2.mergeCells('H7:I7');
  worksheet2.getCell('H7').value = 'Section:';
  worksheet2.getCell('H7').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF334155' } };
  worksheet2.getCell('H7').alignment = { horizontal: 'right' };
  worksheet2.mergeCells('J7:L7');
  worksheet2.getCell('J7').value = data.section;
  worksheet2.getCell('J7').font = { name: 'Arial', size: 9.5, bold: true, color: { argb: darkSlate } };

  worksheet2.getRow(6).height = 18;
  worksheet2.getRow(7).height = 18;
  worksheet2.getRow(8).height = 6;

  // Table Headers
  const styleSwHeader = (cell: ExcelJS.Cell, fontSize = 9) => {
    cell.font = { name: 'Arial', size: fontSize, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: darkNavy } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = hairlineBorder;
  };

  worksheet2.mergeCells('B9:B10');
  worksheet2.getCell('B9').value = 'S.No';
  styleSwHeader(worksheet2.getCell('B9'));

  worksheet2.mergeCells('C9:C10');
  worksheet2.getCell('C9').value = 'Subject Code & Title';
  styleSwHeader(worksheet2.getCell('C9'));

  worksheet2.mergeCells('D9:D10');
  worksheet2.getCell('D9').value = 'Handled By (Faculty)';
  styleSwHeader(worksheet2.getCell('D9'));

  worksheet2.mergeCells('E9:E10');
  worksheet2.getCell('E9').value = 'Strength';
  styleSwHeader(worksheet2.getCell('E9'));

  worksheet2.mergeCells('F9:F10');
  worksheet2.getCell('F9').value = 'Appeared';
  styleSwHeader(worksheet2.getCell('F9'));

  worksheet2.mergeCells('G9:G10');
  worksheet2.getCell('G9').value = 'Absent';
  styleSwHeader(worksheet2.getCell('G9'));

  worksheet2.mergeCells('H9:H10');
  worksheet2.getCell('H9').value = 'Passed';
  styleSwHeader(worksheet2.getCell('H9'));

  worksheet2.mergeCells('I9:I10');
  worksheet2.getCell('I9').value = 'Failed';
  styleSwHeader(worksheet2.getCell('I9'));

  worksheet2.mergeCells('J9:K9');
  worksheet2.getCell('J9').value = 'Pass %';
  styleSwHeader(worksheet2.getCell('J9'));

  worksheet2.getCell('J10').value = 'Str';
  styleSwHeader(worksheet2.getCell('J10'), 8.5);

  worksheet2.getCell('K10').value = 'App';
  styleSwHeader(worksheet2.getCell('K10'), 8.5);

  worksheet2.mergeCells('L9:L10');
  worksheet2.getCell('L9').value = 'Highest Mark';
  styleSwHeader(worksheet2.getCell('L9'));

  worksheet2.getRow(9).height = 18;
  worksheet2.getRow(10).height = 18;

  // Data Rows
  let swRowIdx = 11;
  let totStrength = 0;
  let totAppeared = 0;
  let totAbsent = 0;
  let totPass = 0;
  let totFail = 0;

  data.subjects.forEach((sub, idx) => {
    const swRow = worksheet2.getRow(swRowIdx);
    swRow.height = 20;
    const m = data.subjectMetrics[sub.code];
    const strengthVal = m?.totalStrength ?? data.totalStudentStrength;
    const appearedVal = m?.totalPresent ?? 0;
    const absentVal = m?.totalAbsent ?? 0;
    const passVal = m?.totalPass ?? 0;
    const failVal = m?.totalFail ?? 0;

    totStrength = strengthVal;
    totAppeared += appearedVal;
    totAbsent += absentVal;
    totPass += passVal;
    totFail += failVal;

    swRow.getCell(2).value = idx + 1;
    swRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(2).font = { name: 'Arial', size: 9 };

    swRow.getCell(3).value = `${sub.code} - ${sub.name}`;
    swRow.getCell(3).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    swRow.getCell(3).font = { name: 'Arial', size: 9, bold: true };

    swRow.getCell(4).value = sub.facultyInCharge || m?.facultyInCharge || 'Faculty In-Charge';
    swRow.getCell(4).alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    swRow.getCell(4).font = { name: 'Arial', size: 9 };

    swRow.getCell(5).value = strengthVal;
    swRow.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(5).font = { name: 'Arial', size: 9 };

    swRow.getCell(6).value = appearedVal;
    swRow.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(6).font = { name: 'Arial', size: 9 };

    swRow.getCell(7).value = absentVal;
    swRow.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(7).font = { name: 'Arial', size: 9 };

    swRow.getCell(8).value = passVal;
    swRow.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(8).font = { name: 'Arial', size: 9, bold: true, color: { argb: greenPassColor } };

    swRow.getCell(9).value = failVal;
    swRow.getCell(9).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(9).font = { name: 'Arial', size: 9, bold: true, color: { argb: redFailColor } };

    swRow.getCell(10).value = `${m?.passPercentageStr ?? 0}%`;
    swRow.getCell(10).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(10).font = { name: 'Arial', size: 9 };

    swRow.getCell(11).value = `${m?.passPercentageApp ?? 0}%`;
    swRow.getCell(11).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(11).font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0B1E36' } };

    swRow.getCell(12).value = m?.highestMarks ?? 0;
    swRow.getCell(12).alignment = { horizontal: 'center', vertical: 'middle' };
    swRow.getCell(12).font = { name: 'Arial', size: 9, bold: true };

    for (let c = 2; c <= 12; c++) {
      swRow.getCell(c).border = hairlineBorder;
    }
    swRowIdx++;
  });

  // Summary Row at bottom of table
  const summaryRow = worksheet2.getRow(swRowIdx);
  summaryRow.height = 22;
  worksheet2.mergeCells(`B${swRowIdx}:D${swRowIdx}`);
  const sCell = worksheet2.getCell(`B${swRowIdx}`);
  sCell.value = 'Total / Department Average';
  sCell.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF0B1E36' } };
  sCell.alignment = { horizontal: 'right', vertical: 'middle' };

  summaryRow.getCell(5).value = totStrength;
  summaryRow.getCell(5).font = { name: 'Arial', size: 9.5, bold: true };
  summaryRow.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

  summaryRow.getCell(6).value = totAppeared;
  summaryRow.getCell(6).font = { name: 'Arial', size: 9.5, bold: true };
  summaryRow.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };

  summaryRow.getCell(7).value = totAbsent;
  summaryRow.getCell(7).font = { name: 'Arial', size: 9.5, bold: true };
  summaryRow.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };

  summaryRow.getCell(8).value = totPass;
  summaryRow.getCell(8).font = { name: 'Arial', size: 9.5, bold: true, color: { argb: greenPassColor } };
  summaryRow.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };

  summaryRow.getCell(9).value = totFail;
  summaryRow.getCell(9).font = { name: 'Arial', size: 9.5, bold: true, color: { argb: redFailColor } };
  summaryRow.getCell(9).alignment = { horizontal: 'center', vertical: 'middle' };

  summaryRow.getCell(10).value = '-';
  summaryRow.getCell(10).alignment = { horizontal: 'center', vertical: 'middle' };

  const avgAppVal = data.subjects.length > 0 ? Math.round(data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.passPercentageApp ?? 0), 0) / data.subjects.length) : 0;
  summaryRow.getCell(11).value = `${avgAppVal}%`;
  summaryRow.getCell(11).font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF0B1E36' } };
  summaryRow.getCell(11).alignment = { horizontal: 'center', vertical: 'middle' };

  summaryRow.getCell(12).value = `Overall Pass: ${data.overallPassPercentage.toFixed(1)}%`;
  summaryRow.getCell(12).font = { name: 'Arial', size: 8.5, bold: true };
  summaryRow.getCell(12).alignment = { horizontal: 'center', vertical: 'middle' };

  for (let c = 2; c <= 12; c++) {
    const cCell = summaryRow.getCell(c);
    cCell.border = hairlineBorder;
    cCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  }
  swRowIdx++;

  // Signatures on Sheet 2 (CC | HOD | AC | EXAM CELL COORDINATOR | PRINCIPAL | DIRECTOR)
  swRowIdx += 1;
  const swSignSpaceRow = worksheet2.getRow(swRowIdx);
  swSignSpaceRow.height = 36; // Generous height for physical/digital signatures above lines

  const swSigBlocks = [
    { start: 'B', end: 'C', title: 'CC' },
    { start: 'D', end: 'D', title: 'HOD' },
    { start: 'E', end: 'F', title: 'AC' },
    { start: 'G', end: 'H', title: 'EXAM CELL COORDINATOR' },
    { start: 'I', end: 'J', title: 'PRINCIPAL' },
    { start: 'K', end: 'L', title: 'DIRECTOR' },
  ];

  // Dedicated separate sign lines row
  swRowIdx += 1;
  const swSignLineRow = worksheet2.getRow(swRowIdx);
  swSignLineRow.height = 16;
  swSigBlocks.forEach((block) => {
    if (block.start !== block.end) {
      worksheet2.mergeCells(`${block.start}${swRowIdx}:${block.end}${swRowIdx}`);
    }
    const cell = worksheet2.getCell(`${block.start}${swRowIdx}`);
    cell.value = '__________________';
    cell.alignment = { horizontal: 'center', vertical: 'bottom' };
    cell.font = { name: 'Arial', size: 9, color: { argb: 'FF94A3B8' } };
  });

  // Role titles row
  swRowIdx += 1;
  const swSigRow = worksheet2.getRow(swRowIdx);
  swSigRow.height = 24;

  swSigBlocks.forEach((block) => {
    if (block.start !== block.end) {
      worksheet2.mergeCells(`${block.start}${swRowIdx}:${block.end}${swRowIdx}`);
    }
    const cell = worksheet2.getCell(`${block.start}${swRowIdx}`);
    cell.value = block.title;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
    cell.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF0B1E36' } };
  });

  // =========================================================================
  // SHEET 3+: INDIVIDUAL SUBJECT MARKSHEETS (ACCORDING TO OFFICIAL UPLOADED DESIGN)
  // =========================================================================
  const sheetsToExport = matchingSheetsParam || data.matchingSheets || [];
  sheetsToExport.forEach((sheet, idx) => {
    const rawCode = (sheet.meta.subjectCode || `SUB${idx + 1}`).trim().replace(/[\\/?*:[\]]/g, '_');
    const safeTabName = rawCode.substring(0, 31);
    let uniqueTabName = safeTabName;
    let counter = 1;
    while (workbook.getWorksheet(uniqueTabName)) {
      uniqueTabName = `${safeTabName.substring(0, 27)}_${counter++}`;
    }
    const metrics = computeMetrics(sheet.students, sheet.meta.passMark);
    buildSubjectMarksheetWorksheet(workbook, sheet.meta, sheet.students, metrics, uniqueTabName);
  });

  // Trigger Excel File Download sanitized for Microsoft Excel 2007+
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = await sanitizeExcelFor2007(buffer);
  triggerFileDownload(blob, `${cleanName(data.department)}_${cleanName(data.semester)}_Consolidated_Marksheet.xlsx`);
}

/**
 * 2. EXPORT CONSOLIDATED MARKSHEET TO PDF (.pdf)
 * Page 1: Realigned, compact, hairline-bordered student marksheet with perfectly integrated summary rows
 * Page 2: "Subject Wise Analysis" page matching the uploaded 2nd page photo
 */
export async function exportConsolidatedToPDF(
  data: ConsolidatedData,
  theme: ExportDesignTheme = 'dark-navy',
  matchingSheetsParam?: SubmittedSheetRecord[]
): Promise<void> {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = 297;
  const pageHeight = 210;
  const themeConfig = EXPORT_DESIGN_THEMES[theme] || EXPORT_DESIGN_THEMES['dark-navy'];

  // =========================================================================
  // PAGE 1: STUDENT MARKSHEET (COMPACT, HAIRLINE LINES, PERFECT COLUMN ALIGNMENT)
  // =========================================================================
  renderPdfCollegeHeader(doc, pageWidth);

  // Title Banner with Rounded Pill Capsule Styling
  const titleBannerY = 28;
  doc.setFillColor(...themeConfig.bannerBg);
  doc.roundedRect(12, titleBannerY, pageWidth - 24, 7.5, 2, 2, 'F');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(9);
  doc.setTextColor(...themeConfig.bannerText);
  doc.text(
    `ALL SUBJECTS CONSOLIDATED ASSESSMENT MARKSHEET — ${data.examType.toUpperCase()}`,
    pageWidth / 2,
    titleBannerY + 5.1,
    { align: 'center' }
  );

  // Official 3-Row Metadata Matrix Box (Matching uploaded reference design)
  const metaBoxY = 37.5;
  const metaBoxH = 15;
  doc.setFillColor(248, 250, 252);
  doc.setDrawColor(203, 213, 225);
  doc.setLineWidth(0.18);
  doc.roundedRect(12, metaBoxY, pageWidth - 24, metaBoxH, 1.5, 1.5, 'FD');

  const leftLabelX = 16;
  const leftValX = 42;
  const rightLabelX = pageWidth / 2 + 10;
  const rightValX = rightLabelX + 38;

  // Row 1
  let curMetaY = metaBoxY + 4.2;
  doc.setFontSize(7.5);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Department:', leftLabelX, curMetaY);
  doc.setTextColor(15, 23, 42);
  doc.text(data.department, leftValX, curMetaY);

  doc.setTextColor(71, 85, 105);
  doc.text('Assessment Date:', rightLabelX, curMetaY);
  doc.setTextColor(15, 23, 42);
  doc.text(data.assessmentDate, rightValX, curMetaY);

  // Row 2
  curMetaY += 4.5;
  doc.setTextColor(71, 85, 105);
  doc.text('Semester:', leftLabelX, curMetaY);
  doc.setTextColor(15, 23, 42);
  doc.text(`${data.semester} (${data.academicYear})`, leftValX, curMetaY);

  doc.setTextColor(71, 85, 105);
  doc.text('Section:', rightLabelX, curMetaY);
  doc.setTextColor(15, 23, 42);
  doc.text(data.section || 'A', rightValX, curMetaY);

  // Row 3
  curMetaY += 4.5;
  doc.setTextColor(71, 85, 105);
  doc.text('Total Strength:', leftLabelX, curMetaY);
  doc.setTextColor(15, 23, 42);
  doc.text(`${data.totalStudentStrength} (Present: ${data.totalPresent}  |  Absent: ${data.totalAbsent})`, leftValX, curMetaY);

  doc.setTextColor(71, 85, 105);
  doc.text('Pass in All Subjects:', rightLabelX, curMetaY);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(22, 101, 52);
  doc.text(`${data.passInAllSubjectsCount} / ${data.totalStudentStrength} (${data.overallPassPercentage.toFixed(1)}%)`, rightValX, curMetaY);

  // Table Headers
  const tableHeaders1 = ['#', 'ROLL NO', 'STUDENT NAME'];
  data.subjects.forEach((s) => {
    tableHeaders1.push(`${s.code}\n(${s.shortName})`);
  });
  tableHeaders1.push('TOTAL', '%', 'RES');

  const subCount = data.subjects.length;

  // Build Table Body: Student Rows
  const tableBody1: any[] = data.students.map((st) => {
    const row: any[] = [st.sno.toString(), st.regNo, st.name];
    data.subjects.forEach((s) => {
      const sm = st.subjects[s.code];
      const isAbsent = sm?.attendance === 'ABSENT' || sm?.marksObtained === 'AB';
      const markStr = isAbsent ? 'AB' : sm?.marksObtained !== undefined ? String(sm.marksObtained) : '-';
      const isFail = typeof sm?.marksObtained === 'number' && sm.marksObtained < s.passMark;

      row.push({
        content: markStr,
        styles: {
          halign: 'center',
          fontStyle: isAbsent || isFail ? 'bold' : 'normal',
          textColor: isAbsent || isFail ? [220, 38, 38] : [15, 23, 42],
        },
      });
    });

    row.push({ content: st.totalObtained.toString(), styles: { halign: 'center', fontStyle: 'bold' } });
    row.push({ content: `${st.percentage.toFixed(1)}%`, styles: { halign: 'center', fontStyle: 'bold', textColor: [37, 99, 235] } });
    row.push({
      content: st.overallResult,
      styles: {
        halign: 'center',
        fontStyle: 'bold',
        textColor: st.overallResult === 'P' ? [22, 101, 52] : [220, 38, 38],
        fillColor: st.overallResult === 'P' ? [220, 252, 231] : [254, 226, 226],
      },
    });

    return row;
  });

  // FULL-SCREEN RESPONSIVE COLUMN SIZING (Matches Banner & Metadata card exactly across entire A4 landscape width)
  const fullTableWidth = pageWidth - 24; // 273mm
  const marginSide = 12;

  const snoWidth = 8;
  const rollWidth = 34;
  const totWidth = 16;
  const pctWidth = 14;
  const resWidth = 12;
  const fixedWidthsSum = snoWidth + rollWidth + totWidth + pctWidth + resWidth; // 84mm
  const remainingWidth = fullTableWidth - fixedWidthsSum; // 273 - 84 = 189mm

  let nameWidth = 65;
  let subWidth = 18;

  if (subCount <= 0) {
    nameWidth = remainingWidth;
    subWidth = 0;
  } else if (subCount === 1) {
    // Single subject: fills the landscape page matching screenshot
    nameWidth = 94;
    subWidth = remainingWidth - nameWidth; // 95mm
  } else if (subCount === 2) {
    nameWidth = 75;
    subWidth = (remainingWidth - nameWidth) / 2; // 57mm each
  } else if (subCount === 3) {
    nameWidth = 65;
    subWidth = (remainingWidth - nameWidth) / 3; // 41.33mm each
  } else if (subCount === 4) {
    nameWidth = 57;
    subWidth = (remainingWidth - nameWidth) / 4; // 33mm each
  } else if (subCount === 5) {
    nameWidth = 49;
    subWidth = (remainingWidth - nameWidth) / 5; // 28mm each
  } else if (subCount === 6) {
    nameWidth = 45;
    subWidth = (remainingWidth - nameWidth) / 6; // 24mm each
  } else {
    nameWidth = Math.max(38, 45 - (subCount - 6) * 1.5);
    subWidth = (remainingWidth - nameWidth) / subCount;
  }

  const colStyles1: Record<number, any> = {
    0: { halign: 'center', cellWidth: snoWidth },
    1: { halign: 'center', cellWidth: rollWidth, fontStyle: 'bold' },
    2: { halign: 'left', cellWidth: nameWidth, fontStyle: 'bold' },
  };

  data.subjects.forEach((_, idx) => {
    colStyles1[3 + idx] = { halign: 'center', cellWidth: subWidth };
  });
  colStyles1[3 + data.subjects.length] = { halign: 'center', cellWidth: totWidth, fontStyle: 'bold' };
  colStyles1[4 + data.subjects.length] = { halign: 'center', cellWidth: pctWidth, fontStyle: 'bold' };
  colStyles1[5 + data.subjects.length] = { halign: 'center', cellWidth: resWidth, fontStyle: 'bold' };

  // 1. STUDENTS MARKSHEET TABLE (Contains ONLY student records)
  autoTable(doc, {
    startY: 55.5,
    tableWidth: fullTableWidth,
    margin: { left: marginSide, right: marginSide, bottom: 20 },
    head: [tableHeaders1],
    body: tableBody1,
    theme: 'grid',
    headStyles: {
      fillColor: themeConfig.headFill,
      textColor: themeConfig.headText,
      fontStyle: 'bold',
      fontSize: subCount > 4 ? 6.5 : 7.2,
      halign: 'center',
      valign: 'middle',
      cellPadding: { top: 1.2, bottom: 1.2, left: 0.8, right: 0.8 },
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
    },
    styles: {
      fontSize: subCount > 4 ? 6.5 : 7.2,
      cellPadding: { top: 1.0, bottom: 1.0, left: 0.8, right: 0.8 },
      textColor: [15, 23, 42],
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
      valign: 'middle',
    },
    alternateRowStyles: {
      fillColor: themeConfig.alternateRowBg,
    },
    columnStyles: colStyles1,
  });

  // Space below students table
  const studentsFinalY1 = (doc as any).lastAutoTable?.finalY || 120;
  const summaryStartY1 = studentsFinalY1 + 5.5;

  // 2. TABLE 1: SUBJECT-WISE EVALUATION SUMMARY (Without OVERALL column)
  const summaryHeaders1: any[] = [
    {
      content: 'EVALUATION SUMMARY (SUBJECT-WISE)',
      colSpan: 3,
      styles: { halign: 'left', fontStyle: 'bold' },
    },
    ...data.subjects.map((s) => ({
      content: s.code,
      styles: { halign: 'center', fontStyle: 'bold' },
    })),
  ];

  const summaryBody1: any[] = [];

  const addSummaryTableRow = (
    label: string,
    getVal: (sub: ConsolidatedSubjectInfo) => number | string,
    rowBg: [number, number, number],
    valTextColor: [number, number, number] = [15, 23, 42]
  ) => {
    const r: any[] = [
      {
        content: label,
        colSpan: 3,
        styles: {
          halign: 'right',
          fontStyle: 'bold',
          fillColor: [241, 245, 249],
          textColor: [15, 23, 42],
          fontSize: 6.0,
        },
      },
    ];

    data.subjects.forEach((s) => {
      r.push({
        content: String(getVal(s)),
        styles: {
          halign: 'center',
          fontStyle: 'bold',
          fillColor: rowBg,
          textColor: valTextColor,
          fontSize: 6.2,
        },
      });
    });

    summaryBody1.push(r);
  };

  // 1. Total Student Strength
  addSummaryTableRow('Total Student Strength', (s) => data.subjectMetrics[s.code]?.totalStrength ?? data.totalStudentStrength, [255, 255, 255], [15, 23, 42]);
  // 2. Total Present (Green)
  addSummaryTableRow('Total Present', (s) => data.subjectMetrics[s.code]?.totalPresent ?? 0, [255, 255, 255], [21, 128, 61]);
  // 3. Total Absent (Red)
  addSummaryTableRow('Total Absent', (s) => data.subjectMetrics[s.code]?.totalAbsent ?? 0, [255, 255, 255], [220, 38, 38]);
  // 4. Total Pass (Green)
  addSummaryTableRow('Total Pass', (s) => data.subjectMetrics[s.code]?.totalPass ?? 0, [255, 255, 255], [21, 128, 61]);
  // 5. Total Fail (Red)
  addSummaryTableRow('Total Fail', (s) => data.subjectMetrics[s.code]?.totalFail ?? 0, [255, 255, 255], [220, 38, 38]);
  // 6. Highest marks (Amber)
  addSummaryTableRow('Highest marks', (s) => data.subjectMetrics[s.code]?.highestMarks ?? 0, [255, 255, 255], [180, 83, 9]);

  // 7. Pass (%) (Highlighted Soft Blue Row)
  const passPctRow: any[] = [
    {
      content: 'Pass (%)',
      colSpan: 3,
      styles: {
        halign: 'right',
        fontStyle: 'bold',
        fillColor: [239, 246, 255],
        textColor: [30, 64, 175],
        fontSize: 6.2,
      },
    },
  ];
  data.subjects.forEach((s) => {
    passPctRow.push({
      content: `${data.subjectMetrics[s.code]?.passPercentageStr ?? 0}%`,
      styles: {
        halign: 'center',
        fontStyle: 'bold',
        fillColor: [239, 246, 255],
        textColor: [29, 78, 216],
        fontSize: 6.5,
      },
    });
  });
  summaryBody1.push(passPctRow);

  // 8. Staff Name
  const staffPdfRow: any[] = [
    {
      content: 'Staff Name',
      colSpan: 3,
      styles: {
        halign: 'right',
        fontStyle: 'bold',
        fillColor: [241, 245, 249],
        textColor: [15, 23, 42],
        fontSize: 6.0,
      },
    },
  ];
  data.subjects.forEach((sub) => {
    staffPdfRow.push({
      content: sub.facultyInCharge || 'Faculty',
      styles: {
        halign: 'center',
        fontStyle: 'bold',
        fillColor: [241, 245, 249],
        textColor: [15, 23, 42],
        fontSize: 5.8,
      },
    });
  });
  summaryBody1.push(staffPdfRow);

  const colStylesSummary: Record<number, any> = {
    0: { cellWidth: snoWidth },
    1: { cellWidth: rollWidth },
    2: { cellWidth: nameWidth },
  };
  data.subjects.forEach((_, idx) => {
    colStylesSummary[3 + idx] = { cellWidth: subWidth };
  });

  const t1Width = snoWidth + rollWidth + nameWidth + subWidth * data.subjects.length;

  // AutoTable for Table 1 (Subject-Wise Evaluation Summary)
  autoTable(doc, {
    startY: summaryStartY1,
    tableWidth: t1Width,
    margin: { left: marginSide },
    head: [summaryHeaders1],
    body: summaryBody1,
    theme: 'grid',
    headStyles: {
      fillColor: themeConfig.headFill,
      textColor: themeConfig.headText,
      fontStyle: 'bold',
      fontSize: subCount > 4 ? 6.5 : 7.2,
      halign: 'center',
      valign: 'middle',
      cellPadding: { top: 1.2, bottom: 1.2, left: 0.8, right: 0.8 },
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
    },
    styles: {
      fontSize: subCount > 4 ? 6.5 : 7.2,
      cellPadding: { top: 1.0, bottom: 1.0, left: 0.8, right: 0.8 },
      textColor: [15, 23, 42],
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
      valign: 'middle',
    },
    columnStyles: colStylesSummary,
  });

  // =========================================================================
  // 3. TABLE 2: SEPARATE DETAILED OVERALL RESULT & PERFORMANCE ANALYSIS TABLE
  // =========================================================================
  const t1FinalY = (doc as any).lastAutoTable?.finalY || (summaryStartY1 + 40);
  const t2StartY = t1FinalY + 3.5;

  const appearedCountPdf = data.students.filter((st) =>
    Object.values(st.subjects || {}).some(
      (s) => s && (s.attendance === 'Present' || (typeof s.marksObtained === 'number' && s.marksObtained >= 0))
    )
  ).length || data.totalPresent || data.totalStudentStrength;
  const failedAllCountPdf = Math.max(0, data.totalStudentStrength - data.passInAllSubjectsCount);
  const highestClassTotalPdf = data.students.length > 0 ? Math.max(...data.students.map((s) => s.totalObtained || 0)) : 0;
  const maxClassTotalPdf = data.subjects.length * 100;
  const passAllPctStrPdf = `${((data.passInAllSubjectsCount / (data.totalStudentStrength || 1)) * 100).toFixed(1)}%`;
  const classStatusStrPdf = data.overallPassPercentage >= 75 ? 'First Class / Exemplary' : data.overallPassPercentage >= 50 ? 'Satisfactory' : 'Needs Academic Attention';

  const t2Headers: any[] = [
    {
      content: 'OVERALL RESULT & PERFORMANCE ANALYSIS',
      colSpan: 4,
      styles: { halign: 'center', fontStyle: 'bold', fillColor: themeConfig.headFill, textColor: [255, 255, 255], fontSize: 7.2 },
    },
  ];

  const t2ColWidth = fullTableWidth / 4;
  const t2Body: any[] = [
    // Row 1: Enrolled & Passed All
    [
      { content: 'Total Students Enrolled', styles: { fontStyle: 'bold', fillColor: [241, 245, 249], halign: 'right', textColor: [15, 23, 42], fontSize: 6.0 } },
      { content: String(data.totalStudentStrength), styles: { fontStyle: 'bold', halign: 'center', textColor: [15, 23, 42], fontSize: 6.5 } },
      { content: 'Passed in All Subjects (All Clear)', styles: { fontStyle: 'bold', fillColor: [240, 253, 244], textColor: [22, 101, 52], halign: 'right', fontSize: 6.0 } },
      { content: `${data.passInAllSubjectsCount} (${passAllPctStrPdf})`, styles: { fontStyle: 'bold', halign: 'center', textColor: [21, 128, 61], fillColor: [240, 253, 244], fontSize: 6.5 } },
    ],
    // Row 2: Appeared & Failed
    [
      { content: 'Students Appeared (Present)', styles: { fontStyle: 'bold', fillColor: [241, 245, 249], halign: 'right', textColor: [15, 23, 42], fontSize: 6.0 } },
      { content: String(appearedCountPdf), styles: { fontStyle: 'bold', halign: 'center', textColor: [15, 23, 42], fontSize: 6.5 } },
      { content: 'Failed in One or More Subjects', styles: { fontStyle: 'bold', fillColor: [241, 245, 249], halign: 'right', textColor: [15, 23, 42], fontSize: 6.0 } },
      { content: String(failedAllCountPdf), styles: { fontStyle: 'bold', halign: 'center', textColor: failedAllCountPdf > 0 ? [220, 38, 38] : [15, 23, 42], fontSize: 6.5 } },
    ],
    // Row 3: Highest Total & Status
    [
      { content: 'Class Highest Total Marks', styles: { fontStyle: 'bold', fillColor: [241, 245, 249], halign: 'right', textColor: [15, 23, 42], fontSize: 6.0 } },
      { content: `${highestClassTotalPdf} / ${maxClassTotalPdf}`, styles: { fontStyle: 'bold', halign: 'center', textColor: [180, 83, 9], fontSize: 6.5 } },
      { content: 'Academic Performance Status', styles: { fontStyle: 'bold', fillColor: [241, 245, 249], halign: 'right', textColor: [15, 23, 42], fontSize: 6.0 } },
      { content: classStatusStrPdf, styles: { fontStyle: 'bold', halign: 'center', textColor: [15, 23, 42], fontSize: 6.2 } },
    ],
    // Row 4: Big Highlight KPI Banner
    [
      {
        content: `OVERALL CLASS PASS PERCENTAGE:   ${data.overallPassPercentage.toFixed(1)}%`,
        colSpan: 4,
        styles: {
          halign: 'center',
          fontStyle: 'bold',
          fontSize: 7.5,
          fillColor: [239, 246, 255],
          textColor: [29, 78, 216],
          cellPadding: { top: 1.4, bottom: 1.4 },
        },
      },
    ],
  ];

  autoTable(doc, {
    startY: t2StartY,
    tableWidth: fullTableWidth,
    margin: { left: marginSide, right: marginSide },
    head: [t2Headers],
    body: t2Body,
    theme: 'grid',
    headStyles: {
      fillColor: themeConfig.headFill,
      textColor: [255, 255, 255],
      fontStyle: 'bold',
      fontSize: 7.2,
      halign: 'center',
      valign: 'middle',
      cellPadding: { top: 1.2, bottom: 1.2 },
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
    },
    styles: {
      cellPadding: { top: 1.0, bottom: 1.0, left: 1.0, right: 1.0 },
      textColor: [15, 23, 42],
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
      valign: 'middle',
    },
    columnStyles: {
      0: { cellWidth: t2ColWidth },
      1: { cellWidth: t2ColWidth },
      2: { cellWidth: t2ColWidth },
      3: { cellWidth: t2ColWidth },
    },
  });

  // Signatures at bottom of Page 1 (CC | HOD | AC | EXAM CELL COORDINATOR | PRINCIPAL | DIRECTOR)
  // Ensure generous clearance (at least 20-35mm space above) so authorities have ample room to sign
  const lastY1 = (doc as any).lastAutoTable?.finalY || 150;
  let sigY1 = pageHeight - 16;
  if (sigY1 < lastY1 + 18) {
    sigY1 = lastY1 + 20;
  }
  renderPdfSignatures(doc, pageWidth, marginSide, sigY1, ['CC', 'HOD', 'AC', 'EXAM CELL COORDINATOR', 'PRINCIPAL', 'DIRECTOR']);

  // =========================================================================
  // PAGE 2: "SUBJECT WISE ANALYSIS" (CLEAN & PROPER PROFESSIONAL DESIGN)
  // =========================================================================
  doc.addPage('a4', 'landscape');

  renderPdfCollegeHeader(doc, pageWidth, false);

  doc.setFont('times', 'bold');
  doc.setFontSize(15);
  doc.setTextColor(11, 30, 54);
  doc.text('JAYA SAKTHI ENGINEERING COLLEGE', pageWidth / 2, 9.5, { align: 'center' });

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(51, 65, 85);
  doc.text("St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu", pageWidth / 2, 13.5, { align: 'center' });

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.8);
  doc.setTextColor(71, 85, 105);
  doc.text('Approved by AICTE, New Delhi & Affiliated to Anna University, Chennai', pageWidth / 2, 17.0, { align: 'center' });

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10.5);
  doc.setTextColor(11, 30, 54);
  doc.text(`${data.examType.toUpperCase()} RESULT ANALYSIS — SUBJECT WISE ANALYSIS`, pageWidth / 2, 21.5, { align: 'center' });

  // Metadata Box matching photo
  const swMetaBoxY = 24.5;
  doc.setDrawColor(203, 213, 225);
  doc.setLineWidth(0.1);
  doc.roundedRect(12, swMetaBoxY, pageWidth - 24, 11, 1, 1, 'S');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(71, 85, 105);
  doc.text('Department:', 16, swMetaBoxY + 4.5);
  doc.setTextColor(15, 23, 42);
  doc.text(data.department, 42, swMetaBoxY + 4.5);

  doc.setTextColor(71, 85, 105);
  doc.text('Year/ Semester:', 16, swMetaBoxY + 8.8);
  doc.setTextColor(15, 23, 42);
  doc.text(`${getYearSemRoman(data.semester)} (${data.semester})`, 42, swMetaBoxY + 8.8);

  const rightMetaX = pageWidth - 75;
  doc.setTextColor(71, 85, 105);
  doc.text('Date:', rightMetaX, swMetaBoxY + 4.5);
  doc.setTextColor(15, 23, 42);
  doc.text(data.assessmentDate, rightMetaX + 16, swMetaBoxY + 4.5);

  doc.setTextColor(71, 85, 105);
  doc.text('Section / Batch:', rightMetaX, swMetaBoxY + 8.8);
  doc.setTextColor(15, 23, 42);
  doc.text(`Sec ${data.section} · ${data.academicYear}`, rightMetaX + 26, swMetaBoxY + 8.8);

  // Table matching official standard
  const swTableHeaders = [
    [
      { content: 'S.No', rowSpan: 2 },
      { content: 'Subject Code & Title', rowSpan: 2 },
      { content: 'Handled By (Faculty)', rowSpan: 2 },
      { content: 'Strength', rowSpan: 2 },
      { content: 'Appeared', rowSpan: 2 },
      { content: 'Absent', rowSpan: 2 },
      { content: 'Passed', rowSpan: 2 },
      { content: 'Failed', rowSpan: 2 },
      { content: 'Pass %', colSpan: 2 },
      { content: 'Highest Mark', rowSpan: 2 },
      { content: 'Evaluation Rating', rowSpan: 2 },
    ],
    ['Str', 'App'],
  ];

  const totalSubStrength = data.totalStudentStrength;
  const totalSubAppeared = data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.totalPresent ?? 0), 0);
  const totalSubAbsent = data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.totalAbsent ?? 0), 0);
  const totalSubPass = data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.totalPass ?? 0), 0);
  const totalSubFail = data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.totalFail ?? 0), 0);
  const avgSubApp = data.subjects.length > 0 ? Math.round(data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.passPercentageApp ?? 0), 0) / data.subjects.length) : 0;

  const swTableBody = data.subjects.map((sub, idx) => {
    const m = data.subjectMetrics[sub.code];
    const appPct = m?.passPercentageApp ?? 0;
    let rating = 'Satisfactory';
    if (appPct >= 90) rating = 'Exemplary';
    else if (appPct >= 80) rating = 'Very Good';
    else if (appPct >= 75) rating = 'Target Met';
    else if (appPct < 60) rating = 'Remedial Req.';

    return [
      (idx + 1).toString(),
      `${sub.code} - ${sub.name}`,
      sub.facultyInCharge || m?.facultyInCharge || 'Faculty In-Charge',
      (m?.totalStrength ?? data.totalStudentStrength).toString(),
      (m?.totalPresent ?? 0).toString(),
      (m?.totalAbsent ?? 0).toString(),
      (m?.totalPass ?? 0).toString(),
      (m?.totalFail ?? 0).toString(),
      `${m?.passPercentageStr ?? 0}%`,
      `${appPct}%`,
      (m?.highestMarks ?? 0).toString(),
      rating,
    ];
  });

  // Summary Totals Row at end of Subject Wise Analysis
  swTableBody.push([
    '',
    'Total / Department Average',
    '',
    totalSubStrength.toString(),
    totalSubAppeared.toString(),
    totalSubAbsent.toString(),
    totalSubPass.toString(),
    totalSubFail.toString(),
    '-',
    `${avgSubApp}%`,
    '-',
    `Overall Pass: ${data.overallPassPercentage.toFixed(1)}%`,
  ]);

  autoTable(doc, {
    startY: 37,
    margin: { left: 12, right: 12, bottom: 22 },
    tableWidth: pageWidth - 24,
    head: swTableHeaders as any,
    body: swTableBody,
    theme: 'grid',
    headStyles: {
      fillColor: themeConfig.headFill,
      textColor: themeConfig.headText,
      fontStyle: 'bold',
      fontSize: 7.8,
      halign: 'center',
      valign: 'middle',
      cellPadding: { top: 1.5, bottom: 1.5, left: 1, right: 1 },
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
    },
    styles: {
      fontSize: 7.8,
      cellPadding: { top: 1.5, bottom: 1.5, left: 1, right: 1 },
      textColor: [15, 23, 42],
      lineWidth: 0.08,
      lineColor: themeConfig.borderColor,
      valign: 'middle',
    },
    columnStyles: {
      0: { halign: 'center', cellWidth: 10 },
      1: { halign: 'left', cellWidth: 70, fontStyle: 'bold' },
      2: { halign: 'left', cellWidth: 40 },
      3: { halign: 'center', cellWidth: 16 },
      4: { halign: 'center', cellWidth: 17 },
      5: { halign: 'center', cellWidth: 15 },
      6: { halign: 'center', cellWidth: 16, fontStyle: 'bold', textColor: [22, 101, 52] },
      7: { halign: 'center', cellWidth: 15, fontStyle: 'bold', textColor: [220, 38, 38] },
      8: { halign: 'center', cellWidth: 15 },
      9: { halign: 'center', cellWidth: 16, fontStyle: 'bold', textColor: [37, 99, 235] },
      10: { halign: 'center', cellWidth: 18, fontStyle: 'bold' },
      11: { halign: 'center', cellWidth: 25, fontStyle: 'bold' },
    },
    didParseCell: (hookData) => {
      // Highlight summary row at the bottom
      if (hookData.section === 'body' && hookData.row.index === swTableBody.length - 1) {
        hookData.cell.styles.fontStyle = 'bold';
        hookData.cell.styles.fillColor = [241, 245, 249];
        hookData.cell.styles.textColor = [11, 30, 54];
      }
    },
  });

  // Page 2 Signatures (CC | HOD | AC | EXAM CELL COORDINATOR | PRINCIPAL | DIRECTOR)
  const swFinalY = (doc as any).lastAutoTable?.finalY || 160;
  const swSigY = Math.min(pageHeight - 16, swFinalY + 12);
  renderPdfSignatures(doc, pageWidth, 12, swSigY, ['CC', 'HOD', 'AC', 'EXAM CELL COORDINATOR', 'PRINCIPAL', 'DIRECTOR']);

  // =========================================================================
  // PAGE 3+: INDIVIDUAL SUBJECT MARKSHEETS (ACCORDING TO OFFICIAL UPLOADED DESIGN)
  // =========================================================================
  const sheetsToExport = matchingSheetsParam || data.matchingSheets || [];
  sheetsToExport.forEach((sheet) => {
    const metrics = computeMetrics(sheet.students, sheet.meta.passMark);
    renderSingleSubjectMarksheetToPDF(doc, sheet.meta, sheet.students, metrics, false);
  });

  doc.save(`${cleanName(data.department)}_${cleanName(data.semester)}_Consolidated_Marksheet.pdf`);
}

/**
 * Helper to render PDF Header logos & text
 */
function renderPdfCollegeHeader(doc: jsPDF, pageWidth: number, renderText = true) {
  try {
    if (COLLEGE_LOGO_BASE64) {
      doc.addImage(COLLEGE_LOGO_BASE64, 'PNG', 12, 5.5, 18, 18);
    }
    if (NAAC_LOGO_BASE64) {
      const naacFormat = NAAC_LOGO_BASE64.startsWith('data:image/jpeg') ? 'JPEG' : 'PNG';
      doc.addImage(NAAC_LOGO_BASE64, naacFormat, pageWidth - 12 - 18, 6.5, 18, 14.5);
    }
  } catch (e) {
    console.warn('PDF Header image notice:', e);
  }

  if (renderText) {
    doc.setFont('times', 'bold');
    doc.setFontSize(15);
    doc.setTextColor(11, 30, 54);
    doc.text('JAYA SAKTHI ENGINEERING COLLEGE', pageWidth / 2, 12, { align: 'center' });

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(51, 65, 85);
    doc.text("St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu", pageWidth / 2, 16.6, {
      align: 'center',
    });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(100, 116, 139);
    doc.text('Approved by AICTE, New Delhi & Affiliated to Anna University, Chennai', pageWidth / 2, 21.0, {
      align: 'center',
    });

    // Decorative double line rule (Navy top line + Golden amber bottom line)
    doc.setDrawColor(15, 40, 78);
    doc.setLineWidth(0.7);
    doc.line(12, 24.5, pageWidth - 12, 24.5);
    doc.setDrawColor(217, 119, 6);
    doc.setLineWidth(0.35);
    doc.line(12, 25.6, pageWidth - 12, 25.6);
  }
}

/**
 * Helper to render PDF Signatures row
 */
function renderPdfSignatures(doc: jsPDF, pageWidth: number, margin: number, sigY: number, sigTitles: string[]) {
  const colW = (pageWidth - margin * 2) / sigTitles.length;

  sigTitles.forEach((sig, idx) => {
    const x = margin + colW * (idx + 0.5);
    // Draw clean horizontal signature line
    doc.setDrawColor(148, 163, 184); // Slate-400
    doc.setLineWidth(0.25);
    const lineHalfWidth = Math.min(20, (colW * 0.72) / 2);
    doc.line(x - lineHalfWidth, sigY, x + lineHalfWidth, sigY);

    // Title centered below the line
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(sig.length > 15 ? 7 : 7.8);
    doc.setTextColor(11, 30, 54);
    doc.text(sig, x, sigY + 4.5, { align: 'center' });
  });
}

/**
 * Builds the complete single-subject marksheet children (tables, paragraphs, signatures)
 * strictly matching the uploaded reference design for Microsoft Word (.docx)
 */
function buildSubjectDocxChildren(
  sheet: SubmittedSheetRecord,
  collegeLogoBytes: Uint8Array,
  naacLogoBytes: Uint8Array,
  tableBorders: any,
  noBorders: any
): (Paragraph | Table)[] {
  const meta = sheet.meta;
  const students = sheet.students || [];
  const metrics = computeMetrics(students, meta.passMark);

  // 1. Header Table (College Logo Left | Centered College Title | NAAC Logo Right)
  const headerLeft: Paragraph[] = [];
  if (collegeLogoBytes.length > 0) {
    try {
      headerLeft.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: collegeLogoBytes,
              transformation: { width: 62, height: 62 },
              type: 'png',
            }),
          ],
        })
      );
    } catch {
      headerLeft.push(new Paragraph({}));
    }
  } else {
    headerLeft.push(new Paragraph({}));
  }

  const headerRight: Paragraph[] = [];
  if (naacLogoBytes.length > 0) {
    try {
      headerRight.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: naacLogoBytes,
              transformation: { width: 72, height: 58 },
              type: 'png',
            }),
          ],
        })
      );
    } catch {
      headerRight.push(new Paragraph({}));
    }
  } else {
    headerRight.push(new Paragraph({}));
  }

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: noBorders,
            children: headerLeft,
          }),
          new TableCell({
            width: { size: 70, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: noBorders,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 20 },
                children: [
                  new TextRun({
                    text: 'JAYA SAKTHI ENGINEERING COLLEGE',
                    bold: true,
                    size: 24,
                    font: 'Times New Roman',
                    color: '0B1E36',
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 15 },
                children: [
                  new TextRun({
                    text: "St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu",
                    bold: true,
                    size: 15,
                    font: 'Arial',
                    color: '334155',
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: 'Approved by AICTE, New Delhi & Affiliated to Anna University, Chennai',
                    size: 14,
                    font: 'Arial',
                    color: '64748B',
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: noBorders,
            children: headerRight,
          }),
        ],
      }),
    ],
  });

  // 2. Document Title Banner (Navy Pill Capsule)
  const titleTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: '0B1E36', type: ShadingType.CLEAR },
            borders: noBorders,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 50, after: 50 },
                children: [
                  new TextRun({
                    text: `CONTINUOUS ASSESSMENT MARKSHEET — ${(meta.examType || 'ASSESSMENT').toUpperCase()}`,
                    bold: true,
                    size: 18,
                    font: 'Arial',
                    color: 'FFFFFF',
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  // 3. Metadata Matrix Box
  const metaCell = (label: string, value: string) =>
    new TableCell({
      width: { size: 50, type: WidthType.PERCENTAGE },
      shading: { fill: 'F8FAFC', type: ShadingType.CLEAR },
      borders: tableBorders,
      children: [
        new Paragraph({
          spacing: { before: 30, after: 30 },
          children: [
            new TextRun({ text: `${label}: `, bold: true, size: 16, color: '475569', font: 'Arial' }),
            new TextRun({ text: value, size: 16, bold: true, color: '0F172A', font: 'Arial' }),
          ],
        }),
      ],
    });

  const metaTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: [
      new TableRow({
        children: [
          metaCell('Department', meta.department || ''),
          metaCell('Assessment Date', meta.assessmentDate || ''),
        ],
      }),
      new TableRow({
        children: [
          metaCell('Semester', meta.academicYear ? `${meta.semester || ''} (BATCH: ${meta.academicYear})` : meta.semester || ''),
          metaCell('Faculty In-Charge', meta.facultyInCharge || 'MS. K. SANDHIYA'),
        ],
      }),
      new TableRow({
        children: [
          metaCell('Subject', `${meta.subjectCode} - ${meta.subjectName}`),
          metaCell('Evaluator Name', meta.evaluatorName || 'DR. K.V. ARUN PILLAI'),
        ],
      }),
      new TableRow({
        children: [
          metaCell('Marks Matrix', `Max: ${meta.maxMarks} | Pass Mark: ${meta.passMark}`),
          metaCell('Assessment Type', meta.examType || 'CAT-3'),
        ],
      }),
    ],
  });

  // 4. Student Marksheet Table
  const headers = ['#', 'REGISTER NO', 'STUDENT NAME', 'ATTENDANCE', 'MAX', 'OBT', 'GRADE', 'RESULT'];
  const colWidths = [5, 18, 33, 14, 7, 7, 7, 9];
  const tableRows: TableRow[] = [
    new TableRow({
      children: headers.map((h, i) =>
        new TableCell({
          width: { size: colWidths[i], type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: '0B1E36', type: ShadingType.CLEAR },
          children: [
            new Paragraph({
              alignment: i <= 1 ? AlignmentType.CENTER : i === 2 ? AlignmentType.LEFT : AlignmentType.CENTER,
              spacing: { before: 30, after: 30 },
              children: [new TextRun({ text: h, bold: true, size: 14, color: 'FFFFFF', font: 'Arial' })],
            }),
          ],
        })
      ),
    }),
  ];

  students.forEach((st, idx) => {
    const isAbsent = st.attendance === 'ABSENT';
    const obtStr = isAbsent ? 'AB' : st.marksObtained !== '' ? String(st.marksObtained) : '-';
    const isPass = st.result === 'PASS';

    tableRows.push(
      new TableRow({
        children: [
          new TableCell({
            width: { size: colWidths[0], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(idx + 1), size: 14 })] })],
          }),
          new TableCell({
            width: { size: colWidths[1], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: st.regNo, bold: true, size: 14, font: 'Courier New' })] })],
          }),
          new TableCell({
            width: { size: colWidths[2], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: st.name.toUpperCase(), bold: true, size: 14 })] })],
          }),
          new TableCell({
            width: { size: colWidths[3], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: st.attendance, size: 14 })] })],
          }),
          new TableCell({
            width: { size: colWidths[4], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(st.maxMarks), size: 14 })] })],
          }),
          new TableCell({
            width: { size: colWidths[5], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            shading: isAbsent || !isPass ? { fill: 'FEE2E2', type: ShadingType.CLEAR } : undefined,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: obtStr, bold: true, size: 14, color: isAbsent || !isPass ? 'DC2626' : '0F172A' })] })],
          }),
          new TableCell({
            width: { size: colWidths[6], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: st.grade, bold: true, size: 14 })] })],
          }),
          new TableCell({
            width: { size: colWidths[7], type: WidthType.PERCENTAGE },
            borders: tableBorders,
            shading: { fill: isPass ? 'DCFCE7' : 'FEE2E2', type: ShadingType.CLEAR },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: st.result, bold: true, size: 14, color: isPass ? '15803D' : 'DC2626' })] })],
          }),
        ],
      })
    );
  });

  const studentsTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: tableRows,
  });

  // 5. Evaluation Summary & Performance Metrics Table (Matches reference screenshot)
  const passPct = typeof metrics?.passPercentage === 'number' ? metrics.passPercentage.toFixed(2) : '0.00';
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'auto' };
  const borderOutline = { style: BorderStyle.SINGLE, size: 4, color: 'CBD5E1' };
  const lightDivider = { style: BorderStyle.SINGLE, size: 2, color: 'F1F5F9' };
  const summaryCellMargins = { top: 100, bottom: 100, left: 140, right: 140 };

  const summaryTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: borderOutline,
      bottom: borderOutline,
      left: borderOutline,
      right: borderOutline,
      insideHorizontal: lightDivider,
      insideVertical: noBorder,
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnSpan: 4,
            shading: { fill: '0F2744', type: ShadingType.CLEAR },
            margins: summaryCellMargins,
            borders: {
              top: borderOutline,
              bottom: borderOutline,
              left: borderOutline,
              right: borderOutline,
            },
            children: [
              new Paragraph({
                spacing: { before: 50, after: 50 },
                children: [
                  new TextRun({
                    text: 'EVALUATION SUMMARY & PERFORMANCE METRICS',
                    bold: true,
                    size: 18,
                    color: 'FFFFFF',
                    font: 'Arial',
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      // Row 1: Present & Pass counts
      new TableRow({
        children: [
          // Present Label (30%)
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
            borders: {
              top: borderOutline,
              bottom: lightDivider,
              left: borderOutline,
              right: noBorder,
            },
            children: [
              new Paragraph({
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: 'No of Student Present:', bold: true, size: 18, color: '0F172A', font: 'Arial' }),
                ],
              }),
            ],
          }),
          // Present Value (20%, left aligned)
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
            borders: {
              top: borderOutline,
              bottom: lightDivider,
              left: noBorder,
              right: noBorder,
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: String(metrics?.presentCount ?? 0), bold: true, size: 20, color: '15803D', font: 'Arial' }),
                ],
              }),
            ],
          }),
          // Pass Label (30%)
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
            borders: {
              top: borderOutline,
              bottom: lightDivider,
              left: noBorder,
              right: noBorder,
            },
            children: [
              new Paragraph({
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: 'No of Student Pass:', bold: true, size: 18, color: '0F172A', font: 'Arial' }),
                ],
              }),
            ],
          }),
          // Pass Value (20%, left aligned)
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
            borders: {
              top: borderOutline,
              bottom: lightDivider,
              left: noBorder,
              right: borderOutline,
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: String(metrics?.passCount ?? 0), bold: true, size: 20, color: '15803D', font: 'Arial' }),
                ],
              }),
            ],
          }),
        ],
      }),
      // Row 2: Absent & Fail counts
      new TableRow({
        children: [
          // Absent Label (30%)
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
            borders: {
              top: lightDivider,
              bottom: borderOutline,
              left: borderOutline,
              right: noBorder,
            },
            children: [
              new Paragraph({
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: 'No of Student Absent:', bold: true, size: 18, color: '0F172A', font: 'Arial' }),
                ],
              }),
            ],
          }),
          // Absent Value (20%, left aligned)
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
            borders: {
              top: lightDivider,
              bottom: borderOutline,
              left: noBorder,
              right: noBorder,
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: String(metrics?.absentCount ?? 0), bold: true, size: 20, color: 'DC2626', font: 'Arial' }),
                ],
              }),
            ],
          }),
          // Fail Label (30%)
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
            borders: {
              top: lightDivider,
              bottom: borderOutline,
              left: noBorder,
              right: noBorder,
            },
            children: [
              new Paragraph({
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: 'No of Student Fail:', bold: true, size: 18, color: '0F172A', font: 'Arial' }),
                ],
              }),
            ],
          }),
          // Fail Value (20%, left aligned)
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            margins: summaryCellMargins,
            shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
            borders: {
              top: lightDivider,
              bottom: borderOutline,
              left: noBorder,
              right: borderOutline,
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 40, after: 40 },
                children: [
                  new TextRun({ text: String(metrics?.failCount ?? 0), bold: true, size: 20, color: 'DC2626', font: 'Arial' }),
                ],
              }),
            ],
          }),
        ],
      }),
      // Row 3: Pass Percentage Banner (Centered, Highlighted Soft Blue Card)
      new TableRow({
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            columnSpan: 4,
            shading: { fill: 'EFF6FF', type: ShadingType.CLEAR },
            margins: summaryCellMargins,
            borders: {
              top: borderOutline,
              bottom: borderOutline,
              left: borderOutline,
              right: borderOutline,
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 50, after: 50 },
                children: [
                  new TextRun({
                    text: `Pass percentage:  ${passPct}%`,
                    bold: true,
                    size: 22,
                    color: '1D4ED8',
                    font: 'Arial',
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  // 6. Signatures Table (Evaluator Signature, Exam Cell Coordinator, Principal Signature)
  const sigCell = (title: string, name: string) => {
    const trimmed = name?.trim();
    const children = [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 260, after: 30 },
        children: [new TextRun({ text: '______________________', size: 16, color: '64748B', font: 'Arial', bold: true })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 20, after: 10 },
        children: [new TextRun({ text: title, bold: true, size: 16, color: '0F172A', font: 'Arial' })],
      }),
    ];
    if (trimmed) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 10, after: 20 },
          children: [new TextRun({ text: `(${trimmed})`, size: 14, color: '64748B', font: 'Arial' })],
        })
      );
    }
    return new TableCell({
      width: { size: 33.33, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.BOTTOM,
      borders: noBorders,
      children,
    });
  };

  const sigTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          sigCell(
            'Evaluator Signature',
            meta.evaluatorName?.trim() && meta.evaluatorName?.trim() !== 'Evaluator Signature'
              ? meta.evaluatorName.trim()
              : ''
          ),
          sigCell('Exam Cell Coordinator', meta.examCellCoordinator?.trim() || ''),
          sigCell('Principal Signature', meta.principalName?.trim() || meta.principalSignature?.trim() || ''),
        ],
      }),
    ],
  });

  return [
    headerTable,
    new Paragraph({ spacing: { before: 40, after: 40 } }),
    titleTable,
    new Paragraph({ spacing: { before: 40, after: 40 } }),
    metaTable,
    new Paragraph({ spacing: { before: 60, after: 60 } }),
    studentsTable,
    new Paragraph({ spacing: { before: 80, after: 60 } }),
    summaryTable,
    new Paragraph({ spacing: { before: 120, after: 60 } }),
    sigTable,
  ];
}

/**
 * 3. EXPORT CONSOLIDATED MARKSHEET TO WORD (.docx)
 * Page 1: Realigned, compact, light-design student marksheet
 * Page 2: "Subject Wise Analysis" page matching the uploaded 2nd page photo
 * Page 3+: Every individual subject's complete marksheet page
 */
export async function exportConsolidatedToWord(
  data: ConsolidatedData,
  theme: ExportDesignTheme = 'dark-navy',
  matchingSheetsParam?: SubmittedSheetRecord[]
): Promise<void> {
  const themeConfig = EXPORT_DESIGN_THEMES[theme] || EXPORT_DESIGN_THEMES['dark-navy'];
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'auto' };
  const noBorders = {
    top: noBorder,
    bottom: noBorder,
    left: noBorder,
    right: noBorder,
    insideHorizontal: noBorder,
    insideVertical: noBorder,
  };
  const hairlineBorder = { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1' }; // Hairline border
  const tableBorders = {
    top: hairlineBorder,
    bottom: hairlineBorder,
    left: hairlineBorder,
    right: hairlineBorder,
    insideHorizontal: hairlineBorder,
    insideVertical: hairlineBorder,
  };

  // Prepare Embedded Logos as Uint8Array
  const collegeLogoBytes = base64ToUint8Array(COLLEGE_LOGO_BASE64);
  const naacLogoBytes = base64ToUint8Array(NAAC_LOGO_BASE64);

  // 1. Header Table (College Logo Left | Centered College Title | NAAC Logo Right)
  const headerChildrenLeft: Paragraph[] = [];
  if (collegeLogoBytes.length > 0) {
    try {
      headerChildrenLeft.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: collegeLogoBytes,
              transformation: { width: 62, height: 62 },
              type: 'png',
            }),
          ],
        })
      );
    } catch {
      headerChildrenLeft.push(new Paragraph({}));
    }
  } else {
    headerChildrenLeft.push(new Paragraph({}));
  }

  const headerChildrenRight: Paragraph[] = [];
  if (naacLogoBytes.length > 0) {
    try {
      headerChildrenRight.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: naacLogoBytes,
              transformation: { width: 72, height: 58 },
              type: 'png',
            }),
          ],
        })
      );
    } catch {
      headerChildrenRight.push(new Paragraph({}));
    }
  } else {
    headerChildrenRight.push(new Paragraph({}));
  }

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 14, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: noBorders,
            children: headerChildrenLeft,
          }),
          new TableCell({
            width: { size: 72, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: noBorders,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 20 },
                children: [
                  new TextRun({
                    text: 'JAYA SAKTHI ENGINEERING COLLEGE',
                    bold: true,
                    size: 24,
                    font: 'Times New Roman',
                    color: '0B1E36',
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 15 },
                children: [
                  new TextRun({
                    text: "St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu",
                    bold: true,
                    size: 15,
                    font: 'Arial',
                    color: '334155',
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: 'Approved by AICTE, New Delhi & Affiliated to Anna University, Chennai',
                    size: 14,
                    font: 'Arial',
                    color: '64748B',
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 14, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: noBorders,
            children: headerChildrenRight,
          }),
        ],
      }),
    ],
  });

  // Title Banner (Dark Navy with White Bold Text)
  const titleTable1 = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            shading: { fill: themeConfig.wordHeaderBg, type: ShadingType.CLEAR },
            borders: tableBorders,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 50, after: 50 },
                children: [
                  new TextRun({
                    text: `ALL SUBJECTS CONSOLIDATED ASSESSMENT MARKSHEET — ${data.examType.toUpperCase()}`,
                    bold: true,
                    size: 18,
                    font: 'Arial',
                    color: 'FFFFFF',
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  // Official 3-Row Metadata Matrix Box in Word
  const page1MetaCell = (label: string, value: string) =>
    new TableCell({
      width: { size: 50, type: WidthType.PERCENTAGE },
      shading: { fill: 'F8FAFC', type: ShadingType.CLEAR },
      borders: tableBorders,
      children: [
        new Paragraph({
          spacing: { before: 30, after: 30 },
          children: [
            new TextRun({ text: `${label}: `, bold: true, size: 16, color: '475569', font: 'Arial' }),
            new TextRun({ text: value, size: 16, bold: true, color: '0F172A', font: 'Arial' }),
          ],
        }),
      ],
    });

  const page1MetaTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: [
      new TableRow({
        children: [
          page1MetaCell('Department', data.department),
          page1MetaCell('Assessment Date', data.assessmentDate),
        ],
      }),
      new TableRow({
        children: [
          page1MetaCell('Semester', `${data.semester} (${data.academicYear})`),
          page1MetaCell('Section', data.section || 'A'),
        ],
      }),
      new TableRow({
        children: [
          page1MetaCell('Total Strength', `${data.totalStudentStrength} (Present: ${data.totalPresent}  |  Absent: ${data.totalAbsent})`),
          page1MetaCell('Pass in All Subjects', `${data.passInAllSubjectsCount} / ${data.totalStudentStrength} (${data.overallPassPercentage.toFixed(1)}%)`),
        ],
      }),
    ],
  });

  // Table Headers
  const tableHeaders1 = ['#', 'ROLL NO', 'STUDENT NAME'];
  data.subjects.forEach((s) => tableHeaders1.push(`${s.code}\n(${s.shortName})`));
  tableHeaders1.push('TOTAL', '%', 'RES');

  const subCount = data.subjects.length;

  // Dynamic percentage column widths to ensure full screen width in Word
  const snoPct = 4;
  const rollPct = 14;
  const totPct = 7;
  const pctPct = 7;
  const resPct = 6;
  const fixedPctSum = snoPct + rollPct + totPct + pctPct + resPct; // 38%
  const remPct = 100 - fixedPctSum; // 62%

  let namePct = 26;
  let subPct = 12;

  if (subCount <= 0) {
    namePct = remPct;
    subPct = 0;
  } else if (subCount === 1) {
    namePct = 31;
    subPct = remPct - namePct; // 31%
  } else if (subCount === 2) {
    namePct = 26;
    subPct = (remPct - namePct) / 2; // 18% each
  } else if (subCount === 3) {
    namePct = 23;
    subPct = (remPct - namePct) / 3; // 13% each
  } else if (subCount === 4) {
    namePct = 20;
    subPct = (remPct - namePct) / 4; // 10.5% each
  } else {
    namePct = Math.max(16, 22 - (subCount - 4));
    subPct = (remPct - namePct) / subCount;
  }

  const wordColWidths: number[] = [snoPct, rollPct, namePct];
  for (let i = 0; i < subCount; i++) {
    wordColWidths.push(subPct);
  }
  wordColWidths.push(totPct, pctPct, resPct);

  const docxRows1: TableRow[] = [
    new TableRow({
      children: tableHeaders1.map((h, i) =>
        new TableCell({
          width: { size: wordColWidths[i], type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: themeConfig.wordHeaderBg, type: ShadingType.CLEAR },
          children: [
            new Paragraph({
              alignment: i <= 1 ? AlignmentType.CENTER : i === 2 ? AlignmentType.LEFT : AlignmentType.CENTER,
              spacing: { before: 30, after: 30 },
              children: [
                new TextRun({ text: h, bold: true, size: 13, color: 'FFFFFF', font: 'Arial' }),
              ],
            }),
          ],
        })
      ),
    }),
  ];

  // Student Rows
  data.students.forEach((st, idx) => {
    const rowCells: TableCell[] = [
      new TableCell({
        width: { size: snoPct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(idx + 1), size: 13 })] })],
      }),
      new TableCell({
        width: { size: rollPct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: st.regNo, bold: true, size: 13, font: 'Courier New' })] })],
      }),
      new TableCell({
        width: { size: namePct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: st.name, bold: true, size: 13 })] })],
      }),
    ];

    data.subjects.forEach((s) => {
      const sm = st.subjects[s.code];
      const isAbsent = sm?.attendance === 'ABSENT' || sm?.marksObtained === 'AB';
      const markStr = isAbsent ? 'AB' : sm?.marksObtained !== undefined ? String(sm.marksObtained) : '-';
      const isFail = typeof sm?.marksObtained === 'number' && sm.marksObtained < s.passMark;

      rowCells.push(
        new TableCell({
          width: { size: subPct, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: isAbsent || isFail ? { fill: 'FEE2E2', type: ShadingType.CLEAR } : undefined,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: markStr,
                  bold: true,
                  size: 13,
                  color: isAbsent || isFail ? 'DC2626' : '0F172A',
                }),
              ],
            }),
          ],
        })
      );
    });

    rowCells.push(
      new TableCell({
        width: { size: totPct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(st.totalObtained), size: 13, bold: true })] })],
      }),
      new TableCell({
        width: { size: pctPct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${st.percentage.toFixed(1)}%`, size: 13, bold: true, color: '2563EB' })] })],
      }),
      new TableCell({
        width: { size: resPct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        shading: { fill: st.overallResult === 'P' ? 'DCFCE7' : 'FEE2E2', type: ShadingType.CLEAR },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: st.overallResult,
                bold: true,
                size: 14,
                color: st.overallResult === 'P' ? '15803D' : 'DC2626',
              }),
            ],
          }),
        ],
      })
    );

    docxRows1.push(new TableRow({ children: rowCells }));
  });

  const studentsTable1 = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: docxRows1,
  });



  // 2. TABLE 1: SUBJECT-WISE EVALUATION SUMMARY IN WORD (Without OVERALL column)
  const summaryRows1: TableRow[] = [];

  const summaryHeaderCells: TableCell[] = [
    new TableCell({
      columnSpan: 3,
      width: { size: snoPct + rollPct + namePct, type: WidthType.PERCENTAGE },
      borders: tableBorders,
      shading: { fill: '0F2744', type: ShadingType.CLEAR },
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [new TextRun({ text: 'EVALUATION SUMMARY (SUBJECT-WISE)', bold: true, size: 14, color: 'FFFFFF' })],
        }),
      ],
    }),
  ];

  data.subjects.forEach((s) => {
    summaryHeaderCells.push(
      new TableCell({
        width: { size: subPct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        shading: { fill: '0F2744', type: ShadingType.CLEAR },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: s.code, bold: true, size: 14, color: 'FFFFFF' })],
          }),
        ],
      })
    );
  });

  summaryRows1.push(new TableRow({ children: summaryHeaderCells }));

  const addDocxSummaryRow = (
    label: string,
    getVal: (sub: ConsolidatedSubjectInfo) => string | number,
    valColor: string = '0F172A',
    rowBg: string = 'FFFFFF',
    labelBg: string = 'F1F5F9'
  ) => {
    const cells: TableCell[] = [
      new TableCell({
        columnSpan: 3,
        width: { size: snoPct + rollPct + namePct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        shading: { fill: labelBg, type: ShadingType.CLEAR },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: label, bold: true, size: 13, color: '0F172A' })] })],
      }),
    ];

    data.subjects.forEach((s) => {
      cells.push(
        new TableCell({
          width: { size: subPct, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: rowBg, type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(getVal(s)), bold: true, size: 14, color: valColor })] })],
        })
      );
    });

    summaryRows1.push(new TableRow({ children: cells }));
  };

  addDocxSummaryRow('Total Student Strength', (s) => data.subjectMetrics[s.code]?.totalStrength ?? data.totalStudentStrength, '0F172A', 'FFFFFF', 'F1F5F9');
  addDocxSummaryRow('Total Present', (s) => data.subjectMetrics[s.code]?.totalPresent ?? 0, '15803D', 'FFFFFF', 'F1F5F9');
  addDocxSummaryRow('Total Absent', (s) => data.subjectMetrics[s.code]?.totalAbsent ?? 0, 'DC2626', 'FFFFFF', 'F1F5F9');
  addDocxSummaryRow('Total Pass', (s) => data.subjectMetrics[s.code]?.totalPass ?? 0, '15803D', 'FFFFFF', 'F1F5F9');
  addDocxSummaryRow('Total Fail', (s) => data.subjectMetrics[s.code]?.totalFail ?? 0, 'DC2626', 'FFFFFF', 'F1F5F9');
  addDocxSummaryRow('Highest marks', (s) => data.subjectMetrics[s.code]?.highestMarks ?? 0, 'B45309', 'FFFFFF', 'F1F5F9');
  addDocxSummaryRow('Pass (%)', (s) => `${data.subjectMetrics[s.code]?.passPercentageStr ?? 0}%`, '1D4ED8', 'EFF6FF', 'EFF6FF');

  // Staff Name Row
  const staffDocxCells: TableCell[] = [
    new TableCell({
      columnSpan: 3,
      width: { size: snoPct + rollPct + namePct, type: WidthType.PERCENTAGE },
      borders: tableBorders,
      shading: { fill: 'F8FAFC', type: ShadingType.CLEAR },
      children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Staff Name', bold: true, size: 13, color: '334155' })] })],
    }),
  ];
  data.subjects.forEach((s) => {
    staffDocxCells.push(
      new TableCell({
        width: { size: subPct, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        shading: { fill: 'F8FAFC', type: ShadingType.CLEAR },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: s.facultyInCharge || 'Faculty', bold: true, size: 12, color: '0F172A' })] })],
      })
    );
  });
  summaryRows1.push(new TableRow({ children: staffDocxCells }));

  const summaryTableWidthPct = (snoPct + rollPct + namePct) + subPct * subCount;
  const summaryTable1 = new Table({
    width: { size: summaryTableWidthPct, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: summaryRows1,
  });

  // =========================================================================
  // 3. TABLE 2: SEPARATE DETAILED OVERALL RESULT & PERFORMANCE ANALYSIS TABLE IN WORD
  // =========================================================================
  const appearedCountDocx = data.students.filter((st) =>
    Object.values(st.subjects || {}).some(
      (s) => s && (s.attendance === 'Present' || (typeof s.marksObtained === 'number' && s.marksObtained >= 0))
    )
  ).length || data.totalPresent || data.totalStudentStrength;
  const failedAllCountDocx = Math.max(0, data.totalStudentStrength - data.passInAllSubjectsCount);
  const highestClassTotalDocx = data.students.length > 0 ? Math.max(...data.students.map((s) => s.totalObtained || 0)) : 0;
  const maxClassTotalDocx = data.subjects.length * 100;
  const passAllPctStrDocx = `${((data.passInAllSubjectsCount / (data.totalStudentStrength || 1)) * 100).toFixed(1)}%`;
  const classStatusStrDocx = data.overallPassPercentage >= 75 ? 'First Class / Exemplary' : data.overallPassPercentage >= 50 ? 'Satisfactory' : 'Needs Academic Attention';

  const overallRows1: TableRow[] = [
    // Header Row
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 4,
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: '0F2744', type: ShadingType.CLEAR },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 50, after: 50 },
              children: [new TextRun({ text: 'OVERALL RESULT & PERFORMANCE ANALYSIS', bold: true, size: 14, color: 'FFFFFF' })],
            }),
          ],
        }),
      ],
    }),
    // Row 1: Enrolled & Passed All
    new TableRow({
      children: [
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Total Students Enrolled:', bold: true, size: 13, color: '0F172A' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(data.totalStudentStrength), bold: true, size: 14, color: '0F172A' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'F0FDF4', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Passed in All Subjects (All Clear):', bold: true, size: 13, color: '166534' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'F0FDF4', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${data.passInAllSubjectsCount} (${passAllPctStrDocx})`, bold: true, size: 14, color: '15803D' })] })],
        }),
      ],
    }),
    // Row 2: Appeared & Failed
    new TableRow({
      children: [
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Students Appeared (Present):', bold: true, size: 13, color: '0F172A' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(appearedCountDocx), bold: true, size: 14, color: '0F172A' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Failed in One or More Subjects:', bold: true, size: 13, color: '0F172A' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(failedAllCountDocx), bold: true, size: 14, color: failedAllCountDocx > 0 ? 'DC2626' : '0F172A' })] })],
        }),
      ],
    }),
    // Row 3: Highest Total & Status
    new TableRow({
      children: [
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Class Highest Total Score:', bold: true, size: 13, color: '0F172A' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${highestClassTotalDocx} / ${maxClassTotalDocx}`, bold: true, size: 14, color: 'B45309' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Academic Performance Status:', bold: true, size: 13, color: '0F172A' })] })],
        }),
        new TableCell({
          width: { size: 25, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: classStatusStrDocx, bold: true, size: 13, color: '0F172A' })] })],
        }),
      ],
    }),
    // Row 4: Big Highlight KPI Banner
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 4,
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: tableBorders,
          shading: { fill: 'EFF6FF', type: ShadingType.CLEAR },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 60, after: 60 },
              children: [
                new TextRun({
                  text: `OVERALL CLASS PASS PERCENTAGE:  ${data.overallPassPercentage.toFixed(1)}%`,
                  bold: true,
                  size: 22,
                  color: '1D4ED8',
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];

  const overallTable1 = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: overallRows1,
  });

  // Signatures Table Helper with generous empty space for signing
  const createDocxSigTable = (titles: string[]) => {
    const widthPct = Math.floor(100 / titles.length);
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders,
      rows: [
        new TableRow({
          children: titles.map((t) =>
            new TableCell({
              width: { size: widthPct, type: WidthType.PERCENTAGE },
              borders: noBorders,
              children: [
                // Generous empty signing height (approx 20-25mm) so authorities can put their physical signature
                new Paragraph({ spacing: { before: 800, after: 0 }, children: [] }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: '______________', bold: true, size: 14, color: '64748B' })],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 40 },
                  children: [new TextRun({ text: t, bold: true, size: t.length > 15 ? 12 : 13, color: '0B1E36', font: 'Arial' })],
                }),
              ],
            })
          ),
        }),
      ],
    });
  };

  const signaturesTable1 = createDocxSigTable(['CC', 'HOD', 'AC', 'EXAM CELL COORDINATOR', 'PRINCIPAL', 'DIRECTOR']);

  // =========================================================================
  // PAGE 2 (WORD): "SUBJECT WISE ANALYSIS" MATCHING UPLOADED PHOTO
  // =========================================================================
  const page2Header = new Paragraph({
    pageBreakBefore: true,
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 15 },
    children: [
      new TextRun({
        text: 'Jaya Sakthi Engineering College',
        bold: true,
        size: 24,
        font: 'Times New Roman',
        color: '0B1E36',
      }),
    ],
  });

  const page2SubHeader = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 15 },
    children: [
      new TextRun({
        text: `${data.examType} Result Analysis`,
        bold: true,
        size: 18,
        font: 'Arial',
        color: '0F172A',
      }),
    ],
  });

  const page2Title = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 40 },
    children: [
      new TextRun({
        text: 'Subject Wise Analysis',
        bold: true,
        underline: {},
        size: 17,
        font: 'Arial',
        color: '0F172A',
      }),
    ],
  });

  // Page 2 Metadata Box
  const page2MetaTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: 'Department: ', bold: true, size: 15, color: '475569' }),
                  new TextRun({ text: data.department, bold: true, size: 15, color: '0F172A' }),
                ],
              }),
              new Paragraph({
                spacing: { before: 20 },
                children: [
                  new TextRun({ text: 'Year/ Semester: ', bold: true, size: 15, color: '475569' }),
                  new TextRun({ text: getYearSemRoman(data.semester), bold: true, size: 15, color: '0F172A' }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: tableBorders,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: 'Date: ', bold: true, size: 15, color: '475569' }),
                  new TextRun({ text: data.assessmentDate, bold: true, size: 15, color: '0F172A' }),
                ],
              }),
              new Paragraph({
                spacing: { before: 20 },
                children: [
                  new TextRun({ text: 'Section: ', bold: true, size: 15, color: '475569' }),
                  new TextRun({ text: data.section, bold: true, size: 15, color: '0F172A' }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  // Page 2 Subject Wise Analysis Table (Clean & Professional)
  const swTableHeaders1 = [
    'S.No',
    'Subject Code & Title',
    'Handled By (Faculty)',
    'Strength',
    'Appeared',
    'Absent',
    'Passed',
    'Failed',
    'Pass % Str',
    'Pass % App',
    'Highest Mark',
    'Evaluation Rating',
  ];

  const swDocxRows: TableRow[] = [
    new TableRow({
      children: swTableHeaders1.map((h) =>
        new TableCell({
          borders: tableBorders,
          shading: { fill: '0B1E36', type: ShadingType.CLEAR },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 30, after: 30 },
              children: [new TextRun({ text: h, bold: true, size: 13, color: 'FFFFFF', font: 'Arial' })],
            }),
          ],
        })
      ),
    }),
  ];

  let wTotAppeared = 0;
  let wTotAbsent = 0;
  let wTotPass = 0;
  let wTotFail = 0;

  data.subjects.forEach((sub, idx) => {
    const m = data.subjectMetrics[sub.code];
    const appPct = m?.passPercentageApp ?? 0;
    const appVal = m?.totalPresent ?? 0;
    const absVal = m?.totalAbsent ?? 0;
    const pVal = m?.totalPass ?? 0;
    const fVal = m?.totalFail ?? 0;

    wTotAppeared += appVal;
    wTotAbsent += absVal;
    wTotPass += pVal;
    wTotFail += fVal;

    let rating = 'Satisfactory';
    if (appPct >= 90) rating = 'Exemplary';
    else if (appPct >= 80) rating = 'Very Good';
    else if (appPct >= 75) rating = 'Target Met';
    else if (appPct < 60) rating = 'Remedial Req.';

    swDocxRows.push(
      new TableRow({
        children: [
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(idx + 1), size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: `${sub.code} - ${sub.name}`, bold: true, size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: sub.facultyInCharge || m?.facultyInCharge || 'Faculty In-Charge', size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(m?.totalStrength ?? data.totalStudentStrength), size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(appVal), size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(absVal), size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(pVal), bold: true, size: 13, color: '15803D' })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(fVal), bold: true, size: 13, color: '334155' })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${m?.passPercentageStr ?? 0}%`, size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${appPct}%`, bold: true, size: 13, color: '1E3A8A' })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(m?.highestMarks ?? 0), bold: true, size: 13 })] })],
          }),
          new TableCell({
            borders: tableBorders,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: rating, size: 12, bold: true })] })],
          }),
        ],
      })
    );
  });

  // Summary Row in Word
  const wAvgApp = data.subjects.length > 0 ? Math.round(data.subjects.reduce((sum, s) => sum + (data.subjectMetrics[s.code]?.passPercentageApp ?? 0), 0) / data.subjects.length) : 0;
  swDocxRows.push(
    new TableRow({
      children: [
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '', size: 13 })] })],
        }),
        new TableCell({
          borders: tableBorders,
          columnSpan: 2,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Total / Department Average', bold: true, size: 13, color: '0B1E36' })] })],
        }),
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(data.totalStudentStrength), bold: true, size: 13 })] })],
        }),
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(wTotAppeared), bold: true, size: 13 })] })],
        }),
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(wTotAbsent), bold: true, size: 13 })] })],
        }),
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(wTotPass), bold: true, size: 13, color: '15803D' })] })],
        }),
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(wTotFail), bold: true, size: 13 })] })],
        }),
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '-', size: 13 })] })],
        }),
        new TableCell({
          borders: tableBorders,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${wAvgApp}%`, bold: true, size: 13, color: '0B1E36' })] })],
        }),
        new TableCell({
          borders: tableBorders,
          columnSpan: 2,
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Overall Pass: ${data.overallPassPercentage.toFixed(1)}%`, bold: true, size: 12, color: '0B1E36' })] })],
        }),
      ],
    })
  );

  const swTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: swDocxRows,
  });

  const signaturesTable2 = createDocxSigTable(['CC', 'HOD', 'AC', 'EXAM CELL COORDINATOR', 'PRINCIPAL', 'DIRECTOR']);

  const sheetsToExport = matchingSheetsParam || data.matchingSheets || [];
  const subjectSections = sheetsToExport.map((sheet) => ({
    properties: {
      page: {
        size: { orientation: PageOrientation.PORTRAIT },
        margin: { top: 600, bottom: 600, left: 600, right: 600 },
      },
    },
    children: buildSubjectDocxChildren(sheet, collegeLogoBytes, naacLogoBytes, tableBorders, noBorders),
  }));

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { orientation: PageOrientation.LANDSCAPE },
            margin: { top: 600, bottom: 600, left: 600, right: 600 },
          },
        },
        children: [
          // PAGE 1: CONSOLIDATED STUDENT MARKSHEET
          headerTable,
          new Paragraph({ spacing: { before: 40, after: 40 } }),
          titleTable1,
          new Paragraph({ spacing: { before: 40, after: 40 } }),
          page1MetaTable,
          new Paragraph({ spacing: { before: 60, after: 60 } }),
          studentsTable1,
          new Paragraph({ spacing: { before: 140, after: 60 } }),
          summaryTable1,
          new Paragraph({ spacing: { before: 120, after: 60 } }),
          overallTable1,
          new Paragraph({ spacing: { before: 450, after: 120 } }),
          signaturesTable1,

          // PAGE 2: SUBJECT WISE ANALYSIS
          page2Header,
          page2SubHeader,
          page2Title,
          page2MetaTable,
          new Paragraph({ spacing: { before: 60, after: 60 } }),
          swTable,
          new Paragraph({ spacing: { before: 100, after: 60 } }),
          signaturesTable2,
        ],
      },
      ...subjectSections,
    ],
  });

  const rawBlob = await Packer.toBlob(doc);
  let finalBlob = rawBlob;

  try {
    const zip = await JSZip.loadAsync(rawBlob);

    // 1. Sanitize word/document.xml
    let docXml = await zip.file('word/document.xml')?.async('string');
    if (docXml) {
      // (a) Fix Markup Compatibility ignorable namespaces on <w:document>
      docXml = docXml.replace(
        /<w:document\b([^>]*?)mc:Ignorable="([^"]*)"/g,
        '<w:document $1mc:Ignorable="w14 w15 w16 w16cex w16cid w16se w16sdtdh wp14 wpc wpg wpi wps cx cx1 cx2 cx3 cx4 cx5 cx6 cx7 cx8 aink am3d"'
      );

      // (b) Fix <wp:docPr> (Word 2007 does not support 'title' attribute; name must be non-empty)
      let docPrCount = 0;
      docXml = docXml.replace(/<wp:docPr\b[^>]*\/?>/g, () => {
        docPrCount++;
        const name = `Picture_${docPrCount}`;
        return `<wp:docPr id="${docPrCount}" name="${name}" descr="${name}"/>`;
      });

      // (c) Fix <pic:cNvPr> (id must be > 0 and name non-empty)
      let cNvPrCount = 0;
      docXml = docXml.replace(/<pic:cNvPr\b[^>]*\/?>/g, () => {
        cNvPrCount++;
        const name = `Picture_${cNvPrCount}`;
        return `<pic:cNvPr id="${cNvPrCount + 100}" name="${name}"/>`;
      });

      // (d) Strip any residual inside vertical lines on borderless tables
      docXml = docXml.replace(/<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/g, (match, content) => {
        if (content.includes('w:top w:val="none"') || content.includes('w:insideV w:val="none"')) {
          content = content.replace(/<w:insideV w:val="[^"]*"/g, '<w:insideV w:val="none"');
          if (content.includes('w:top w:val="none"') || content.includes('w:insideH w:val="none"')) {
            content = content.replace(/<w:insideH w:val="[^"]*"/g, '<w:insideH w:val="none"');
          }
          return `<w:tblBorders>${content}</w:tblBorders>`;
        }
        return match;
      });

      zip.file('word/document.xml', docXml);
    }

    // 2. Sanitize word/settings.xml for Word 2007 (strip unsupported Office 2010+ compatSetting tags)
    let settingsXml = await zip.file('word/settings.xml')?.async('string');
    if (settingsXml) {
      settingsXml = settingsXml.replace(/<w:compatSetting\b[^>]*\/?>/g, '');
      settingsXml = settingsXml.replace(/<w:compat>\s*<\/w:compat>/g, '');
      settingsXml = settingsXml.replace(/w:val="1[45]"/g, 'w:val="12"');
      zip.file('word/settings.xml', settingsXml);
    }

    // 3. Sanitize docProps/app.xml for standard Word application headers
    let appXml = await zip.file('docProps/app.xml')?.async('string');
    if (appXml) {
      if (!appXml.includes('<Application>')) {
        if (appXml.includes('/>')) {
          appXml = appXml.replace(
            /<Properties([^>]*)\/>/,
            '<Properties$1><Application>Microsoft Office Word</Application><AppVersion>12.0000</AppVersion></Properties>'
          );
        } else {
          appXml = appXml.replace(
            /<\/Properties>/,
            '<Application>Microsoft Office Word</Application><AppVersion>12.0000</AppVersion></Properties>'
          );
        }
        zip.file('docProps/app.xml', appXml);
      }
    }

    finalBlob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  } catch (e) {
    console.warn('Word zip sanitization notice:', e);
  }

  triggerFileDownload(finalBlob, `${cleanName(data.department)}_${cleanName(data.semester)}_Consolidated_Marksheet.docx`);
}

function getColumnLetter(colIndex: number): string {
  let temp: number;
  let letter = '';
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = Math.floor((colIndex - temp - 1) / 26);
  }
  return letter;
}

function cleanName(str: string): string {
  return String(str || 'JSEC').split(' - ')[0].replace(/[^a-zA-Z0-9_-]/g, '_');
}

function triggerFileDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  setTimeout(() => {
    try {
      if (link.parentNode) {
        document.body.removeChild(link);
      }
      URL.revokeObjectURL(url);
    } catch (e) {
      console.warn('URL revoke notice:', e);
    }
  }, 4000);
}
