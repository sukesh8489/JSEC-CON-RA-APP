import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { AssessmentMeta, StudentMarkRecord, EvaluationMetrics } from '../types';
import { COLLEGE_LOGO_BASE64 } from '../assets/logoBase64';
import { NAAC_LOGO_BASE64 } from '../assets/naacLogoBase64';

/**
 * Builds an official Marksheet worksheet for a single subject into the provided ExcelJS Workbook
 * strictly adhering to the user's uploaded reference screenshot:
 * - Institutional branding & centered letterhead (Navy)
 * - Decorative double horizontal rule (Navy top + Amber gold bottom)
 * - Deep navy capsule title banner: CONTINUOUS ASSESSMENT MARKSHEET — [EXAM_TYPE]
 * - 3-row clean metadata matrix (Department, Semester, Subject, Date, Faculty, Marks Matrix)
 * - Column Headers: #, Register No, Student Name, Attendance, Max, Obt, Grade, Result
 * - Formatted student rows with green PASS and red FAIL
 * - EVALUATION SUMMARY & PERFORMANCE METRICS card (Present/Pass in green, Absent/Fail, Pass percentage pill)
 * - 3 Official Signatures: Faculty In-charge Signature, Head of the Department, Principal Signature
 */
export function buildSubjectMarksheetWorksheet(
  workbook: ExcelJS.Workbook,
  meta: AssessmentMeta,
  students: StudentMarkRecord[],
  metrics: EvaluationMetrics,
  sheetName = 'Assessment Marksheet'
): ExcelJS.Worksheet {
  // Sanitize worksheet name (max 31 chars, no invalid chars)
  const safeTitle = sheetName.replace(/[\\/?*:[\]]/g, '_').substring(0, 31) || 'Marksheet';
  const worksheet = workbook.addWorksheet(safeTitle, {
    views: [{ showGridLines: true }],
    pageSetup: {
      paperSize: 9, // A4
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: {
        left: 0.5,
        right: 0.5,
        top: 0.5,
        bottom: 0.5,
        header: 0.3,
        footer: 0.3,
      },
    },
  });

  // Column definitions (8 columns: A through H matching the 8 screenshot columns, full-screen width)
  worksheet.columns = [
    { key: 'sno', width: 8 },          // A: #
    { key: 'regNo', width: 24 },       // B: Register No
    { key: 'name', width: 40 },        // C: Student Name
    { key: 'attendance', width: 16 },  // D: Attendance
    { key: 'max', width: 12 },         // E: Max
    { key: 'obt', width: 12 },         // F: Obt
    { key: 'grade', width: 12 },       // G: Grade
    { key: 'result', width: 16 },      // H: Result
  ];

  // Colors & styles
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  const navyDark = 'FF0B1E36';
  const slate900 = 'FF0F172A';
  const slate700 = 'FF334155';
  const slate500 = 'FF64748B';
  const greenPass = 'FF15803D';
  const redFail = 'FFDC2626';

  // Embed College Seal Logo (Left)
  try {
    if (COLLEGE_LOGO_BASE64) {
      const cleanBase64 = COLLEGE_LOGO_BASE64.replace(/^data:image\/\w+;base64,/, '');
      const imageId = workbook.addImage({
        base64: cleanBase64,
        extension: 'png',
      });
      worksheet.addImage(imageId, {
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
      worksheet.addImage(naacImgId, {
        tl: { col: 6.9, row: 0.1 },
        ext: { width: 72, height: 58 },
      });
    }
  } catch (err) {
    console.warn('Could not add logos to Excel:', err);
  }

  // Row 1: College Name
  const row1 = worksheet.getRow(1);
  row1.height = 26;
  worksheet.mergeCells('A1:H1');
  const cellA1 = worksheet.getCell('A1');
  cellA1.value = 'JAYA SAKTHI ENGINEERING COLLEGE';
  cellA1.font = { name: 'Times New Roman', size: 16, bold: true, color: { argb: navyDark } };
  cellA1.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 2: Address
  const row2 = worksheet.getRow(2);
  row2.height = 17;
  worksheet.mergeCells('A2:H2');
  const cellA2 = worksheet.getCell('A2');
  cellA2.value = "St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu";
  cellA2.font = { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } };
  cellA2.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 3: AICTE Affiliation
  const row3 = worksheet.getRow(3);
  row3.height = 17;
  worksheet.mergeCells('A3:H3');
  const cellA3 = worksheet.getCell('A3');
  cellA3.value = 'Approved by AICTE, New Delhi & Affiliated to Anna University, Chennai';
  cellA3.font = { name: 'Arial', size: 8.5, color: { argb: slate500 } };
  cellA3.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 4: Decorative dual rule (Navy top line + Golden amber bottom line)
  const row4 = worksheet.getRow(4);
  row4.height = 6;
  for (let c = 1; c <= 8; c++) {
    const cell = row4.getCell(c);
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0B1E36' } },
      bottom: { style: 'thin', color: { argb: 'FFD97706' } },
    };
  }

  // Row 5: Assessment Title Banner
  const row5 = worksheet.getRow(5);
  row5.height = 30;
  worksheet.mergeCells('A5:H5');
  const cellA5 = worksheet.getCell('A5');
  cellA5.value = `CONTINUOUS ASSESSMENT MARKSHEET — ${meta.examType.toUpperCase()}`;
  cellA5.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  cellA5.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: navyDark } };
  cellA5.alignment = { horizontal: 'center', vertical: 'middle' };
  for (let c = 1; c <= 8; c++) {
    row5.getCell(c).border = {
      top: { style: 'thin', color: { argb: navyDark } },
      bottom: { style: 'thin', color: { argb: navyDark } },
      left: { style: 'thin', color: { argb: navyDark } },
      right: { style: 'thin', color: { argb: navyDark } },
    };
  }

  // Row 6: Spacing
  worksheet.getRow(6).height = 6;

  // 3. METADATA SUMMARY BOX (Matches uploaded screenshot exactly: 3 clean rows)
  const metaFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF8FAFC' },
  };

  const styleMetaRow = (rowNum: number, h: number) => {
    const r = worksheet.getRow(rowNum);
    r.height = h;
    for (let c = 1; c <= 8; c++) {
      const cell = r.getCell(c);
      cell.fill = metaFill;
      cell.border = thinBorder;
      cell.alignment = { vertical: 'middle' };
    }
  };

  // Row 7: Department (Left) & Assessment Date (Right)
  styleMetaRow(7, 21);
  worksheet.mergeCells('A7:D7');
  worksheet.mergeCells('E7:H7');
  const cellDept = worksheet.getCell('A7');
  cellDept.value = {
    richText: [
      { text: 'Department:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: meta.department, font: { name: 'Arial', size: 9, bold: true, color: { argb: slate900 } } },
    ],
  };
  cellDept.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  const cellDate = worksheet.getCell('E7');
  cellDate.value = {
    richText: [
      { text: 'Assessment Date:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: meta.assessmentDate || new Date().toISOString().split('T')[0], font: { name: 'Arial', size: 9, color: { argb: slate900 } } },
    ],
  };
  cellDate.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // Row 8: Semester (Left) & Faculty In-Charge (Right)
  styleMetaRow(8, 21);
  worksheet.mergeCells('A8:D8');
  worksheet.mergeCells('E8:H8');
  const cellSem = worksheet.getCell('A8');
  const semText = meta.academicYear ? `${meta.semester} (BATCH: ${meta.academicYear})` : meta.semester;
  cellSem.value = {
    richText: [
      { text: 'Semester:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: semText, font: { name: 'Arial', size: 9, bold: true, color: { argb: slate900 } } },
    ],
  };
  cellSem.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  const cellFaculty = worksheet.getCell('E8');
  cellFaculty.value = {
    richText: [
      { text: 'Faculty In-Charge:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: meta.facultyInCharge || 'Faculty In-Charge', font: { name: 'Arial', size: 9, color: { argb: slate900 } } },
    ],
  };
  cellFaculty.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // Row 9: Subject (Left) & Evaluator Name (Right)
  styleMetaRow(9, 21);
  worksheet.mergeCells('A9:D9');
  worksheet.mergeCells('E9:H9');
  const cellSubj = worksheet.getCell('A9');
  cellSubj.value = {
    richText: [
      { text: 'Subject:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: `${meta.subjectCode} - ${meta.subjectName}`.toUpperCase(), font: { name: 'Arial', size: 9, bold: true, color: { argb: slate900 } } },
    ],
  };
  cellSubj.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  const cellEval = worksheet.getCell('E9');
  cellEval.value = {
    richText: [
      { text: 'Evaluator Name:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: meta.evaluatorName || 'DR. K.V. ARUN PILLAI', font: { name: 'Arial', size: 9, color: { argb: slate900 } } },
    ],
  };
  cellEval.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // Row 10: Marks Matrix (Left) & Assessment Type (Right)
  styleMetaRow(10, 21);
  worksheet.mergeCells('A10:D10');
  worksheet.mergeCells('E10:H10');
  const cellMarks = worksheet.getCell('A10');
  cellMarks.value = {
    richText: [
      { text: 'Marks Matrix:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: `Max: ${meta.maxMarks}  |  Pass Mark: ${meta.passMark}`, font: { name: 'Arial', size: 9, color: { argb: slate900 } } },
    ],
  };
  cellMarks.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  const cellType = worksheet.getCell('E10');
  cellType.value = {
    richText: [
      { text: 'Assessment Type:  ', font: { name: 'Arial', size: 9, bold: true, color: { argb: slate700 } } },
      { text: meta.examType, font: { name: 'Arial', size: 9, color: { argb: slate900 } } },
    ],
  };
  cellType.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

  // Row 11: Spacing
  worksheet.getRow(11).height = 6;

  // 4. STUDENT DETAILED MARKSHEET TABLE (#, Register No, Student Name, Attendance, Max, Obt, Grade, Result)
  const headerRow = worksheet.getRow(12);
  headerRow.height = 24;
  const headers = ['#', 'Register No', 'Student Name', 'Attendance', 'Max', 'Obt', 'Grade', 'Result'];

  headers.forEach((hdr, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = hdr;
    cell.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: navyDark } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = thinBorder;
  });

  // Student Data Rows
  let currentRowNum = 13;
  students.forEach((student, index) => {
    const row = worksheet.getRow(currentRowNum);
    row.height = 19;

    const isEven = index % 2 === 1;
    const rowBg = isEven ? 'FFF8FAFC' : 'FFFFFFFF';
    const isPass = student.result === 'PASS';
    const isAbsent = student.attendance === 'ABSENT';

    // Col 1: #
    const c1 = row.getCell(1);
    c1.value = index + 1;
    c1.alignment = { horizontal: 'center', vertical: 'middle' };
    c1.font = { name: 'Arial', size: 9, color: { argb: slate500 } };

    // Col 2: Register No
    const c2 = row.getCell(2);
    c2.value = student.regNo;
    c2.alignment = { horizontal: 'center', vertical: 'middle' };
    c2.font = { name: 'Arial', size: 9, bold: true, color: { argb: slate900 } };

    // Col 3: Student Name
    const c3 = row.getCell(3);
    c3.value = student.name.toUpperCase();
    c3.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    c3.font = { name: 'Arial', size: 9, bold: true, color: { argb: slate900 } };

    // Col 4: Attendance
    const c4 = row.getCell(4);
    c4.value = student.attendance;
    c4.alignment = { horizontal: 'center', vertical: 'middle' };
    c4.font = {
      name: 'Arial',
      size: 8.5,
      bold: true,
      color: { argb: isAbsent ? 'FFB45309' : slate900 },
    };

    // Col 5: Max
    const c5 = row.getCell(5);
    c5.value = student.maxMarks;
    c5.alignment = { horizontal: 'center', vertical: 'middle' };
    c5.font = { name: 'Arial', size: 9, color: { argb: slate700 } };

    // Col 6: Obt
    const c6 = row.getCell(6);
    c6.value = isAbsent ? 'AB' : student.marksObtained !== '' ? Number(student.marksObtained) : '-';
    c6.alignment = { horizontal: 'center', vertical: 'middle' };
    c6.font = {
      name: 'Arial',
      size: 9,
      bold: true,
      color: { argb: isAbsent ? 'FFB45309' : slate900 },
    };

    // Col 7: Grade
    const c7 = row.getCell(7);
    c7.value = student.grade;
    c7.alignment = { horizontal: 'center', vertical: 'middle' };
    c7.font = { name: 'Arial', size: 9, bold: true, color: { argb: slate900 } };

    // Col 8: Result (PASS in green, FAIL in red)
    const c8 = row.getCell(8);
    c8.value = student.result;
    c8.alignment = { horizontal: 'center', vertical: 'middle' };
    c8.font = {
      name: 'Arial',
      size: 9,
      bold: true,
      color: { argb: isPass ? greenPass : redFail },
    };

    for (let c = 1; c <= 8; c++) {
      const cell = row.getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      cell.border = thinBorder;
    }

    currentRowNum++;
  });

  // Spacing after table
  worksheet.getRow(currentRowNum).height = 8;
  currentRowNum++;

  // 5. EVALUATION SUMMARY & PERFORMANCE METRICS
  const evalStartRow = currentRowNum;
  // Header row
  const rEvalH = worksheet.getRow(evalStartRow);
  rEvalH.height = 20;
  worksheet.mergeCells(`A${evalStartRow}:H${evalStartRow}`);
  const cEvalH = worksheet.getCell(`A${evalStartRow}`);
  cEvalH.value = 'EVALUATION SUMMARY & PERFORMANCE METRICS';
  cEvalH.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
  cEvalH.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  cEvalH.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F2744' } };
  for (let c = 1; c <= 8; c++) {
    rEvalH.getCell(c).border = thinBorder;
    rEvalH.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F2744' } };
  }

  // Row 1: Present & Pass
  currentRowNum++;
  const rEval1 = worksheet.getRow(currentRowNum);
  rEval1.height = 20;
  worksheet.mergeCells(`A${currentRowNum}:B${currentRowNum}`);
  worksheet.mergeCells(`C${currentRowNum}:D${currentRowNum}`);
  worksheet.mergeCells(`E${currentRowNum}:F${currentRowNum}`);
  worksheet.mergeCells(`G${currentRowNum}:H${currentRowNum}`);

  const cPresL = worksheet.getCell(`A${currentRowNum}`);
  cPresL.value = 'No of student Present:';
  cPresL.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0F172A' } };
  cPresL.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  cPresL.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  const cPresV = worksheet.getCell(`C${currentRowNum}`);
  cPresV.value = metrics.presentCount;
  cPresV.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF15803D' } };
  cPresV.alignment = { horizontal: 'left', vertical: 'middle' };
  cPresV.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  const cPassL = worksheet.getCell(`E${currentRowNum}`);
  cPassL.value = 'No of student Pass:';
  cPassL.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0F172A' } };
  cPassL.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  cPassL.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  const cPassV = worksheet.getCell(`G${currentRowNum}`);
  cPassV.value = metrics.passCount;
  cPassV.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF15803D' } };
  cPassV.alignment = { horizontal: 'left', vertical: 'middle' };
  cPassV.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  for (let c = 1; c <= 8; c++) {
    rEval1.getCell(c).border = thinBorder;
  }

  // Row 2: Absent & Fail
  currentRowNum++;
  const rEval2 = worksheet.getRow(currentRowNum);
  rEval2.height = 20;
  worksheet.mergeCells(`A${currentRowNum}:B${currentRowNum}`);
  worksheet.mergeCells(`C${currentRowNum}:D${currentRowNum}`);
  worksheet.mergeCells(`E${currentRowNum}:F${currentRowNum}`);
  worksheet.mergeCells(`G${currentRowNum}:H${currentRowNum}`);

  const cAbsL = worksheet.getCell(`A${currentRowNum}`);
  cAbsL.value = 'No of student Absent:';
  cAbsL.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0F172A' } };
  cAbsL.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  cAbsL.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  const cAbsV = worksheet.getCell(`C${currentRowNum}`);
  cAbsV.value = metrics.absentCount;
  cAbsV.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFDC2626' } };
  cAbsV.alignment = { horizontal: 'left', vertical: 'middle' };
  cAbsV.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  const cFailL = worksheet.getCell(`E${currentRowNum}`);
  cFailL.value = 'no of student Fail:';
  cFailL.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0F172A' } };
  cFailL.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  cFailL.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };

  const cFailV = worksheet.getCell(`G${currentRowNum}`);
  cFailV.value = metrics.failCount;
  cFailV.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFDC2626' } };
  cFailV.alignment = { horizontal: 'left', vertical: 'middle' };
  cFailV.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  for (let c = 1; c <= 8; c++) {
    rEval2.getCell(c).border = thinBorder;
  }

  // Row 3: Centered Pass percentage - Highlighted Soft Blue Card
  currentRowNum++;
  const rPill = worksheet.getRow(currentRowNum);
  rPill.height = 22;
  worksheet.mergeCells(`A${currentRowNum}:H${currentRowNum}`);
  const cPill = worksheet.getCell(`A${currentRowNum}`);
  cPill.value = `Pass percentage: ${metrics.passPercentage.toFixed(2)}%`;
  cPill.font = { name: 'Arial', size: 10.5, bold: true, color: { argb: 'FF1D4ED8' } };
  cPill.alignment = { horizontal: 'center', vertical: 'middle' };
  cPill.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
  for (let c = 1; c <= 8; c++) {
    rPill.getCell(c).border = thinBorder;
  }

  // 6. SIGNATURE SECTION (3 Signatures with separate dedicated lines, balanced column alignment)
  // Evaluator: Cols A:B (width 32)
  // Exam Cell Coordinator: Cols C:E (width 68) — gives full room for title & full name without truncation
  // Principal: Cols F:H (width 40)
  currentRowNum += 3;
  const signLineEvaluator = '________________________';
  const signLineCoordinator = '________________________________________';
  const signLinePrincipal = '________________________';

  // Row 1: Separate dedicated lines for each signature
  const sigLineRow = worksheet.getRow(currentRowNum);
  sigLineRow.height = 16;
  worksheet.mergeCells(`A${currentRowNum}:B${currentRowNum}`);
  worksheet.mergeCells(`C${currentRowNum}:E${currentRowNum}`);
  worksheet.mergeCells(`F${currentRowNum}:H${currentRowNum}`);

  const cSigL1 = worksheet.getCell(`A${currentRowNum}`);
  cSigL1.value = signLineEvaluator;
  cSigL1.alignment = { horizontal: 'center', vertical: 'bottom' };
  cSigL1.font = { name: 'Arial', size: 9, color: { argb: 'FF94A3B8' } };

  const cSigL2 = worksheet.getCell(`C${currentRowNum}`);
  cSigL2.value = signLineCoordinator;
  cSigL2.alignment = { horizontal: 'center', vertical: 'bottom' };
  cSigL2.font = { name: 'Arial', size: 9, color: { argb: 'FF94A3B8' } };

  const cSigL3 = worksheet.getCell(`F${currentRowNum}`);
  cSigL3.value = signLinePrincipal;
  cSigL3.alignment = { horizontal: 'center', vertical: 'bottom' };
  cSigL3.font = { name: 'Arial', size: 9, color: { argb: 'FF94A3B8' } };

  currentRowNum++;

  // Row 2: Signature Titles (Exam Cell Coordinator gets width 68 across C, D, E so it displays fully without any truncation)
  const sigTitleRow = worksheet.getRow(currentRowNum);
  sigTitleRow.height = 20;
  worksheet.mergeCells(`A${currentRowNum}:B${currentRowNum}`);
  worksheet.mergeCells(`C${currentRowNum}:E${currentRowNum}`);
  worksheet.mergeCells(`F${currentRowNum}:H${currentRowNum}`);

  const cSigT1 = worksheet.getCell(`A${currentRowNum}`);
  cSigT1.value = 'Evaluator Signature';
  cSigT1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  cSigT1.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: slate700 } };

  const cSigT2 = worksheet.getCell(`C${currentRowNum}`);
  cSigT2.value = 'Exam Cell Coordinator';
  cSigT2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  cSigT2.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: slate700 } };

  const cSigT3 = worksheet.getCell(`F${currentRowNum}`);
  cSigT3.value = 'Principal Signature';
  cSigT3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  cSigT3.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: slate700 } };

  currentRowNum++;

  // Row 3: Name strings below signatures (Exam Cell Coordinator name displays fully in 68-width merged C:E block)
  const sigNameRow = worksheet.getRow(currentRowNum);
  sigNameRow.height = 18;
  worksheet.mergeCells(`A${currentRowNum}:B${currentRowNum}`);
  worksheet.mergeCells(`C${currentRowNum}:E${currentRowNum}`);
  worksheet.mergeCells(`F${currentRowNum}:H${currentRowNum}`);

  const cSigN1 = worksheet.getCell(`A${currentRowNum}`);
  cSigN1.value = meta.evaluatorName?.trim() ? `(${meta.evaluatorName.trim()})` : '';
  cSigN1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  cSigN1.font = { name: 'Arial', size: 8.5, color: { argb: slate500 } };

  const cSigN2 = worksheet.getCell(`C${currentRowNum}`);
  cSigN2.value = meta.examCellCoordinator?.trim() ? `(${meta.examCellCoordinator.trim()})` : '';
  cSigN2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  cSigN2.font = { name: 'Arial', size: 8.5, color: { argb: slate500 } };

  const cSigN3 = worksheet.getCell(`F${currentRowNum}`);
  const enteredPrincipal = meta.principalName?.trim() || meta.principalSignature?.trim();
  cSigN3.value = enteredPrincipal ? `(${enteredPrincipal})` : '';
  cSigN3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  cSigN3.font = { name: 'Arial', size: 8.5, color: { argb: slate500 } };

  return worksheet;
}

/**
 * Sanitize an ExcelJS generated XLSX package to ensure 100% full compatibility with
 * Microsoft Excel 2007 (and Excel 2010/2013/2016/2019/365).
 * - Replaces 32-bit unsigned DPI values (4294967295) with standard 300 DPI (prevents schema parser crash in 2007)
 * - Strips Office 2014 extension a16:creationId from DrawingML drawing*.xml (prevents unreadable content alert)
 * - Cleans Office 2010/2013 extension blocks in xl/styles.xml
 * - Normalizes docProps/app.xml to Excel 2007 (AppVersion 12.0000)
 */
export async function sanitizeExcelFor2007(buffer: ArrayBuffer | Uint8Array): Promise<Blob> {
  try {
    const zip = await JSZip.loadAsync(buffer);

    for (const fileName of Object.keys(zip.files)) {
      if (fileName.startsWith('xl/worksheets/sheet') && fileName.endsWith('.xml')) {
        let sheetXml = await zip.file(fileName)!.async('string');
        sheetXml = sheetXml.replace(/horizontalDpi="4294967295"/g, 'horizontalDpi="300"');
        sheetXml = sheetXml.replace(/verticalDpi="4294967295"/g, 'verticalDpi="300"');
        sheetXml = sheetXml.replace(/<mergeCell\s+ref="([A-Z]+[0-9]+):\1"\s*\/>/g, '');
        zip.file(fileName, sheetXml);
      }
      if (fileName.startsWith('xl/drawings/drawing') && fileName.endsWith('.xml')) {
        let drawingXml = await zip.file(fileName)!.async('string');
        drawingXml = drawingXml.replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g, '');
        zip.file(fileName, drawingXml);
      }
      if (fileName === 'xl/styles.xml') {
        let stylesXml = await zip.file(fileName)!.async('string');
        stylesXml = stylesXml.replace(/<extLst>[\s\S]*?<\/extLst>/g, '');
        zip.file(fileName, stylesXml);
      }
      if (fileName === 'docProps/app.xml') {
        let appXml = await zip.file(fileName)!.async('string');
        appXml = appXml.replace(/<AppVersion>[^<]*<\/AppVersion>/g, '<AppVersion>12.0000</AppVersion>');
        zip.file(fileName, appXml);
      }
    }

    return await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (err) {
    console.warn('Excel 2007 zip sanitization notice:', err);
    return new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }
}

/**
 * Generate and download an official Marksheet Excel (.xlsx) document for a single subject
 */
export async function exportMarksheetToExcel(
  meta: AssessmentMeta,
  students: StudentMarkRecord[],
  metrics: EvaluationMetrics
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Jaya Sakthi Engineering College';
  workbook.lastModifiedBy = 'Central Examination Cell';
  workbook.created = new Date();
  workbook.modified = new Date();

  buildSubjectMarksheetWorksheet(
    workbook,
    meta,
    students,
    metrics,
    `${meta.subjectCode || 'Marksheet'}`.substring(0, 31)
  );

  const cleanSubject = (meta.subjectCode || 'Marksheet').replace(/[^a-zA-Z0-9_-]/g, '_');
  const cleanSem = (meta.semester || 'Assessment').replace(/[^a-zA-Z0-9_-]/g, '_');
  const filename = `${cleanSubject}_${cleanSem}_Marksheet.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = await sanitizeExcelFor2007(buffer);

  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  setTimeout(() => {
    try {
      if (link.parentNode) {
        document.body.removeChild(link);
      }
      URL.revokeObjectURL(downloadUrl);
    } catch {
      // ignore
    }
  }, 1500);
}
