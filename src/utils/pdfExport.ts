import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { AssessmentMeta, StudentMarkRecord, EvaluationMetrics } from '../types';
import { COLLEGE_LOGO_BASE64 } from '../assets/logoBase64';
import { NAAC_LOGO_BASE64 } from '../assets/naacLogoBase64';

/**
 * Render institutional letterhead & document title banner matching the uploaded screenshot:
 * - Circular College Seal on the left
 * - Multi-tier centered institutional letterhead (Navy uppercase)
 * - NAAC 'A' Grade Medal on the right
 * - Double accent stripe (Dark Navy top rule + Golden Amber bottom rule)
 * - Deep navy rounded capsule/pill title banner with bold white text
 */
export function renderVectorHeaderFallback(
  doc: jsPDF,
  meta: AssessmentMeta,
  pageWidth: number,
  margin: number
): void {
  // 1. Official College Seal Logo (Left) and NAAC Medal (Right)
  try {
    if (COLLEGE_LOGO_BASE64) {
      doc.addImage(COLLEGE_LOGO_BASE64, 'PNG', 13, 8, 18, 18);
    }
    if (NAAC_LOGO_BASE64) {
      const naacFormat = NAAC_LOGO_BASE64.startsWith('data:image/jpeg') ? 'JPEG' : 'PNG';
      doc.addImage(NAAC_LOGO_BASE64, naacFormat, pageWidth - margin - 18, 9, 18, 14.5);
    }
  } catch (imgErr) {
    console.warn('Could not render college/naac logo in PDF:', imgErr);
  }

  // 2. Main College Name
  doc.setFont('times', 'bold');
  doc.setFontSize(16);
  doc.setTextColor(11, 30, 54); // Dark Navy #0b1e36
  doc.text('JAYA SAKTHI ENGINEERING COLLEGE', pageWidth / 2, 15.5, { align: 'center' });

  // Address directly below college name
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(51, 65, 85); // Slate-700
  doc.text(
    "St. Mary's Nagar, Thiruninravur, Chennai - 602024, Tamil Nadu",
    pageWidth / 2,
    20.2,
    { align: 'center' }
  );

  // AICTE & Anna University Affiliation
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor(100, 116, 139); // Slate-500
  doc.text(
    'Approved by AICTE, New Delhi & Affiliated to Anna University, Chennai',
    pageWidth / 2,
    24.5,
    { align: 'center' }
  );

  // 3. Decorative double line rule (Navy top line + Golden amber bottom line)
  doc.setDrawColor(15, 40, 78); // Navy
  doc.setLineWidth(0.7);
  doc.line(margin, 28, pageWidth - margin, 28);
  doc.setDrawColor(217, 119, 6); // Amber gold #d97706
  doc.setLineWidth(0.35);
  doc.line(margin, 29.2, pageWidth - margin, 29.2);

  // 4. DOCUMENT TITLE BANNER (Pill rounded corners exactly like the reference image)
  const bannerY = 32.5;
  const bannerHeight = 8.5;
  doc.setFillColor(11, 30, 54); // Dark Navy background #0b1e36
  doc.roundedRect(margin, bannerY, pageWidth - margin * 2, bannerHeight, 2.5, 2.5, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.setTextColor(255, 255, 255);
  doc.text(
    `CONTINUOUS ASSESSMENT MARKSHEET — ${meta.examType.toUpperCase()}`,
    pageWidth / 2,
    bannerY + 5.8,
    { align: 'center' }
  );
}

/**
 * Render a complete official Marksheet section for a single subject
 * exactly matching the user's uploaded reference design:
 * - Institutional Letterhead with college seal logo and double rule (Navy + Gold)
 * - Deep navy rounded capsule/pill title banner
 * - 2-column Metadata Box (Department, Semester, Subject, Date, Faculty, Marks Matrix)
 * - Detailed Student Marksheet Table (#, Register No, Student Name, Attendance, Max, Obt, Grade, Result)
 * - Evaluation Summary & Performance Metrics Card (Present & Pass in green, Absent & Fail in red/slate, Pass percentage pill)
 * - 3-Column Signatures: Faculty In-charge Signature, Head of the Department, Principal Signature
 */
export function renderSingleSubjectMarksheetToPDF(
  doc: jsPDF,
  meta: AssessmentMeta,
  students: StudentMarkRecord[],
  metrics: EvaluationMetrics,
  isFirstPage = true
): void {
  if (!isFirstPage) {
    doc.addPage('a4', 'portrait');
  }

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12;

  // 1. INSTITUTIONAL LETTERHEAD & BANNER
  renderVectorHeaderFallback(doc, meta, pageWidth, margin);

  // 2. METADATA SUMMARY BOX (Matches uploaded changes3.jpg: exact 4 rows)
  const metaBoxY = 43.5;
  const metaBoxHeight = 24.5;
  doc.setFillColor(248, 250, 252); // Slate-50
  doc.setDrawColor(203, 213, 225); // Slate-300
  doc.setLineWidth(0.2);
  doc.roundedRect(margin, metaBoxY, pageWidth - margin * 2, metaBoxHeight, 1.5, 1.5, 'FD');

  const col1LabelX = margin + 3.5;
  const col1ValX = margin + 26;
  const col2LabelX = margin + 102;
  const col2ValX = margin + 132;

  doc.setFontSize(8);

  // Row 1: Department & Assessment Date
  let curY = metaBoxY + 4.8;
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105); // Slate-700
  doc.text('Department:', col1LabelX, curY);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(15, 23, 42); // Slate-900
  doc.text(meta.department, col1ValX, curY);

  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Assessment Date:', col2LabelX, curY);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text(meta.assessmentDate || new Date().toISOString().split('T')[0], col2ValX, curY);

  // Row 2: Semester & Faculty In-Charge
  curY += 4.8;
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Semester:', col1LabelX, curY);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(15, 23, 42);
  doc.text(
    meta.academicYear ? `${meta.semester} (BATCH: ${meta.academicYear})` : meta.semester,
    col1ValX,
    curY
  );

  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Faculty In-Charge:', col2LabelX, curY);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text(meta.facultyInCharge || 'Faculty In-Charge', col2ValX, curY);

  // Row 3: Subject & Evaluator Name
  curY += 4.8;
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Subject:', col1LabelX, curY);
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(15, 23, 42);
  const subjNameFormatted = `${meta.subjectCode} - ${meta.subjectName}`.toUpperCase();
  const maxSubjWidth = col2LabelX - col1ValX - 4;
  const subjTruncated = doc.splitTextToSize(subjNameFormatted, maxSubjWidth)[0];
  doc.text(subjTruncated, col1ValX, curY);

  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Evaluator Name:', col2LabelX, curY);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text(meta.evaluatorName || 'DR. K.V. ARUN PILLAI', col2ValX, curY);

  // Row 4: Marks Matrix & Assessment Type
  curY += 4.8;
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Marks Matrix:', col1LabelX, curY);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text(`Max: ${meta.maxMarks} | Pass Mark: ${meta.passMark}`, col1ValX, curY);

  doc.setFont('helvetica', 'bold');
  doc.setTextColor(71, 85, 105);
  doc.text('Assessment Type:', col2LabelX, curY);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(15, 23, 42);
  doc.text(meta.examType, col2ValX, curY);

  // 3. DETAILED STUDENT MARKSHEET TABLE (#, Register No, Student Name, Attendance, Max, Obt, Grade, Result)
  const tableData = students.map((s, idx) => {
    const isAbsent = s.attendance === 'ABSENT';
    const obtStr = isAbsent ? 'AB' : s.marksObtained !== '' ? s.marksObtained.toString() : '-';
    return [
      (idx + 1).toString(),
      s.regNo,
      s.name.toUpperCase(),
      s.attendance,
      s.maxMarks.toString(),
      obtStr,
      s.grade,
      s.result,
    ];
  });

  autoTable(doc, {
    startY: metaBoxY + metaBoxHeight + 3.5,
    margin: { left: margin, right: margin, bottom: 26 },
    tableWidth: pageWidth - margin * 2,
    head: [
      ['#', 'Register No', 'Student Name', 'Attendance', 'Max', 'Obt', 'Grade', 'Result'],
    ],
    body: tableData,
    theme: 'grid',
    headStyles: {
      fillColor: [11, 30, 54], // Dark navy #0B1E36
      textColor: [255, 255, 255],
      fontStyle: 'bold',
      fontSize: 8,
      halign: 'center',
      cellPadding: 2,
    },
    styles: {
      fontSize: 8,
      cellPadding: 1.8,
      textColor: [30, 41, 59],
      lineWidth: 0.15,
      lineColor: [226, 232, 240],
    },
    alternateRowStyles: {
      fillColor: [248, 250, 252],
    },
    columnStyles: {
      0: { halign: 'center', cellWidth: 9 },
      1: { halign: 'center', cellWidth: 32, fontStyle: 'bold' },
      2: { halign: 'left', cellWidth: 57, fontStyle: 'bold' },
      3: { halign: 'center', cellWidth: 22 },
      4: { halign: 'center', cellWidth: 15 },
      5: { halign: 'center', cellWidth: 15, fontStyle: 'bold' },
      6: { halign: 'center', cellWidth: 15, fontStyle: 'bold' },
      7: { halign: 'center', cellWidth: 21, fontStyle: 'bold' },
    },
    didParseCell: (data) => {
      if (data.section === 'body') {
        const rowData = students[data.row.index];
        if (!rowData) return;

        // Result column (index 7): PASS in green, FAIL in red
        if (data.column.index === 7) {
          if (rowData.result === 'PASS') {
            data.cell.styles.textColor = [21, 128, 61]; // #15803D Green
          } else if (rowData.result === 'FAIL') {
            data.cell.styles.textColor = [220, 38, 38]; // #DC2626 Red
          }
        }
        // Attendance column (index 3)
        if (data.column.index === 3 && rowData.attendance === 'ABSENT') {
          data.cell.styles.textColor = [180, 83, 9]; // Amber
          data.cell.styles.fontStyle = 'bold';
        }
        // Marks Obtained (index 5)
        if (data.column.index === 5 && rowData.attendance === 'ABSENT') {
          data.cell.styles.textColor = [180, 83, 9];
        }
      }
    },
  });

  // 4. EVALUATION SUMMARY & PERFORMANCE METRICS (Matches uploaded screenshot exactly)
  const lastTableY = (doc as any).lastAutoTable.finalY || 180;
  const evalCardHeight = 26;
  const sigSectionHeight = 18;
  let evalBoxY = lastTableY + 4;

  // Page break if necessary
  if (evalBoxY + evalCardHeight + sigSectionHeight + 8 > pageHeight - 12) {
    doc.addPage('a4', 'portrait');
    evalBoxY = 16;
  }

  const evalBoxWidth = pageWidth - margin * 2;

  // Outer evaluation card container
  doc.setFillColor(255, 255, 255);
  doc.setDrawColor(203, 213, 225); // Slate-300
  doc.setLineWidth(0.3);
  doc.roundedRect(margin, evalBoxY, evalBoxWidth, evalCardHeight, 2, 2, 'FD');

  // Title: "EVALUATION SUMMARY & PERFORMANCE METRICS"
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(15, 23, 42); // Slate-900
  doc.text('EVALUATION SUMMARY & PERFORMANCE METRICS', margin + 4, evalBoxY + 5.2);

  // Subtle divider line
  doc.setDrawColor(226, 232, 240); // Slate-200
  doc.setLineWidth(0.25);
  doc.line(margin + 3.5, evalBoxY + 7.2, margin + evalBoxWidth - 3.5, evalBoxY + 7.2);

  // Metrics grid coordinates
  const leftColLabelX = margin + 6;
  const leftColValX = margin + 50;
  const rightColLabelX = margin + 84;
  const rightColValX = margin + 130;

  const row1Y = evalBoxY + 11.5;
  const row2Y = evalBoxY + 15.8;

  doc.setFontSize(8);

  // Row 1: "No of student Present:" (green) & "No of student Pass:" (green)
  doc.setFont('helvetica', 'bold');
  doc.setTextColor(30, 41, 59);
  doc.text('No of student Present:', leftColLabelX, row1Y);
  doc.setTextColor(21, 128, 61); // Green #15803D
  doc.text(metrics.presentCount.toString(), leftColValX, row1Y);

  doc.setTextColor(30, 41, 59);
  doc.text('No of student Pass:', rightColLabelX, row1Y);
  doc.setTextColor(21, 128, 61); // Green #15803D
  doc.text(metrics.passCount.toString(), rightColValX, row1Y);

  // Row 2: "No of student Absent:" (slate) & "no of student Fail:" (red/slate)
  doc.setTextColor(30, 41, 59);
  doc.text('No of student Absent:', leftColLabelX, row2Y);
  doc.setTextColor(15, 23, 42);
  doc.text(metrics.absentCount.toString(), leftColValX, row2Y);

  doc.setTextColor(30, 41, 59);
  doc.text('no of student Fail:', rightColLabelX, row2Y);
  doc.setTextColor(metrics.failCount > 0 ? 220 : 15, metrics.failCount > 0 ? 38 : 23, metrics.failCount > 0 ? 38 : 42);
  doc.text(metrics.failCount.toString(), rightColValX, row2Y);

  // Centered Pass percentage banner pill
  const pillY = evalBoxY + 18.5;
  const pillHeight = 5.2;
  const pillMargin = 5;
  const pillWidth = evalBoxWidth - pillMargin * 2;
  doc.setFillColor(238, 242, 255); // Indigo/Blue-50
  doc.roundedRect(margin + pillMargin, pillY, pillWidth, pillHeight, 1.5, 1.5, 'F');

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8.5);
  doc.setTextColor(29, 78, 216); // Blue #1D4ED8
  doc.text(
    `Pass percentage: ${metrics.passPercentage.toFixed(2)}%`,
    margin + evalBoxWidth / 2,
    pillY + 3.8,
    { align: 'center' }
  );

  // 5. SIGNATURE SECTION (3 Signatures exactly matching uploaded screenshot)
  const sigY = evalBoxY + evalCardHeight + 11;
  const sigBlockWidth = (pageWidth - margin * 2) / 3;

  const sigs = [
    {
      title: 'Evaluator Signature',
      name: meta.evaluatorName ? `(${meta.evaluatorName})` : '(DR. K.V. ARUN PILLAI)',
      x: margin + sigBlockWidth * 0.5,
    },
    {
      title: 'Exam Cell Coordinator',
      name: meta.examCellCoordinator?.trim() ? `(${meta.examCellCoordinator.trim()})` : '',
      x: margin + sigBlockWidth * 1.5,
    },
    {
      title: 'Principal Signature',
      name: (meta.principalName?.trim() || meta.principalSignature?.trim())
        ? `(${meta.principalName?.trim() || meta.principalSignature?.trim()})`
        : '',
      x: margin + sigBlockWidth * 2.5,
    },
  ];

  sigs.forEach((s) => {
    // Signature underline
    doc.setDrawColor(100, 116, 139); // Slate-500
    doc.setLineWidth(0.4);
    doc.line(s.x - 26, sigY, s.x + 26, sigY);

    // Title
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor(30, 41, 59);
    doc.text(s.title, s.x, sigY + 4.5, { align: 'center' });

    // Optional Name
    if (s.name) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(6.8);
      doc.setTextColor(100, 116, 139);
      doc.text(s.name, s.x, sigY + 8.2, { align: 'center' });
    }
  });
}

/**
 * Generate and download an official Marksheet PDF document
 * using jsPDF and jspdf-autotable.
 */
export async function exportMarksheetToPDF(
  meta: AssessmentMeta,
  students: StudentMarkRecord[],
  metrics: EvaluationMetrics
): Promise<string> {
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4',
  });

  renderSingleSubjectMarksheetToPDF(doc, meta, students, metrics, true);

  // Footer on all pages
  const totalPages = doc.getNumberOfPages();
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12;

  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.5);
    doc.setTextColor(148, 163, 184);

    const printDate = new Date().toLocaleString('en-IN', {
      timeZone: 'Asia/Kolkata',
      dateStyle: 'medium',
      timeStyle: 'short',
    });
    doc.text(
      `Generated by JSEC Central Examination Cell on ${printDate} IST`,
      margin,
      pageHeight - 6
    );
    doc.text(`Page ${i} of ${totalPages}`, pageWidth - margin, pageHeight - 6, {
      align: 'right',
    });
  }

  const cleanSubject = (meta.subjectCode || 'Marksheet').replace(/[^a-zA-Z0-9_-]/g, '_');
  const cleanSem = (meta.semester || 'Assessment').replace(/[^a-zA-Z0-9_-]/g, '_');
  const filename = `${cleanSubject}_${cleanSem}_Marksheet.pdf`;

  try {
    doc.save(filename);
  } catch {
    const blob = doc.output('blob');
    const blobUrl = URL.createObjectURL(blob);
    const downloadLink = document.createElement('a');
    downloadLink.href = blobUrl;
    downloadLink.download = filename;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    setTimeout(() => {
      document.body.removeChild(downloadLink);
      URL.revokeObjectURL(blobUrl);
    }, 1500);
  }

  return filename;
}
