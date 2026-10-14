// Builds the MediKiosk mini-project report (.docx) in the KSRCE Design Studio format.
// Usage: node build_report.js [pages.json]
//   pages.json (optional) maps TOC keys to page labels; produced by the second pass.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, LevelFormat, WidthType, BorderStyle, ShadingType,
  Footer, PageNumber, NumberFormat, VerticalAlign, TabStopType,
} = require('docx');

const A = (f) => path.join(__dirname, 'assets', f);
const PAGES = process.argv[2] && fs.existsSync(process.argv[2]) ? JSON.parse(fs.readFileSync(process.argv[2])) : {};

// ---------------------------------------------------------------- details to edit
const INFO = {
  title: 'AI POWERED PATIENT CASE-TAKING KIOSK FOR HOSPITALS',
  project: 'MediKiosk',
  course: '24CBP34 - DESIGN STUDIO - I LABORATORY',
  student: 'STUDENT NAME',
  regno: 'REGISTER NUMBER',
  month: 'October 2026',
  monthShort: 'Oct 2026',
  hod: 'Mrs. K.Sudha, M.E, Ph.D',
  supervisor: 'Ms. A.Priyadharshini, M.E',
};
const STUDENT_LINE = `${INFO.student} (${INFO.regno})`;

// ---------------------------------------------------------------- layout constants
const FONT = 'Times New Roman';
const PAGE_W = 11906, PAGE_H = 16838;
const M_LEFT = 2160, M_RIGHT = 1440, M_TOP = 1440, M_BOTTOM = 1440;
const TEXT_W = PAGE_W - M_LEFT - M_RIGHT; // 8306 twips ≈ 554 px
const TEXT_PX = Math.floor(TEXT_W / 15);
const LINE = 360; // 1.5 line spacing

// ---------------------------------------------------------------- inline markup: **bold**, *italic*
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ text: t.slice(2, -2), ...base, bold: true }));
    else out.push(new TextRun({ text: t.slice(1, -1), ...base, italics: true }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}

// ---------------------------------------------------------------- block helpers
const para = (text, o = {}) => new Paragraph({
  alignment: o.align ?? AlignmentType.JUSTIFIED,
  indent: o.indent === false ? undefined : { firstLine: o.indent ?? 720 },
  spacing: { line: LINE, after: o.after ?? 120, before: o.before ?? 0 },
  keepNext: o.keepNext,
  children: runs(text, o.run || {}),
});
const plain = (text, o = {}) => para(text, { ...o, indent: false });
const center = (text, o = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: o.line ?? LINE, before: o.before ?? 0, after: o.after ?? 120 },
  children: runs(text, { size: o.size ?? 24, bold: o.bold, italics: o.italics }),
});
const chapter = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, pageBreakBefore: true,
  spacing: { before: 240, after: 480, line: LINE },
  children: [new TextRun({ text, bold: true, size: 32, font: FONT })],
});
const frontTitle = (text, o = {}) => new Paragraph({
  alignment: AlignmentType.CENTER, pageBreakBefore: o.pageBreak ?? true,
  spacing: { before: o.before ?? 240, after: o.after ?? 360, line: LINE },
  children: [new TextRun({ text, bold: true, size: 28, font: FONT })],
});
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2, keepNext: true,
  spacing: { before: 280, after: 120, line: LINE },
  children: [new TextRun({ text, bold: true, size: 24, font: FONT })],
});
const h3 = (text) => new Paragraph({
  keepNext: true, spacing: { before: 160, after: 80, line: LINE },
  children: [new TextRun({ text, bold: true, size: 24 })],
});
const bullet = (text, level = 0) => new Paragraph({
  numbering: { reference: 'bullets', level },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: 60 },
  children: runs(text),
});
const numbered = (ref) => (text) => new Paragraph({
  numbering: { reference: ref, level: 0 },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: 100 },
  children: runs(text),
});
// numbered heading + body, the "1. Lack of ... / indented paragraph" pattern of the template
const point = (n, title, body) => [h3(`${n}. ${title}`), para(body)];

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
function figure(file, caption, widthPx = TEXT_PX) {
  const { w, h } = pngSize(A(file));
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 200, after: 80 },
      children: [new ImageRun({ type: 'png', data: fs.readFileSync(A(file)),
        transformation: { width: widthPx, height: Math.round(widthPx * h / w) } })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 240 },
      children: [new TextRun({ text: caption, bold: true, size: 22 })],
    }),
  ];
}
const spacer = (pts = 1) => new Paragraph({ spacing: { after: 0, line: 240 * pts }, children: [] });

// ---------------------------------------------------------------- tables
const thin = { style: BorderStyle.SINGLE, size: 4, color: '999999' };
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const allBorders = { top: thin, bottom: thin, left: thin, right: thin };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };

function cell(content, width, o = {}) {
  const children = (Array.isArray(content) ? content : [content]).map((c) =>
    c instanceof Paragraph ? c : new Paragraph({
      alignment: o.align ?? AlignmentType.LEFT, spacing: { before: 40, after: 40, line: o.line ?? 276 },
      children: runs(String(c), { bold: o.bold, size: o.size ?? 24 }),
    }));
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders: o.borders ?? allBorders,
    verticalAlign: o.valign ?? VerticalAlign.TOP,
    shading: o.shade ? { fill: o.shade, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children,
  });
}
function grid(widths, header, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (r, isHead) => new TableRow({
    tableHeader: isHead, cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], { bold: isHead || (o.boldFirst && i === 0), size: o.size, shade: isHead ? o.headShade : undefined })),
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    alignment: AlignmentType.CENTER,
    rows: [mk(header, true), ...rows.map((r) => mk(r, false))],
  });
}

// ================================================================ FRONT MATTER
const logo = (file, px) => {
  const { w, h } = pngSize(A(file));
  return new Paragraph({ alignment: AlignmentType.CENTER, children: [
    new ImageRun({ type: 'png', data: fs.readFileSync(A(file)), transformation: { width: px, height: Math.round(px * h / w) } })] });
};

const cover = [
  center(`**${INFO.title}**`, { size: 28, before: 720, after: 960 }),
  center(`**${INFO.course}**`, { size: 26, after: 160 }),
  center('**A MINI PROJECT REPORT**', { size: 26, after: 840 }),
  center('Submitted by', { size: 26, after: 480, bold: true, italics: true }),
  center(`**${INFO.student}  -  (${INFO.regno})**`, { size: 26, after: 240 }),
  center('*in partial fulfilment for the award of the degree of*', { size: 26, after: 480 }),
  center('**BACHELOR OF ENGINEERING**', { size: 26, after: 60 }),
  center('in', { size: 26, after: 240, bold: true, italics: true }),
  center('**COMPUTER SCIENCE AND ENGINEERING**', { size: 26, after: 0 }),
  center('**(CYBER SECURITY)**', { size: 26, after: 1200 }),
  new Table({
    width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: [1700, 4906, 1700],
    rows: [new TableRow({ children: [
      cell(logo('ksrce_logo.png', 96), 1700, { borders: noBorders, valign: VerticalAlign.CENTER }),
      cell([
        center('**K.S.R COLLEGE OF ENGINEERING**', { size: 26, after: 0 }),
        center('**(Autonomous)**', { size: 26, after: 0 }),
        center('**TIRUCHENGODE 637 215**', { size: 26, after: 0 }),
      ], 4906, { borders: noBorders, valign: VerticalAlign.CENTER }),
      cell(logo('anna_univ_logo.png', 96), 1700, { borders: noBorders, valign: VerticalAlign.CENTER }),
    ] })],
  }),
  center('**ANNA UNIVERSITY :: CHENNAI 600 025**', { size: 26, before: 60, after: 120 }),
  center(`**${INFO.month}**`, { size: 26 }),
];

const sigCol = (role, name, desig) => [
  plain('**SIGNATURE**', { after: 0 }),
  plain(name, { after: 0 }),
  plain(`**${role}**`, { after: 0 }),
  plain(desig, { after: 0 }),
  plain('Department of CSE(Cyber Security)', { after: 0 }),
  plain('K.S.R College of Engineering', { after: 0 }),
  plain('Tiruchengode - 637 215', { after: 0 }),
];
const bonafide = [
  frontTitle('K.S.R. COLLEGE OF ENGINEERING (AUTONOMOUS)', { after: 0, before: 360 }),
  center('**TIRUCHENGODE 637215**', { size: 28, after: 360 }),
  center('**ANNA UNIVERSITY:: CHENNAI 600025**', { size: 28, after: 480 }),
  center('**BONAFIDE CERTIFICATE**', { size: 28, after: 360 }),
  plain(`Certified that this report “**${INFO.title}**” is the bonafide work of **${STUDENT_LINE}** who carried out the project work for Design Studio I under my supervision.`, { after: 2400 }),
  new Table({
    width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: [4153, 4153],
    rows: [new TableRow({ children: [
      cell(sigCol('HEAD OF THE DEPARTMENT', INFO.hod, 'Professor & Head,'), 4153, { borders: noBorders }),
      cell(sigCol('SUPERVISOR', INFO.supervisor, 'Assistant Professor,'), 4153, { borders: noBorders }),
    ] })],
  }),
  new Paragraph({ border: { bottom: { style: BorderStyle.DASHED, size: 8, color: '000000', space: 4 } }, spacing: { after: 240 }, children: [] }),
  plain('Submitted for the project viva-voce examination held on ………………………', { after: 960 }),
  new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W }],
    children: [new TextRun({ text: 'Internal Examiner', bold: true }), new TextRun({ text: '\tExternal Examiner', bold: true })],
  }),
];

const acknowledgement = [
  frontTitle('ACKNOWLEDGEMENT'),
  para('We feel highly honored to extend our sincere gratitude to our beloved Founder Theivathiru Lion **Dr. K. S. Rangasamy, MJF.** and our Chairman **Mr.R.Srinivasan, B.B.M., M.I.S.T.E.,** for giving us an opportunity and facilities for the completion of this project.'),
  para('We express our heartfelt gratitude to our respected Dean **Dr. M.Venkatesan, M.E., Ph.D.,** for the continuous support, valuable guidance, and encouragement extended to us in all our academic endeavors.'),
  para('We express our deep sense of gratitude to our beloved Principal **Dr. P.Meenakshi Devi, M.E., Ph.D.,** who supported us in all our endeavors and has been responsible for motivating us throughout our career.'),
  para('We feel highly elated to thank our respectable Head of the Department **Mrs. K.Sudha M.E., Ph.D** who guided us and was a pillar of support for the successful completion of the project and our guide **Ms. A.Priyadharshini, M.E.,** who has guided us by providing fruitful ideas in various stages of our project and helping us in completing this project.'),
  para('It is a pleasure to express our gratitude to our beloved parents for providing their support and confidence to us for the completion of the project and our heartfelt thanks to our entire department faculty members, non-teaching staff and beloved friends, directly and indirectly who helped us during the tenure of the project.'),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 360 }, children: [new TextRun({ text: STUDENT_LINE, bold: true })] }),
];

const declaration = [
  frontTitle('DEPARTMENT OF COMPUTER SCIENCE AND', { after: 0, before: 360 }),
  center('**ENGINEERING (CYBER SECURITY)**', { size: 28, after: 240 }),
  center('**K.S.R. COLLEGE OF ENGINEERING (Autonomous)**', { size: 28, after: 0 }),
  center('**TIRUCHENGODE 637 215**', { size: 28, after: 0 }),
  center(`**${INFO.monthShort}**`, { size: 28, after: 0 }),
  center('**DECLARATION**', { size: 28, after: 360 }),
  para(`We affirm that the project report titled being “**${INFO.title}**” submitted in partial fulfilment of the requirements for the award of the degree of Bachelor of Engineering is the original work carried out by us. It has not formed the part of any other project report or dissertation on the basis of which a degree or award was conferred on an earlier occasion on this or any other candidates.`, { after: 1440 }),
  plain(`**${STUDENT_LINE}**`, { after: 1440 }),
  plain('To verify that the declaration made by the above candidates is true to the best of my knowledge.', { after: 1200 }),
  new Paragraph({
    tabStops: [{ type: TabStopType.LEFT, position: 2900 }],
    children: [new TextRun('Date:'), new TextRun('\tName and Signature of the supervisor with seal')],
  }),
];

const abstract = [
  frontTitle('ABSTRACT', { pageBreak: false, before: 720, after: 480 }),
  para(`Outpatient departments in Indian hospitals, including Ayush institutions such as the All India Institute of Ayurveda, handle very large patient volumes, and a significant share of every consultation is spent on routine case-taking: recording the chief complaint, history of present illness, past illness, allergies and lifestyle, and reading through the paper prescriptions and lab reports that patients carry with them. This process is repetitive, often incomplete, frequently hampered by language barriers, and leaves the physician with less time for examination and counselling. **${INFO.project}**, developed for Smart India Hackathon problem statement SIH26047, is an AI-powered self-service kiosk that conducts this clinical intake before the patient meets the doctor. The patient identifies themselves through an ABHA number or QR code, chooses one of eight Indian languages, gives consent in line with the Digital Personal Data Protection Act 2023, and then answers a conversational, doctor-like interview by voice or touch. The interview supports two clinical modes: an allopathic history structured on the SOCRATES framework and an Ayush mode based on Ayurvedic Dashavidha Pariksha (Agni, Koshtha, Prakriti, Ahara and Vihara). A rule-based red-flag engine detects emergency patterns such as severe chest pain with sweating or breathlessness and immediately escalates the patient to a triage nurse. Prior medical documents are scanned, their key findings are extracted and abnormal values are flagged, and the system finally generates a structured case summary that the patient reviews and the physician receives as an editable draft, ready to be linked to the patient’s ABHA record through HL7 FHIR. A working web prototype has been built with HTML5, CSS3, JavaScript and the Web Speech API, and was verified through fourteen automated functional tests covering the full five-step flow on desktop and mobile screens, all of which passed. ${INFO.project} demonstrates how conversational AI, multilingual speech and interoperable digital health standards can shorten OPD waiting time, improve the completeness of clinical histories and support safer, faster care across both modern and traditional systems of medicine.`),
];

// ---------------------------------------------------------------- TOC
const TOC_ROWS = [
  ['', 'ABSTRACT', 'abstract', true],
  ['', 'LIST OF ABBREVIATIONS', 'abbrev', true],
  ['1', 'INTRODUCTION', 'ch1', true],
  ['', '1.1 INTRODUCTION OF THE PROJECT', '1.1'],
  ['', '1.2 NEED OF THE PROJECT', '1.2'],
  ['', '1.3 OBJECTIVE OF THE PROJECT', '1.3'],
  ['', '1.4 THEME OF THE PROJECT', '1.4'],
  ['', '1.5 ORGANIZATION', '1.5'],
  ['', '1.6 SUMMARY', '1.6'],
  ['2', 'LITERATURE REVIEW', 'ch2', true],
  ['', '2.1 INTRODUCTION', '2.1'],
  ['', '2.2 EXISTING SYSTEMS AND THEIR LIMITATIONS', '2.2'],
  ['', '2.3 SCOPE OF THE PROJECT', '2.3'],
  ['', '2.4 SUMMARY', '2.4'],
  ['3', 'SYSTEM DESIGN AND DEVELOPMENT', 'ch3', true],
  ['', '3.1 FUNDAMENTALS OF AI-BASED CLINICAL INTAKE SYSTEMS', '3.1'],
  ['', '3.2 PROPOSED SYSTEM ARCHITECTURE', '3.2'],
  ['', '3.3 SYSTEM WORKFLOW', '3.3'],
  ['', '3.4 IMPLEMENTATION DETAILS', '3.4'],
  ['', '3.5 SUMMARY', '3.5'],
  ['4', 'RESULT AND DISCUSSION', 'ch4', true],
  ['', '4.1 FUNCTIONAL ANALYSIS', '4.1'],
  ['', '4.2 EXPERIMENTAL ANALYSIS', '4.2'],
  ['', '4.3 SUMMARY', '4.3'],
  ['5', 'CONCLUSIONS', 'ch5', true],
  ['', '5.1 CONCLUSION', '5.1'],
  ['', '5.2 FUTURE SCOPE', '5.2'],
  ['', '5.3 APPLICATIONS', '5.3'],
  ['', '5.4 SUMMARY', '5.4'],
  ['', 'REFERENCES', 'refs'],
];
const tocTable = new Table({
  width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: [1800, 5006, 1500],
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell('CHAPTER NO', 1800, { borders: noBorders, bold: true }),
      cell('TITLE', 5006, { borders: noBorders, bold: true }),
      cell('PAGE NO', 1500, { borders: noBorders, bold: true, align: AlignmentType.CENTER }),
    ] }),
    ...TOC_ROWS.map(([no, title, key, bold]) => new TableRow({ cantSplit: true, children: [
      cell(no, 1800, { borders: noBorders, bold: true, line: LINE }),
      cell(title, 5006, { borders: noBorders, bold, line: LINE }),
      cell(PAGES[key] ?? '0', 1500, { borders: noBorders, align: AlignmentType.CENTER, line: LINE }),
    ] })),
  ],
});
const toc = [frontTitle('TABLE OF CONTENTS'), tocTable];

const ABBR = [
  ['ABDM', 'Ayushman Bharat Digital Mission'],
  ['ABHA', 'Ayushman Bharat Health Account'],
  ['AI', 'Artificial Intelligence'],
  ['AIIA', 'All India Institute of Ayurveda'],
  ['API', 'Application Programming Interface'],
  ['ASR', 'Automatic Speech Recognition'],
  ['AYUSH', 'Ayurveda, Yoga & Naturopathy, Unani, Siddha, Sowa-Rigpa and Homoeopathy'],
  ['CC', 'Chief Complaint'],
  ['DPDP', 'Digital Personal Data Protection'],
  ['EMR', 'Electronic Medical Record'],
  ['FHIR', 'Fast Healthcare Interoperability Resources'],
  ['HIS', 'Hospital Information System'],
  ['HL7', 'Health Level Seven'],
  ['HPI', 'History of Present Illness'],
  ['HTML', 'HyperText Markup Language'],
  ['LLM', 'Large Language Model'],
  ['ML', 'Machine Learning'],
  ['NLP', 'Natural Language Processing'],
  ['OCR', 'Optical Character Recognition'],
  ['OPD', 'Out-Patient Department'],
  ['ROS', 'Review of Systems'],
  ['SIH', 'Smart India Hackathon'],
  ['TTS', 'Text-To-Speech'],
  ['UI', 'User Interface'],
];
const abbreviations = [
  frontTitle('LIST OF ABBREVIATIONS'),
  grid([2400, 5400], ['TERMS', 'EXPANSION'], ABBR),
];

// ================================================================ CHAPTER 1
const P = INFO.project;
const ch1 = [
  chapter('CHAPTER 1 – INTRODUCTION'),
  h2('1.1 INTRODUCTION OF THE PROJECT'),
  para('Healthcare delivery in India is under constant pressure from the sheer number of patients that visit public and teaching hospitals every day. A single outpatient department may register thousands of patients in a morning, and each of them must be interviewed before a diagnosis can be made. Case-taking — the structured collection of the chief complaint, the history of the present illness, past medical history, medication and allergy details, family and personal history — is the foundation of every clinical decision, yet it is also one of the most time-consuming and repetitive parts of a consultation.'),
  para('In practice, case-taking in a busy OPD is often rushed, incomplete and recorded on paper. Patients arrive with bundles of old prescriptions and lab reports in different formats and languages, and the physician has to read through them while simultaneously interviewing the patient. Language differences between patients and doctors, especially in multilingual states and in national institutes that receive patients from across the country, add further delay and risk of misunderstanding. In Ayush hospitals the challenge is even greater, because an Ayurvedic assessment requires additional parameters such as Prakriti (constitution), Agni (digestive capacity) and Koshtha (bowel nature) that are rarely captured in a structured digital form.'),
  para(`To address these challenges, the project **${P}** introduces an AI-powered patient case-taking kiosk, developed for Smart India Hackathon problem statement SIH26047 proposed under the Ministry of Ayush / All India Institute of Ayurveda. The kiosk is placed in the hospital waiting area. While the patient waits for their turn, it identifies them through their ABHA (Ayushman Bharat Health Account) number, talks to them in their preferred language, conducts a doctor-like interview by voice or touch, scans their previous medical documents, and produces a concise, structured case summary that reaches the physician before the patient enters the consultation room.`),
  para(`The kiosk supports two clinical modes. The **Allopathy mode** follows the widely taught SOCRATES approach to symptom analysis, while the **Ayush mode** follows the classical Ayurvedic Dashavidha Pariksha, capturing digestion, bowel habit, constitution, diet and daily routine. A built-in red-flag engine continuously watches the answers for emergency patterns, such as severe chest pain with sweating or breathlessness, and immediately escalates such patients to a triage nurse rather than leaving them in the queue.`),
  para(`In essence, ${P} acts as a digital clinical assistant that prepares the ground for the doctor. By automating routine history-taking and document review, it allows the physician to spend more of the consultation on examination, diagnosis and counselling, while giving patients a faster, more comfortable and language-inclusive experience. Because every summary can be linked to the patient’s ABHA record, the system also contributes to the longitudinal digital health record envisioned by the Ayushman Bharat Digital Mission.`),

  h2('1.2 NEED OF THE PROJECT'),
  para('Indian hospitals face a combination of high patient load, limited consultation time and fragmented medical records. Despite rapid digitisation in registration and billing, the clinical history itself is still largely captured verbally and on paper. This leads to incomplete histories, repeated questioning, long waiting times and avoidable clinical risk.'),
  plain('The need for this project arises from the following critical gaps in the current outpatient process:'),
  ...point(1, 'High Patient Volume and Short Consultation Time', 'Doctors in government and teaching hospitals often have only a few minutes per patient. A large part of this time is consumed by basic questions that could be collected in advance, leaving less time for examination and explanation.'),
  ...point(2, 'Incomplete and Unstructured Case Histories', 'Histories written quickly on paper frequently miss important details such as allergies, previous illnesses or the onset and severity of symptoms. Unstructured notes are also difficult to search, share or analyse later.'),
  ...point(3, 'Language Barriers', 'India has 22 scheduled languages and many dialects. National institutes receive patients from all over the country, and a patient who cannot explain their problem in the doctor’s language is at a clear disadvantage.'),
  ...point(4, 'Scattered Paper Records', 'Patients carry old prescriptions, discharge summaries and lab reports from different hospitals. Reading and correlating these documents during the consultation is slow, and abnormal values are easily overlooked.'),
  ...point(5, 'Delayed Recognition of Emergencies', 'In a crowded waiting hall, a patient with a potentially serious condition such as cardiac chest pain may wait in the general queue. An automated screening step that flags such symptoms early can save lives.'),
  ...point(6, 'Lack of Structured Ayush Case-Taking', 'Ayurvedic diagnosis depends on parameters such as Prakriti, Agni and Koshtha. These are rarely captured in a standardised digital form, which limits research, outcome tracking and integration of Ayush with mainstream digital health.'),
  ...point(7, 'Need for Interoperable Digital Health Records', 'The Ayushman Bharat Digital Mission envisions a longitudinal health record for every citizen. Case histories collected at the point of care must be structured and standards-compliant (FHIR) so that they can be linked to ABHA and reused across facilities.'),

  h2('1.3 OBJECTIVE OF THE PROJECT'),
  para(`The primary objective of the ${P} project is to develop an AI-powered, multilingual self-service kiosk that collects a complete, structured clinical history from patients before consultation, flags emergencies early, and delivers an interoperable case summary to the physician.`),
  plain('The major objectives of the project are as follows:'),
  ...point(1, 'To Automate Patient Case-Taking', 'To conduct a guided, conversational interview that captures the chief complaint, history of present illness, past history, allergies and lifestyle factors without the physician’s direct involvement.'),
  ...point(2, 'To Support Both Allopathy and Ayush Systems', 'To provide a SOCRATES-based allopathic interview and a Dashavidha Pariksha based Ayurvedic interview within the same kiosk, so that it can be used in integrated hospitals.'),
  ...point(3, 'To Provide Multilingual Voice Interaction', 'To let patients interact in their preferred Indian language using speech or touch, with spoken feedback for patients with limited literacy.'),
  ...point(4, 'To Detect Red-Flag Symptoms Early', 'To identify emergency patterns during the interview and immediately alert triage staff and move the patient to a priority queue.'),
  ...point(5, 'To Digitise Previous Medical Documents', 'To scan prescriptions, lab reports and discharge summaries, extract key findings and highlight abnormal values for the physician.'),
  ...point(6, 'To Generate a Structured Clinical Summary', 'To convert the interview and document data into a concise summary (CC, HPI, past history, allergy, ROS or Ayush parameters, investigations) that the doctor can review and edit.'),
  ...point(7, 'To Integrate with the National Digital Health Ecosystem', 'To identify patients through ABHA and design the summary to be shared as HL7 FHIR resources with the Hospital Information System and the patient’s ABDM health record.'),
  ...point(8, 'To Ensure Privacy and Consent', 'To capture explicit patient consent, follow the principles of the Digital Personal Data Protection Act 2023, and clear session data from the kiosk after each patient.'),

  h2('1.4 THEME OF THE PROJECT'),
  para(`The theme of ${P} is **AI for Accessible and Efficient Healthcare**. It applies conversational artificial intelligence, speech technology and interoperable health data standards to one of the most fundamental and time-consuming steps in medical practice: taking a patient’s history.`),
  para('At its core, the project emphasises **Digital Health Innovation**, enabling patients to describe their problem comfortably in their own language and enabling doctors to start every consultation with an organised, complete picture of the patient. It integrates multiple information sources — the conversational interview, scanned documents and the patient’s ABHA identity — into a single summary that supports informed clinical decisions.'),
  plain('This project aligns with national priorities like:'),
  bullet('**Ayushman Bharat Digital Mission (ABDM)**'),
  bullet('**Digital India and Bhashini (National Language Translation Mission)**'),
  bullet('**Integration of Ayush with Mainstream Healthcare**'),
  bullet('**Digital Personal Data Protection Act, 2023**'),
  bullet('**AI for Social Good in Healthcare**'),
  plain('By combining technology with established clinical practice, the theme focuses on:'),
  h3('• Empowering Patients with Technology'),
  plain('Allowing every patient, regardless of language or literacy, to communicate their health problem clearly through voice and simple touch options.'),
  h3('• Reducing the Burden on Doctors'),
  plain('Automating repetitive history-taking and document review so that physicians can focus on examination, diagnosis and counselling.'),
  h3('• Improving Patient Safety'),
  plain('Detecting emergency symptoms early and ensuring that critical information such as allergies and abnormal lab values is never missed.'),
  h3('• Bridging Traditional and Modern Medicine'),
  plain('Capturing Ayurvedic assessment parameters in a structured digital form alongside conventional clinical history.'),
  h3('• Supporting Interoperable, Consent-Based Records'),
  plain('Producing summaries designed for ABHA linkage and FHIR exchange, with patient consent at the centre of every data flow.'),

  h2('1.5 ORGANIZATION'),
  plain('This report is organized into five chapters:'),
  bullet('**Chapter 1** introduces the project, its need, objectives and theme.'),
  bullet('**Chapter 2** reviews existing systems and related work on digital case-taking, telemedicine and clinical decision support, and identifies their limitations.'),
  bullet('**Chapter 3** details the system design, proposed architecture, workflow and implementation methodology.'),
  bullet('**Chapter 4** presents the results, functional analysis, experimental test results and discussion.'),
  bullet('**Chapter 5** concludes the report and highlights possible future improvements and applications.'),

  h2('1.6 SUMMARY'),
  para(`Chapter 1 presented an overview of **${P} – AI Powered Patient Case-Taking Kiosk**, highlighting the motivation, relevance and foundational concepts behind the project. The chapter explained why manual case-taking in busy outpatient departments is slow, incomplete and affected by language barriers, and why an intelligent, multilingual intake system is valuable for hospitals across India, including Ayush institutions.`),
  para('The need for the project was discussed in terms of patient load, incomplete histories, scattered records, delayed emergency recognition and the absence of structured Ayush case-taking. The objectives were clearly defined, and the theme was linked to national missions such as ABDM and Bhashini. Finally, the organisation of the report was outlined, preparing the reader for the detailed review, design and analysis that follow.'),
];

// ================================================================ CHAPTER 2
const ch2 = [
  chapter('CHAPTER 2 – LITERATURE REVIEW'),
  h2('2.1 INTRODUCTION'),
  para('A literature review is essential for understanding the existing research, technologies and digital platforms related to patient case-taking and clinical decision support. It shows how hospitals currently record patient histories, how digital health platforms and AI tools have evolved, and where present systems still fall short. This chapter evaluates previously developed systems and research directions to identify the gaps that the proposed kiosk addresses.'),
  para('Over the last decade, India has made major progress in digital health. Hospital Information Systems manage registration and billing, telemedicine services connect patients with doctors remotely, ABDM provides a unique health ID and a framework for sharing records, and Bhashini offers speech and translation services for Indian languages. In parallel, research in natural language processing and large language models has shown that AI can understand clinical text and conduct structured medical conversations. However, these capabilities have not yet been combined into a single, patient-facing tool that performs complete case-taking at the point of care.'),
  para(`This chapter reviews existing approaches to clinical history collection and the technologies that support it. It establishes the foundation for understanding how **${P}** advances beyond these systems through multilingual conversational intake, dual allopathy/Ayush support, automatic red-flag triage and standards-based record integration.`),

  h2('2.2 EXISTING SYSTEMS AND THEIR LIMITATIONS'),
  para('Various systems are currently used to record patient information and to support doctors. The key existing systems and their limitations are summarised below:'),
  h3('1. Manual Paper-Based Case-Taking in the OPD'),
  plain('(Case sheets and OPD cards filled in by the doctor or a junior resident)'),
  plain('**Limitations:**'),
  bullet('Time-consuming and dependent on the doctor’s available time'),
  bullet('Often incomplete, illegible and not searchable'),
  bullet('Not shareable across departments or hospitals'),
  h3('2. Hospital Information Systems / EMR Platforms'),
  plain('(e.g., NIC e-Hospital, state HMIS and private EMR software)'),
  plain('**Limitations:**'),
  bullet('Focus mainly on registration, billing, pharmacy and laboratory workflows'),
  bullet('Clinical history is still typed by the doctor during the consultation'),
  bullet('No patient-facing, multilingual interview capability'),
  h3('3. Telemedicine Platforms'),
  plain('(e.g., eSanjeevani national telemedicine service)'),
  plain('**Limitations:**'),
  bullet('History is taken verbally by the doctor during the video consultation'),
  bullet('Not designed for in-hospital pre-consultation intake'),
  bullet('Do not automatically digitise patients’ paper documents'),
  h3('4. AI Symptom Checker Applications'),
  plain('(e.g., consumer apps such as Ada Health and similar chatbots)'),
  plain('**Limitations:**'),
  bullet('Built for self-assessment, not for passing a structured history to a treating doctor'),
  bullet('Limited support for Indian languages and low-literacy users'),
  bullet('No Ayush assessment and no ABHA / hospital integration'),
  h3('5. ABHA-Based Registration (Scan and Share)'),
  plain('(QR-code based OPD registration under ABDM)'),
  plain('**Limitations:**'),
  bullet('Speeds up demographic registration only'),
  bullet('Does not capture symptoms or clinical history'),
  h3('6. Ayush Hospital Management Systems'),
  plain('(e.g., Ayush Health Management Information System – AHMIS)'),
  plain('**Limitations:**'),
  bullet('Record Ayush consultations but rely on the physician for data entry'),
  bullet('No guided, patient-facing Dashavidha Pariksha interview'),
  h3('Key Limitations Across All Existing Systems'),
  bullet('Case-taking is still performed manually by the doctor'),
  bullet('No multilingual, voice-based patient interaction'),
  bullet('No automatic early detection of red-flag symptoms in the waiting area'),
  bullet('Paper documents are not digitised and correlated with the history'),
  bullet('Ayurvedic parameters are not captured in a structured form'),
  bullet('Limited end-to-end integration with ABHA and FHIR-based records'),
  para(`These limitations justify the need for an integrated, AI-driven, multilingual case-taking kiosk such as **${P}**.`, { before: 120 }),

  h2('2.3 SCOPE OF THE PROJECT'),
  para(`The scope of ${P} extends across multiple dimensions of hospital workflow and digital health:`, { indent: 0 }),
  h3('1. Pre-Consultation Clinical Intake'),
  plain('Designed to be placed in OPD waiting areas of government, private and Ayush hospitals to collect patient histories before consultation.'),
  h3('2. Dual-System Case-Taking'),
  plain('Supports a SOCRATES-based allopathic interview and a Dashavidha Pariksha based Ayurvedic interview, making it suitable for integrative hospitals such as AIIA.'),
  h3('3. Multilingual and Voice-Enabled Interaction'),
  plain('Offers eight Indian languages in the prototype, with voice input and spoken feedback, and can be extended to all languages supported by Bhashini.'),
  h3('4. Patient Safety Through Triage'),
  plain('Screens every interview for red-flag symptoms and escalates emergencies to triage staff.'),
  h3('5. Document Digitisation'),
  plain('Scans prescriptions, lab reports and discharge summaries and highlights abnormal findings.'),
  h3('6. Interoperable Records'),
  plain('Produces structured summaries designed to be shared as FHIR resources and linked to the patient’s ABHA record.'),
  h3('7. Expandable Architecture'),
  plain('Future modules can include LLM-driven adaptive questioning, vitals capture through connected devices, ICD-11 / NAMASTE coding and analytics dashboards for hospital administrators.'),
  para('The scope ensures that the system is not limited to a single hospital or medical system but can be scaled to healthcare facilities across the country.', { before: 120, indent: 0 }),

  h2('2.4 SUMMARY'),
  para('This chapter reviewed the existing methods and systems used for clinical history collection, including manual case sheets, hospital information systems, telemedicine platforms, symptom checkers, ABHA-based registration and Ayush management systems. While each of these provides value, none of them performs complete, multilingual, patient-facing case-taking with early triage and structured, interoperable output.'),
  para(`The review clearly shows how **${P}** addresses these gaps by combining conversational AI, speech technology, document digitisation, red-flag detection and ABDM/FHIR integration in a single kiosk. The insights gained from this chapter provide the foundation for the system design and methodology described in the next chapter.`),
];

// ================================================================ CHAPTER 3
const ch3 = [
  chapter('CHAPTER 3 – SYSTEM DESIGN AND DEVELOPMENT'),
  h2('3.1 FUNDAMENTALS OF AI-BASED CLINICAL INTAKE SYSTEMS'),
  para('AI-based clinical intake systems are patient-facing tools that collect medical information before a consultation and present it to the clinician in a structured form. They work on the principle of breaking a complex clinical interview into a sequence of simple, guided questions, adapting the next question to the previous answers, and organising the result into the sections a doctor expects to see.'),
  plain('Fundamental characteristics include:'),
  h3('1. Patient-Centric Access'),
  para('The interface must be usable by people of all ages and literacy levels. A kiosk with a large touch screen, simple step-by-step navigation, voice interaction and support for regional languages removes the need for the patient to own a smartphone or to type.'),
  h3('2. Structured Clinical Interviewing'),
  para('Clinical histories follow well-established frameworks. For symptom analysis, the SOCRATES mnemonic (Site, Onset, Character, Radiation, Associated symptoms, Time course, Exacerbating/relieving factors, Severity) is widely taught. In Ayurveda, Dashavidha Pariksha (the ten-fold examination) assesses the patient through parameters such as Prakriti, Sara, Samhanana, Satmya, Sattva, Ahara Shakti, Vyayama Shakti and Vaya, complemented by assessment of Agni and Koshtha. An intake system encodes these frameworks as question flows.'),
  h3('3. Safety and Triage'),
  para('Because the system interacts with patients who have not yet seen a doctor, it must recognise danger signs and escalate them immediately. Triage rules are designed to be conservative: when in doubt, the patient is prioritised.'),
  plain('Typical outputs of the intake process are:'),
  bullet('**Chief complaint and history of present illness**'),
  bullet('**Past history, medication and allergies**'),
  bullet('**Review of systems or Ayush assessment parameters**'),
  bullet('**Summary of prior investigations**'),
  h3('4. Integration with Health Information Systems'),
  plain('Modern intake systems rely on standard interfaces such as:'),
  bullet('**ABDM / ABHA APIs** for patient identity and consent-based record sharing'),
  bullet('**HL7 FHIR resources** for exchanging structured clinical data'),
  bullet('**Speech and translation APIs** such as Bhashini for Indian languages'),
  h3('5. Privacy, Consent and Presentation'),
  plain('Health data is sensitive personal data, so the system must:'),
  bullet('**Capture explicit consent before collecting data**'),
  bullet('**Show the patient what will be shared**'),
  bullet('**Present a clear, editable summary to the doctor**'),
  bullet('**Clear data from the shared kiosk after each session**'),
  para(`${P} follows all these fundamentals and enhances them with AI-driven conversation, OCR and red-flag detection.`, { before: 120, indent: 0 }),

  h2('3.2 PROPOSED SYSTEM ARCHITECTURE'),
  para(`The proposed architecture of ${P} is designed using modular layers to ensure scalability, reliability and smooth integration with hospital systems. Fig 3.1 shows the overall architecture, which consists of the following layers:`, { indent: 0 }),
  ...figure('arch.png', `Fig 3.1 Proposed System Architecture of ${P}`),
  h3('1. Kiosk User Interface Layer'),
  bullet('Touch-screen, five-step wizard (Identify, Converse, Scan, Summary, Consult)'),
  bullet('Language selection for eight Indian languages'),
  bullet('Voice input and spoken feedback'),
  bullet('ABHA QR scan and document scanner tray'),
  h3('2. Application / Session Layer'),
  plain('This layer handles:'),
  bullet('Navigation between steps and progress display'),
  bullet('Storing the current session’s answers, documents and flags'),
  bullet('Recording patient consent'),
  bullet('Resetting all data at the end of each session'),
  h3('3. AI Processing Layer'),
  plain('The intelligent decision-making component:'),
  bullet('Speech recognition and text-to-speech for Indian languages'),
  bullet('Conversational case-taking engine with allopathy and Ayush question flows'),
  bullet('Red-flag triage rules applied to every answer'),
  bullet('OCR and extraction of key findings from medical documents'),
  bullet('Clinical summary generator'),
  h3('4. Data Integration Layer'),
  plain('Connects the kiosk to the wider health ecosystem:'),
  bullet('ABDM / ABHA for patient identity and health records'),
  bullet('HL7 FHIR R4 resources for structured data exchange'),
  bullet('Hospital Information System for token generation and doctor assignment'),
  h3('5. Output Layer'),
  plain('Produces:'),
  bullet('Consultation token and audio confirmation for the patient'),
  bullet('Editable case summary on the physician’s screen'),
  bullet('Priority alert to the triage nurse for red-flag cases'),
  para('This architecture ensures fast, safe and interoperable clinical intake, while allowing each AI component to be upgraded independently.', { before: 120, indent: 0 }),

  h2('3.3 SYSTEM WORKFLOW'),
  plain(`**The workflow of ${P} follows a sequential, systematic process:**`),
  h3('Step 1: Identify'),
  bullet('The patient enters their ABHA number or scans their ABHA QR code.'),
  bullet('The patient selects a preferred language.'),
  bullet('Consent for data capture and sharing is recorded.'),
  h3('Step 2: Conversational Interview'),
  bullet('**The AI asks questions one at a time, like a doctor.**'),
  bullet('**The patient answers by voice or by tapping options:**'),
  bullet('Allopathy mode: complaint, onset, severity, past illness, allergies, habits', 1),
  bullet('Ayush mode: complaint, Agni, Koshtha, Prakriti, Ahara Shakti, Vihara', 1),
  bullet('**Every answer is checked against red-flag rules; an emergency triggers a triage alert.**'),
  h3('Step 3: Document Scan'),
  bullet('**Previous prescriptions, lab reports and discharge summaries are scanned.**'),
  bullet('**Key findings are extracted and tagged:**'),
  bullet('Medicines and diagnoses', 1),
  bullet('Laboratory values', 1),
  bullet('Abnormal results highlighted', 1),
  h3('Step 4: Summary Generation'),
  bullet('**The system builds a structured case summary with:**'),
  bullet('Chief complaint and HPI', 1),
  bullet('Past history and drug allergy', 1),
  bullet('Review of systems or Dashavidha parameters', 1),
  bullet('Prior investigations', 1),
  bullet('The patient reviews the summary and may go back to edit it.'),
  h3('Step 5: Consult'),
  bullet('A token is generated and the summary is sent to the assigned doctor.'),
  bullet('The physician sees an editable draft and confirms it before it is saved to ABHA.'),
  bullet('All session data is cleared from the kiosk.'),
  ...figure('workflow.png', `Fig 3.2 Workflow Diagram of ${P}`, 330),

  h2('3.4 IMPLEMENTATION DETAILS'),
  para(`This section describes the technologies, tools and modules used to build the ${P} prototype. The prototype is a fully working front-end that demonstrates the complete patient journey; external services such as Bhashini, OCR, ABDM and the HIS are represented by simulated responses so that the system can be demonstrated without hospital infrastructure.`, { indent: 0 }),
  h3('Frontend Development'),
  bullet('HTML5, CSS3 and vanilla JavaScript in a single self-contained page (patientportal.html)'),
  bullet('Responsive layout using CSS Flexbox and Grid, usable on kiosk, laptop, tablet and mobile screens'),
  bullet('Large touch targets, colour-coded progress bar and high-contrast alerts'),
  h3('Conversational Case-Taking Engine'),
  bullet('Question flows stored as JavaScript data for the Allopathy (6 questions) and Ayush (6 questions) modes'),
  bullet('Each answer is appended to the chat and stored in the session answer list'),
  bullet('Next question is asked automatically after each answer'),
  h3('Speech'),
  bullet('Web Speech API (SpeechSynthesis) used to speak back each answer with an Indian English voice (en-IN)'),
  bullet('Voice-input control included in the UI; production speech recognition is planned through Bhashini ASR'),
  h3('Red-Flag Triage'),
  bullet('Rule-based detection of severe pain, sweating and breathlessness in answers'),
  bullet('Raises a single, prominent alert and marks the patient for the priority queue'),
  h3('Document Module'),
  bullet('Upload area for prescriptions, lab reports and discharge summaries'),
  bullet('Extracted findings shown with “OCR OK” or “ABNORMAL” tags (simulated OCR output in the prototype; planned engine: Tesseract / Indic OCR)'),
  h3('Summary and Physician View'),
  bullet('Structured table: Chief Complaint, HPI, Past History, Drug / Allergy, ROS or Ayush parameters, Prior Investigations'),
  bullet('Physician screen preview listing the key points and red-flag status as an editable draft'),
  h3('Integration Design'),
  bullet('Patient identity through ABHA number / QR'),
  bullet('Summary designed for HL7 FHIR R4 resources (Patient, Condition, AllergyIntolerance, Observation, DocumentReference)'),
  bullet('Consent capture following the DPDP Act 2023; session reset clears all data'),
  h3('Hosting'),
  bullet('Static web application with no server dependency; runs in any modern browser and can be hosted on GitHub Pages, Netlify or a hospital intranet kiosk'),
  plain('The main functions of the prototype are summarised in Table 3.1.', { before: 120 }),
  new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 120 }, children: [new TextRun({ text: 'Table 3.1 Major Functions of the Prototype', bold: true, size: 22 })] }),
  grid([2300, 6006], ['Function', 'Purpose'], [
    ['goStep(n)', 'Shows step n, updates the progress bar and fills the physician view at step 5'],
    ['setMode(m)', 'Switches between Allopathy and Ayush question flows'],
    ['startChat() / ask()', 'Starts the interview and renders the next question with answer options'],
    ['answer(txt)', 'Records the answer, checks red-flag rules, speaks feedback and moves to the next question'],
    ['redFlagFlag()', 'Displays the emergency alert once and marks the session for priority triage'],
    ['speak(t)', 'Uses the Web Speech API to give spoken confirmation to the patient'],
    ['addDoc()', 'Adds a scanned document with its extracted findings and status tag'],
    ['generateSummary()', 'Builds the structured case summary table from the session data'],
    ['resetAll()', 'Clears answers, documents and identity and returns to step 1 for the next patient'],
  ], { boldFirst: true, size: 22, headShade: 'E7E6E6' }),
  para('This implementation delivers a smooth, fast and easy-to-use intake experience and provides a clear foundation for connecting real AI services in the next phase.', { before: 200, indent: 0 }),

  h2('3.5 SUMMARY'),
  para(`This chapter described the system design and development process behind **${P}**. It introduced the fundamentals of AI-based clinical intake systems and presented the five-layer architecture that connects the kiosk interface, session management, AI processing, data integration and output. The workflow and implementation details explained how the system identifies the patient, conducts the interview, screens for emergencies, digitises documents and delivers a structured summary to the doctor.`),
  para('The next chapter presents the results, outputs and analysis of the system’s performance.'),
];

// ================================================================ CHAPTER 4
const TESTS = [
  ['Identification Test', 'ABHA QR scan fills the ABHA number', 'Tap “Scan ABHA QR”', 'ABHA 91-4521-8934-7712 filled', 'Passed'],
  ['Language Test', 'Selecting a preferred language', 'Tap “தமிழ்” (Tamil)', 'Only Tamil highlighted as selected', 'Passed'],
  ['Navigation Test', 'Moving from Identify to Converse', 'Tap “Start Case-Taking”', 'Step 2 shown; progress bar updated', 'Passed'],
  ['Red-Flag Test', 'Emergency detection in Allopathy mode', 'Chest pain, since 2 hours, sweating & breathlessness', 'One red-flag alert; priority triage', 'Passed'],
  ['Interview Test', 'Completion of all six questions', 'Six answers', 'Completion message displayed', 'Passed'],
  ['Document Test', 'Upload and tagging of documents', 'Four upload attempts (3 samples)', '3 documents listed, 1 tagged ABNORMAL; extra attempt handled', 'Passed'],
  ['Summary Test', 'Structured summary from answers', 'Allopathy session above', 'CC, HPI, past history and PRIORITY banner correct', 'Passed'],
  ['Physician View Test', 'Summary sent to doctor screen', 'Tap “Submit to Doctor”', 'Complaint and RED FLAG shown on physician view', 'Passed'],
  ['Session Reset Test', 'Privacy after each patient', 'Tap “New Patient”', 'ABHA, documents cleared; back to step 1', 'Passed'],
  ['Ayush Mode Test', 'Dashavidha interview flow', 'Acidity, sharp Agni, hard Koshtha, Pitta, etc.', 'No false red flag; interview completed', 'Passed'],
  ['Ayush Summary Test', 'Ayush parameters in summary', 'Ayush session above', 'Agni, Koshtha and Prakriti row present', 'Passed'],
  ['Non-Emergency Test', 'No escalation for moderate symptoms', 'Fever since 3 days, moderate', 'No red-flag alert raised', 'Passed'],
  ['Mobile Layout Test', 'Responsive design on phone width', '390 × 844 viewport', 'No horizontal scrolling; all steps usable', 'Passed'],
  ['Performance Test', 'Page load and summary generation time', 'Local load in Chromium', 'Page loaded in < 0.1 s; summary in < 0.1 s', 'Passed'],
];
const ch4 = [
  chapter('CHAPTER 4 – RESULT AND DISCUSSION'),
  h2('4.1 FUNCTIONAL ANALYSIS'),
  para(`The functional analysis focuses on how effectively the ${P} prototype performs its intended operations and fulfils the needs of patients and physicians. The system was evaluated on correctness of each step, safety behaviour, responsiveness, usability and the quality of the final summary.`),
  h3('1. Patient Identification and Consent'),
  plain('The system correctly captures:'),
  bullet('ABHA number by manual entry or QR scan'),
  bullet('Preferred language from eight Indian languages'),
  bullet('Explicit consent for data capture and sharing'),
  h3('2. Conversational Interview'),
  plain('Functional testing showed that the interview:'),
  bullet('Presents one question at a time with large answer buttons'),
  bullet('Records each answer and gives spoken confirmation'),
  bullet('Supports both Allopathy and Ayush (Dashavidha Pariksha) flows'),
  h3('3. Red-Flag Triage'),
  plain('The safety function was tested for:'),
  bullet('Detection of severe pain, sweating and breathlessness'),
  bullet('Exactly one alert per session, even if several danger signs are reported'),
  bullet('No false alerts for moderate, non-emergency symptoms'),
  h3('4. Document Digitisation'),
  plain('The document module lists each uploaded document with its extracted findings and clearly marks abnormal reports, such as a high cholesterol value, so that they are not missed by the physician.'),
  h3('5. Output Presentation'),
  plain('Results are displayed in a simple, readable format. The interface:'),
  bullet('Shows a structured summary table for patient review'),
  bullet('Displays a priority banner for red-flag cases'),
  bullet('Generates a consultation token and a physician-screen preview'),
  h3('6. Privacy and Usability'),
  plain('The session reset clears identity, answers and documents after each patient. The interface has been verified on desktop and mobile screen sizes, and its step-by-step design is suitable for patients with varied digital literacy.'),
  para(`Overall, the functional analysis shows that ${P} performs its core tasks effectively and meets the intended project objectives.`, { before: 120, indent: 0 }),

  h2('4.2 EXPERIMENTAL ANALYSIS'),
  para('The prototype was tested with an automated end-to-end test script (Playwright, headless Chromium) that drives the interface exactly as a patient would — clicking buttons, answering questions and uploading documents — and checks the result of each step. Fourteen test cases were executed on a 1280 × 860 desktop viewport and a 390 × 844 mobile viewport. The results are summarised in Table 4.1.'),
  new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 120 }, children: [new TextRun({ text: 'Table 4.1 Experimental Test Results', bold: true, size: 22 })] }),
  grid([1500, 1900, 1800, 2206, 900], ['Test Category', 'Description', 'Input', 'Output', 'Result'], TESTS, { boldFirst: true, size: 20, headShade: 'E7E6E6' }),
  para('All fourteen test cases passed. The prototype responded instantly at every step because all processing runs locally in the browser. During testing, the following points were also observed and are carried forward as improvements in Chapter 5: the follow-up questions are currently a fixed sequence rather than generated from the previous answer; the clinical mode should be chosen before the interview begins, since switching modes mid-interview continues from the same question number; and document extraction results are simulated in the prototype.', { before: 200 }),
  plain('**PROTOTYPE WEB DESKTOP:**', { before: 240 }),
  ...figure('d1_identify.png', 'Fig 4.1 Step 1 – Patient Identification, Language Selection and Consent'),
  ...figure('d2_converse.png', 'Fig 4.2 Step 2 – Conversational Case-Taking Interview (Allopathy Mode)'),
  ...figure('d3_scan.png', 'Fig 4.3 Step 3 – Scanned Documents with Extracted Findings'),
  ...figure('d4_summary.png', 'Fig 4.4 Step 4 – Structured Case Summary with Red-Flag Priority'),
  ...figure('d7_ayush_summary.png', 'Fig 4.5 Case Summary in Ayush Mode (Dashavidha Pariksha Parameters)'),
  ...figure('d5_consult.png', 'Fig 4.6 Step 5 – Token Generation and Physician Screen Preview'),
  plain('**PROTOTYPE MOBILE VIEW:**', { before: 240 }),
  ...figure('mobile_combo.png', 'Fig 4.7 Mobile View – Identification, Interview and Summary Screens', 520),

  h2('4.3 SUMMARY'),
  para(`This chapter presented the results and detailed analysis of the **${P}** prototype. Functional analysis demonstrated that the system correctly identifies the patient, records consent and language preference, conducts the case-taking interview in both Allopathy and Ayush modes, detects red-flag symptoms, lists scanned documents with abnormal findings and generates an accurate structured summary for the physician.`),
  para('The experimental analysis, carried out through fourteen automated end-to-end test cases on desktop and mobile screen sizes, confirmed the correctness and responsiveness of every step, with all tests passing. The observations made during testing identify clear directions for the next version.'),
  para(`Overall, the results show that **${P}** is a dependable, user-friendly intake tool capable of reducing consultation time and improving the completeness and safety of patient histories. The outcomes of Chapter 4 provide the foundation for the conclusions presented in the next chapter.`),
];

// ================================================================ CHAPTER 5
const fut = numbered('future');
const ch5 = [
  chapter('CHAPTER 5 – CONCLUSION AND FUTURE SCOPE'),
  h2('5.1 CONCLUSION'),
  para(`**${P} – AI Powered Patient Case-Taking Kiosk** successfully demonstrates how routine clinical history-taking can be moved out of the consultation room and performed by an intelligent, multilingual self-service system. Through a guided conversational interview, red-flag triage, document digitisation and structured summary generation, the system gives the physician a complete picture of the patient before the consultation begins.`),
  para('The project achieves its primary objectives of reducing the doctor’s documentation burden, improving the completeness of case histories, recognising emergencies earlier and making the process accessible to patients in their own language. Its support for both allopathic SOCRATES-based history and Ayurvedic Dashavidha Pariksha makes it particularly suitable for integrative institutions such as the All India Institute of Ayurveda. Designed around ABHA identity, FHIR-ready output and explicit consent, the system is a practical and scalable step towards smarter, patient-centred hospitals in India.'),

  h2('5.2 FUTURE SCOPE'),
  para('The system has strong potential for expansion and can be enhanced with several additional features to increase its usefulness for hospitals nationwide:', { indent: 0 }),
  fut('**LLM-Driven Adaptive Questioning** – Generating each follow-up question from the patient’s previous answers using a medically tuned large language model, instead of a fixed sequence.'),
  fut('**Live Bhashini Speech Integration** – Real-time speech recognition, translation and text-to-speech in all 22 scheduled Indian languages.'),
  fut('**Real OCR for Handwritten Documents** – Integrating Indic OCR and medical entity extraction for printed and handwritten prescriptions and reports.'),
  fut('**ABDM and FHIR Integration** – Connecting to the ABDM sandbox and hospital HIS to push summaries as FHIR resources linked to ABHA.'),
  fut('**Vitals Capture** – Connecting blood pressure, SpO₂, temperature and weight devices to the kiosk to include vital signs in the summary.'),
  fut('**Clinical Coding** – Automatic mapping of diagnoses to ICD-11 and Ayush terminologies (NAMASTE codes) for analytics and research.'),
  fut('**Expanded Triage Rules** – Validated red-flag protocols for paediatric, obstetric, stroke and sepsis emergencies, reviewed by clinicians.'),
  fut('**Analytics Dashboard** – Hospital-level insights on patient load, common complaints and waiting times.'),
  para('These enhancements can significantly improve the system’s clinical value and its adoption across India.', { before: 120, indent: 0 }),

  h2('5.3 APPLICATIONS'),
  para(`The ${P} system can be used in a variety of real-world healthcare contexts:`, { indent: 0 }),
  h3('1. Government and Teaching Hospitals'),
  plain('To reduce OPD waiting time and the documentation load on doctors and residents.'),
  h3('2. Ayush Hospitals and Integrative Medicine Centres'),
  plain('To capture structured Ayurvedic assessments alongside conventional history.'),
  h3('3. Primary Health Centres and Ayushman Arogya Mandirs'),
  plain('To support limited medical staff with pre-collected, structured patient information.'),
  h3('4. Private Clinics and Diagnostic Centres'),
  plain('To speed up patient onboarding and digitise previous reports.'),
  h3('5. Telemedicine Hubs'),
  plain('To prepare a structured history before a remote consultation begins.'),
  h3('6. Medical Education and Research'),
  plain('To teach structured history-taking and to build anonymised datasets for clinical research.'),

  h2('5.4 SUMMARY'),
  para(`This chapter summarised the achievements, future potential and real-world applications of the **${P}** system. The project demonstrates how conversational AI, multilingual speech, document digitisation and interoperable health standards can make clinical intake faster, safer and more inclusive. The conclusion reinforces that the system is practical, scalable and aligned with India’s mission for digital health. Future enhancements can make ${P} even more impactful and widely accessible to hospitals and patients across the nation.`),
];

// ================================================================ REFERENCES
const ref = numbered('refs');
const references = [
  new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { after: 360 },
    children: [new TextRun({ text: 'REFERENCES', size: 28, font: FONT })] }),
  ...[
    'National Health Authority, *Ayushman Bharat Digital Mission (ABDM) – Sandbox and Integration Documentation*, Government of India, 2023.',
    'Ministry of Health and Family Welfare, *National Digital Health Blueprint*, Government of India, 2019.',
    'Ministry of Electronics and Information Technology, *The Digital Personal Data Protection Act, 2023*, Gazette of India, 2023.',
    'HL7 International, *FHIR Release 4 (R4) Specification*, Health Level Seven International, 2019.',
    'Ministry of Electronics and Information Technology, *Bhashini – National Language Translation Mission*, Digital India, 2022.',
    'Bickley, L. S., Szilagyi, P. G. & Hoffman, R. M., *Bates’ Guide to Physical Examination and History Taking*, 13th ed., Wolters Kluwer, 2021.',
    'Innes, J. A., Dover, A. R. & Fairhurst, K., *Macleod’s Clinical Examination*, 14th ed., Elsevier, 2018.',
    'Agnivesha, *Charaka Samhita* (Vimana Sthana, Chapter 8 – Dashavidha Pariksha), translated by R. K. Sharma & Bhagwan Dash, Chowkhamba Sanskrit Series Office, Varanasi.',
    'Ministry of Ayush, *NAMASTE Portal – National Ayush Morbidity and Standardized Terminologies Electronic Portal*, Government of India, 2017.',
    'World Health Organization, *International Classification of Diseases, 11th Revision (ICD-11)*, WHO, 2019.',
    'Topol, E. J., “High-performance medicine: the convergence of human and artificial intelligence,” *Nature Medicine*, vol. 25, pp. 44–56, 2019.',
    'Singhal, K. et al., “Large language models encode clinical knowledge,” *Nature*, vol. 620, pp. 172–180, 2023.',
    'Semigran, H. L., Linder, J. A., Gidengil, C. & Mehrotra, A., “Evaluation of symptom checkers for self diagnosis and triage,” *BMJ*, vol. 351, h3480, 2015.',
    'Smith, R., “An Overview of the Tesseract OCR Engine,” *Proc. 9th International Conference on Document Analysis and Recognition (ICDAR)*, IEEE, pp. 629–633, 2007.',
    'Mackway-Jones, K., Marsden, J. & Windle, J., *Emergency Triage: Manchester Triage Group*, 3rd ed., Wiley-Blackwell, 2014.',
    'W3C, *Web Speech API – Draft Community Group Report*, World Wide Web Consortium, 2023.',
    'W3C, *Web Content Accessibility Guidelines (WCAG) 2.1*, World Wide Web Consortium, 2018.',
    'Ministry of Health and Family Welfare, *eSanjeevani – National Telemedicine Service*, Government of India, 2020.',
    'Jiang, F. et al., “Artificial intelligence in healthcare: past, present and future,” *Stroke and Vascular Neurology*, vol. 2, no. 4, pp. 230–243, 2017.',
    'Ministry of Education & AICTE, *Smart India Hackathon – Problem Statement SIH26047*, Government of India, 2026.',
  ].map(ref),
];

// ================================================================ ASSEMBLE
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 22 })] })] });
const emptyFooter = new Footer({ children: [new Paragraph({ children: [] })] });
const pageProps = (extra = {}) => ({
  page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: M_TOP, bottom: M_BOTTOM, left: M_LEFT, right: M_RIGHT, footer: 708 }, ...extra },
});

const numCfg = (ref, fmt, text, indentL = 720) => ({
  reference: ref,
  levels: [{ level: 0, format: fmt, text, alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: indentL, hanging: 360 } } } }],
});

const doc = new Document({
  creator: INFO.student,
  title: `${INFO.title} – Mini Project Report`,
  styles: {
    default: { document: { run: { font: FONT, size: 24 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: '000000' }, paragraph: { outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: '000000' }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: 'o', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      ] },
      numCfg('future', LevelFormat.DECIMAL, '%1.'),
      numCfg('refs', LevelFormat.DECIMAL, '%1.'),
    ],
  },
  sections: [
    { properties: pageProps(), footers: { default: emptyFooter }, children: [...cover, ...bonafide, ...acknowledgement, ...declaration] },
    { properties: pageProps({ pageNumbers: { start: 5, formatType: NumberFormat.LOWER_ROMAN } }), footers: { default: footer },
      children: [...abstract, ...toc, ...abbreviations] },
    { properties: pageProps({ pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } }), footers: { default: footer },
      children: [...ch1, ...ch2, ...ch3, ...ch4, ...ch5, ...references] },
  ],
});

const out = path.join(__dirname, 'MediKiosk_Project_Report.docx');
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log('wrote', out); });
